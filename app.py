"""Flask dashboard for merging POS decline, POS and ATM reports.

Everything runs in memory: uploaded files are read into RAM, the merged
workbook is generated in RAM and served for download, and nothing is ever
written to disk (the in-memory download cache expires after 30 minutes).
"""

from __future__ import annotations

import time
import uuid
from io import BytesIO

from flask import Flask, jsonify, render_template, request, send_file

from merger import MODES, MergeResult, merge_reports, build_filtered_workbook, build_workbook
from ips_report import (
    parse_ips_report,
    collect_ips_dates,
    filter_ips_by_dates,
    build_ips_workbook,
)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 800 * 1024 * 1024  # 800 MB request cap

# In-memory cache of generated workbooks:
# token -> (created_at, filename, bytes, records, columns, mode_key,
#           duplicate_records)
_CACHE: dict[str, tuple[float, str, bytes, list, list, str, list]] = {}
_CACHE_TTL_SECONDS = 30 * 60
_MAX_FILES = 50
_MAX_BYTES_PER_FILE = 800 * 1024 * 1024

# Parsed IPS records awaiting a date selection:
# token -> {created_at, records, dates, per_file, warnings}
_IPS_CACHE: dict[str, dict] = {}


def _sweep_cache() -> None:
    now = time.time()
    for tok in [t for t, entry in _CACHE.items() if now - entry[0] > _CACHE_TTL_SECONDS]:
        _CACHE.pop(tok, None)
    for tok in [t for t, e in _IPS_CACHE.items() if now - e["created_at"] > _CACHE_TTL_SECONDS]:
        _IPS_CACHE.pop(tok, None)


@app.get("/")
def index():
    return render_template("index.html")


@app.post("/merge")
def merge():
    mode_key = request.form.get("mode", "pos_decline")
    if mode_key not in MODES:
        return jsonify({"error": f"Unknown report mode '{mode_key}'."}), 400

    sort_by = request.form.get("sort_by", "")
    sort_dir = request.form.get("sort_dir", "asc")
    if sort_dir not in ("asc", "desc"):
        sort_dir = "asc"

    dedupe = request.form.get("dedupe", "") in ("1", "true", "on", "yes")

    uploads = [f for f in request.files.getlist("files") if f and f.filename]
    if not uploads:
        return jsonify({"error": "No files were uploaded."}), 400
    if len(uploads) > _MAX_FILES:
        return jsonify({"error": f"Too many files (maximum is {_MAX_FILES})."}), 400

    payloads: list[tuple[str, bytes]] = []
    for f in uploads:
        data = f.read()
        if len(data) > _MAX_BYTES_PER_FILE:
            return jsonify(
                {"error": f"'{f.filename}' exceeds the 800 MB per-file size limit."}
            ), 400
        payloads.append((f.filename, data))

    try:
        result: MergeResult = merge_reports(
            payloads,
            mode_key=mode_key,
            sort_by=sort_by,
            sort_dir=sort_dir,
            dedupe=dedupe,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    token = uuid.uuid4().hex
    _CACHE[token] = (
        time.time(),
        result.filename,
        result.workbook_bytes,
        result.records,
        list(result.records[0].keys()) if result.records else [],
        result.mode_key,
        result.duplicate_records,
    )
    _sweep_cache()

    # Compute unique values per column for filter dropdowns
    unique_values = {}
    if result.records:
        columns = list(result.records[0].keys())
        for col in columns:
            vals = sorted(
                {str(r.get(col, "")).strip() for r in result.records
                 if r.get(col, "") not in ("", None)}
            )
            unique_values[col] = vals

    return jsonify(
        {
            "token": token,
            "filename": result.filename,
            "mode": result.mode_key,
            "mode_label": result.mode_label,
            "columns": list(result.records[0].keys()) if result.records else [],
            "total_rows": result.total_rows,
            "duplicate_count": len(result.duplicate_records),
            "from_date": result.from_date,
            "to_date": result.to_date,
            "per_file": result.per_file,
            "preview": result.records[:50],
            "unique_values": unique_values,
            "resp_counts": result.resp_counts,
            "warnings": result.warnings,
            "sort_by": result.sort_by,
            "sort_dir": result.sort_dir,
        }
    )

@app.post("/ips-analyze")
def ips_analyze():
    """Parse raw IPS exports and return the dates found across all sheets.

    The parsed records are cached under a token so the client can then ask
    for a merged workbook filtered to the dates it selects.
    """
    uploads = [f for f in request.files.getlist("files") if f and f.filename]
    if not uploads:
        return jsonify({"error": "No files were uploaded."}), 400
    if len(uploads) > _MAX_FILES:
        return jsonify({"error": f"Too many files (maximum is {_MAX_FILES})."}), 400

    parsed = []
    for f in uploads:
        data = f.read()
        if len(data) > _MAX_BYTES_PER_FILE:
            return jsonify(
                {"error": f"'{f.filename}' exceeds the 800 MB per-file size limit."}
            ), 400
        try:
            parsed.append(parse_ips_report(data, f.filename))
        except ValueError as exc:
            return jsonify({"error": str(exc)}), 400
        except Exception as exc:  # noqa: BLE001
            return jsonify({"error": f"Failed to read '{f.filename}': {exc}"}), 400

    all_records = [r for p in parsed for r in p["records"]]
    dates = collect_ips_dates(parsed)
    warnings = [f"{p['filename']}: {w}" for p in parsed for w in p["warnings"]]
    per_file = [
        {
            "filename": p["filename"], "status": "ok", "sheet": "all sheets",
            "raw_rows": p["rows_scanned"], "data_rows": len(p["records"]),
            "columns_kept": len(p["columns"]),
            "column_order": p["columns"], "order_mismatch": False,
            "blank_columns": [], "dropped_columns": [], "extra_columns": [],
            "warnings": p["warnings"],
        }
        for p in parsed
    ]

    if not all_records:
        return jsonify({
            "error": "No IPS transaction rows were found in the uploaded file(s).",
            "warnings": warnings,
        }), 400

    token = uuid.uuid4().hex
    _IPS_CACHE[token] = {
        "created_at": time.time(),
        "records": all_records,
        "dates": dates,
        "per_file": per_file,
        "warnings": warnings,
    }
    _sweep_cache()

    return jsonify({
        "token": token,
        "dates": dates,
        "per_file": per_file,
        "warnings": warnings,
        "total_rows": len(all_records),
        "columns": list(all_records[0].keys()),
    })


@app.post("/ips-merge")
def ips_merge():
    """Filter cached IPS records by the selected date(s) and build the workbook."""
    token = request.form.get("token") or (request.get_json(silent=True) or {}).get("token")
    if not token:
        return jsonify({"error": "Missing token."}), 400

    cached = _IPS_CACHE.get(token)
    if cached is None:
        return jsonify({"error": "IPS analysis expired. Please upload the file(s) again."}), 404

    raw_dates = request.form.get("dates") or (request.get_json(silent=True) or {}).get("dates")
    if isinstance(raw_dates, str):
        text = raw_dates.strip()
        if text.startswith("["):
            try:
                raw_dates = __import__("json").loads(text)
            except Exception:
                return jsonify({"error": "Invalid date selection."}), 400
        else:
            raw_dates = [d.strip() for d in text.split(",") if d.strip()]
    if not isinstance(raw_dates, list) or not raw_dates:
        return jsonify({"error": "Select at least one date."}), 400

    labels = {d["key"]: d["label"] for d in cached["dates"]}
    keys = sorted(str(k) for k in raw_dates)
    records = filter_ips_by_dates(cached["records"], keys)

    from_date = keys[0]
    to_date = keys[-1]
    from_lbl = labels.get(from_date, from_date)
    to_lbl = labels.get(to_date, to_date)
    if from_lbl == to_lbl:
        filename = f"IPS_Transactions_{from_lbl}_Merged.xlsx"
    else:
        filename = f"IPS_Transactions_{from_lbl}_to_{to_lbl}_Merged.xlsx"

    workbook_bytes = build_ips_workbook(records, from_date, to_date)

    resp = {}
    for r in records:
        status = str(r.get("STATUS", "") or "").strip() or "(blank)"
        resp[status] = resp.get(status, 0) + 1
    resp = dict(sorted(resp.items(), key=lambda kv: -kv[1]))

    merge_token = uuid.uuid4().hex
    _CACHE[merge_token] = (
        time.time(),
        filename,
        workbook_bytes,
        records,
        list(records[0].keys()) if records else [],
        "ips",
        [],
    )
    _sweep_cache()

    unique_values = {}
    if records:
        for col in records[0].keys():
            unique_values[col] = sorted(
                {str(r.get(col, "")).strip() for r in records
                 if r.get(col, "") not in ("", None)}
            )

    return jsonify({
        "token": merge_token,
        "filename": filename,
        "mode": "ips",
        "mode_label": "IPS",
        "columns": list(records[0].keys()) if records else [],
        "total_rows": len(records),
        "duplicate_count": 0,
        "from_date": from_date,
        "to_date": to_date,
        "per_file": cached["per_file"],
        "preview": records[:50],
        "unique_values": unique_values,
        "resp_counts": resp,
        "warnings": cached["warnings"],
        "sort_by": "date_time",
        "sort_dir": "asc",
    })


@app.get("/download/<token>")
def download(token: str):
    entry = _CACHE.get(token)
    if entry is None:
        return (
            "The merged report is no longer available - it is kept only in memory "
            "and expires after 30 minutes. Please merge again.",
            404,
        )
    _, filename, data = entry[0], entry[1], entry[2]
    return send_file(
        BytesIO(data),
        as_attachment=True,
        download_name=filename,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


@app.get("/download-duplicates/<token>")
def download_duplicates(token: str):
    """Serve a workbook containing only the fully-duplicate rows."""
    entry = _CACHE.get(token)
    if entry is None:
        return (
            "The merged report is no longer available - it is kept only in memory "
            "and expires after 30 minutes. Please merge again.",
            404,
        )
    _, _, _, records, columns, mode_key, duplicate_records = entry
    if not duplicate_records:
        return jsonify({"error": "No duplicate rows were found in the merged data."}), 400

    mode = MODES.get(mode_key)
    if mode is None:
        return jsonify({"error": f"Unknown mode '{mode_key}'."}), 500

    from_date, to_date = "", ""
    try:
        wb_bytes = build_workbook(duplicate_records, from_date, to_date, mode)
    except Exception as exc:
        return jsonify({"error": f"Failed to build duplicates workbook: {exc}"}), 500

    base = mode.output_prefix
    out_name = f"{base}_Duplicates.xlsx"
    return send_file(
        BytesIO(wb_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


@app.post("/filter-download")
def filter_download():
    token = request.form.get("token") or (request.get_json(silent=True) or {}).get("token")
    if not token:
        return jsonify({"error": "Missing token."}), 400

    entry = _CACHE.get(token)
    if entry is None:
        return jsonify({"error": "Merge result expired. Please merge again."}), 404

    _, _, _, records, columns, mode_key, _ = entry
    mode = MODES.get(mode_key)
    if mode is None:
        return jsonify({"error": f"Unknown mode '{mode_key}'."}), 500

    # Parse sheet definitions from JSON
    raw = request.form.get("sheets") or (request.get_json(silent=True) or {}).get("sheets")
    if not raw:
        return jsonify({"error": "No sheet definitions provided."}), 400
    try:
        sheet_defs = raw if isinstance(raw, list) else __import__("json").loads(raw)
    except Exception:
        return jsonify({"error": "Invalid sheet definitions JSON."}), 400

    if not sheet_defs:
        return jsonify({"error": "At least one filter sheet is required."}), 400

    if len(sheet_defs) > 20:
        return jsonify({"error": "Maximum 20 filter sheets allowed."}), 400

    try:
        wb_bytes = build_filtered_workbook(records, columns, sheet_defs, mode)
    except Exception as exc:
        return jsonify({"error": f"Failed to build filtered workbook: {exc}"}), 500

    # Build output filename
    base = mode.output_prefix
    if len(sheet_defs) == 1:
        sheet_label = sheet_defs[0].get("name", "Filtered")
        out_name = f"{base}_{sheet_label}_Filtered.xlsx"
    else:
        out_name = f"{base}_Filtered_{len(sheet_defs)}_Sheets.xlsx"

    return send_file(
        BytesIO(wb_bytes),
        as_attachment=True,
        download_name=out_name,
        mimetype=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )


def main() -> None:
    app.run(host="127.0.0.1", port=5000, debug=False)


if __name__ == "__main__":
    main()
