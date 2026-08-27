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

from collections import Counter
from merger import MODES, MergeResult, merge_reports, build_filtered_workbook

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 64 * 1024 * 1024  # 64 MB request cap

# In-memory cache of generated workbooks:
# token -> (created_at, filename, bytes, records, columns, mode_key)
_CACHE: dict[str, tuple[float, str, bytes, list, list, str]] = {}
_CACHE_TTL_SECONDS = 30 * 60
_MAX_FILES = 50
_MAX_BYTES_PER_FILE = 50 * 1024 * 1024


def _sweep_cache() -> None:
    now = time.time()
    for tok in [t for t, entry in _CACHE.items() if now - entry[0] > _CACHE_TTL_SECONDS]:
        _CACHE.pop(tok, None)


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
                {"error": f"'{f.filename}' exceeds the 50 MB per-file size limit."}
            ), 400
        payloads.append((f.filename, data))

    try:
        result: MergeResult = merge_reports(
            payloads,
            mode_key=mode_key,
            sort_by=sort_by,
            sort_dir=sort_dir,
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


@app.post("/filter-download")
def filter_download():
    """Build and return a multi-sheet workbook from filter definitions."""
    token = request.form.get("token") or (request.get_json(silent=True) or {}).get("token")
    if not token:
        return jsonify({"error": "Missing token."}), 400

    entry = _CACHE.get(token)
    if entry is None:
        return jsonify({"error": "Merge result expired. Please merge again."}), 404

    _, _, _, records, columns, mode_key = entry
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
