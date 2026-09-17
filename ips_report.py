"""Parser for raw IPS (P2P) transaction exports.

The sample report ("SEPT 06-16,2026, IPS transactions.xlsx") is a raw
transaction-level export whose 8 columns come from the SQL that generated it:

    Destination_Bank | Source_bank | TRX_DATE | DBTR_ACCT | CDTR_ACCT |
    amount | tx_id | STATUS

TRX_DATE looks like "16-SEP-26 03.27.38.850000000 PM" (DD-MON-YY, with the
time and a thousands-of-milliseconds fraction). The workbook often has no
header row (the sample puts the data in "Sheet10", with empty Sheet1-Sheet9
plus an "Export Worksheet" / "SQL" sheet), so the parser scans EVERY sheet,
collects the distinct transaction dates found across ALL sheets and lets the
caller filter records by one or more selected dates.
"""

from __future__ import annotations

import re
from collections import Counter
from typing import Any

from merger import (
    IPS_CANONICAL_COLUMNS,
    IPS_HEADER_ALIASES,
    IPS_MODE,
    _cell_value,
    _detect_engine,
    _is_empty,
    _read_grids,
    build_workbook,
    sort_records,
)

_FULL_MONTHS = {
    "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
    "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
}

# TRX_DATE values like '16-SEP-26 03.27.38.850000000 PM'
_TRX_DATE_RE = re.compile(r"^(\d{1,2})-([A-Za-z]{3})-(\d{2,4})")


def parse_ips_trx_date(value: Any) -> tuple[str, str] | None:
    """Parse an IPS TRX_DATE into (iso_date_key, display_label).

    '16-SEP-26 03.27.38.850000000 PM' -> ('2026-09-16', '16-SEP-26').
    Returns None when the value has no recognizable DD-MON-YY prefix.
    """
    if _is_empty(value):
        return None
    m = _TRX_DATE_RE.match(str(value).strip())
    if not m:
        return None
    day = int(m.group(1))
    mon = m.group(2).upper()
    yy = m.group(3)
    month = _FULL_MONTHS.get(mon)
    if not month or not (1 <= day <= 31):
        return None
    year = int(yy)
    if year < 100:
        year += 2000
    key = f"{year:04d}-{month:02d}-{day:02d}"
    label = f"{day:02d}-{mon}-{str(year)[2:]}"
    return key, label


def _find_header_row(grid: list[list[Any]]) -> int:
    """Locate a row whose cells match known IPS header tokens, or -1."""
    for r in range(min(30, len(grid))):
        cells = [str(v).strip().upper() for v in grid[r] if not _is_empty(v)]
        if any(c in IPS_HEADER_ALIASES for c in cells):
            return r
    return -1


def _col_map_from_header(cells: list[Any]) -> dict[int, str]:
    mapping: dict[int, str] = {}
    for idx, v in enumerate(cells):
        if _is_empty(v):
            continue
        key = IPS_HEADER_ALIASES.get(str(v).strip().upper())
        if key is not None:
            mapping[idx] = key
    return mapping


def parse_ips_report(data: bytes, filename: str = "") -> dict:
    """Parse an IPS export, scanning every sheet.

    Returns a dict with:
      records      - list of dicts keyed by IPS_CANONICAL_COLUMNS
      rows_scanned - total non-empty rows checked across all sheets
      sheets       - {sheet_name: rows_kept} for every sheet
      dates        - sorted list of {key, label, count}
      warnings     - human-friendly notes (empty sheets, no header, ...)
      filename     - original file name
    """
    engine = _detect_engine(data)
    grids = _read_grids(data, engine)

    files_dates: Counter = Counter()
    records: list[dict] = []
    sheets_found: dict[str, int] = {}
    warnings: list[str] = []
    rows_scanned = 0

    for sheet_name, grid in grids.items():
        sheets_found[sheet_name] = 0
        if not grid:
            warnings.append(f"Sheet '{sheet_name}' is empty.")
            continue

        header_row_idx = _find_header_row(grid)
        col_map: dict[int, str] = {}
        start = 0
        if header_row_idx >= 0:
            col_map = _col_map_from_header(grid[header_row_idx])
            start = header_row_idx + 1
        else:
            # No header row: rely on the SQL column order.
            col_map = dict(enumerate(IPS_CANONICAL_COLUMNS))

        sheet_records = 0
        for row in grid[start:]:
            if not any(not _is_empty(v) for v in row):
                continue
            rows_scanned += 1
            values: dict[str, Any] = {}
            for idx, col in col_map.items():
                if idx < len(row):
                    values[col] = _cell_value(row[idx])
            trx = values.get("TRX_DATE")
            parsed = parse_ips_trx_date(trx)
            if parsed is None:
                continue
            key, label = parsed
            rec = {col: values.get(col, "") for col in IPS_CANONICAL_COLUMNS}
            rec["TRX_DATE"] = str(trx).strip()
            records.append(rec)
            files_dates[key] += 1
            sheet_records += 1
        sheets_found[sheet_name] = sheet_records

    if not records:
        warnings.append(
            "No IPS transaction rows found - expected rows with a "
            "TRX_DATE like '16-SEP-26 03.27.38.850000000 PM'."
        )
    elif header_row_idx < 0:
        warnings.append(
            "No header row detected; columns were mapped positionally from "
            "the SQL order (Destination_Bank, Source_bank, TRX_DATE, DBTR_ACCT, "
            "CDTR_ACCT, AMOUNT, TX_ID, STATUS)."
        )

    dates = [
        {"key": k, "label": _label_for_key(k), "count": c}
        for k, c in sorted(files_dates.items())
    ]

    return {
        "records": records,
        "rows_scanned": rows_scanned,
        "sheets": sheets_found,
        "dates": dates,
        "warnings": warnings,
        "filename": filename,
        "columns": list(IPS_CANONICAL_COLUMNS),
    }


def _label_for_key(key: str) -> str:
    """'2026-09-16' -> '16-SEP-26'."""
    year, month, day = key.split("-")
    mon = {v: k for k, v in _FULL_MONTHS.items()}[int(month)]
    return f"{day}-{mon}-{year[2:]}"


def collect_ips_dates(parsed_files: list[dict]) -> list[dict]:
    """Merge date catalogs from several parsed IPS files.

    Each entry: {key, label, count, files: list of source file names}.
    """
    counter: Counter = Counter()
    files_by_key: dict[str, list[str]] = {}
    for parsed in parsed_files:
        for d in parsed.get("dates", []):
            counter[d["key"]] += d["count"]
            files_by_key.setdefault(d["key"], []).append(parsed.get("filename", ""))
    return [
        {
            "key": k,
            "label": _label_for_key(k),
            "count": counter[k],
            "files": sorted(set(files_by_key.get(k, []))),
        }
        for k in sorted(counter)
    ]


def filter_ips_by_dates(records: list[dict], date_keys: list[str]) -> list[dict]:
    """Keep only records whose TRX_DATE falls on one of the selected days."""
    wanted = set(date_keys)
    return [r for r in records if parse_ips_trx_date(r.get("TRX_DATE"))[0] in wanted]


def build_ips_workbook(records: list[dict], from_date: str = "", to_date: str = "") -> bytes:
    """Build the downloadable IPS workbook (header in row 1, plain sheet).

    Records are sorted by TRX_DATE so rows group by day. Returns the raw
    workbook bytes, ready to save or hand to the front end.
    """
    sorted_records = sort_records(records, "TRX_DATE")
    return build_workbook(sorted_records, from_date, to_date, IPS_MODE)