"""NBE Institution Report generator for POS (Daily) and ATM (Daily) merged reports.

Normalises ISSUER and ACQUIRER bank names to canonical NBE institution names,
filters by transaction type (POS purchase vs ATM cash withdrawal) and response code (-1, -1.0),
aggregates transaction count and total monetary value, and outputs structured DataFrames
and styled Excel workbooks matching NBE REPORT layout.
"""

from __future__ import annotations

import io
from typing import Any

import pandas as pd
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# Standard NBE Institution Order (31 Institutions)
# ---------------------------------------------------------------------------
STANDARD_NBE_BANKS = [
    "Abay Bank",
    "Addis Bank",
    "Ahadu Bank",
    "Amhara Bank",
    "Awash Bank",
    "Birhan Bank",
    "BOA",
    "Bunna Bank",
    "CBE",
    "CBO",
    "Global Bank",
    "Dashen Bank",
    "Enat Bank",
    "Gadda Bank",
    "Goh Betoch Bank",
    "Hijra Bank",
    "Lion Bank",
    "Nib Bank",
    "Oromia Bank",
    "Rammis Bank",
    "Santim Pay",
    "Sinqee Bank",
    "Sidama Bank",
    "Siket Bank",
    "Tseday Bank",
    "Tsehay Bank",
    "United Bank",
    "Wegagen Bank",
    "Yagout Pay",
    "Zamzam Bank",
    "Zemen Bank",
]

# ---------------------------------------------------------------------------
# Comprehensive Bank Name Alias Mapping
# ---------------------------------------------------------------------------
_NBE_ALIAS_MAP: dict[str, str] = {
    # Abay
    "abay": "Abay Bank",
    "abay bank": "Abay Bank",
    # Addis
    "addis": "Addis Bank",
    "addis bank": "Addis Bank",
    "addis int": "Addis Bank",
    "addis int bank": "Addis Bank",
    "addis international": "Addis Bank",
    "addis international bank": "Addis Bank",
    # Ahadu
    "ahadu": "Ahadu Bank",
    "ahadu bank": "Ahadu Bank",
    "ahadu ebirr": "Ahadu Bank",
    # Amhara
    "amhara": "Amhara Bank",
    "amhara bank": "Amhara Bank",
    "amharaethbirr": "Amhara Bank",
    # Awash
    "awash": "Awash Bank",
    "awash bank": "Awash Bank",
    "aib": "Awash Bank",
    # Birhan / Berhan
    "birhan": "Birhan Bank",
    "birhan bank": "Birhan Bank",
    "berhan": "Birhan Bank",
    "berhan bank": "Birhan Bank",
    # Bank of Abyssinia / BOA
    "boa": "BOA",
    "abyssinia": "BOA",
    "abyssinia bank": "BOA",
    "bank of abyssinia": "BOA",
    # Bunna
    "bunna": "Bunna Bank",
    "bunna bank": "Bunna Bank",
    "buna": "Bunna Bank",
    "buna bank": "Bunna Bank",
    "bunna int": "Bunna Bank",
    "bunna int bank": "Bunna Bank",
    "bunna international bank": "Bunna Bank",
    # CBE
    "cbe": "CBE",
    "commercial bank": "CBE",
    "commercial bank of ethiopia": "CBE",
    "cbé": "CBE",
    "cbébirr": "CBE",
    "cbébírr": "CBE",
    # CBO / Coop
    "cbo": "CBO",
    "cbo switch": "CBO",
    "coop": "CBO",
    "coop bank": "CBO",
    "coop bank of oromia": "CBO",
    "cooperative bank of oromia": "CBO",
    "coopay-e-birr": "CBO",
    "coopay": "CBO",
    # Global / Debub
    "global": "Global Bank",
    "global bank": "Global Bank",
    "debub": "Global Bank",
    "debub bank": "Global Bank",
    "dedebit": "Global Bank",
    # Dashen
    "dashen": "Dashen Bank",
    "dashen bank": "Dashen Bank",
    "db": "Dashen Bank",
    # Enat
    "enat": "Enat Bank",
    "enat bank": "Enat Bank",
    # Gadda / Gadaa
    "gadda": "Gadda Bank",
    "gadda bank": "Gadda Bank",
    "gadaa": "Gadda Bank",
    "gadaa bank": "Gadda Bank",
    "gada": "Gadda Bank",
    "gedaa": "Gadda Bank",
    "gedaa bank": "Gadda Bank",
    # Goh Betoch
    "goh": "Goh Betoch Bank",
    "goh betoch": "Goh Betoch Bank",
    "goh betoch bank": "Goh Betoch Bank",
    # Hijra
    "hijra": "Hijra Bank",
    "hijra bank": "Hijra Bank",
    # Lion
    "lion": "Lion Bank",
    "lion bank": "Lion Bank",
    "lion int": "Lion Bank",
    "lion int bank": "Lion Bank",
    "lion international bank": "Lion Bank",
    "lib": "Lion Bank",
    # Nib
    "nib": "Nib Bank",
    "nib bank": "Nib Bank",
    "nib int": "Nib Bank",
    "nib int bank": "Nib Bank",
    "nib international": "Nib Bank",
    "nib international bank": "Nib Bank",
    "nibbirr": "Nib Bank",
    # Oromia
    "oromia": "Oromia Bank",
    "oromia bank": "Oromia Bank",
    "oib": "Oromia Bank",
    # Rammis
    "rammis": "Rammis Bank",
    "rammis bank": "Rammis Bank",
    "raamis": "Rammis Bank",
    "raammis bank": "Rammis Bank",
    "ramis": "Rammis Bank",
    "ramis bank": "Rammis Bank",
    # Santim Pay
    "santimpay": "Santim Pay",
    "santim pay": "Santim Pay",
    "santim": "Santim Pay",
    # Sinqee
    "sinqee": "Sinqee Bank",
    "sinqee bank": "Sinqee Bank",
    "siinqee": "Sinqee Bank",
    "siinqee bank": "Sinqee Bank",
    "siinqee wallet": "Sinqee Bank",
    # Sidama
    "sidama": "Sidama Bank",
    "sidama bank": "Sidama Bank",
    # Siket
    "siket": "Siket Bank",
    "siket bank": "Siket Bank",
    # Tseday
    "tseday": "Tseday Bank",
    "tseday bank": "Tseday Bank",
    "tsedey": "Tseday Bank",
    "tsedey bank": "Tseday Bank",
    # Tsehay
    "tsehay": "Tsehay Bank",
    "tsehay bank": "Tsehay Bank",
    # United / Hibret
    "united": "United Bank",
    "united bank": "United Bank",
    "hibret": "United Bank",
    "hibret bank": "United Bank",
    "ub": "United Bank",
    "h-cash": "United Bank",
    # Wegagen
    "wegagen": "Wegagen Bank",
    "wegagen bank": "Wegagen Bank",
    "wb": "Wegagen Bank",
    "wegagen e-birr": "Wegagen Bank",
    # Yagout Pay
    "yagoutpay": "Yagout Pay",
    "yagout pay": "Yagout Pay",
    "yagout": "Yagout Pay",
    # Zamzam
    "zamzam": "Zamzam Bank",
    "zamzam bank": "Zamzam Bank",
    "zam zam": "Zamzam Bank",
    "zam zam bank": "Zamzam Bank",
    # Zemen
    "zemen": "Zemen Bank",
    "zemen bank": "Zemen Bank",
    "zb": "Zemen Bank",
}


def normalize_nbe_bank(name: Any) -> str:
    """Return the canonical NBE display name for a bank, or title-cased cleaned value."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).strip()
    if not s:
        return ""
    key = s.lower()
    if key in _NBE_ALIAS_MAP:
        return _NBE_ALIAS_MAP[key]

    # Try removing trailing suffixes
    for suffix in (" bank", " int bank", " international bank", " ebirr", " e-birr", " wallet"):
        if key.endswith(suffix):
            base_key = key[:-len(suffix)].strip()
            if base_key in _NBE_ALIAS_MAP:
                return _NBE_ALIAS_MAP[base_key]

    return s.title()


def _is_success_resp(val: Any) -> bool:
    """Return True if response code is -1 or -1.0."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return False
    s = str(val).strip()
    if s in ("-1", "-1.0"):
        return True
    try:
        return float(s) == -1.0
    except (ValueError, TypeError):
        return False


# NBE Mode configuration: label, valid TRANS_TYPE values, whether only
# successful (-1 / -1.0) transactions are counted, and whether monetary
# amounts are reported (balance inquiries have no amount).
_NBE_MODE_CONFIGS: dict[str, dict] = {
    "pos": {
        "label": "PURCHASE",
        "valid_types": ("pos purchase", "purchase"),
        "success_only": True,
        "include_amount": True,
    },
    "atm": {
        "label": "CASH WITHDRAWAL",
        "valid_types": ("atm cash withdrawal", "cash withdrawal"),
        "success_only": True,
        "include_amount": True,
    },
    "pos_decline": {
        "label": "PURCHASE",
        "valid_types": ("pos purchase", "purchase"),
        "success_only": False,
        "include_amount": True,
    },
    "balance_inquiry": {
        "label": "BALANCE INQUIRY",
        "valid_types": ("pos balance inquiry", "balance inquiry"),
        "success_only": False,
        "include_amount": False,
    },
}


def generate_nbe_report(records: list[dict[str, Any]], mode_key: str) -> pd.DataFrame:
    """Generate NBE Institution Summary DataFrame for POS / ATM / POS Decline /
    Balance Inquiry records.

    Filters:
    - pos:           TRANS_TYPE in ('pos purchase', 'purchase'), RESP in (-1, -1.0)
    - atm:           TRANS_TYPE in ('atm cash withdrawal', 'cash withdrawal'), RESP in (-1, -1.0)
    - pos_decline:   TRANS_TYPE in ('pos purchase', 'purchase'), any RESP (declined files)
    - balance_inquiry: TRANS_TYPE in ('pos balance inquiry', 'balance inquiry'), counts only
    """
    cfg = _NBE_MODE_CONFIGS.get(mode_key)
    if cfg is None:
        raise ValueError(
            f"NBE report is only supported for 'pos', 'atm', 'pos_decline' and "
            f"'balance_inquiry' modes, got '{mode_key}'"
        )

    valid_types = set(cfg["valid_types"])
    trans_label = cfg["label"]
    success_only = cfg["success_only"]
    include_amount = cfg["include_amount"]

    # Data aggregators per institution
    stats: dict[str, dict[str, float]] = {}

    def _get_bank_stat(b: str) -> dict[str, float]:
        if b not in stats:
            stats[b] = {
                "issuer_count": 0,
                "issuer_amount": 0.0,
                "acquirer_count": 0,
                "acquirer_amount": 0.0,
            }
        return stats[b]

    # Process matching records
    for r in records:
        # Check transaction type
        t_type = str(r.get("TRANS_TYPE", "")).strip().lower()
        if t_type not in valid_types:
            continue

        # Check response code (-1 / -1.0) only for success-filtered modes.
        # pos_decline and balance_inquiry count every matching transaction.
        resp_val = r.get("RESP")
        if resp_val is None and "RESP_CODE" in r:
            resp_val = r.get("RESP_CODE")
        if resp_val is None and "RESP CODE" in r:
            resp_val = r.get("RESP CODE")
        if resp_val is None and "STATUS" in r:
            resp_val = r.get("STATUS")

        if success_only and not _is_success_resp(resp_val):
            continue

        # Amount (only tracked for modes that report monetary values)
        amt = 0.0
        if include_amount:
            try:
                amt = float(r.get("AMOUNT", 0) or 0)
            except (ValueError, TypeError):
                amt = 0.0

        iss = normalize_nbe_bank(r.get("ISSUER"))
        acq = normalize_nbe_bank(r.get("ACQUIRER"))

        if iss:
            st_iss = _get_bank_stat(iss)
            st_iss["issuer_count"] += 1
            st_iss["issuer_amount"] += amt

        if acq:
            st_acq = _get_bank_stat(acq)
            st_acq["acquirer_count"] += 1
            st_acq["acquirer_amount"] += amt

    # Build full institution list: Standard 31 NBE banks first, followed by any extras
    bank_list = list(STANDARD_NBE_BANKS)
    seen_banks = set(bank_list)
    extra_banks = sorted([b for b in stats.keys() if b and b not in seen_banks])
    bank_list.extend(extra_banks)

    rows = []
    tot_iss_cnt = 0
    tot_iss_amt = 0.0
    tot_acq_cnt = 0
    tot_acq_amt = 0.0

    for idx, b in enumerate(bank_list, start=1):
        s = stats.get(b, {"issuer_count": 0, "issuer_amount": 0.0, "acquirer_count": 0, "acquirer_amount": 0.0})
        i_cnt = int(s["issuer_count"])
        i_amt = round(float(s["issuer_amount"]), 2)
        a_cnt = int(s["acquirer_count"])
        a_amt = round(float(s["acquirer_amount"]), 2)

        tot_iss_cnt += i_cnt
        if include_amount:
            tot_iss_amt += i_amt
        tot_acq_cnt += a_cnt
        if include_amount:
            tot_acq_amt += a_amt

        row = {
            "S/N": idx,
            "BANKS": b,
            f"{trans_label} As Issuer (Count)": i_cnt,
            f"{trans_label} As Acquirer (Count)": a_cnt,
        }
        if include_amount:
            row[f"{trans_label} As Issuer (Amount ETB)"] = i_amt
            row[f"{trans_label} As Acquirer (Amount ETB)"] = a_amt
        rows.append(row)

    # Summary row
    total_row = {
        "S/N": "",
        "BANKS": "Total",
        f"{trans_label} As Issuer (Count)": tot_iss_cnt,
        f"{trans_label} As Acquirer (Count)": tot_acq_cnt,
    }
    if include_amount:
        total_row[f"{trans_label} As Issuer (Amount ETB)"] = round(tot_iss_amt, 2)
        total_row[f"{trans_label} As Acquirer (Amount ETB)"] = round(tot_acq_amt, 2)
    rows.append(total_row)

    return pd.DataFrame(rows)


def build_nbe_report_excel(df: pd.DataFrame, mode_key: str) -> bytes:
    """Build formatted Excel file matching NBE Report layout.

    ``mode_key`` selects the report type and therefore the column layout:
    - pos / pos_decline: PURCHASE with Count + Amount (6 columns)
    - atm:               CASH WITHDRAWAL with Count + Amount (6 columns)
    - balance_inquiry:   BALANCE INQUIRY counts only (4 columns)
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "NBE Report"

    cfg = _NBE_MODE_CONFIGS.get(mode_key)
    if cfg is None:
        raise ValueError(f"Unknown NBE mode '{mode_key}'.")

    trans_label = cfg["label"]
    include_amount = cfg["include_amount"]
    mode_label = mode_key.upper().replace("_", " ")

    # Styling definitions
    font_family = "Arial"

    header_title_font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    header_title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    sub_header_font = Font(name=font_family, size=11, bold=True, color="FFFFFF")
    sub_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")

    col_hdr_font = Font(name=font_family, size=10, bold=True, color="1F4E78")
    col_hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    data_font = Font(name=font_family, size=10)
    total_font = Font(name=font_family, size=11, bold=True)
    total_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    thick_top_double_bottom = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

    n_cols = 6 if include_amount else 4

    # Row 1: Report Main Header
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = f"NBE REPORT - {mode_label} ({trans_label})"
    title_cell.font = header_title_font
    title_cell.fill = header_title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Row 2: Section Header (BANKS | As Issuer | As Acquirer)
    ws["A2"] = "S/N"
    ws["B2"] = "BANKS"
    if include_amount:
        ws.merge_cells("C2:D2")
        ws["C2"] = f"{trans_label} As Issuer"
        ws.merge_cells("E2:F2")
        ws["E2"] = f"{trans_label} As Acquirer"
    else:
        ws["C2"] = f"{trans_label} As Issuer"
        ws["D2"] = f"{trans_label} As Acquirer"

    for col in range(1, n_cols + 1):
        cell = ws.cell(row=2, column=col)
        cell.font = sub_header_font
        cell.fill = sub_header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 24

    # Row 3: Sub-columns (S/N | BANKS | Count | [Amount] | Count | [Amount])
    ws["A3"] = "S/N"
    ws["B3"] = "BANKS"
    ws["C3"] = "Count"
    if include_amount:
        ws["D3"] = "Amount (ETB)"
        ws["E3"] = "Count"
        ws["F3"] = "Amount (ETB)"
    else:
        ws["D3"] = "Count"

    for col in range(1, n_cols + 1):
        cell = ws.cell(row=3, column=col)
        cell.font = col_hdr_font
        cell.fill = col_hdr_fill
        cell.alignment = Alignment(horizontal="center" if col in (1, 3, n_cols) else ("left" if col == 2 else "right"), vertical="center")
    ws.row_dimensions[3].height = 20

    # Data Rows
    current_row = 4

    for idx, row in df.iterrows():
        is_total = (idx == len(df) - 1)
        row_num = current_row

        ws.cell(row=row_num, column=1, value=row.iloc[0])
        ws.cell(row=row_num, column=2, value=row.iloc[1])

        c_cnt = ws.cell(row=row_num, column=3, value=row.iloc[2])
        a_cnt = ws.cell(row=row_num, column=4 if not include_amount else 5, value=row.iloc[3] if not include_amount else row.iloc[4])
        c_cnt.number_format = "#,##0"
        a_cnt.number_format = "#,##0"

        if include_amount:
            c_amt = ws.cell(row=row_num, column=4, value=row.iloc[3])
            a_amt = ws.cell(row=row_num, column=6, value=row.iloc[5])
            c_amt.number_format = "#,##0.00"
            a_amt.number_format = "#,##0.00"

        for col in range(1, n_cols + 1):
            cell = ws.cell(row=row_num, column=col)
            if is_total:
                cell.font = total_font
                cell.fill = total_fill
                cell.border = thick_top_double_bottom
            else:
                cell.font = data_font
                cell.border = thin_border

            if col == 1:
                cell.alignment = Alignment(horizontal="center")
            elif col == 2:
                cell.alignment = Alignment(horizontal="left")
            else:
                cell.alignment = Alignment(horizontal="right")

        ws.row_dimensions[row_num].height = 20
        current_row += 1

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws.column_dimensions["B"].width = 24
    if include_amount:
        ws.column_dimensions["D"].width = 20
        ws.column_dimensions["F"].width = 20
    else:
        ws.column_dimensions["C"].width = 22
        ws.column_dimensions["D"].width = 22

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


# ---------------------------------------------------------------------------
# Sett(Sum) report - merges the two settlement sheets (ISS_BANKS and
# ACQ_BANKS) into one row per bank, summing every numeric column.
# ---------------------------------------------------------------------------
_SETT_METRIC_COLUMNS = (
    "CASH_WITHDRAWAL",
    "AMOUNT_CW",
    "BALANCE_INQUIRY",
    "PURCHASE",
    "AMOUNT_POS",
    "STATEMENT",
)

# Columns that hold counts (no decimals) vs amounts (2 decimals)
_SETT_COUNT_COLUMNS = set("CASH_WITHDRAWAL BALANCE_INQUIRY PURCHASE STATEMENT".split())

# Normalised header -> canonical column (first column of a sheet is the bank name)
_SETT_HEADER_ALIASES = {
    "iss_banks": "BANKS",
    "acq_banks": "BANKS",
    "banks": "BANKS",
    "bank": "BANKS",
    "cash_withdrawal": "CASH_WITHDRAWAL",
    "cash withdrawal": "CASH_WITHDRAWAL",
    "amount_cw": "AMOUNT_CW",
    "amount cw": "AMOUNT_CW",
    "balance_inquiry": "BALANCE_INQUIRY",
    "balance inquiry": "BALANCE_INQUIRY",
    "purchase": "PURCHASE",
    "purchases": "PURCHASE",
    "amount_pos": "AMOUNT_POS",
    "amount pos": "AMOUNT_POS",
    "statement": "STATEMENT",
    "stau": "STATEMENT",
    "statu": "STATEMENT",
    "status": "STATEMENT",
}


def _sett_header_name(raw: Any) -> str:
    """Lower-case, whitespace-trimmed header cell value."""
    if raw is None:
        return ""
    return str(raw).strip().lower()


def _sett_num(value: Any) -> float:
    """Convert a settlement cell to a float (0.0 when missing / unparseable)."""
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).strip()
    if not s:
        return 0.0
    try:
        return float(s)
    except (ValueError, TypeError):
        return 0.0


def _sett_sheet_to_grid(ws: Any) -> tuple[list[str], list[list[Any]]]:
    """Split one settlement sheet into (canonical columns, data rows).

    The header row is located by looking for a row whose cells map to at
    least one bank-name column ("ISS_BANKS" / "ACQ_BANKS") and two metric
    columns. Returns the canonical column names in source order plus every
    data row as raw values.
    """
    rows = [list(r) for r in ws.iter_rows(values_only=True)] or []

    header_idx = None
    for i, row in enumerate(rows):
        norm = {c: _sett_header_name(v) for c, v in enumerate(row)}
        nonempty = [(c, n) for c, n in norm.items() if n]
        if not nonempty:
            continue
        has_bank = any(n in ("iss_banks", "acq_banks", "banks", "bank") for _, n in nonempty)
        metric_hits = sum(
            1 for _, n in nonempty
            if n in _SETT_HEADER_ALIASES and _SETT_HEADER_ALIASES[n] != "BANKS"
        )
        if has_bank and metric_hits >= 2:
            header_idx = i
            break
    if header_idx is None:
        return [], []

    header = rows[header_idx]
    col_map: list[str] = []  # canonical name per column ("" = skip)
    bank_col: int | None = None
    for c, v in enumerate(header):
        canon = _SETT_HEADER_ALIASES.get(_sett_header_name(v))
        if canon is None:
            col_map.append("")
            continue
        if canon == "BANKS":
            if bank_col is None:
                bank_col = c
            col_map.append("BANKS")
        else:
            col_map.append(canon)

    data_rows: list[list[Any]] = []
    for row in rows[header_idx + 1:]:
        if len(row) <= bank_col or _sett_header_name(row[bank_col]) == "":
            continue  # skip blank rows and the SUM total row
        data_rows.append(list(row))

    return col_map, data_rows


def generate_sett_sum_report(file_bytes: bytes) -> pd.DataFrame:
    """Generate a Sett(Sum) DataFrame from a bini-style settlement workbook.

    The workbook holds two sheets (ISS_BANKS and ACQ_BANKS). The first
    column of each sheet is the bank name; the remaining columns are
    per-bank metrics (CASH_WITHDRAWAL, AMOUNT_CW, BALANCE_INQUIRY,
    PURCHASE, AMOUNT_POS, STATEMENT). Rows carrying the same bank name
    (after NBE normalisation) are merged into a single row and every
    matching numeric column is summed across both sheets, so the output
    has exactly one row per institution plus a Total row.
    """
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001 - surface a friendly message
        raise ValueError(f"Could not read the settlement workbook ({exc}).") from exc

    # Read every sheet into memory before closing the workbook (read-only
    # worksheets cannot be iterated after the archive is closed).
    try:
        grids = [_sett_sheet_to_grid(ws) for ws in wb.worksheets]
    finally:
        wb.close()

    sheets = [(cm, dr) for cm, dr in grids if cm and dr]

    if not sheets:
        raise ValueError("The settlement workbook contains no readable sheets.")

    stats: dict[str, dict[str, float]] = {}

    for col_map, data_rows in sheets:
        if not col_map or not data_rows:
            continue
        bank_col = col_map.index("BANKS")
        metric_cols = [c for c, name in enumerate(col_map) if name not in ("", "BANKS")]
        if not metric_cols:
            continue

        for row in data_rows:
            bank = normalize_nbe_bank(row[bank_col] if bank_col < len(row) else None)
            if not bank:
                continue
            st = stats.setdefault(bank, {m: 0.0 for m in _SETT_METRIC_COLUMNS})
            for c in metric_cols:
                if c >= len(row):
                    continue
                canon = col_map[c]
                st[canon] = st.get(canon, 0.0) + _sett_num(row[c])

    if not stats:
        raise ValueError(
            "No bank rows found in the settlement workbook (expected two sheets "
            "whose first column is ISS_BANKS / ACQ_BANKS)."
        )

    bank_list = [b for b in STANDARD_NBE_BANKS if b in stats]
    extra_banks = sorted([b for b in stats.keys() if b not in set(STANDARD_NBE_BANKS)])
    bank_list.extend(extra_banks)

    rows = []
    totals = {m: 0.0 for m in _SETT_METRIC_COLUMNS}
    for idx, b in enumerate(bank_list, start=1):
        s = stats[b]
        row: dict[str, Any] = {"S/N": idx, "BANKS": b}
        for m in _SETT_METRIC_COLUMNS:
            v = s.get(m, 0.0)
            row[m] = int(round(v)) if m in _SETT_COUNT_COLUMNS else round(v, 2)
            totals[m] += v
        rows.append(row)

    total_row: dict[str, Any] = {"S/N": "", "BANKS": "Total"}
    for m in _SETT_METRIC_COLUMNS:
        total_row[m] = (
            int(round(totals[m])) if m in _SETT_COUNT_COLUMNS else round(totals[m], 2)
        )
    rows.append(total_row)

    return pd.DataFrame(rows)


def build_sett_sum_report_excel(df: pd.DataFrame) -> bytes:
    """Build a formatted Excel workbook for the Sett(Sum) report."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sett(Sum)"

    font_family = "Arial"

    header_title_font = Font(name=font_family, size=14, bold=True, color="FFFFFF")
    header_title_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")

    col_hdr_font = Font(name=font_family, size=10, bold=True, color="1F4E78")
    col_hdr_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")

    data_font = Font(name=font_family, size=10)
    total_font = Font(name=font_family, size=11, bold=True)
    total_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    thick_top_double_bottom = Border(
        top=Side(style="thin", color="000000"),
        bottom=Side(style="double", color="000000"),
    )

    n_cols = len(df.columns)

    # Row 1: Report Title
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "SETT(SUM) - SETTLEMENT SUMMARY"
    title_cell.font = header_title_font
    title_cell.fill = header_title_fill
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Row 2: Column Headers
    for col_idx, col_name in enumerate(df.columns, start=1):
        cell = ws.cell(row=2, column=col_idx, value=col_name)
        cell.font = col_hdr_font
        cell.fill = col_hdr_fill
        cell.alignment = Alignment(
            horizontal="center" if col_idx == 1 else ("left" if col_idx == 2 else "right"),
            vertical="center",
        )
    ws.row_dimensions[2].height = 20

    # Data Rows
    current_row = 3
    for idx, row in df.iterrows():
        is_total = (idx == len(df) - 1)
        row_num = current_row
        for col_idx, col_name in enumerate(df.columns, start=1):
            cell = ws.cell(row=row_num, column=col_idx, value=row.iloc[col_idx - 1])
            if col_name in _SETT_METRIC_COLUMNS:
                cell.number_format = "#,##0" if col_name in _SETT_COUNT_COLUMNS else "#,##0.00"
            if is_total:
                cell.font = total_font
                cell.fill = total_fill
                cell.border = thick_top_double_bottom
            else:
                cell.font = data_font
                cell.border = thin_border
            if col_idx == 1:
                cell.alignment = Alignment(horizontal="center")
            elif col_idx == 2:
                cell.alignment = Alignment(horizontal="left")
            else:
                cell.alignment = Alignment(horizontal="right")
        ws.row_dimensions[row_num].height = 20
        current_row += 1

    # Auto-adjust column widths
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 4, 12)

    ws.column_dimensions["B"].width = 24
    for m in ("AMOUNT_CW", "AMOUNT_POS"):
        if m in df.columns:
            col_idx = list(df.columns).index(m) + 1
            ws.column_dimensions[get_column_letter(col_idx)].width = 20

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()
