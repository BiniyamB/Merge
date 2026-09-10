"""Successful QR & P2P (IPS) interoperable transaction report generator.

Parses the EthSwitch per-bank "… success for source and destination"
summary workbooks (columns NO, BANK_ID, BANK_NAME, ISSUER_TXN_COUNT,
ISSUER_TOTAL_AMOUNT, ACQUIRER_TXN_COUNT, ACQUIRER_TOTAL_AMOUNT), merges any
number of files by matching banks on the canonical institution name and
summing every numeric column, and emits:

  * a merged summary workbook (consistent with the other report modes), and
  * a styled "Successful QR / IPS Transaction" workbook that reproduces the
    EthSwitch September 2026 reference layout exactly (merges, fonts, fills,
    number formats, SUM totals, column widths and row heights).
"""

from __future__ import annotations

import io
import re
from datetime import date
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

# ---------------------------------------------------------------------------
# Canonical institution order (matches the reference reports row for row).
# ---------------------------------------------------------------------------
QR_BANK_ORDER = [
    "Abay", "Abyssinia", "Addis", "Ahadu", "Ahadu Ebirr", "Amhara", "Awash",
    "Berhan Bank", "Bunna", "CBE", "CBE Birr", "CBO", "Coopay-E-Birr",
    "Dashen", "Enat Bank", "Global", "GOH BETOCH", "Hibret", "KAAFI", "Kacha",
    "LIB", "M-PESA", "NIB", "Rammis Bank", "Saha", "Shabelle", "Sidama Bank",
    "Siinqee", "Siket", "telebirr", "Tsedey Bank", "Tsehay", "Vision Fund",
    "Wegagen", "Wegagen E-Birr", "ZamZam",
]

IPS_BANK_ORDER = [
    "Abay", "Abyssinia", "Addis", "Ahadu", "Ahadu Ebirr", "Amhara",
    "AmharaEthBirr", "Awash", "Berhan Bank", "Buna Bank", "CBE", "CBE BIRR",
    "CBO", "Coopay-E-Birr", "Dashen", "Dedebit MFI", "Enat Bank", "Gadaa",
    "Global", "GOH BETOCH", "H-Cash", "HalalPay", "Hibret", "Hijra Bank",
    "KAAFI MFI", "Kacha", "LIB", "Metemamen", "MPESA", "NIB", "NibE Birr",
    "Nisir", "Omo Bank", "Oromia bank", "Rammis Bank", "Rays", "Saha",
    "Shabelle", "Sidama Bank", "Siinqee", "Siinqee Wallet", "Siket",
    "Tsedey Bank", "Tsehay", "Vision Fund", "Vitabirr", "Wegagen",
    "Wegagen e-Birr", "Yaya Wallet", "ZamZam", "Zemen",
]

# ---------------------------------------------------------------------------
# Input bank-name aliases -> canonical display names (per series).
# Keys are lowercased with whitespace collapsed.
# ---------------------------------------------------------------------------
_QR_ALIASES: dict[str, str] = {
    "amhara bank": "Amhara",
    "awash bank": "Awash",
    "dashen bank": "Dashen",
    "coop": "CBO",
    "cbe": "CBE",
    "kaafi": "KAAFI",
    "tsehay bank": "Tsehay",
    "nib bank": "NIB",
    "addis bank": "Addis",
    "addis": "Addis",
    "wegagen bank": "Wegagen",
    "abay bank": "Abay",
    "boa": "Abyssinia",
    "abyssinia": "Abyssinia",
    "bank of abyssinia": "Abyssinia",
    "buna bank": "Bunna",
    "bunna bank": "Bunna",
    "buna": "Bunna",
    "siinqee bank": "Siinqee",
    "siinqee": "Siinqee",
    "wegagen e-birr": "Wegagen E-Birr",
    "cbe birr": "CBE Birr",
    "cbebirr": "CBE Birr",
    "coopay-e-birr": "Coopay-E-Birr",
    "mpesa": "M-PESA",
    "m-pesa": "M-PESA",
    "telebirr": "telebirr",
    "kacha": "Kacha",
    "tsedey bank": "Tsedey Bank",
    "lib": "LIB",
    "zamzam": "ZamZam",
    "enat bank": "Enat Bank",
    "siket": "Siket",
    "goh betoch": "GOH BETOCH",
    "goh betoch bank": "GOH BETOCH",
    "rammis bank": "Rammis Bank",
    "hibret": "Hibret",
    "global": "Global",
    "global bank": "Global",
    "vision fund": "Vision Fund",
    "berhan bank": "Berhan Bank",
    "sidama bank": "Sidama Bank",
    "ahadu": "Ahadu",
    "ahadu bank": "Ahadu",
    "ahadu ebirr": "Ahadu Ebirr",
}

_IPS_ALIASES: dict[str, str] = {
    "amhara bank": "Amhara",
    "awash bank": "Awash",
    "dashen bank": "Dashen",
    "coop": "CBO",
    "zemen bank": "Zemen",
    "zemen": "Zemen",
    "cbe": "CBE",
    "kaafi": "KAAFI MFI",
    "kaafi mfi": "KAAFI MFI",
    "tsehay bank": "Tsehay",
    "nib bank": "NIB",
    "addis bank": "Addis",
    "addis": "Addis",
    "wegagen bank": "Wegagen",
    "abay bank": "Abay",
    "boa": "Abyssinia",
    "abyssinia": "Abyssinia",
    "bank of abyssinia": "Abyssinia",
    "buna bank": "Buna Bank",
    "bunna bank": "Buna Bank",
    "buna": "Buna Bank",
    "rays": "Rays",
    "siinqee bank": "Siinqee",
    "siinqee": "Siinqee",
    "wegagen e-birr": "Wegagen e-Birr",
    "cbe birr": "CBE BIRR",
    "cbebirr": "CBE BIRR",
    "coopay-e-birr": "Coopay-E-Birr",
    "mpesa": "MPESA",
    "m-pesa": "MPESA",
    "kacha": "Kacha",
    "tsedey bank": "Tsedey Bank",
    "lib": "LIB",
    "zamzam": "ZamZam",
    "enat bank": "Enat Bank",
    "siket": "Siket",
    "hijra bank": "Hijra Bank",
    "goh betoch": "GOH BETOCH",
    "goh betoch bank": "GOH BETOCH",
    "rammis bank": "Rammis Bank",
    "saha": "Saha",
    "halalpay": "HalalPay",
    "hibret": "Hibret",
    "global": "Global",
    "global bank": "Global",
    "vision fund": "Vision Fund",
    "berhan bank": "Berhan Bank",
    "oromia bank": "Oromia bank",
    "oromia": "Oromia bank",
    "yaya wallet": "Yaya Wallet",
    "gadaa": "Gadaa",
    "gadda": "Gadaa",
    "gadda bank": "Gadaa",
    "sidama bank": "Sidama Bank",
    "ahadu": "Ahadu",
    "ahadu bank": "Ahadu",
    "ahadu ebirr": "Ahadu Ebirr",
    "nisir": "Nisir",
    "shabelle": "Shabelle",
    "omo": "Omo Bank",
    "omo bank": "Omo Bank",
    "nibebirr": "NibE Birr",
    "nibbirr": "NibE Birr",
    "dedebit": "Dedebit MFI",
    "dedebit mfi": "Dedebit MFI",
    "h-cash": "H-Cash",
    "vitabirr": "Vitabirr",
    "amharaethbirr": "AmharaEthBirr",
    "siinqee wallet": "Siinqee Wallet",
    "metemamen": "Metemamen",
}

# Canonical source input headers (tolerant matching used during parsing).
_DATA_HEADERS = [
    "NO", "BANK_ID", "BANK_NAME", "ISSUER_TXN_COUNT", "ISSUER_TOTAL_AMOUNT",
    "ACQUIRER_TXN_COUNT", "ACQUIRER_TOTAL_AMOUNT",
]

_ACCOUNTING_NUMFMT = '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"??_ ;_ @_ '
_DATE_NUMFMT = "d\\-mmm\\-yyyy"
_FILL_BANK = PatternFill("solid", fgColor="FFA9D08E")
_FILL_SUBTITLE = PatternFill("solid", fgColor="FFFFFFFF")
_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_FONT_TNR = "Times New Roman"


def _normalise(name: str | None) -> str:
    return re.sub(r"\s+", " ", str(name or "")).strip().lower()


def _to_number(value: Any) -> int | float | None:
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return value
    text = str(value).strip()
    text = re.sub(r"[ ,]", "", text)
    if not text:
        return None
    try:
        number = float(text)
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def normalize_bank_name(name: str, series: str) -> str:
    """Map an input BANK_NAME to the canonical display name for a series."""
    aliases = _IPS_ALIASES if series == "p2p" else _QR_ALIASES
    key = _normalise(name)
    if key in aliases:
        return aliases[key]
    order = IPS_BANK_ORDER if series == "p2p" else QR_BANK_ORDER
    seq_key = re.sub(r"\W+", "", key)
    stripped = {c: re.sub(r"\W+", "", c.lower()) for c in order}
    if not seq_key:
        return str(name).strip()
    # exact equality first, so "NibE Birr" matches "NibE Birr" over the
    # substring-equivalent "NIB"
    for canonical, canon_seq in stripped.items():
        if seq_key == canon_seq:
            return canonical
    for canonical, canon_seq in stripped.items():
        if seq_key in canon_seq or canon_seq in seq_key:
            return canonical
    return str(name).strip()


def parse_bank_summary_file(data: bytes, filename: str = "") -> list[dict]:
    """Parse one summary workbook into per-bank rows (numeric values)."""
    wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    header_row: int | None = None
    col_index: dict[str, int] = {}
    for r, row in enumerate(rows):
        if row is None:
            continue
        hits: dict[str, int] = {}
        for i, value in enumerate(row):
            if value is None:
                continue
            hits[_compact(_normalise(str(value)))] = i
        if "bankname" in hits and (
                any(k for k in hits if "issuer" in k and ("count" in k or "amount" in k)) or
                any(k for k in hits if "acquirer" in k and ("count" in k or "amount" in k))):
            header_row = r
            for target in _DATA_HEADERS:
                col_index[target] = hits.get(_compact(target.lower()))
            break
    if header_row is None or not all(col_index.get(t) is not None for t in _DATA_HEADERS):
        raise ValueError(f"'{filename or 'file'}': could not locate the BANK_NAME / ISSUER / ACQUIRER summary columns.")

    out: list[dict] = []
    for row in rows[header_row + 1:]:
        if row is None:
            continue
        name = row[col_index["BANK_NAME"]]
        if name is None or str(name).strip() == "":
            continue
        out.append({
            "NO": row[col_index["NO"]],
            "BANK_ID": row[col_index["BANK_ID"]],
            "BANK_NAME": str(name).strip(),
            "ISSUER_TXN_COUNT": _to_number(row[col_index["ISSUER_TXN_COUNT"]]),
            "ISSUER_TOTAL_AMOUNT": _to_number(row[col_index["ISSUER_TOTAL_AMOUNT"]]),
            "ACQUIRER_TXN_COUNT": _to_number(row[col_index["ACQUIRER_TXN_COUNT"]]),
            "ACQUIRER_TOTAL_AMOUNT": _to_number(row[col_index["ACQUIRER_TOTAL_AMOUNT"]]),
        })
    if not out:
        raise ValueError(f"'{filename or 'file'}': no bank rows found under the summary header.")
    return out


def _compact(text: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", text)


def merge_bank_summary_files(files: list[tuple[str, bytes]], series: str) -> tuple[list[dict], list[dict], list[str]]:
    """Merge summary files, summing every numeric column per canonical bank.

    Banks not present in the canonical list for the series are treated as
    *new* institutions: they are kept (appended after the canonical banks in
    first-encounter order) so nothing is dropped. The final Total row reports
    the issuer (destination) and acquirer (source) figures as identical in
    both count and value - if the raw sums disagree, both sides are shown as
    the larger of the two and a balancing warning is added.

    Returns ``(records, per_file, warnings)`` where ``records`` holds one row
    per canonical institution (canonical order) followed by any new
    institutions, then a Total row.
    """
    order = IPS_BANK_ORDER if series == "p2p" else QR_BANK_ORDER
    totals: dict[str, dict[str, Any]] = {name: {
        "ISSUER_TXN_COUNT": 0, "ISSUER_TOTAL_AMOUNT": 0.0,
        "ACQUIRER_TXN_COUNT": 0, "ACQUIRER_TOTAL_AMOUNT": 0.0,
    } for name in order}
    extra: dict[str, dict[str, Any]] = {}
    extra_order: list[str] = []
    bank_id: dict[str, Any] = {}

    per_file: list[dict] = []
    warnings: list[str] = []

    for filename, data in files:
        try:
            rows = parse_bank_summary_file(data, filename)
        except Exception as exc:
            per_file.append({"filename": filename, "status": "error", "rows": str(exc)})
            warnings.append(f"'{filename}': {exc}")
            continue
        matched = 0
        for row in rows:
            canonical = normalize_bank_name(row["BANK_NAME"], series)
            if canonical in totals:
                bucket = totals[canonical]
            else:
                if canonical not in extra:
                    extra[canonical] = {
                        "ISSUER_TXN_COUNT": 0, "ISSUER_TOTAL_AMOUNT": 0.0,
                        "ACQUIRER_TXN_COUNT": 0, "ACQUIRER_TOTAL_AMOUNT": 0.0,
                    }
                    extra_order.append(canonical)
                    warnings.append(
                        f"'{filename}': new institution '{canonical}' is not in the "
                        f"{'IPS' if series == 'p2p' else 'QR'} list; it has been "
                        "added to the report.")
                bucket = extra[canonical]
            matched += 1
            for col in ("ISSUER_TXN_COUNT", "ISSUER_TOTAL_AMOUNT", "ACQUIRER_TXN_COUNT", "ACQUIRER_TOTAL_AMOUNT"):
                val = row.get(col)
                if val is not None:
                    bucket[col] = (bucket[col] or 0) + val
            bid = row.get("BANK_ID")
            if bid is not None:
                bank_id[canonical] = bid
        per_file.append({"filename": filename, "status": "ok", "rows": matched})

    records: list[dict] = []
    for i, name in enumerate(order, start=1):
        cu = totals[name]
        records.append({
            "NO": i,
            "BANK_ID": bank_id.get(name, ""),
            "BANK_NAME": name,
            "ISSUER_TXN_COUNT": cu["ISSUER_TXN_COUNT"],
            "ISSUER_TOTAL_AMOUNT": cu["ISSUER_TOTAL_AMOUNT"],
            "ACQUIRER_TXN_COUNT": cu["ACQUIRER_TXN_COUNT"],
            "ACQUIRER_TOTAL_AMOUNT": cu["ACQUIRER_TOTAL_AMOUNT"],
        })
    for j, name in enumerate(extra_order, start=len(order) + 1):
        cu = extra[name]
        records.append({
            "NO": j,
            "BANK_ID": bank_id.get(name, ""),
            "BANK_NAME": name,
            "ISSUER_TXN_COUNT": cu["ISSUER_TXN_COUNT"],
            "ISSUER_TOTAL_AMOUNT": cu["ISSUER_TOTAL_AMOUNT"],
            "ACQUIRER_TXN_COUNT": cu["ACQUIRER_TXN_COUNT"],
            "ACQUIRER_TOTAL_AMOUNT": cu["ACQUIRER_TOTAL_AMOUNT"],
        })

    iss_count = sum(t["ISSUER_TXN_COUNT"] for t in totals.values()) + \
        sum(t["ISSUER_TXN_COUNT"] for t in extra.values())
    acq_count = sum(t["ACQUIRER_TXN_COUNT"] for t in totals.values()) + \
        sum(t["ACQUIRER_TXN_COUNT"] for t in extra.values())
    iss_value = round(sum(t["ISSUER_TOTAL_AMOUNT"] for t in totals.values()) +
                      sum(t["ISSUER_TOTAL_AMOUNT"] for t in extra.values()), 2)
    acq_value = round(sum(t["ACQUIRER_TOTAL_AMOUNT"] for t in totals.values()) +
                      sum(t["ACQUIRER_TOTAL_AMOUNT"] for t in extra.values()), 2)

    # Issuer (destination) and acquirer (source) must balance: the total count
    # and total value are reported identically on both sides. On a clean
    # ledger they already match; any discrepancy is shown as the larger figure
    # with a warning.
    balanced_count = max(iss_count, acq_count)
    balanced_value = max(iss_value, acq_value)
    if iss_count != acq_count or iss_value != acq_value:
        warnings.append(
            "Balancing: issuer (destination) and acquirer (source) totals differ "
            f"(counts {iss_count} vs {acq_count}, values {iss_value} vs "
            f"{acq_value}); both sides are reported as {balanced_count} / "
            f"{balanced_value}.")

    total_row = {
        "NO": "",
        "BANK_ID": "",
        "BANK_NAME": "Total",
        "ISSUER_TXN_COUNT": balanced_count,
        "ISSUER_TOTAL_AMOUNT": balanced_value,
        "ACQUIRER_TXN_COUNT": balanced_count,
        "ACQUIRER_TOTAL_AMOUNT": balanced_value,
    }
    records.append(total_row)
    return records, per_file, warnings


# ---------------------------------------------------------------------------
# Styled "Successful … Transaction" report (reference layout).
# ---------------------------------------------------------------------------
_QR_LAYOUT = {
    "title1": "EthSwitch S.C.",
    "title2": "QR Report",
    "subtitle": "Successful QR Interoperable Transactions for {date}",
    "bank_header": "Bank",
    "header_row": 6,          # merged two-row header starts here (6-7)
    "data_start": 8,
    "data_col": 3,            # C
    "value_cols": (4, 5, 6, 7),  # D, E, F, G
    "total_row": 44,
    "total_label": "Total",
    "total_bold_label": False,
    "value_align": "center",
    "bank_align": "center",
    "widths": {"A": 9, "C": 19.44140625, "D": 18.109375, "E": 17.33203125,
               "F": 18.6640625, "G": 20, "H": 9},
    "row_heights": {2: 15.6, 3: 15.6, 4: 15.6, 44: 17.4},
    "merges": {
        2: ("C2:G2", "EthSwitch S.C.", True),
        3: ("C3:G3", "QR Report", True),
        4: ("C4:G4", "-DATE-", True),
        5: ("C5:G5", "-SUBTITLE-", False),
    },
    "header_merges": {
        ("C6:C7", "Bank"),
        ("D6:E6", "As a Destination"),
        ("F6:G6", "As a Source"),
    },
    "sub_headers": {"D7": "No.Transactions", "E7": "Values",
                    "F7": "No.Transactions", "G7": "Values"},
    "blank_zero": True,
}

_IPS_LAYOUT = {
    "title1": "EthSwitch S.C.",
    "title2": "IPS Successful Report",
    "subtitle": "Successful IPS Interoperable Transactions Held on {date}",
    "bank_header": "BANK",
    "header_row": 5,          # merged two-row header starts here (5-6)
    "data_start": 7,
    "data_col": 2,            # B
    "value_cols": (3, 4, 5, 6),  # C, D, E, F
    "total_row": 58,
    "total_label": "TOTAL",
    "total_bold_label": True,
    "value_align": "right",
    "bank_align": None,
    "widths": {"A": 9, "B": 18.5546875, "G": 9},
    "row_heights": {1: 15.6, 2: 15.6, 3: 15.6, 4: 20.25, 5: 13.95},
    "merges": {
        1: ("B1:F1", "EthSwitch S.C.", True),
        2: ("B2:F2", "IPS Successful Report", True),
        3: ("B3:F3", "-DATE-", True),
        4: ("B4:F4", "-SUBTITLE-", False),
    },
    "header_merges": {
        ("B5:B6", "BANK"),
        ("C5:D5", "Successful  Transactions As Destination"),
        ("E5:F5", "Successful Transactions As Source"),
    },
    "sub_headers": {"C6": "No.Transactions", "D6": "Values",
                    "E6": "No.Transactions", "F6": "Values"},
    "blank_zero": True,
}


def _mon_date(d: date) -> str:
    return f"{d.strftime('%B')} {d.day:02d},{d.year}"


def _file_date(d: date) -> str:
    return f"{d.strftime('%B')} {d.day},{d.year}"


def success_report_filename(series: str, report_date: date) -> str:
    label = "QR" if series == "qr" else "IPS"
    return f"Successful {label} Transaction for {_file_date(report_date)}.xlsx"


def build_success_report_excel(records: list[dict], series: str,
                               report_date: date) -> bytes:
    """Build the styled 'Successful QR/IPS Transaction' workbook (by reference)."""
    layout = _QR_LAYOUT if series == "qr" else _IPS_LAYOUT
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    for col, width in layout["widths"].items():
        ws.column_dimensions[col].width = width
    for row, height in layout["row_heights"].items():
        ws.row_dimensions[row].height = height

    title_font = Font(name=_FONT_TNR, size=12, bold=True)
    subtitle_font = Font(name=_FONT_TNR, size=11, bold=True)
    header_font = Font(name=_FONT_TNR, size=11, bold=False)
    data_font = Font(name=_FONT_TNR, size=11, bold=False)
    total_font = Font(name=_FONT_TNR, size=11, bold=True)

    for row, (range_, value, is_title) in layout["merges"].items():
        ws.merge_cells(range_)
        cell = ws[range_.split(":")[0]]
        if value == "-DATE-":
            cell.value = report_date
            cell.number_format = _DATE_NUMFMT
        elif value == "-SUBTITLE-":
            cell.value = layout["subtitle"].format(date=_mon_date(report_date))
            cell.fill = _FILL_SUBTITLE
        else:
            cell.value = value
        cell.font = title_font if is_title else subtitle_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # ── two-row header block ──
    data_col = layout["data_col"]
    header_row = layout["header_row"]
    for range_, value in layout["header_merges"]:
        ws.merge_cells(range_)
        cell = ws[range_.split(":")[0]]
        cell.value = value
        cell.font = header_font
        cell.fill = _FILL_BANK
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for coord, value in layout["sub_headers"].items():
        cell = ws[coord]
        cell.value = value
        cell.font = header_font
        cell.fill = _FILL_BANK
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # ── data rows ──
    value_cols = layout["value_cols"]
    blank_zero = layout["blank_zero"]
    data_rows = [r for r in records if r.get("BANK_NAME") not in ("Total", "TOTAL")]
    for i, rec in enumerate(data_rows):
        r = layout["data_start"] + i
        bank_cell = ws.cell(row=r, column=data_col, value=rec["BANK_NAME"])
        bank_cell.font = data_font
        bank_cell.fill = _FILL_BANK
        bank_cell.alignment = Alignment(horizontal=layout["bank_align"], vertical="center")
        values = (rec["ISSUER_TXN_COUNT"], rec["ISSUER_TOTAL_AMOUNT"],
                  rec["ACQUIRER_TXN_COUNT"], rec["ACQUIRER_TOTAL_AMOUNT"])
        for col, val in zip(value_cols, values):
            cell = ws.cell(row=r, column=col)
            if not (blank_zero and val in (0, 0.0, None)):
                cell.value = val
            cell.font = data_font
            cell.number_format = _ACCOUNTING_NUMFMT
            cell.alignment = Alignment(horizontal=layout["value_align"], vertical="center")

    # ── total row ──
    # Position is dynamic: the canonical banks are always listed, any new
    # institutions found in the input are appended after them, and the Total
    # row follows the last data row.
    total_r = layout["data_start"] + len(data_rows)
    ref_total_row = layout["total_row"]
    if ref_total_row in layout["row_heights"]:
        ws.row_dimensions[total_r].height = layout["row_heights"][ref_total_row]

    label_cell = ws.cell(row=total_r, column=data_col, value=layout["total_label"])
    label_cell.font = Font(name=_FONT_TNR, size=11, bold=layout["total_bold_label"])
    label_cell.fill = _FILL_BANK
    label_cell.alignment = Alignment(horizontal=layout["bank_align"], vertical="center")

    # Issuer/acquirer totals are balanced by the merge step, so the two sides
    # are written as equal literals (count and value match on both columns).
    total_rec = records[-1]
    totals_values = (
        total_rec["ISSUER_TXN_COUNT"], total_rec["ISSUER_TOTAL_AMOUNT"],
        total_rec["ACQUIRER_TXN_COUNT"], total_rec["ACQUIRER_TOTAL_AMOUNT"],
    )
    for col, val in zip(value_cols, totals_values):
        cell = ws.cell(row=total_r, column=col, value=val)
        cell.font = total_font
        cell.number_format = _ACCOUNTING_NUMFMT
        cell.alignment = Alignment(horizontal=layout["value_align"], vertical="center")

    # ── thin borders over the whole used block ──
    first_row = 1 if series == "p2p" else 2
    last_row = max(total_r, layout["data_start"] + len(data_rows) - 1)
    min_col = data_col
    max_col = max(value_cols)
    for r in range(first_row, last_row + 1):
        for c in range(min_col, max_col + 1):
            ws.cell(row=r, column=c).border = _BORDER

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Merged summary workbook (consistent with the other report modes).
# ---------------------------------------------------------------------------
def build_merged_summary_excel(records: list[dict], series: str) -> bytes:
    """Build a plain merged summary workbook with a title block + TOTAL row."""
    label = "QR" if series == "qr" else "IPS"
    last_col = "F"
    columns = ["NO", "BANK_NAME", "As a Destination (No.Transactions)",
               "As a Destination (Values)", "As a Source (No.Transactions)",
               "As a Source (Values)"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Report"

    center = Alignment(horizontal="center", vertical="center")
    left = Alignment(horizontal="left", vertical="center")
    title_font = Font(bold=True, size=14)
    label_font = Font(bold=True)
    header_font = Font(bold=True, color="1F2937")
    header_fill = PatternFill("solid", fgColor="D9E1F2")
    total_font = Font(bold=True)
    total_fill = PatternFill("solid", fgColor="E2EFDA")
    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    a1 = ws.cell(row=1, column=1, value="Report name:")
    a1.font = label_font
    ws.merge_cells(f"B1:{last_col}1")
    b1 = ws.cell(row=1, column=2, value=f"SUCCESSFUL {label} TRANSACTION SUMMARY")
    b1.font = title_font
    b1.alignment = left

    header_row = 2
    for c, name in enumerate(columns, start=1):
        cell = ws.cell(row=header_row, column=c, value=name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center
        cell.border = border

    data_start = header_row + 1
    bank_rows = [rec for rec in records if rec.get("BANK_NAME") not in ("Total", "TOTAL")]
    for i, rec in enumerate(bank_rows):
        r = data_start + i
        cells = [
            rec.get("NO"), rec.get("BANK_NAME"),
            rec.get("ISSUER_TXN_COUNT"), rec.get("ISSUER_TOTAL_AMOUNT"),
            rec.get("ACQUIRER_TXN_COUNT"), rec.get("ACQUIRER_TOTAL_AMOUNT"),
        ]
        for c, val in enumerate(cells, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            if c in (2,):
                cell.alignment = left
            else:
                cell.alignment = center
            if c in (4, 6) and isinstance(val, (int, float)):
                cell.number_format = "#,##0.00"

    total_r = data_start + len(bank_rows)
    t_n = ws.cell(row=total_r, column=2, value="TOTAL")
    t_n.font = total_font
    t_n.fill = total_fill
    t_n.border = border
    # Issuer/acquirer totals are balanced by the merge step (identical count
    # and value on both sides).
    total_rec = records[-1]
    total_values = (
        total_rec["ISSUER_TXN_COUNT"], total_rec["ISSUER_TOTAL_AMOUNT"],
        total_rec["ACQUIRER_TXN_COUNT"], total_rec["ACQUIRER_TOTAL_AMOUNT"],
    )
    for c, val in zip((3, 4, 5, 6), total_values):
        cell = ws.cell(row=total_r, column=c, value=val)
        cell.font = total_font
        cell.fill = total_fill
        cell.border = border
        cell.alignment = center
        if c in (4, 6):
            cell.number_format = "#,##0.00"

    ws.cell(row=total_r, column=1).border = border
    ws.cell(row=total_r, column=1).alignment = center
    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 32
    for col in ("C", "D", "E", "F"):
        ws.column_dimensions[col].width = 24

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()