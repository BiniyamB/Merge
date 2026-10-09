"""ATM average (success rate) report built from several daily ATM reports.

Each source file is one day of the EthSwitch "ATM Declined Transaction
Report": a sheet whose row 4 holds `RC/BANK NAME` followed by one column per
bank (and a trailing `Total`), whose lower rows list decline response codes
with their counts, and whose `Issu. SUCC. RATE (%)` row carries the success
rate each bank achieved that day::

    B4=CBE  C4=BOA  D4=ABAY ... Y4=GADAA  Z4=Total
    B5..Y34       decline counts per response code
    B35..Y35      =SUM(B5:B34)                       TOTAL
    B36..Y36      =B42  (Issu. SUCC. RATE (%))
    B37..Y37      =B35*88/12                          expected transactions
    B38..Y38      =B5+B15+B17+...                     counted declines
    B39..Y39      =B35-B38                            other declines
    B40..Y40      =B37+B38
    B41..Y41      =B39+B40
    B42..Y42      =B40/B41                            the success rate

Uploading several of those files produces one average report in the shape of
`cfd.xlsx`: a row per day, a column per bank, and a bottom `Average` row of
live `=AVERAGE(...)` formulas. Every rate cell is filled by its tier:

===============  =======
Success rate     Fill
===============  =======
98.5% - 100%     Green  (#00B050)
90% - 98.4%      Yellow (#FFFF00)
80% - 89%        Amber  (#FFC000)
<= 79%           Red    (#FF0000)
===============  =======

Three details the daily files make necessary:

- **Both dates are read and compared.** The date in the file name
  (`ATM Declined Transaction Report for Oct 6,2026.xlsx`) and the one in the
  report title (`EthSwitch / October 6,2026 ...`) are parsed separately; a
  file whose two disagree is reported rather than silently resolved, and the
  report title wins as the authoritative date.
- **Banks are unioned, not intersected.** An institution that appears in one
  day's file but not another still gets a column; days that do not report it
  are left blank and excluded from that bank's average rather than counted
  as a zero.
- **The rate is read, not re-derived.** Excel's own cached result for
  `Issu. SUCC. RATE (%)` is used as published. A report never opened in Excel
  has no cached result, so the short `=B40/B41` arithmetic behind the row is
  resolved from the counts underneath instead of being reported as missing.
"""

from __future__ import annotations

import io
import re
from dataclasses import dataclass, field
from datetime import date
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

__all__ = [
    "ATM_RATE_TIERS",
    "DailyReport",
    "ParsedDay",
    "build_atm_average_excel",
    "build_average_frame",
    "duplicate_days",
    "rate_fill",
    "read_daily_report",
    "read_daily_reports",
    "report_filename",
]

# --------------------------------------------------------------------------
# Success rate colour tiers (the cfd.xlsx palette)
# --------------------------------------------------------------------------

#: ``(label, lower bound, fill colour, font colour)``, highest band first.
#: A rate is placed in the first band whose lower bound it reaches, so
#: 98.5% is green and 98.49% is yellow. The colours carry the leading ``FF``
#: alpha channel, the way ``cfd.xlsx`` stores them.
ATM_RATE_TIERS: tuple[tuple[str, float, str, str], ...] = (
    ("98.5% - 100%", 0.985, "FF00B050", "FF000000"),
    ("90% - 98.4%", 0.90, "FFFFFF00", "FF000000"),
    ("80% - 89%", 0.80, "FFFFC000", "FF000000"),
    ("<= 79%", 0.0, "FFFF0000", "FFFFFFFF"),
)

def _round_rate(rate: float | None, digits: int = 5) -> float | None:
    """Round a rate to the precision the workbook stores.

    Rounding before the average is taken keeps the written ``=AVERAGE()`` and
    the fill colour chosen for the Average row in step with the daily cells
    they are computed from.
    """
    if rate is None:
        return None
    return round(float(rate), digits)


def _tier_for(rate: float) -> tuple[str, float, str, str]:
    """Return the tier a success rate falls in."""
    for tier in ATM_RATE_TIERS:
        if rate >= tier[1]:
            return tier
    return ATM_RATE_TIERS[-1]


def rate_fill(rate: float | None) -> PatternFill | None:
    """Return the tier fill for a success rate, or None when there is none.

    A missing rate (a bank a given day did not report) stays unfilled
    rather than being drawn as a zero, which would read as a total failure.
    """
    if rate is None:
        return None
    colour = _tier_for(rate)[2]
    return PatternFill(start_color=colour, end_color=colour, fill_type="solid")


def rate_tier_label(rate: float | None) -> str:
    """Return the human-readable band a rate falls in."""
    if rate is None:
        return ""
    return _tier_for(rate)[0]


# --------------------------------------------------------------------------
# Locating the pieces of a daily report
# --------------------------------------------------------------------------

_HEADER_LABELS = {"rc/bank name", "rc/bankname", "rc/bank", "bank name"}
_RATE_LABELS = ("issu. succ. rate", "issuing success rate", "success rate",
                "succ. rate")

_DATE_RE = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*,\s*(\d{4})")
_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _parse_date_text(text: str | None) -> date | None:
    """Pull a date out of free text such as ``October 6,2026``.

    Returns None when the text holds no readable date, so a caller can tell
    "no date" apart from "wrong date".
    """
    if not text:
        return None
    match = _DATE_RE.search(str(text))
    if not match:
        return None
    month_name, day, year = match.groups()
    month = _MONTHS.get(month_name[:3].lower())
    if not month:
        return None
    try:
        return date(int(year), month, int(day))
    except ValueError:
        return None


def _find_header_row(ws) -> int | None:
    """Return the 1-based row whose first cell is the ``RC/BANK NAME`` label."""
    for row in range(1, min(ws.max_row, 40) + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and value.strip().lower() in _HEADER_LABELS:
            return row
    return None


def _find_rate_row(ws, first_data_row: int) -> int | None:
    """Return the row labelled with the success rate, e.g. ``Issu. SUCC. RATE (%)``."""
    for row in range(first_data_row, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if not isinstance(value, str):
            continue
        text = value.strip().lower()
        if any(text.startswith(label) or label in text for label in _RATE_LABELS):
            return row
    return None


def _to_float(value: Any) -> float | None:
    """Coerce a cell to float, treating blanks and text as missing."""
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def _resolve_formula(ws_values, ws_formulas, column: int, row: int,
                     depth: int = 0) -> float | None:
    """Read a cell, preferring Excel's own cached result.

    A daily report expresses the success rate as ``=B42``, which is itself
    ``=B40/B41``, built from ``=B37+B38`` over ``=B39+B40`` -- every operand a
    cell reference or a literal. When the report has been opened in Excel its
    cached value is used as-is; when it has not, openpyxl hands back only the
    formula string, so the arithmetic is resolved here instead of reporting
    the rate as missing. ``depth`` stops the reference cycles a hand-edited
    report can contain.
    """
    formula = ws_formulas.cell(row=row, column=column).value
    if not isinstance(formula, str) or not formula.startswith("="):
        cached = ws_values.cell(row=row, column=column).value
        return _to_float(cached if cached is not None else formula)
    if depth > 20:
        return None

    expression = formula[1:].strip()

    def _sum_range(m: re.Match) -> str:
        letter, start, end = m.group(1), int(m.group(2)), int(m.group(4))
        if end < start:
            return "0.0"
        column_index = openpyxl.utils.cell.column_index_from_string(letter)
        total = 0.0
        for row_index in range(start, end + 1):
            part = _resolve_formula(ws_values, ws_formulas, column_index,
                                    row_index, depth + 1)
            if part is not None:
                total += part
        return repr(total)

    # Expand SUM(<col><first>:<col><last>) before the plain cell references.
    expression = re.sub(
        r"SUM\(\s*([A-Z]{1,3})(\d+)\s*:\s*([A-Z]{1,3})(\d+)\s*\)",
        lambda m: _sum_range(m) if m.group(1) == m.group(3) else "0.0",
        expression,
    )

    def _cell_ref(m: re.Match) -> str:
        column_index = openpyxl.utils.cell.column_index_from_string(m.group(1))
        part = _resolve_formula(ws_values, ws_formulas, column_index,
                                int(m.group(2)), depth + 1)
        return repr(part if part is not None else 0.0)

    expression = re.sub(r"\b([A-Z]{1,3})(\d+)\b", _cell_ref, expression)
    if not re.fullmatch(r"[\d\s+\-*/().%]*", expression):
        return None
    try:
        return float(eval(expression, {"__builtins__": {}}, {}))  # noqa: S307
    except (ZeroDivisionError, SyntaxError, ValueError, TypeError, NameError):
        return None


# --------------------------------------------------------------------------
# Reading one day
# --------------------------------------------------------------------------


@dataclass
class ParsedDay:
    """One day's success rate per bank."""

    day: date
    banks: list[str]
    rates: dict[str, float]
    source: str = ""

    def rate_for(self, bank: str) -> float | None:
        return self.rates.get(bank)


@dataclass
class DailyReport:
    """A parsed daily ATM report and anything worth telling the user about it."""

    parsed: ParsedDay | None = None
    source: str = ""
    date_from_name: date | None = None
    date_in_report: date | None = None
    date_matches: bool | None = None
    error: str | None = None
    notes: list[str] = field(default_factory=list)


def read_daily_report(data: bytes, source: str = "") -> DailyReport:
    """Read one daily ATM report workbook.

    Both the date in the file name and the one in the report title are
    parsed and compared; ``date_matches`` is None when either could not be
    read, which is different from the two disagreeing.
    """
    report = DailyReport(source=source, date_from_name=_parse_date_text(source))

    # Loaded twice: `values` carries the results Excel cached, `formulas` the
    # formulas themselves. A report that has never been opened in Excel has
    # no cached results, so the formulas are the only way to read the rate.
    try:
        values_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        formulas_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    except Exception as exc:  # noqa: BLE001 - reported per file, never fatal
        report.error = f"could not be opened ({exc})"
        return report

    try:
        ws = None
        ws_formulas = None
        header_row = None
        for sheet, sheet_formulas in zip(values_wb.worksheets,
                                         formulas_wb.worksheets):
            found = _find_header_row(sheet)
            if found:
                ws, ws_formulas, header_row = sheet, sheet_formulas, found
                break
        if ws is None or header_row is None:
            report.error = "no 'RC/BANK NAME' header row was found"
            return report

        # The title carries the day the report is about.
        for row in range(1, header_row):
            value = ws.cell(row=row, column=1).value
            if isinstance(value, str) and "decline" in value.lower():
                report.date_in_report = _parse_date_text(value)
                break
        if report.date_in_report is None:
            report.date_in_report = _parse_date_text(
                ws.cell(row=1, column=1).value)

        rate_row = _find_rate_row(ws, header_row + 1)
        if rate_row is None:
            report.error = "no 'Issu. SUCC. RATE (%)' row was found"
            return report

        # Banks are every headed column between the label and `Total`.
        banks: list[str] = []
        for column in range(2, ws.max_column + 1):
            name = ws.cell(row=header_row, column=column).value
            if not isinstance(name, str) or not name.strip():
                continue
            name = name.strip()
            if name.lower() in ("total", "totals", "grand total"):
                continue
            # A header repeated down the columns would otherwise give the
            # bank two columns of the same name.
            if name in banks:
                continue
            banks.append(name)

        rates: dict[str, float] = {}
        for offset, bank in enumerate(banks):
            value = _resolve_formula(ws, ws_formulas, 2 + offset, rate_row)
            if value is not None:
                rates[bank] = value

        if not rates:
            report.error = "the success rate row held no readable values"
            return report

        if report.date_in_report is None:
            report.error = "no date could be read from the report"
            return report

        if report.date_from_name is None:
            report.notes.append(
                f"no date in the file name; using {report.date_in_report:%d %b %Y} "
                "from the report")
        elif report.date_from_name != report.date_in_report:
            report.notes.append(
                f"date difference: file name says {report.date_from_name:%d %b %Y} "
                f"but the report says {report.date_in_report:%d %b %Y}; "
                "the report date is used")
            report.date_matches = False
        else:
            report.date_matches = True

        report.parsed = ParsedDay(day=report.date_in_report, banks=banks,
                                  rates=rates, source=source)
        return report
    finally:
        values_wb.close()
        formulas_wb.close()


def read_daily_reports(payloads: list[tuple[str, bytes]]) -> list[DailyReport]:
    """Read several uploaded daily reports, keeping per-file diagnostics."""
    return [read_daily_report(data, name) for name, data in payloads]


# --------------------------------------------------------------------------
# Building the average table
# --------------------------------------------------------------------------


def _union_banks(days: list[ParsedDay]) -> list[str]:
    """Every bank seen in any day, in the order the reports list them.

    A bank that only appears in some of the uploaded files still gets a
    column; the days missing it stay blank and drop out of its average
    instead of dragging it down as if the bank had scored zero.
    """
    banks: list[str] = []
    seen: set[str] = set()
    for day in days:
        for bank in day.banks:
            if bank not in seen:
                seen.add(bank)
                banks.append(bank)
    return banks


def duplicate_days(days: list[ParsedDay]) -> dict[date, list[str]]:
    """Group the days that more than one uploaded file claims.

    Two files for the same date would each get a row and be counted twice in
    every average, quietly over-weighting that day, so the caller surfaces it
    instead.
    """
    by_day: dict[date, list[str]] = {}
    for day in days:
        by_day.setdefault(day.day, []).append(day.source or day.day.isoformat())
    return {day: sorted(sources) for day, sources in by_day.items() if len(sources) > 1}


def build_average_frame(days: list[ParsedDay]):
    """Return the average table: one row per day plus an ``Average`` row.

    The index holds the days in date order with ``Average`` last; the columns
    are the union of the banks across every day.
    """
    import pandas as pd

    ordered = sorted(days, key=lambda d: d.day)
    banks = _union_banks(ordered)

    index: list[Any] = [d.day for d in ordered]
    rows: list[list[float | None]] = []
    for day in ordered:
        # Rounded to the precision the workbook stores, so the Average row
        # here matches the =AVERAGE() written over those same cells.
        rows.append([_round_rate(day.rates.get(bank)) for bank in banks])

    if rows:
        averages = []
        for column in range(len(banks)):
            values = [row[column] for row in rows if row[column] is not None]
            averages.append(sum(values) / len(values) if values else None)
        index.append("Average")
        rows.append(averages)

    return pd.DataFrame(rows, index=index, columns=banks)


# --------------------------------------------------------------------------
# Writing the styled workbook
# --------------------------------------------------------------------------

_THIN = Side(style="thin")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)
_HDR_FONT = Font(name="Calibri", size=11, bold=True)
_CELL_FONT = Font(name="Calibri", size=11)
_DATE_FONT = Font(name="Calibri", size=11)


def build_atm_average_excel(days: list[ParsedDay], sheet_title: str = "ATM AVERAGE") -> bytes:
    """Build the average report workbook in the shape of ``cfd.xlsx``.

    Column A stays empty and the dates sit in column B, as in the reference;
    the header is row 2, one row per day follows, and the last row holds live
    ``=AVERAGE(...)`` formulas over that bank's daily rates.
    """
    import pandas as pd

    frame = build_average_frame(days)
    banks = list(frame.columns)
    date_rows = [idx for idx in frame.index if isinstance(idx, date)]
    if not date_rows:
        raise ValueError("no dated days to average")

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet_title

    first_bank_col = 3  # column C, leaving A and B for the layout of cfd.xlsx
    last_bank_col = first_bank_col + len(banks) - 1

    # Row 2: the bank names.
    for offset, bank in enumerate(banks):
        cell = ws.cell(row=2, column=first_bank_col + offset, value=bank)
        cell.font = _HDR_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _BORDER
    ws.row_dimensions[2].height = 16.8

    # One row per day, then the Average row.
    first_rate_row = 3
    average_row = first_rate_row + len(date_rows)
    for row_offset, idx in enumerate(frame.index):
        row = first_rate_row + row_offset
        is_average = idx == "Average"

        label = ws.cell(row=row, column=2,
                        value="Average" if is_average else pd.Timestamp(idx).to_pydatetime())
        label.number_format = "General" if is_average else "d-mmm"
        label.font = _DATE_FONT
        # No explicit alignment: Excel right-aligns the dates and left-aligns
        # the "Average" label on its own, as the reference does.
        label.border = _BORDER

        for offset, bank in enumerate(banks):
            value = frame.iat[row_offset, offset]
            # AVERAGE over a column of blanks is a #DIV/0! in Excel, so a
            # bank no day reported a rate for is left empty instead.
            column_has_rates = any(
                _row_rate(frame, r, offset) is not None
                for r in range(len(date_rows)))
            if is_average:
                letter = get_column_letter(first_bank_col + offset)
                cell = ws.cell(
                    row=row, column=first_bank_col + offset,
                    value=(f"=AVERAGE({letter}{first_rate_row}:"
                           f"{letter}{average_row - 1})" if column_has_rates
                           else None))
            elif value is None or (isinstance(value, float) and value != value):
                # A day that does not report this bank stays blank and is
                # left out of the average above rather than counted as zero.
                cell = ws.cell(row=row, column=first_bank_col + offset)
            else:
                cell = ws.cell(row=row, column=first_bank_col + offset,
                               value=float(value))

            cell.number_format = "0.00%"
            shown = _row_rate(frame, row_offset, offset)
            fill = rate_fill(shown)
            if fill is not None:
                cell.fill = fill
            _, _, _, font_colour = _tier_for(shown if shown is not None else 0.0)
            cell.font = Font(name="Calibri", size=11, color=font_colour)
            if not is_average:
                # The daily rates are aligned right; the Average row keeps the
                # reference's default alignment.
                cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = _BORDER

    # The bank names are all short enough for Excel's default column width,
    # which is what cfd.xlsx leaves them at.

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _row_rate(frame, row_offset: int, column: int) -> float | None:
    """The rate a cell displays, so its fill matches what the user sees.

    The Average row shows an Excel formula, so the value that decides its
    colour is the one that formula will evaluate to.
    """
    value = frame.iat[row_offset, column]
    if value is None:
        return None
    try:
        value = float(value)
    except (TypeError, ValueError):
        return None
    if value != value:  # NaN
        return None
    return value


def report_filename(days: list[ParsedDay]) -> str:
    """Name the download after the days it covers, e.g. ``ATM_Average_Success_Rate_06-Oct-2026_to_08-Oct-2026.xlsx``."""
    if not days:
        return "ATM_Average_Success_Rate.xlsx"
    ordered = sorted(d.day for d in days)
    first, last = ordered[0], ordered[-1]
    if first == last:
        return f"ATM_Average_Success_Rate_{first:%d-%b-%Y}.xlsx"
    return (f"ATM_Average_Success_Rate_{first:%d-%b-%Y}"
            f"_to_{last:%d-%b-%Y}.xlsx")