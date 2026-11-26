"""Daily POS & ATM success-rate comparison report.

The two EthSwitch daily summaries are not transaction lists: they are
``RC/BANK NAME`` matrices whose last meaningful row carries the success rate
each bank reached that day (``success rate`` in the POS workbook,
``Issu. SUCC. RATE (%)`` in the ATM workbook). This module reads those two
workbooks, checks that they describe the same day, and produces the compact
side-by-side sheet modelled on ``sample.xlsx``::

                    Daily POS and ATM Success rate Nov 16,2026
                 |      POS      |      |      ATM      |
                 | Abay     95%  |      | CBE     99.3%  |
                 | BOA      99%  |      | BOA     99.6%  |
                 | ...           |      | ...           |
                 | Total    98%  |      | Total   99.1%  |

The rates are published percentages, so they are read (Excel's cached result,
or the simple arithmetic behind the formula when there is none) rather than
recalculated from the decline counts. POS is stored to two decimals and shown
as ``0%``; ATM is stored to three decimals and shown as ``0.00%``, exactly as
in the reference workbook. Both use the same four colour tiers as the ATM
average report.
"""

from __future__ import annotations

import calendar
import html
import io
import re
from dataclasses import dataclass
from datetime import date
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Border, Color, Font, PatternFill, Side

__all__ = [
    "EXCLUDED_DECLINE_CODES",
    "GREEN_THRESHOLD",
    "below_green_declines",
    "DailySuccessReport",
    "build_preview_html",
    "build_success_rate_excel",
    "combine",
    "declines_text",
    "POS_RATE_TIERS",
    "parse_report_date",
    "pretty_date",
    "preview_height",
    "read_reports",
    "report_filename",
]

_HEADER_LABELS = {"rc/bank name", "rc/bankname", "rc/bank", "bank name"}
_TOTAL_LABELS = {"total", "totals", "grand total"}
_ATM_RATE_LABELS = ("issu. succ. rate", "issuing success rate", "issuer success rate")
_POS_RATE_LABELS = ("success rate", "succ. rate", "pos success rate")

_POS_DIGITS = 2
_ATM_DIGITS = 3

#: Anything below the top ATM band is "out of green".
GREEN_THRESHOLD = 0.985

#: Cardholder-related codes the ATM decline breakdown deliberately leaves out.
EXCLUDED_DECLINE_CODES = (821, 901, 904, 906, 911, 912, 914, 915)

#: The reference workbook bands the coarse, whole-percent POS rates a little
#: lower than the fine-grained ATM rates. ``(label, lower bound, fill)``.
POS_RATE_TIERS = (
    ("96% - 100%", 0.96, "FF00B050"),
    ("88% - 95%", 0.88, "FFFFFF00"),
    ("80% - 87%", 0.80, "FFFFC000"),
    ("< 80%", 0.0, "FFFF0000"),
)

#: The ATM bands, kept here so this module has no dependency on any other
#: report module (a stale deploy of one must not break this report).
_ATM_RATE_TIERS = (
    ("98.5% - 100%", 0.985, "FF00B050", "FF000000"),
    ("90% - 98.4%", 0.90, "FFFFFF00", "FF000000"),
    ("80% - 89%", 0.80, "FFFFC000", "FF000000"),
    ("<= 79%", 0.0, "FFFF0000", "FFFFFFFF"),
)

_DATE_RE = re.compile(r"([A-Za-z]{3,9})\.?\s+(\d{1,2})\s*,\s*(\d{4})")
_MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


def _parse_date_text(text: str | None) -> date | None:
    """Pull a date out of free text such as ``November 16,2026``."""
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


_EXCEL_FUNC_RE = re.compile(
    r"\b(IFERROR|AVERAGE|ROUND|SUM|MIN|MAX|IF)\s*\(",
    re.IGNORECASE,
)
_CELL_RANGE_RE = re.compile(
    r"^\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)$"
)
_CELL_REF_RE = re.compile(r"\$?\b([A-Z]{1,3})\$?(\d+)\b")


def _find_matching_paren(text: str, open_idx: int) -> int:
    """Return the index of the closing ``)`` matching ``text[open_idx]``."""
    depth = 0
    for idx in range(open_idx, len(text)):
        ch = text[idx]
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
            if depth == 0:
                return idx
    return -1


def _split_top_level_args(text: str) -> list[str]:
    """Split ``text`` on commas outside parentheses."""
    args: list[str] = []
    depth = 0
    start = 0
    for idx, ch in enumerate(text):
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        elif ch == "," and depth == 0:
            args.append(text[start:idx].strip())
            start = idx + 1
    tail = text[start:].strip()
    if tail or args:
        args.append(tail)
    return args


def _eval_excel_condition(cond: str, ws_values, ws_formulas,
                          depth: int) -> bool | None:
    """Evaluate a simple comparison such as ``B41=0`` or ``B41>0``."""
    for op in ("<>", "!=", "<=", ">=", "=", "<", ">"):
        # Split on the first top-level occurrence of op.
        p_depth = 0
        for idx in range(len(cond) - len(op) + 1):
            ch = cond[idx]
            if ch == "(":
                p_depth += 1
            elif ch == ")":
                p_depth -= 1
            elif p_depth == 0 and cond[idx:idx + len(op)] == op:
                lhs = _eval_excel_expr(cond[:idx], ws_values, ws_formulas, depth)
                rhs = _eval_excel_expr(cond[idx + len(op):], ws_values,
                                       ws_formulas, depth)
                if lhs is None or rhs is None:
                    return None
                if op == "=":
                    return lhs == rhs
                if op in ("<>", "!="):
                    return lhs != rhs
                if op == "<=":
                    return lhs <= rhs
                if op == ">=":
                    return lhs >= rhs
                if op == "<":
                    return lhs < rhs
                if op == ">":
                    return lhs > rhs
    val = _eval_excel_expr(cond, ws_values, ws_formulas, depth)
    return bool(val) if val is not None else None


def _eval_excel_func(name: str, arg_str: str, ws_values, ws_formulas,
                     depth: int) -> float | None:
    func = name.upper()
    args = _split_top_level_args(arg_str)

    if func == "IF":
        if len(args) < 2:
            return None
        cond_val = _eval_excel_condition(args[0], ws_values, ws_formulas, depth)
        if cond_val is None:
            return None
        branch = args[1] if cond_val else (args[2] if len(args) >= 3 else "0")
        return _eval_excel_expr(branch, ws_values, ws_formulas, depth)

    if func == "IFERROR":
        if len(args) < 2:
            return None
        val = _eval_excel_expr(args[0], ws_values, ws_formulas, depth)
        if val is not None:
            return val
        return _eval_excel_expr(args[1], ws_values, ws_formulas, depth)

    if func in ("SUM", "AVERAGE", "MIN", "MAX"):
        values: list[float] = []
        for arg in args:
            if not arg:
                continue
            range_m = _CELL_RANGE_RE.match(arg.strip())
            if range_m:
                first_col = openpyxl.utils.cell.column_index_from_string(
                    range_m.group(1))
                last_col = openpyxl.utils.cell.column_index_from_string(
                    range_m.group(3))
                start, end = int(range_m.group(2)), int(range_m.group(4))
                for col_index in range(first_col, last_col + 1):
                    for row_index in range(start, end + 1):
                        if (ws_formulas.cell(row_index, col_index).value is None
                                and ws_values.cell(row_index, col_index).value is None):
                            continue
                        part = _resolve_formula(ws_values, ws_formulas,
                                                col_index, row_index, depth + 1)
                        if part is not None:
                            values.append(part)
            else:
                val = _eval_excel_expr(arg, ws_values, ws_formulas, depth)
                if val is not None:
                    values.append(val)
        if func == "SUM":
            return float(sum(values))
        if not values:
            return None
        if func == "AVERAGE":
            return float(sum(values) / len(values))
        return float(min(values) if func == "MIN" else max(values))

    if func == "ROUND":
        if len(args) != 2:
            return None
        val = _eval_excel_expr(args[0], ws_values, ws_formulas, depth)
        digits = _eval_excel_expr(args[1], ws_values, ws_formulas, depth)
        if val is None or digits is None:
            return None
        return float(round(val, int(digits)))

    return None


def _eval_excel_expr(expression: str, ws_values, ws_formulas,
                     depth: int) -> float | None:
    expr = expression.strip()
    if not expr:
        return None

    while True:
        match = _EXCEL_FUNC_RE.search(expr)
        if not match:
            break
        open_idx = match.end() - 1
        close_idx = _find_matching_paren(expr, open_idx)
        if close_idx == -1:
            return None
        func_val = _eval_excel_func(
            match.group(1),
            expr[open_idx + 1:close_idx],
            ws_values,
            ws_formulas,
            depth,
        )
        if func_val is None:
            return None
        expr = f"{expr[:match.start()]}{repr(func_val)}{expr[close_idx + 1:]}"

    missing_ref = False

    def _cell_ref(match: re.Match) -> str:
        nonlocal missing_ref
        col_index = openpyxl.utils.cell.column_index_from_string(match.group(1))
        row_index = int(match.group(2))
        if (ws_formulas.cell(row_index, col_index).value is None
                and ws_values.cell(row_index, col_index).value is None):
            return "0.0"
        part = _resolve_formula(ws_values, ws_formulas, col_index,
                                row_index, depth + 1)
        if part is None:
            missing_ref = True
            return "0.0"
        return repr(part)

    expr = _CELL_REF_RE.sub(_cell_ref, expr)
    if missing_ref:
        return None
    expr = re.sub(r"(\d+(?:\.\d+)?)\s*%", r"(\1/100.0)", expr)
    if not re.fullmatch(r"[\d\s+\-*/().eE]*", expr):
        return None
    try:
        return float(eval(expr, {"__builtins__": {}}, {}))  # noqa: S307
    except (ZeroDivisionError, SyntaxError, ValueError, TypeError, NameError):
        return None


def _resolve_formula(ws_values, ws_formulas, column: int, row: int,
                     depth: int = 0) -> float | None:
    """Read a rate cell, preferring Excel's cached result.

    The rate rows are formulas such as ``=B45/B46``, ``=B42``, or
    ``=IF(B41=0,1,B40/B41)``. When the workbook has been opened in Excel the
    cached value is used; otherwise the formula is resolved from the cells it
    points at. ``depth`` guards against cycles in hand-edited reports.
    """
    cached = ws_values.cell(row=row, column=column).value
    cached_num = _to_float(cached)
    if cached_num is not None:
        return cached_num

    formula = ws_formulas.cell(row=row, column=column).value
    if not isinstance(formula, str) or not formula.startswith("="):
        return _to_float(formula)
    if depth > 20:
        return None

    return _eval_excel_expr(formula[1:], ws_values, ws_formulas, depth)


def _tiers_for(kind: str) -> tuple:
    return POS_RATE_TIERS if kind == "POS" else _ATM_RATE_TIERS


def _fill_for(rate: float, kind: str) -> PatternFill:
    for tier in _tiers_for(kind):
        if rate >= tier[1]:
            colour = tier[2]
            return PatternFill(start_color=colour, end_color=colour,
                               fill_type="solid")
    colour = _tiers_for(kind)[-1][2]
    return PatternFill(start_color=colour, end_color=colour, fill_type="solid")


@dataclass
class SourceReport:
    """One uploaded workbook, with the pieces needed for the comparison."""

    source: str
    kind: str
    day: date | None
    title: str
    banks: list[str]
    rates: dict[str, float]
    total: float | None
    error: str | None = None
    notes: list[str] | None = None
    declines: dict[str, dict[str, int]] | None = None
    rc_descriptions: dict[str, str] | None = None

    def __post_init__(self):
        if self.notes is None:
            self.notes = []


@dataclass
class DailySuccessReport:
    """The two validated rate lists, ready to write."""

    day: date
    pos: list[tuple[str, float]]
    atm: list[tuple[str, float]]
    sources: list[str]
    atm_declines: dict[str, dict[str, int]] | None = None
    rc_descriptions: dict[str, str] | None = None

    @property
    def pos_total(self) -> float:
        return self.pos[-1][1]

    @property
    def atm_total(self) -> float:
        return self.atm[-1][1]

    @property
    def title(self) -> str:
        return f"Daily POS and ATM Success rate {pretty_date(self.day)}"


# --------------------------------------------------------------------------
# Dates
# --------------------------------------------------------------------------


def parse_report_date(text: str | None) -> date | None:
    """Parse a date out of a report title or file name."""
    return _parse_date_text(text)


def pretty_date(day: date) -> str:
    """The date form used in the reference title, e.g. ``Nov 16,2026``."""
    return f"{calendar.month_abbr[day.month]} {day.day},{day.year}"


def _to_float(value: Any) -> float | None:
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", ""))
    except (TypeError, ValueError):
        return None


def _rc_key(value: Any) -> str | None:
    """Normalise a response code (``503``, ``'503'``, ``503.0``) to ``'503'``."""
    number = _to_float(value)
    if number is None:
        return None
    return str(int(number)) if number == int(number) else str(number)


def _read_atm_declines(ws, header_row: int, rate_row: int,
                       banks: list[str], columns: list[int]) -> dict[str, dict[str, int]]:
    """Decline counts per bank and response code, less the excluded codes."""
    excluded = {str(code) for code in EXCLUDED_DECLINE_CODES}
    declines: dict[str, dict[str, int]] = {bank: {} for bank in banks}
    for row in range(header_row + 1, rate_row):
        code = _rc_key(ws.cell(row=row, column=1).value)
        if code is None or code in excluded:
            continue
        for bank, column in zip(banks, columns):
            count = _to_float(ws.cell(row=row, column=column).value)
            if count and count > 0:
                declines[bank][code] = int(count)
    return declines


def _read_rc_descriptions(ws, rate_row: int) -> dict[str, str]:
    """RC -> short description from the legend below the rate row."""
    descriptions: dict[str, str] = {}
    for row in range(rate_row + 1, ws.max_row + 1):
        code = _rc_key(ws.cell(row=row, column=1).value)
        if code is None:
            continue
        short = ws.cell(row=row, column=2).value
        remark = ws.cell(row=row, column=5).value
        text = str(short).strip() if short is not None else ""
        if not text:
            text = str(remark).strip() if remark is not None else ""
        if text:
            descriptions[code] = text
    return descriptions


# --------------------------------------------------------------------------
# Reading one source workbook
# --------------------------------------------------------------------------


def _find_header_row(ws) -> int | None:
    for row in range(1, min(ws.max_row, 40) + 1):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and value.strip().lower() in _HEADER_LABELS:
            return row
    return None


def _find_rate_row(ws, first_data_row: int) -> tuple[int, str] | None:
    """Return the rate row and its kind (``POS`` or ``ATM``)."""
    rows: list[tuple[int, str]] = []
    for row in range(first_data_row, ws.max_row + 1):
        value = ws.cell(row=row, column=1).value
        if not isinstance(value, str):
            continue
        text = value.strip().lower()
        if any(label in text for label in _ATM_RATE_LABELS):
            rows.append((row, "ATM"))
        elif any(label in text for label in _POS_RATE_LABELS):
            rows.append((row, "POS"))
    if not rows:
        return None
    # An ATM workbook labels its rate row with the issuer wording, which is
    # more specific than the POS workbook's plain "success rate".
    for row, kind in rows:
        if kind == "ATM":
            return row, kind
    return rows[0]


def _find_title(ws, header_row: int) -> str:
    for row in range(1, header_row):
        value = ws.cell(row=row, column=1).value
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ""


def _read_one(data: bytes, source: str = "") -> SourceReport:
    name = source or "uploaded workbook"
    try:
        values_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        formulas_wb = openpyxl.load_workbook(io.BytesIO(data), data_only=False)
    except Exception as exc:  # noqa: BLE001 - reported per file, never fatal
        return SourceReport(name, "", None, "", [], {}, None,
                            error=f"could not be opened ({exc})")

    try:
        ws = ws_formulas = None
        header_row = None
        for sheet, sheet_formulas in zip(values_wb.worksheets,
                                         formulas_wb.worksheets):
            found = _find_header_row(sheet)
            if found:
                ws, ws_formulas, header_row = sheet, sheet_formulas, found
                break
        if ws is None or header_row is None:
            return SourceReport(name, "", None, "", [], {}, None,
                                error="no 'RC/BANK NAME' header row was found")

        title = _find_title(ws, header_row)
        rate_row_info = _find_rate_row(ws, header_row + 1)
        if rate_row_info is None:
            return SourceReport(
                name, "", parse_report_date(title), title, [], {}, None,
                error="no success-rate row was found")
        rate_row, kind = rate_row_info

        banks: list[str] = []
        columns: list[int] = []
        total_column: int | None = None
        for column in range(2, ws.max_column + 1):
            value = ws.cell(row=header_row, column=column).value
            if not isinstance(value, str) or not value.strip():
                continue
            label = value.strip()
            if label.lower() in _TOTAL_LABELS:
                total_column = column
                continue
            if label in banks:
                continue
            banks.append(label)
            columns.append(column)
        if not banks:
            return SourceReport(name, kind, parse_report_date(title), title,
                                [], {}, None,
                                error="the header row listed no banks")

        rates: dict[str, float] = {}
        for bank, column in zip(banks, columns):
            value = _to_float(_resolve_formula(ws, ws_formulas, column,
                                               rate_row))
            if value is not None:
                rates[bank] = value
        if not rates:
            return SourceReport(name, kind, parse_report_date(title), title,
                                banks, {}, None,
                                error="the success-rate row held no readable values")

        total_value = None
        if total_column is not None:
            total_value = _to_float(_resolve_formula(
                ws, ws_formulas, total_column, rate_row))
        if total_value is None:
            total_value = sum(rates.values()) / len(rates)

        declines = None
        rc_descriptions = None
        if kind == "ATM":
            declines = _read_atm_declines(ws, header_row, rate_row,
                                          banks, columns)
            rc_descriptions = _read_rc_descriptions(ws, rate_row)

        title_date = parse_report_date(title)
        name_date = parse_report_date(source)
        day = title_date or name_date
        notes: list[str] = []
        if title_date and name_date and title_date != name_date:
            notes.append(
                f"date difference: file name says {pretty_date(name_date)} "
                f"but the report title says {pretty_date(title_date)}; "
                "the report title is used")
        if day is None:
            return SourceReport(name, kind, None, title, banks, rates,
                                total_value,
                                error="no date could be read from the report")

        return SourceReport(name, kind, day, title, banks, rates, total_value,
                            notes=notes, declines=declines,
                            rc_descriptions=rc_descriptions)
    finally:
        values_wb.close()
        formulas_wb.close()


def read_reports(payloads: list[tuple[str, bytes]]) -> list[SourceReport]:
    """Read the uploaded workbooks, keeping per-file diagnostics."""
    return [_read_one(data, name) for name, data in payloads]


# --------------------------------------------------------------------------
# Validating and combining the two reports
# --------------------------------------------------------------------------


def _rate_list(report: SourceReport, digits: int) -> list[tuple[str, float]]:
    rows = [(bank, round(report.rates[bank], digits))
            for bank in report.banks if bank in report.rates]
    rows.append(("Total", round(report.total or 0.0, digits)))
    return rows


def combine(reports: list[SourceReport]) -> DailySuccessReport:
    """Validate the two uploads and build the side-by-side report.

    ``ValueError`` is raised with a ``date difference:`` message when the two
    workbooks describe different days, and with a descriptive message for any
    other problem (unreadable file, wrong report types, missing date).
    """
    for report in reports:
        if report.error:
            raise ValueError(f"{report.source}: {report.error}")
        for note in report.notes or []:
            if note.startswith("date difference:"):
                raise ValueError(f"{report.source}: {note}")

    by_kind: dict[str, SourceReport] = {}
    for report in reports:
        if report.kind in by_kind:
            raise ValueError(
                f"expected one POS report and one ATM report, but two "
                f"{report.kind} reports were uploaded "
                f"({by_kind[report.kind].source} and {report.source})")
        by_kind[report.kind] = report

    missing = {"POS", "ATM"} - set(by_kind)
    if missing:
        raise ValueError(
            "expected one POS report and one ATM report; missing "
            + " and ".join(sorted(missing)))

    pos, atm = by_kind["POS"], by_kind["ATM"]
    if pos.day != atm.day:
        raise ValueError(
            f"date difference: the POS report is dated {pretty_date(pos.day)} "
            f"({pos.source}) but the ATM report is dated "
            f"{pretty_date(atm.day)} ({atm.source})")

    return DailySuccessReport(
        day=pos.day,
        pos=_rate_list(pos, _POS_DIGITS),
        atm=_rate_list(atm, _ATM_DIGITS),
        sources=[pos.source, atm.source],
        atm_declines=atm.declines or {},
        rc_descriptions=atm.rc_descriptions or {},
    )


def below_green_declines(report: DailySuccessReport) -> list[tuple]:
    """ATM banks below 98.5% with their remaining decline codes and counts.

    Each item is ``(bank, rate, [(code, count, description), ...])``. Codes in
    ``EXCLUDED_DECLINE_CODES`` were already dropped when the ATM workbook was
    read, and banks already in the green band are skipped entirely.
    """
    declines = report.atm_declines or {}
    descriptions = report.rc_descriptions or {}
    out: list[tuple] = []
    for bank, rate in report.atm:
        if bank == "Total" or rate >= GREEN_THRESHOLD:
            continue
        codes = declines.get(bank, {})
        items = [(code, count, descriptions.get(code, ""))
                 for code, count in sorted(codes.items(),
                                           key=lambda pair: int(pair[0]))]
        out.append((bank, rate, items))
    return out


def excluded_codes_text() -> str:
    return ", ".join(str(code) for code in EXCLUDED_DECLINE_CODES)


def _join_codes(codes: list[str]) -> str:
    if len(codes) == 1:
        return codes[0]
    if len(codes) == 2:
        return f"{codes[0]} and {codes[1]}"
    return ", ".join(codes[:-1]) + " and " + codes[-1]


def declines_text(report: DailySuccessReport,
                  issues: dict[str, str] | None = None) -> str:
    """One short prose line per bank, ready to copy into a report.

    Banks with a user-written specific issue show that sentence verbatim.
    Every other bank is described by the response codes left after the
    cardholder-related exclusions, e.g. ``Bunna: They had 905 and 917
    response code declines.``
    """
    issues = issues or {}
    declines = report.atm_declines or {}
    below = below_green_declines(report)

    banks = [bank for bank, _, _ in below]
    for bank in issues:
        if bank not in banks:
            banks.append(bank)

    lines: list[str] = []
    for bank in banks:
        note = str(issues.get(bank, "") or "").strip()
        if note:
            if not note.endswith("."):
                note += "."
            lines.append(f"{bank}: {note}")
            continue
        codes = sorted(declines.get(bank, {}), key=int)
        if codes:
            lines.append(f"{bank}: They had {_join_codes(codes)} "
                         "response code declines.")
        else:
            lines.append(f"{bank}: They had no counted declines.")
    return "\n".join(lines)


def preview_height(report: DailySuccessReport) -> int:
    """Height for the Streamlit iframe holding only the rate table."""
    return 130 + max(len(report.pos), len(report.atm)) * 28


# --------------------------------------------------------------------------
# Writing the styled workbook (the shape of sample.xlsx)
# --------------------------------------------------------------------------

_TITLE_FONT = Font(name="Calibri", size=11, bold=True, color=Color(theme=1))
_TITLE_FILL = PatternFill(fill_type="solid", fgColor=Color(theme=5))
_HEADER_FONT = Font(name="Calibri", size=11, bold=True, color=Color(theme=1))
_LABEL_FONT = Font(name="Calibri", size=11, bold=True, color="FF000000")
_POS_FONT = Font(name="Calibri", size=11, bold=True, color="FF000000")
_ATM_FONT = Font(name="Arial", size=10, bold=True, color="FF000000")
_THIN = Side(style="thin")


def _border(left=False, right=False, top=False, bottom=False) -> Border:
    return Border(left=_THIN if left else None,
                  right=_THIN if right else None,
                  top=_THIN if top else None,
                  bottom=_THIN if bottom else None)


def build_success_rate_excel(report: DailySuccessReport) -> bytes:
    """Build the two-column comparison workbook modelled on ``sample.xlsx``.

    Columns A and B stay empty, the title sits in ``C3:G3``, bank names and
    rates run down C/D (POS) and F/G (ATM), and column E is a blank divider.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    pos_last = 4 + len(report.pos)
    atm_last = 4 + len(report.atm)
    last_row = max(pos_last, atm_last)

    ws.merge_cells("C3:G3")
    title_cell = ws["C3"]
    title_cell.value = report.title
    title_cell.font = _TITLE_FONT
    title_cell.fill = _TITLE_FILL
    title_cell.alignment = Alignment(horizontal="center")
    for column in range(3, 8):
        ws.cell(row=3, column=column).border = _border(
            left=column == 3, right=column == 7, top=True, bottom=True)

    ws.merge_cells("C4:D4")
    ws.merge_cells("F4:G4")
    ws.merge_cells(start_row=4, start_column=5, end_row=last_row, end_column=5)
    for coordinate, text in (("C4", "POS "), ("F4", "ATM ")):
        cell = ws[coordinate]
        cell.value = text
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for column in range(3, 8):
        if column == 5:
            divider = ws.cell(row=4, column=column)
            divider.border = _border(top=True, bottom=True)
            divider.fill = PatternFill(fill_type="solid", fgColor=Color(theme=0))
        else:
            ws.cell(row=4, column=column).border = _border(
                left=column in (3, 6), right=column in (4, 7), top=True)

    for offset, (bank, rate) in enumerate(report.pos):
        row = 5 + offset
        name = ws.cell(row=row, column=3, value=bank)
        name.font = _LABEL_FONT
        name.alignment = Alignment(horizontal="left", vertical="center",
                                   wrap_text=True)
        name.border = _border(left=True, right=True, top=True, bottom=True)
        _write_rate(ws.cell(row=row, column=4), rate, "0%", "POS")

    for offset, (bank, rate) in enumerate(report.atm):
        row = 5 + offset
        name = ws.cell(row=row, column=6, value=bank)
        name.font = _LABEL_FONT
        name.alignment = Alignment(horizontal="left", vertical="center",
                                   wrap_text=True)
        name.border = _border(left=True, right=True, top=True, bottom=True)
        _write_rate(ws.cell(row=row, column=7), rate, "0.00%", "ATM")

    # The shorter side ends in an empty bordered block, as in the reference.
    if pos_last < atm_last:
        first_col, short_last = 3, pos_last
    elif atm_last < pos_last:
        first_col, short_last = 6, atm_last
    else:
        first_col, short_last = 0, 0
    if first_col:
        ws.merge_cells(start_row=short_last + 1, start_column=first_col,
                       end_row=last_row, end_column=first_col + 1)
        anchor = ws.cell(row=short_last + 1, column=first_col)
        anchor.border = _border(left=True, bottom=True)
        anchor.font = Font(name="Calibri", size=11, bold=True)
        anchor.fill = PatternFill(fill_type="solid", fgColor=Color(theme=0))
    ws.cell(row=last_row, column=5).border = _border(bottom=True)

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output.getvalue()


def _write_rate(cell, rate: float, number_format: str, kind: str) -> None:
    cell.value = float(rate)
    cell.number_format = number_format
    cell.font = _POS_FONT if kind == "POS" else _ATM_FONT
    cell.alignment = Alignment(horizontal="left", wrap_text=True)
    cell.fill = _fill_for(rate, kind)
    cell.border = _border(left=True, right=True, top=True, bottom=True)


# --------------------------------------------------------------------------
# HTML preview with PNG export
# --------------------------------------------------------------------------


def _html_colour(rate: float, kind: str) -> str:
    tiers = _tiers_for(kind)
    for tier in tiers:
        if rate >= tier[1]:
            return "#" + tier[2][2:]
    return "#" + tiers[-1][2][2:]


def _rate_cell(rate: float, kind: str) -> str:
    text = f"{rate * 100:.0f}%" if kind == "POS" else f"{rate * 100:.2f}%"
    font = "posrate" if kind == "POS" else "atmrate"
    return (f"<td class='rate {font}' style='background:{_html_colour(rate, kind)}'>"
            f"{text}</td>")


def _side_cell(entry: tuple[str, float | None], kind: str, last: bool) -> str:
    """One two-column side of a row, or the empty block when a side ended."""
    if entry[1] is None:
        return ("<td class='emptybox last' colspan='2'></td>" if last
                else "<td class='emptybox' colspan='2'></td>")
    return (f"<td class='bank'>{html.escape(str(entry[0]))}</td>"
            f"{_rate_cell(entry[1], kind)}")


def build_preview_html(report: DailySuccessReport) -> str:
    """Return the sample-styled HTML table with a PNG download button."""
    rows: list[str] = []
    row_count = max(len(report.pos), len(report.atm))
    for index in range(row_count):
        pos = report.pos[index] if index < len(report.pos) else ("", None)
        atm = report.atm[index] if index < len(report.atm) else ("", None)
        last = index == row_count - 1
        divider = "<td class='divider last'></td>" if last else "<td class='divider'></td>"
        left = _side_cell(pos, "POS", last)
        right = _side_cell(atm, "ATM", last)
        rows.append(f"<tr>{left}{divider}{right}</tr>")

    png_name = report_filename(report).replace(".xlsx", ".png")
    return (
        "<!DOCTYPE html><html><head><meta charset='utf-8'/><style>"
        "body{margin:0;font-family:Arial,sans-serif;color:#000;background:#fff;}"
        ".toolbar{padding:8px;text-align:right;}"
        ".toolbar button{background:#2563eb;color:#fff;border:0;border-radius:5px;"
        "padding:7px 14px;font-size:13px;font-weight:700;cursor:pointer;}"
        "#report-content{display:inline-block;padding:0;margin:0;background:#fff;}"
        "table{border-collapse:collapse;margin:0;}"
        "td{padding:3px 8px;height:20px;}"
        ".title{border:1px solid #000;background:#ED7D31;color:#000;text-align:center;"
        "font-family:Calibri,sans-serif;font-size:11px;font-weight:700;}"
        ".head{border:1px solid #000;border-bottom:0;text-align:center;"
        "font-family:Calibri,sans-serif;font-size:11px;font-weight:700;}"
        ".bank{border:1px solid #000;font-family:Calibri,sans-serif;font-size:11px;"
        "font-weight:700;text-align:left;min-width:86px;}"
        ".rate{border:1px solid #000;text-align:left;min-width:56px;}"
        ".posrate{font-family:Calibri,sans-serif;font-size:11px;font-weight:700;}"
        ".atmrate{font-family:Arial,sans-serif;font-size:10px;font-weight:700;}"
        ".divider{border-left:0;border-right:0;padding:0 6px;width:18px;}"
        ".divider.last{border-bottom:1px solid #000;}"
        ".head-divider{border-top:1px solid #000;border-bottom:1px solid #000;padding:0;}"
        ".emptybox{border-left:1px solid #000;}"
        ".emptybox.last{border-bottom:1px solid #000;}"
        "</style>"
        "<script src='https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'></script>"
        "</head><body>"
        "<div class='toolbar'><button onclick='capturePNG()'>Download PNG</button></div>"
        "<div id='report-content'>"
        "<table>"
        f"<tr><td class='title' colspan='5'>{html.escape(report.title)}</td></tr>"
        "<tr>"
        "<td class='head' colspan='2'>POS </td>"
        "<td class='divider head-divider'></td>"
        "<td class='head' colspan='2'>ATM </td>"
        "</tr>"
        + "".join(rows) +
        "</table></div>"
        "<script>"
        "function capturePNG(){var e=document.getElementById('report-content');"
        "var w=e.scrollWidth,h=e.scrollHeight;"
        "html2canvas(e,{scale:2,useCORS:true,backgroundColor:'#ffffff',"
        "width:w,height:h,scrollX:0,scrollY:0,windowWidth:w,windowHeight:h})"
        ".then(function(c){var a=document.createElement('a');"
        f"a.download='{html.escape(png_name)}';a.href=c.toDataURL('image/png');a.click();}})"
        ".catch(function(err){"
        "alert('PNG export needs internet (html2canvas loads from CDN): '+err.message);});}"
        "</script></body></html>"
    )


def report_filename(report: DailySuccessReport) -> str:
    return f"Daily_POS_and_ATM_Success_Rate_{report.day:%d-%b-%Y}.xlsx"
