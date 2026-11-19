"""Unit tests for the daily POS & ATM success-rate report."""

import calendar
import io
from datetime import date
from pathlib import Path

import openpyxl
import pytest

import daily_success_rate as dsr

ROOT = Path(__file__).resolve().parents[1]
POS_FILE = "POS SUCCESS RATE November 16,2026.xlsx"
ATM_FILE = "ATM Declined Transaction Report for Nov 16,2026.xlsx"
SAMPLE_FILE = "sample.xlsx"


def workbook_bytes(workbook) -> bytes:
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def synthetic(kind, day, banks, rates, total, formula=False):
    """Build a small daily summary laid out like the EthSwitch workbooks."""
    wb = openpyxl.Workbook()
    ws = wb.active
    month = calendar.month_name[day.month]
    if kind == "POS":
        ws["A1"] = f"{month} {day.day},{day.year} Pos Transaction Decline Response Summary"
    else:
        ws["A1"] = f"EthSwitch \n{month} {day.day},{day.year}  Transaction Decline Response Summary"
    ws["A4"] = "RC/BANK NAME"
    for offset, bank in enumerate(banks):
        ws.cell(row=4, column=2 + offset, value=bank)
    total_column = 2 + len(banks)
    ws.cell(row=4, column=total_column, value="Total")
    ws.cell(row=6, column=1,
            value="success rate" if kind == "POS" else "Issu. SUCC. RATE (%)")
    for offset, rate in enumerate(rates):
        column = 2 + offset
        if formula:
            ws.cell(row=8, column=column, value=round(rate * 1000))
            ws.cell(row=9, column=column, value=1000)
            letter = openpyxl.utils.get_column_letter(column)
            ws.cell(row=6, column=column, value=f"={letter}8/{letter}9")
        else:
            ws.cell(row=6, column=column, value=rate)
    if formula:
        ws.cell(row=8, column=total_column, value=round(total * 1000))
        ws.cell(row=9, column=total_column, value=1000)
        letter = openpyxl.utils.get_column_letter(total_column)
        ws.cell(row=6, column=total_column, value=f"={letter}8/{letter}9")
    else:
        ws.cell(row=6, column=total_column, value=total)
    return workbook_bytes(wb)


def combined(formula=False, pos_day=None, atm_day=None):
    pos_day = pos_day or date(2026, 11, 16)
    atm_day = atm_day or pos_day
    payloads = [
        ("POS SUCCESS RATE.xlsx",
         synthetic("POS", pos_day, ["Abay", "BOA", "BIB"],
                   [0.949152, 0.988947, 0.875], 0.976192, formula)),
        ("ATM Declined Transaction Report.xlsx",
         synthetic("ATM", atm_day, ["CBE", "BOA", "LION"],
                   [0.993492, 0.996298, 0.86225], 0.991169, formula)),
    ]
    return dsr.combine(dsr.read_reports(payloads))


def sample_payloads():
    out = []
    for name in (POS_FILE, ATM_FILE):
        path = ROOT / name
        if not path.exists():
            pytest.skip(f"{name} is not present")
        out.append((name, path.read_bytes()))
    return out


# ── Reading and validation ─────────────────────────────────────────────────

def test_combines_the_two_reports():
    report = combined()
    assert report.day == date(2026, 11, 16)
    assert report.title == "Daily POS and ATM Success rate Nov 16,2026"
    assert report.pos == [("Abay", 0.95), ("BOA", 0.99), ("BIB", 0.88),
                          ("Total", 0.98)]
    assert report.atm == [("CBE", 0.993), ("BOA", 0.996), ("LION", 0.862),
                          ("Total", 0.991)]
    assert report.pos_total == 0.98
    assert report.atm_total == 0.991


def test_reads_rates_written_as_formulas():
    assert combined(formula=True).pos == [("Abay", 0.95), ("BOA", 0.99),
                                          ("BIB", 0.88), ("Total", 0.98)]


def test_different_dates_raise_a_date_difference_error():
    with pytest.raises(ValueError, match="date difference"):
        combined(atm_day=date(2026, 11, 17))


def test_two_reports_of_the_same_kind_are_rejected():
    wb = synthetic("POS", date(2026, 11, 16), ["Abay"], [0.95], 0.95)
    with pytest.raises(ValueError, match="one POS report and one ATM report"):
        dsr.combine(dsr.read_reports([("first.xlsx", wb), ("second.xlsx", wb)]))


def test_a_report_without_a_rate_row_reports_an_error():
    wb = openpyxl.Workbook()
    wb.active["A4"] = "RC/BANK NAME"
    report = dsr.read_reports([("bad.xlsx", workbook_bytes(wb))])[0]
    assert report.error
    with pytest.raises(ValueError, match="bad.xlsx"):
        dsr.combine([report])


# ── Output ────────────────────────────────────────────────────────────────

def test_preview_html_offers_a_png_download():
    html = dsr.build_preview_html(combined())
    assert "Download PNG" in html
    assert "html2canvas" in html
    assert "Daily_POS_and_ATM_Success_Rate_16-Nov-2026.png" in html
    assert "#00B050" in html
    assert "95%" in html and "99.30%" in html


def test_excel_filename():
    assert dsr.report_filename(combined()) == (
        "Daily_POS_and_ATM_Success_Rate_16-Nov-2026.xlsx")


def test_generated_workbook_matches_the_reference_sample():
    """Every value, fill, format, font, border and merge matches sample.xlsx."""
    path = ROOT / SAMPLE_FILE
    if not path.exists():
        pytest.skip(f"{SAMPLE_FILE} is not present")
    report = dsr.combine(dsr.read_reports(sample_payloads()))

    generated = openpyxl.load_workbook(
        io.BytesIO(dsr.build_success_rate_excel(report))).active
    expected = openpyxl.load_workbook(path).active

    assert generated.dimensions == expected.dimensions
    assert {str(m) for m in generated.merged_cells.ranges} == {
        str(m) for m in expected.merged_cells.ranges}

    def describe(cell):
        fill = cell.fill
        colour = None
        if fill and fill.fill_type == "solid" and fill.fgColor:
            colour = (fill.fgColor.type,
                      fill.fgColor.rgb if fill.fgColor.type == "rgb"
                      else fill.fgColor.theme)
        border = tuple(
            getattr(getattr(cell.border, side), "style", None)
            for side in ("left", "right", "top", "bottom"))
        return (cell.value, colour, cell.number_format, cell.font.name,
                cell.font.sz, cell.font.bold, border)

    for row in range(1, max(generated.max_row, expected.max_row) + 1):
        for column in range(1, max(generated.max_column, expected.max_column) + 1):
            assert describe(generated.cell(row, column)) == describe(
                expected.cell(row, column)), f"cell {generated.cell(row, column).coordinate}"
