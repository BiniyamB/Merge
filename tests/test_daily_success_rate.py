"""Unit tests for the daily POS & ATM success-rate report."""

import calendar
import io
import re
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
    assert "#ED7D31" in html
    assert "border:1px solid #000" in html
    assert "#report-content{display:inline-block;padding:0" in html
    assert "width:w,height:h,scrollX:0,scrollY:0" in html
    assert "95%" in html and "99.30%" in html


def test_preview_rows_stay_aligned_when_one_side_is_shorter():
    payloads = [
        ("POS.xlsx", synthetic("POS", date(2026, 11, 16), ["Abay"], [0.95], 0.95)),
        ("ATM.xlsx", synthetic("ATM", date(2026, 11, 16),
                               ["CBE", "BOA", "LION"],
                               [0.993, 0.996, 0.862], 0.991)),
    ]
    html = dsr.build_preview_html(dsr.combine(dsr.read_reports(payloads)))
    table = html[html.index("<table>"):html.index("</table>")]
    for row in table.split("<tr>")[3:]:
        cells = re.findall(r"<td([^>]*)>", row)
        columns = sum(int(re.search(r"colspan='(\d+)'", cell).group(1))
                      if "colspan" in cell else 1 for cell in cells)
        assert columns == 5, row


def test_excel_filename():
    assert dsr.report_filename(combined()) == (
        "Daily_POS_and_ATM_Success_Rate_16-Nov-2026.xlsx")


def test_below_green_declines_exclude_the_cardholder_codes():
    report = dsr.combine(dsr.read_reports(sample_payloads()))
    detail = dsr.below_green_declines(report)
    banks = {bank for bank, _, _ in detail}
    assert banks == {"BUNNA", "ENAT", "LION", "OIB", "ZAMZAM", "HIJRA",
                     "SINQEE", "AMHARA", "Tseday"}
    excluded = {str(code) for code in dsr.EXCLUDED_DECLINE_CODES}
    for _, _, items in detail:
        assert all(code not in excluded for code, _, _ in items)
    oib = {code: count for bank, _, items in detail if bank == "OIB"
           for code, count, _ in items}
    assert oib == {"801": 159, "802": 357, "939": 15, "905": 2, "917": 24}


def test_declines_are_copyable_text_and_not_in_excel_or_png():
    report = dsr.combine(dsr.read_reports(sample_payloads()))
    text = dsr.declines_text(report)
    lines = text.splitlines()
    assert lines[0] == "BUNNA: They had 905 and 917 response code declines."
    assert lines[2] == "LION: They had 801, 905 and 917 response code declines."
    assert "Tseday: They had 801, 886 and 917 response code declines." in lines
    assert all(line.endswith(".") for line in lines)
    assert "159" not in text

    ws = openpyxl.load_workbook(
        io.BytesIO(dsr.build_success_rate_excel(report))).active
    values = [cell.value for row in ws.iter_rows() for cell in row
              if cell.value is not None]
    assert 159 not in values

    html = dsr.build_preview_html(report)
    assert "response code declines" not in html
    assert "159" not in html


def test_a_specific_issue_replaces_the_response_codes_for_that_bank():
    report = dsr.combine(dsr.read_reports(sample_payloads()))
    text = dsr.declines_text(report, {
        "Tseday": "They had a CBS performance issue.",
    })
    assert "Tseday: They had a CBS performance issue." in text.splitlines()
    assert "Tseday: They had 801" not in text
    assert "OIB: They had 801, 802, 905, 917 and 939 response code declines." in text
    assert "Oromia: Their card service was not working." in dsr.declines_text(
        report, {"Oromia": "Their card service was not working"})


def test_generated_workbook_matches_the_reference_sample():
    """Every value, fill, format, font, border and merge matches sample.xlsx."""
    path = ROOT / SAMPLE_FILE
    if not path.exists():
        pytest.skip(f"{SAMPLE_FILE} is not present")
    report = dsr.combine(dsr.read_reports(sample_payloads()))

    generated = openpyxl.load_workbook(
        io.BytesIO(dsr.build_success_rate_excel(report))).active
    expected = openpyxl.load_workbook(path).active

    expected_merges = {str(m) for m in expected.merged_cells.ranges}
    assert expected_merges <= {str(m) for m in generated.merged_cells.ranges}

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

    for row in range(1, expected.max_row + 1):
        for column in range(1, expected.max_column + 1):
            assert describe(generated.cell(row, column)) == describe(
                expected.cell(row, column)), f"cell {generated.cell(row, column).coordinate}"
