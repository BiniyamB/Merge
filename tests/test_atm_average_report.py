"""Unit tests for atm_average_report."""

import io
from datetime import date, datetime
from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from atm_average_report import (
    ATM_RATE_TIERS,
    ParsedDay,
    build_atm_average_excel,
    build_average_frame,
    duplicate_days,
    rate_fill,
    read_daily_report,
    read_daily_reports,
    report_filename,
)

ROOT = Path(__file__).resolve().parents[1]
DAILY_FILES = [
    "ATM Declined Transaction Report for Oct 6,2026.xlsx",
    "ATM Declined Transaction Report for Oct 7,2026.xlsx",
    "ATM Declined Transaction Report for Oct 8,2026.xlsx",
]


def payloads():
    """The three sample daily reports, skipped when they are not present."""
    out = []
    for name in DAILY_FILES:
        path = ROOT / name
        if not path.exists():
            pytest.skip(f"{name} is not present")
        out.append((name, path.read_bytes()))
    return out


def workbook_bytes(workbook) -> bytes:
    buffer = io.BytesIO()
    workbook.save(buffer)
    return buffer.getvalue()


def synthetic_daily(banks, rates, day, title=None):
    """Build a daily ATM report laid out like the EthSwitch ones.

    The success-rate row is written as formulas (``=B40/B41`` over the counts
    above it) with no cached results, which is what openpyxl produces and
    what the parser has to resolve on its own.
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = title or f"EthSwitch \n{day:%B %-d,%Y}  Transaction Decline Response Summary"

    first_rc_row = 5
    ws.cell(row=4, column=1, value="RC/BANK NAME")
    for offset, bank in enumerate(banks):
        ws.cell(row=4, column=2 + offset, value=bank)
        ws.cell(row=first_rc_row, column=2 + offset, value=100)
    total_row = first_rc_row + 1
    rate_row = total_row + 2
    ws.cell(row=total_row, column=1, value="TOTAL")
    ws.cell(row=rate_row, column=1, value="Issu. SUCC. RATE (%)")

    for offset, rate in enumerate(rates):
        column = 2 + offset
        letter = openpyxl.utils.get_column_letter(column)
        ws.cell(row=total_row, column=column,
                value=f"=SUM({letter}{first_rc_row}:{letter}{first_rc_row})")
        ws.cell(row=rate_row, column=column, value=rate)
    return workbook_bytes(wb)


# ── Success rate tiers ─────────────────────────────────────────────────────

@pytest.mark.parametrize("rate,colour", [
    (1.00, "FF00B050"),
    (0.985, "FF00B050"),
    (0.9849, "FFFFFF00"),
    (0.98, "FFFFFF00"),
    (0.90, "FFFFFF00"),
    (0.8999, "FFFFC000"),
    (0.85, "FFFFC000"),
    (0.80, "FFFFC000"),
    (0.7999, "FFFF0000"),
    (0.50, "FFFF0000"),
    (0.0, "FFFF0000"),
])
def test_rate_fills_follow_the_tier_bounds(rate, colour):
    """Each band starts exactly at its stated bound."""
    assert rate_fill(rate).start_color.rgb == colour


def test_a_missing_rate_is_left_unfilled():
    """A day that does not report a bank must not read as a zero."""
    assert rate_fill(None) is None


def test_the_four_tiers_are_distinct_and_reachable():
    """The four documented bands each map to their own fill."""
    assert len(ATM_RATE_TIERS) == 4
    colours = {rate_fill(rate).start_color.rgb for rate in (1.0, 0.95, 0.85, 0.5)}
    assert len(colours) == 4


# ── Reading one daily report ───────────────────────────────────────────────

def test_reads_the_date_from_both_the_file_name_and_the_report():
    """Both dates are parsed, and reported as agreeing."""
    name, data = payloads()[0]
    report = read_daily_report(data, name)

    assert report.error is None
    assert report.parsed is not None
    assert report.date_from_name == date(2026, 10, 6)
    assert report.date_in_report == date(2026, 10, 6)
    assert report.date_matches is True
    assert report.notes == []


def test_reads_a_rate_for_every_bank_and_skips_the_total_column():
    """All 24 bank columns are read; `Total` is not one of them."""
    name, data = payloads()[0]
    parsed = read_daily_report(data, name).parsed

    assert len(parsed.banks) == 24
    assert parsed.banks[0] == "CBE"
    assert parsed.banks[-1] == "GADAA"
    assert "Total" not in parsed.banks
    assert set(parsed.rates) == set(parsed.banks)
    assert all(0.0 <= value <= 1.0 for value in parsed.rates.values())


def test_the_rate_is_the_reports_own_cached_value():
    """The rate is read, not re-derived by a different rule."""
    name, data = payloads()[0]
    parsed = read_daily_report(data, name).parsed

    wb = openpyxl.load_workbook(ROOT / name, data_only=True)
    ws = wb["Sheet1"]
    try:
        assert parsed.rates["CBE"] == pytest.approx(ws.cell(row=36, column=2).value)
        assert parsed.rates["BOA"] == pytest.approx(ws.cell(row=36, column=3).value)
        assert parsed.rates["GADAA"] == pytest.approx(ws.cell(row=36, column=25).value)
    finally:
        wb.close()


def test_all_three_daily_reports_read_cleanly():
    """The sample trio yields three days, no errors and no notes."""
    reports = read_daily_reports(payloads())

    assert len(reports) == 3
    assert all(r.error is None for r in reports)
    assert all(r.notes == [] for r in reports)
    assert [r.parsed.day for r in reports] == [
        date(2026, 10, 6), date(2026, 10, 7), date(2026, 10, 8)]
    assert all(r.date_matches for r in reports)


def test_resolves_a_formula_success_row_with_no_cached_values():
    """A report never opened in Excel still yields its rates.

    The daily reports express the rate as `=B40/B41` built from `=B37+B38`
    over `=B39+B40`; with nothing cached those have to be resolved from the
    counts underneath rather than reported as missing.
    """
    banks = ["CBE", "BOA"]
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "EthSwitch \nMarch 3,2026  Transaction Decline Response Summary"
    ws["A4"] = "RC/BANK NAME"
    ws["B4"], ws["C4"] = "CBE", "BOA"
    # Two decline codes with counts, then the report's calculation chain.
    ws["A5"], ws["A6"] = 901, 915
    ws["B5"], ws["B6"] = 10, 5        # CBE declines -> 15
    ws["C5"], ws["C6"] = 20, 4        # BOA declines -> 24
    ws["A7"] = "TOTAL"
    ws["B7"], ws["C7"] = "=SUM(B5:B6)", "=SUM(C5:C6)"
    # expected = total*9 + first decline; rate = expected / (total + expected)
    ws["B8"], ws["C8"] = "=B7*9+B5", "=C7*9+C5"
    ws["A9"] = "Issu. SUCC. RATE (%)"
    ws["B9"], ws["C9"] = "=B8/(B7+B8)", "=C8/(C7+C8)"

    parsed = read_daily_report(workbook_bytes(wb),
                               "ATM Declined Transaction Report for Mar 3,2026.xlsx").parsed

    assert parsed is not None
    assert parsed.rates["CBE"] == pytest.approx(145 / 160)
    assert parsed.rates["BOA"] == pytest.approx(236 / 260)


def test_flags_a_file_name_and_report_date_that_disagree():
    """A mismatched date is reported, and the report date is used."""
    _, data = payloads()[0]
    report = read_daily_report(data, "ATM Declined Transaction Report for Oct 9,2026.xlsx")

    assert report.date_matches is False
    assert report.parsed.day == date(2026, 10, 6)
    assert any("date difference" in note for note in report.notes)


def test_falls_back_to_the_report_date_when_the_name_has_none():
    """No date in the file name is a note, not a mismatch."""
    _, data = payloads()[0]
    report = read_daily_report(data, "atm_daily_copy.xlsx")

    assert report.date_from_name is None
    assert report.date_matches is None
    assert report.parsed.day == date(2026, 10, 6)
    assert any("no date in the file name" in note for note in report.notes)


def test_reads_the_date_from_an_abbreviated_month_in_the_title():
    """`Oct 6,2026` and `October 6,2026` resolve to the same day."""
    _, data = payloads()[0]
    for title in ("EthSwitch \nOct 6,2026  Transaction Decline Response Summary",
                  "EthSwitch \nOctober 6,2026  Transaction Decline Response Summary"):
        blob = synthetic_daily(["CBE"], [0.99], date(2026, 10, 6), title=title)
        report = read_daily_report(blob, "ATM Declined Transaction Report for Oct 6,2026.xlsx")
        assert report.date_matches is True, title


def test_reports_a_workbook_that_is_not_a_daily_atm_report():
    """An unrelated file fails with a message instead of raising."""
    wb = openpyxl.Workbook()
    wb.active["A1"] = "some other report"

    report = read_daily_report(workbook_bytes(wb), "other.xlsx")

    assert report.parsed is None
    assert "RC/BANK NAME" in report.error


def test_reports_an_unreadable_file_instead_of_raising():
    """A corrupt upload is reported per file so the others still merge."""
    report = read_daily_report(b"not a workbook", "broken.xlsx")

    assert report.parsed is None
    assert report.error


# ── The average table ──────────────────────────────────────────────────────

def test_rows_are_the_days_in_date_order_and_the_average_is_last():
    days = [ParsedDay(date(2026, 10, 8), ["CBE"], {"CBE": 0.90}),
            ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 1.00}),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.50})]
    frame = build_average_frame(days)

    assert list(frame.index) == [date(2026, 10, 6), date(2026, 10, 7),
                                 date(2026, 10, 8), "Average"]
    assert frame.loc["Average", "CBE"] == pytest.approx(0.80)


def test_the_average_is_a_simple_mean_of_the_daily_rates():
    """Each day weighs the same, as in cfd.xlsx."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.99}),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.91})]
    frame = build_average_frame(days)

    assert frame.loc["Average", "CBE"] == pytest.approx((0.99 + 0.91) / 2)


def test_a_bank_seen_in_only_some_days_still_gets_a_column():
    """A new institution keeps its column and its own average."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE", "NEWBANK"],
                      {"CBE": 0.99, "NEWBANK": 0.95}),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.97})]
    frame = build_average_frame(days)

    assert "NEWBANK" in frame.columns
    assert pd.isna(frame.loc[date(2026, 10, 7), "NEWBANK"])
    # The day that does not report NEWBANK is left out rather than counted
    # as a zero, so the average is 95% and not 47.5%.
    assert frame.loc["Average", "NEWBANK"] == pytest.approx(0.95)


def test_bank_order_follows_the_reports_then_the_extra_ones():
    days = [ParsedDay(date(2026, 10, 6), ["CBE", "BOA"], {"CBE": 0.9, "BOA": 0.9}),
            ParsedDay(date(2026, 10, 7), ["BOA", "ZAMZAM"],
                      {"BOA": 0.9, "ZAMZAM": 0.9})]
    frame = build_average_frame(days)

    assert list(frame.columns) == ["CBE", "BOA", "ZAMZAM"]


def test_a_single_day_is_averaged_against_itself():
    frame = build_average_frame([ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.97})])

    assert frame.loc["Average", "CBE"] == pytest.approx(0.97)


def test_the_three_sample_days_average_across_all_24_banks():
    """End to end over the real files: 3 day rows plus an Average row."""
    days = [r.parsed for r in read_daily_reports(payloads())]
    frame = build_average_frame(days)

    assert len(frame) == 4
    assert list(frame.columns)[:3] == ["CBE", "BOA", "ABAY"]
    assert frame.loc["Average", "CBE"] == pytest.approx(
        (0.995619936164989 + 0.9951165913807838 + 0.9950829947765525) / 3)


# ── The written workbook ───────────────────────────────────────────────────

def test_the_workbook_matches_the_shape_of_cfd():
    """Dates in column B, banks across row 2, Average last, all as formulas."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE", "BOA"], {"CBE": 0.99, "BOA": 0.95}),
            ParsedDay(date(2026, 10, 7), ["CBE", "BOA"], {"CBE": 0.97, "BOA": 0.93})]
    ws = openpyxl.load_workbook(
        io.BytesIO(build_atm_average_excel(days))).active

    assert [ws.cell(row=2, column=c).value for c in (1, 2, 3, 4)] == \
        [None, None, "CBE", "BOA"]
    assert ws["B3"].value == datetime(2026, 10, 6)
    assert ws["B3"].number_format == "d-mmm"
    assert ws["B4"].number_format == "d-mmm"
    assert ws["B5"].value == "Average"
    assert ws["C3"].value == pytest.approx(0.99)
    assert ws["C3"].number_format == "0.00%"
    assert ws["C5"].value == "=AVERAGE(C3:C4)"
    assert ws["D5"].value == "=AVERAGE(D3:D4)"


def test_the_average_row_is_filled_by_the_rate_it_shows():
    """The fill on the Average row matches the average, not the last day."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 1.00}),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.60})]
    ws = openpyxl.load_workbook(
        io.BytesIO(build_atm_average_excel(days))).active

    assert ws["C3"].fill.start_color.rgb == "FF00B050"  # day one is green
    assert ws["C4"].fill.start_color.rgb == "FFFF0000"  # day two is red
    assert ws["C5"].fill.start_color.rgb == "FFFFC000"  # their 80% mean is amber


def test_a_day_missing_a_bank_leaves_the_cell_blank():
    """The blank is excluded from =AVERAGE(), which ignores blanks in Excel."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE", "BOA"], {"CBE": 0.99, "BOA": 0.95}),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.97})]
    ws = openpyxl.load_workbook(
        io.BytesIO(build_atm_average_excel(days))).active

    assert ws["D4"].value is None
    assert ws["D4"].fill.patternType is None
    assert ws["D5"].value == "=AVERAGE(D3:D4)"


def test_every_bank_column_of_the_sample_trio_is_written():
    """All 24 banks from the real files land in the workbook, coloured."""
    days = [r.parsed for r in read_daily_reports(payloads())]
    ws = openpyxl.load_workbook(
        io.BytesIO(build_atm_average_excel(days))).active

    assert ws["C2"].value == "CBE"
    assert ws.cell(row=2, column=26).value == "GADAA"
    assert ws["B3"].value == datetime(2026, 10, 6)
    assert ws["B6"].value == "Average"
    assert ws["C6"].value == "=AVERAGE(C3:C5)"
    # HIJRA averages 76.7% across the three days, which is in the red band.
    assert ws.cell(row=6, column=21).fill.start_color.rgb == "FFFF0000"


def test_a_bank_no_day_reported_leaves_no_average_formula():
    """AVERAGE over a blank column would show #DIV/0! in Excel."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE", "GHOST"], {"CBE": 0.99}),
            ParsedDay(date(2026, 10, 7), ["CBE", "GHOST"], {"CBE": 0.97})]
    ws = openpyxl.load_workbook(
        io.BytesIO(build_atm_average_excel(days))).active

    assert ws["C5"].value == "=AVERAGE(C3:C4)"
    assert ws["D5"].value is None


def test_a_repeated_bank_header_does_not_produce_two_columns():
    """A name repeated down the header row is one bank, not two."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws["A1"] = "EthSwitch \nMarch 3,2026  Transaction Decline Response Summary"
    ws["A4"] = "RC/BANK NAME"
    ws["B4"], ws["C4"], ws["D4"] = "CBE", "BOA", "CBE"
    ws["A9"] = "Issu. SUCC. RATE (%)"
    ws["B9"], ws["C9"], ws["D9"] = 0.99, 0.95, 0.50

    parsed = read_daily_report(
        workbook_bytes(wb),
        "ATM Declined Transaction Report for Mar 3,2026.xlsx").parsed

    assert parsed.banks == ["CBE", "BOA"]


def test_two_files_for_one_day_are_reported():
    """A repeated date would double-count that day in every average."""
    days = [ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.99}, source="a.xlsx"),
            ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.97}, source="b.xlsx"),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.95}, source="c.xlsx")]

    assert duplicate_days(days) == {date(2026, 10, 6): ["a.xlsx", "b.xlsx"]}


def test_duplicate_days_is_empty_for_distinct_dates():
    days = [ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.99}, source="a.xlsx"),
            ParsedDay(date(2026, 10, 7), ["CBE"], {"CBE": 0.97}, source="b.xlsx")]

    assert duplicate_days(days) == {}


def test_the_download_name_holds_the_dates_it_covers():
    days = [ParsedDay(date(2026, 10, 6), ["CBE"], {"CBE": 0.99}),
            ParsedDay(date(2026, 10, 8), ["CBE"], {"CBE": 0.99})]

    assert report_filename(days) == \
        "ATM_Average_Success_Rate_06-Oct-2026_to_08-Oct-2026.xlsx"
    assert report_filename([days[0]]) == \
        "ATM_Average_Success_Rate_06-Oct-2026.xlsx"