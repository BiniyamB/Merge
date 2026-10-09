"""Regression tests for the daily "Financial & Decline Transaction" sheet.

They run against the reference workbook that ships in ``28.09.2026`` and are
skipped when that sample data is not present, so the suite still passes on a
checkout without the real reports.
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from datetime import date
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import daily_report as dr  # noqa: E402

SAMPLE = ROOT / "28.09.2026"
BOOK = SAMPLE / "September_Successful_Financial_and_Decline_Transaction_Report_28.xlsx"
TARGET = "28.09.2026"
DONOR = "27.09.2026"
NEW_DAY = "29.09.2026"

SOURCE_FOR_KEY = {
    "iss_success": "Iss_Report Sucess.xlsx",
    "acq_success": "Acq_Report Sucess.xlsx",
    "iss_decline": "Iss_Report Decline.xlsx",
    "acq_decline": "Acq_Report Decline.xlsx",
    "ips_success": "IPS success for source and destination.xlsx",
    "ips_decline": "IPS Declined for source and destination.xlsx",
    "qr_success": "QR success for source and destination.xlsx",
    "qr_decline": "QR Declined for source and destination.xlsx",
}

needs_sample = pytest.mark.skipif(
    not (BOOK.exists() and all((SAMPLE / n).exists() for n in SOURCE_FOR_KEY.values())),
    reason="reference workbook or source exports are not available",
)


# ── helpers ─────────────────────────────────────────────────────────────────
def sheet_part(zf: zipfile.ZipFile, name: str) -> str:
    workbook = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = re.search(rf'<sheet name="{re.escape(name)}"[^>]*?r:id="([^"]+)"', workbook).group(1)
    return "xl/" + re.search(rf'Id="{rid}"[^>]*?Target="([^"]+)"', rels).group(1)


def read_part(book: bytes, name: str) -> str:
    with zipfile.ZipFile(io.BytesIO(book)) as zf:
        return zf.read(sheet_part(zf, name)).decode("utf-8")


def shared_strings(book: bytes) -> str:
    with zipfile.ZipFile(io.BytesIO(book)) as zf:
        return zf.read("xl/sharedStrings.xml").decode("utf-8")


def cell_map(xml: str, sst_xml: str) -> dict[str, tuple[str, str]]:
    """``A1 -> (value, formula)`` for every cell that has content."""
    sst = dr.SharedStrings(sst_xml)
    out: dict[str, tuple[str, str]] = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>(.*?)</c>)', xml, re.S):
        ref, attrs, inner = m.group(1), m.group(2), m.group(3) or ""
        formula = re.search(r"<f[^>]*>(.*?)</f>", inner, re.S)
        value = re.search(r"<v>(.*?)</v>", inner, re.S)
        text = re.search(r"<t[^>]*>(.*?)</t>", inner, re.S)
        if text is not None:
            out[ref] = (text.group(1), "")
        elif value is not None:
            raw = value.group(1)
            out[ref] = (sst.text(int(raw)) if 't="s"' in attrs else raw,
                        formula.group(1) if formula else "")
        else:
            out[ref] = ("", formula.group(1) if formula else "")
    return out


def block_values(xml: str, sst_xml: str) -> dict[str, dict[str, list[float]]]:
    """``{bank: [numbers...]}`` per block, keyed by normalised bank name."""
    cells = cell_map(xml, sst_xml)
    out: dict[str, dict[str, list[float]]] = {}
    for block in dr.BLOCKS:
        bank_col = dr.col_letter(block.bank_col)
        rows: dict[str, list[float]] = {}
        number = dr.FIRST_DATA_ROW
        while True:
            bank = cells.get(f"{bank_col}{number}")
            if bank is None or bank[0] in ("", "Total"):
                break
            rows[dr.normalize_bank(bank[0])] = [
                float(cells.get(f"{dr.col_letter(col)}{number}", ("0", ""))[0] or 0)
                for col in range(block.first_col, block.last_col + 1)
            ]
            number += 1
        out[block.key] = rows
    return out


def styles_map(xml: str) -> dict[str, str]:
    return {m.group(1): (m.group(2) or "0")
            for m in re.finditer(r'<c r="([A-Z]+\d+)"[^>]*?\ss="(\d+)"', xml)}


def merges_of(xml: str) -> set[str]:
    return set(re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"', xml))


# ── fixtures ────────────────────────────────────────────────────────────────
@pytest.fixture(scope="module")
def template() -> bytes:
    return BOOK.read_bytes()


@pytest.fixture(scope="module")
def sources() -> dict[str, bytes]:
    return {key: (SAMPLE / name).read_bytes() for key, name in SOURCE_FOR_KEY.items()}


@pytest.fixture(scope="module")
def reference_footer(template: bytes) -> dr.MonthlyFooter:
    return dr.read_monthly_footer(read_part(template, TARGET), shared_strings(template))


# ── pure logic: no sample data needed ───────────────────────────────────────
def test_pick_donor_sheet_skips_the_replaced_sheet():
    names = ["01.09.2026", TARGET, "27.09.2026", DONOR, "not-a-date"]
    # the day being rebuilt is never its own donor
    assert dr.pick_donor_sheet(names, TARGET) == DONOR
    # a brand new day takes the most recent dated sheet
    assert dr.pick_donor_sheet(names, "30.09.2026") == TARGET
    # a workbook with a single sheet still has to work, so it falls back to it
    assert dr.pick_donor_sheet([TARGET], TARGET) == TARGET


def test_month_and_sheet_date():
    assert dr.sheet_date(TARGET) == date(2026, 9, 28)
    assert dr.month_of(TARGET) == (2026, 9)
    assert dr.sheet_date("Summary") is None
    assert dr.month_of("Summary") is None


def test_bank_matcher_folds_appends_and_ignores():
    matcher = dr.BankMatcher(["Bunna International Bank", "Siket"])
    mapping = matcher.resolve_all([
        "bunna  international  bank",   # same bank, odd case and spacing
        "Bunna International Bnk",       # near miss, above the fuzzy cutoff
        "Goh Betoch",                    # genuinely new
        "SYSTEM INSTITUTION",            # not a real institution
    ])
    assert mapping["bunna  international  bank"] == "Bunna International Bank"
    assert mapping["Bunna International Bnk"] == "Bunna International Bank"
    assert mapping["Goh Betoch"] == "Goh Betoch"
    assert mapping["SYSTEM INSTITUTION"] is None
    assert matcher.new_banks == ["Goh Betoch"]
    assert matcher.ignored == ["SYSTEM INSTITUTION"]


def test_system_institution_is_ignored():
    assert dr.IGNORED_BANKS == {"systeminstitution"}
    assert dr.normalize_bank("SYSTEM INSTITUTION") in dr.IGNORED_BANKS


def test_grand_total_excludes_npg_and_rtp():
    labels = [spec.label for spec in dr.FOOTER_ROWS]
    excluded = [labels[i] for i in range(len(labels)) if i not in dr.FOOTER_TOTAL_INCLUDED]
    assert "NPG-card and NPG-online" in excluded
    assert "IPS RTP" in excluded


# ── the real workbook ───────────────────────────────────────────────────────
@needs_sample
def test_read_monthly_footer_matches_reference(reference_footer: dr.MonthlyFooter):
    assert reference_footer.rtp_count == 7
    assert reference_footer.rtp_value == 502
    assert reference_footer.plans[0] == pytest.approx(347974.3)
    assert reference_footer.plans[4] == dr.PLAN_PLACEHOLDER
    assert reference_footer.plans[6] == dr.PLAN_PLACEHOLDER


@needs_sample
def test_regenerating_the_reference_reproduces_it(template, sources, reference_footer):
    """Self-donor rebuild: values, formulas, styles, merges and row layout."""
    diagnostics = dr.Diagnostics()
    built = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=date(2026, 9, 28), sources=sources,
        footer=reference_footer, donor_name=TARGET, diagnostics=diagnostics)

    expected_xml = read_part(template, TARGET)
    got_xml = read_part(built, TARGET)
    sst = shared_strings(built)

    expected, got = cell_map(expected_xml, shared_strings(template)), cell_map(got_xml, sst)
    assert set(expected) == set(got)

    # The total row's achievement and decline cells are ratios of the total
    # row's own figures, so they are the one place the rebuild reproduces the
    # reference only in style: the reference adds the line percentages there.
    total_row = (dr.DonorSheet.parse(expected_xml, dr.SharedStrings(shared_strings(template)))
                 .footer_start + 2 + len(dr.FOOTER_ROWS))
    ratio_refs = {f"{cell}{total_row}" for cell, _, _ in dr.TOTAL_RATIO_CELLS}
    for ref, (want, _f) in expected.items():
        if ref in ratio_refs:
            continue
        try:
            assert float(got[ref][0]) == pytest.approx(float(want))
        except (TypeError, ValueError):
            assert got[ref][0] == want

    for cell, numerator, denominator in dr.TOTAL_RATIO_CELLS:
        ref = f"{cell}{total_row}"
        assert got[ref][1] == f"{numerator}{total_row}/{denominator}{total_row}"
        assert float(got[ref][0]) == pytest.approx(
            float(expected[f"{numerator}{total_row}"][0])
            / float(expected[f"{denominator}{total_row}"][0]))

    assert styles_map(got_xml) == styles_map(expected_xml)
    assert merges_of(got_xml) == merges_of(expected_xml)
    assert re.search(r'<dimension ref="([^"]+)"', got_xml).group(1) == \
        re.search(r'<dimension ref="([^"]+)"', expected_xml).group(1)
    assert re.findall(r"<row r=\"\d+\"[^>]*>", got_xml) == \
        re.findall(r"<row r=\"\d+\"[^>]*>", expected_xml)
    assert "SYSTEM INSTITUTION" in diagnostics.ignored


@needs_sample
def test_real_donor_matches_reference_data(template, sources, reference_footer):
    """The everyday case: 27.09.2026 is the donor, 28.09.2026 the target."""
    diagnostics = dr.Diagnostics()
    built = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=date(2026, 9, 28), sources=sources,
        footer=reference_footer, donor_name=DONOR, diagnostics=diagnostics)

    expected = block_values(read_part(template, TARGET), shared_strings(template))
    got = block_values(read_part(built, TARGET), shared_strings(built))
    assert set(expected) == set(got)
    for key, rows in expected.items():
        for bank, values in rows.items():
            assert got[key][bank] == pytest.approx(values, rel=1e-9, abs=1e-6)

    # banks the donor has never seen are appended, not dropped
    appended = {bank for banks in diagnostics.new_banks.values() for bank in banks}
    assert "Goh Betoch" in appended
    assert "Siket" in appended
    assert "SYSTEM INSTITUTION" not in appended


@needs_sample
def test_new_day_is_appended_last_and_package_stays_valid(template, sources,
                                                           reference_footer):
    built = dr.build_daily_sheet(
        template, sheet_name=NEW_DAY, day=date(2026, 9, 29), sources=sources,
        footer=reference_footer)

    with zipfile.ZipFile(io.BytesIO(built)) as zf:
        assert zf.testzip() is None
        names = set(zf.namelist())
        assert "xl/calcChain.xml" not in names, "a stale calcChain breaks Excel"
        order = re.findall(r'<sheet name="([^"]+)"', zf.read("xl/workbook.xml").decode())
        assert len(order) == 29
        assert order[-1] == NEW_DAY

        # every relationship still points at a part that exists
        for rels_name in (n for n in names if n.endswith(".rels")):
            base = rels_name.rsplit("_rels/", 1)[0]
            for target in re.findall(r'Target="([^"]+)"', zf.read(rels_name).decode()):
                if target.startswith(("http", "mailto", "file:")):
                    continue
                stack: list[str] = []
                for item in [p for p in (base + target).split("/") if p not in ("", ".")]:
                    stack.pop() if item == ".." else stack.append(item)
                assert "/".join(stack) in names, f"{rels_name} -> {target}"

        app = zf.read("docProps/app.xml").decode("utf-8")
        count = re.search(r"Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)",
                          app).group(1)
        assert int(count) == 29
        assert len(re.findall(r"<vt:lpstr>", app)) - 1 == 29

        sheet = dr.DonorSheet.parse(zf.read(sheet_part(zf, NEW_DAY)).decode(),
                                    dr.SharedStrings(zf.read("xl/sharedStrings.xml").decode()))
    assert sheet.footer_start == 55, "the date row moved, so the footer would be misaligned"
    # the label lives in the shared string table, so resolve it rather than
    # searching the raw XML
    assert cell_map(read_part(built, NEW_DAY), shared_strings(built))["A55"][0] == \
        "Date 29.09.2026"


@needs_sample
def test_replacing_a_day_keeps_its_tab_position(template, sources, reference_footer):
    built = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=date(2026, 9, 28), sources=sources,
        footer=reference_footer)
    with zipfile.ZipFile(io.BytesIO(built)) as zf:
        order = re.findall(r'<sheet name="([^"]+)"', zf.read("xl/workbook.xml").decode())
    assert len(order) == 28
    assert order[-1] == TARGET


@needs_sample
def test_first_day_of_a_new_month_accepts_blank_footer(template, sources):
    """The footer form starts empty on a month with no sheet yet."""
    footer = dr.MonthlyFooter(plans=[dr.PLAN_PLACEHOLDER] * len(dr.FOOTER_ROWS))
    built = dr.build_daily_sheet(
        template, sheet_name=NEW_DAY, day=date(2026, 9, 29), sources=sources, footer=footer)
    cells = cell_map(read_part(built, NEW_DAY), shared_strings(built))
    assert cells["A55"][0] == "Date 29.09.2026"
    # no plans: every plan cell keeps the reference placeholder and the
    # achievement column stays at zero instead of dividing by a blank
    assert cells["B57"][0] == dr.PLAN_PLACEHOLDER
    assert cells["E57"][0] == "0"
    # the grand total still chains the five financial lines, skipping NPG
    # (row 61) and IPS RTP (row 63)
    assert cells["A64"][0] == "Total interbank (Financial only)"
    assert cells["B64"][1] == "B57+B58+B59+B60+B62"
    assert cells["F64"][1] == "F57+F58+F59+F60+F62"


@needs_sample
def test_total_row_rates_are_ratios_of_the_totals(template, sources, reference_footer):
    """The total's achievement and decline rate are C64/B64 and F64/H64.

    The reference workbook drags one ``B57+B58+...`` across the whole total
    row, so its E64 and I64 *add* the line percentages together. The sheet
    recomputes both from the total row's own figures instead.
    """
    built = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=date(2026, 9, 28), sources=sources,
        footer=reference_footer, donor_name=TARGET)
    cells = cell_map(read_part(built, TARGET), shared_strings(built))
    total_row = (dr.DonorSheet.parse(read_part(template, TARGET),
                                     dr.SharedStrings(shared_strings(template)))
                 .footer_start + 2 + len(dr.FOOTER_ROWS))
    for cell, numerator, denominator in dr.TOTAL_RATIO_CELLS:
        assert cells[f"{cell}{total_row}"][1] == \
            f"{numerator}{total_row}/{denominator}{total_row}"
        assert float(cells[f"{cell}{total_row}"][0]) == pytest.approx(
            float(cells[f"{numerator}{total_row}"][0])
            / float(cells[f"{denominator}{total_row}"][0]))
    # sanity: the line percentages are not merely summed any more
    line_sum = sum(float(cells[f"E{r}"][0]) for r in range(57, 63)
                   if cells.get(f"E{r}", ("", ""))[0] not in ("", None))
    assert float(cells[f"E{total_row}"][0]) != pytest.approx(line_sum)


@needs_sample
def test_output_opens_in_openpyxl(template, sources, reference_footer):
    openpyxl = pytest.importorskip("openpyxl")
    built = dr.build_daily_sheet(
        template, sheet_name=NEW_DAY, day=date(2026, 9, 29), sources=sources,
        footer=reference_footer)
    book = openpyxl.load_workbook(io.BytesIO(built))
    try:
        assert NEW_DAY in book.sheetnames
        assert book[NEW_DAY]["A5"].value  # the first bank row has a name
    finally:
        book.close()
