"""Rebuild 28.09.2026 from the 27.09.2026 donor sheet and diff against the
reference sheet that ships with the workbook.

Run:  .venv\\Scripts\\python.exe tools\\check_daily_report.py
"""

from __future__ import annotations

import io
import re
import sys
import zipfile
from collections.abc import Mapping
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl

import daily_report as dr

ROOT = Path(__file__).resolve().parents[1]
BOOK = ROOT / "28.09.2026" / "September_Successful_Financial_and_Decline_Transaction_Report_28.xlsx"
FOLDER = ROOT / "28.09.2026"
TARGET = "28.09.2026"
DONOR = "27.09.2026"

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


def sheet_part(zf: zipfile.ZipFile, name: str) -> str:
    workbook = zf.read("xl/workbook.xml").decode("utf-8")
    rels = zf.read("xl/_rels/workbook.xml.rels").decode("utf-8")
    rid = re.search(rf'<sheet name="{re.escape(name)}"[^>]*?r:id="([^"]+)"', workbook).group(1)
    target = re.search(rf'Id="{rid}"[^>]*?Target="([^"]+)"', rels).group(1)
    return "xl/" + target


def validate_package(book: bytes, label: str) -> list[str]:
    """Every part well-formed, every relationship target present."""
    import xml.etree.ElementTree as ET

    problems: list[str] = []
    with zipfile.ZipFile(io.BytesIO(book)) as zf:
        names = set(zf.namelist())
        if zf.testzip() is not None:
            problems.append(f"{label}: corrupt zip entry {zf.testzip()}")
        for name in sorted(names):
            if not name.endswith((".xml", ".rels", ".vml")):
                continue
            try:
                ET.fromstring(zf.read(name))
            except ET.ParseError as exc:
                problems.append(f"{label}: {name} is not well-formed ({exc})")
        for name in sorted(n for n in names if n.endswith(".rels")):
            base = name.rsplit("_rels/", 1)[0]
            for target in re.findall(r'Target="([^"]+)"', zf.read(name).decode("utf-8")):
                if target.startswith(("http://", "https://", "mailto:", "file:")):
                    continue
                parts = [p for p in (base + target).split("/") if p not in ("", ".")]
                stack: list[str] = []
                for item in parts:
                    if item == "..":
                        if stack:
                            stack.pop()
                    else:
                        stack.append(item)
                resolved = "/".join(stack)
                if resolved not in names:
                    problems.append(f"{label}: {name} -> missing part {resolved}")
    return problems


def new_day_pass(template: bytes, sources: Mapping[str, bytes], footer) -> list[str]:
    """A day that is not in the workbook yet must be appended cleanly."""
    label = "new day 29.09.2026"
    built = dr.build_daily_sheet(template, sheet_name="29.09.2026",
                                 day=dr.date(2026, 9, 29), sources=sources,
                                 footer=footer, donor_name=TARGET)
    problems = validate_package(built, label)
    with zipfile.ZipFile(io.BytesIO(built)) as zf:
        names = zf.namelist()
        if "xl/calcChain.xml" in names:
            problems.append(f"{label}: calcChain was not removed")
        order = re.findall(r'<sheet name="([^"]+)"', zf.read("xl/workbook.xml").decode())
        if "29.09.2026" not in order:
            problems.append(f"{label}: sheet missing from workbook.xml")
        if len(order) != 29:
            problems.append(f"{label}: expected 29 sheets, found {len(order)}")
        if order[-1] != "29.09.2026":
            problems.append(f"{label}: new day is not the last tab ({order[-1]})")
        rels = zf.read("xl/_rels/workbook.xml.rels").decode()
        if len(re.findall(r"/worksheet\"", rels)) != len(order):
            problems.append(f"{label}: worksheet relationship count mismatch")
        app = zf.read("docProps/app.xml").decode("utf-8")
        count = re.search(r"Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)",
                          app).group(1)
        titles = len(re.findall(r"<vt:lpstr>", app)) - 1
        if int(count) != 29 or titles != 29:
            problems.append(f"{label}: app.xml count={count} titles={titles}")
        part = sheet_part(zf, "29.09.2026")
        xml = zf.read(part).decode("utf-8")
        sheet = dr.DonorSheet.parse(xml, dr.SharedStrings(
            zf.read("xl/sharedStrings.xml").decode("utf-8")))
        print(f"  {label}: sheet part {part}, date row {sheet.footer_start}, "
              f"blocks " + ", ".join(f"{k.split('_')[0]}:{len(v)}"
                                     for k, v in sheet.banks.items()))
        if sheet.footer_start != 55:
            problems.append(f"{label}: date row {sheet.footer_start}, expected 55")
        if not re.search(r"<v>7</v>", xml) and "Date 29.09.2026" not in xml:
            problems.append(f"{label}: date label not written")
    return problems


def total_ratio_deviations(expected_xml: str, sst_xml: str) -> dict[str, tuple[str, str]]:
    """What the total row's ratio cells should be, versus the reference.

    The reference drags one sum across the whole total row, so its E and I add
    the line percentages together; the sheet writes ``C/B`` and ``F/H`` of the
    total row instead. Returns ``{ref: (value, formula)}`` for ``compare``.
    """
    total_row = (dr.DonorSheet.parse(expected_xml, dr.SharedStrings(sst_xml)).footer_start
                 + 2 + len(dr.FOOTER_ROWS))
    cells = cell_map(expected_xml, sst_xml)
    out: dict[str, tuple[str, str]] = {}
    for cell, numerator, denominator in dr.TOTAL_RATIO_CELLS:
        top = float(cells[f"{numerator}{total_row}"][0])
        bottom = float(cells[f"{denominator}{total_row}"][0])
        ref = f"{cell}{total_row}"
        if bottom:
            out[ref] = (repr(top / bottom),
                        f"{numerator}{total_row}/{denominator}{total_row}")
        else:
            out[ref] = ("0", "")
    return out


def main() -> int:
    template = BOOK.read_bytes()
    sources = {key: (FOLDER / name).read_bytes() for key, name in SOURCE_FOR_KEY.items()}

    with zipfile.ZipFile(io.BytesIO(template)) as zf:
        sst_xml = zf.read("xl/sharedStrings.xml").decode("utf-8")
        donor_xml = zf.read(sheet_part(zf, DONOR)).decode("utf-8")
        expected_xml = zf.read(sheet_part(zf, TARGET)).decode("utf-8")

    # The monthly footer of the day we are reproducing, read from the reference.
    footer = dr.read_monthly_footer(expected_xml, sst_xml)
    print("footer plans:", footer.plans)
    print("footer rtp:", footer.rtp_count, footer.rtp_value)

    # ── 1. exact reproduction, using the target sheet itself as the donor ────
    print("\n=== pass 1: self-donor (regenerate 28.09.2026 from itself) ===")
    diagnostics = dr.Diagnostics()
    built = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=dr.date(2026, 9, 28), sources=sources,
        footer=footer, donor_name=TARGET, diagnostics=diagnostics)
    print("ignored:", diagnostics.ignored, "new banks:", diagnostics.new_banks,
          "missing:", diagnostics.missing_sources)

    with zipfile.ZipFile(io.BytesIO(built)) as zf:
        names = zf.namelist()
        got_part = sheet_part(zf, TARGET)
        got_xml = zf.read(got_part).decode("utf-8")
        got_sst = zf.read("xl/sharedStrings.xml").decode("utf-8")
        print("new part:", got_part, "| calcChain removed:",
              "xl/calcChain.xml" not in names)
        rels = zf.read(f"xl/worksheets/_rels/{got_part.split('/')[-1]}.rels").decode("utf-8")
        print("logo rels:", len(re.findall(r"<Relationship ", rels)))
        for target_ in re.findall(r'Target="([^"]+)"', rels):
            resolved = "xl/" + target_.replace("../", "")
            if resolved not in names:
                print("  MISSING PART:", resolved)
        workbook = zf.read("xl/workbook.xml").decode("utf-8")
        order = re.findall(r'<sheet name="([^"]+)"', workbook)
        print("sheets:", len(order), "| position of target:", order.index(TARGET) + 1,
              "of", len(order))
        app = zf.read("docProps/app.xml").decode("utf-8")
        print("app titles:", len(re.findall(r"<vt:lpstr>", app)) - 1,
              "| count:",
              re.search(r"Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>(\d+)",
                        app).group(1))

    exact = compare(expected_xml, got_xml, got_sst,
                    formulas_expected=formula_map(template, TARGET),
                    formulas_got=formula_map(built, TARGET),
                    deviations=total_ratio_deviations(expected_xml, sst_xml))

    # ── 2. cross-donor: the real workflow, 27.09.2026 supplies the layout ───
    print("\n=== pass 2: 27.09.2026 donor, data compared by bank ===")
    diagnostics2 = dr.Diagnostics()
    built2 = dr.build_daily_sheet(
        template, sheet_name=TARGET, day=dr.date(2026, 9, 28), sources=sources,
        footer=footer, donor_name=DONOR, diagnostics=diagnostics2)
    print("ignored:", diagnostics2.ignored)
    print("new banks:", {k: v for k, v in diagnostics2.new_banks.items()})
    with zipfile.ZipFile(io.BytesIO(built2)) as zf:
        got2 = zf.read(sheet_part(zf, TARGET)).decode("utf-8")
        got2_sst = zf.read("xl/sharedStrings.xml").decode("utf-8")
    same = compare_data(expected_xml, got2, got2_sst)

    # ── 3. package validity and a day that does not exist yet ───────────────
    print("\n=== pass 3: package validity + appending a new day ===")
    problems = validate_package(built, "self-donor")
    problems += new_day_pass(template, sources, footer)
    if problems:
        print(f"\n{len(problems)} package problem(s):")
        for line in problems[:30]:
            print("  -", line)
    else:
        print("\npackage ok: every part well-formed, every relationship resolves")

    return 0 if (exact and same and not problems) else 1


def block_values(xml: str, sst_xml: str) -> dict[str, dict[str, list[float]]]:
    """``{bank: [numbers...]}`` per block, keyed by normalised bank name."""
    cells = cell_map(xml, sst_xml)
    out: dict[str, dict[str, list[float]]] = {}
    for block in dr.BLOCKS:
        bank_col = dr.col_letter(block.bank_col)
        first = dr.col_letter(block.first_col)
        last = dr.col_letter(block.last_col)
        rows: dict[str, list[float]] = {}
        number = dr.FIRST_DATA_ROW
        while True:
            bank = cells.get(f"{bank_col}{number}")
            if bank is None or bank[0] in ("", "Total"):
                break
            values = []
            for col in range(block.first_col, block.last_col + 1):
                cell = cells.get(f"{dr.col_letter(col)}{number}", ("0", ""))
                values.append(float(cell[0] or 0))
            rows[dr.normalize_bank(bank[0])] = values
            number += 1
        out[block.key] = rows
    return out


def compare_data(expected: str, got: str, sst_xml: str, tol: float = 1e-6) -> bool:
    exp, act = block_values(expected, sst_xml), block_values(got, sst_xml)
    problems: list[str] = []
    for key in exp:
        exp_rows, act_rows = exp[key], act.get(key, {})
        for bank in sorted(set(exp_rows) | set(act_rows)):
            if bank not in exp_rows:
                problems.append(f"{key}: extra bank {bank}")
            elif bank not in act_rows:
                problems.append(f"{key}: missing bank {bank}")
            else:
                for i, (e, a) in enumerate(zip(exp_rows[bank], act_rows[bank])):
                    field = dr.col_letter(dr.BLOCK_BY_KEY[key].first_col + i)
                    if abs(e - a) > tol * max(1.0, abs(e)):
                        problems.append(f"{key} {bank} {field}: {a} != {e}")
    if problems:
        print(f"\n{len(problems)} data difference(s):")
        for line in problems[:40]:
            print("  -", line)
        return False
    print("\nall four blocks match the reference values (bank order aside)")
    return True


def cell_map(xml: str, sst_xml: str) -> dict[str, tuple[str, str]]:
    """``A1 -> (value, formula)`` for every cell with content."""
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


def formula_map(book: bytes, sheet_name: str) -> dict[str, str]:
    """``{ref: formula}`` as Excel sees it (shared formulas expanded)."""
    wb = openpyxl.load_workbook(io.BytesIO(book), data_only=False)
    ws = wb[sheet_name]
    out = {}
    for row in ws.iter_rows():
        for cell in row:
            if isinstance(cell.value, str) and cell.value.startswith("="):
                out[cell.coordinate] = cell.value[1:]
    wb.close()
    return out


def styles_map(xml: str) -> dict[str, str]:
    out = {}
    for m in re.finditer(r'<c r="([A-Z]+\d+)"([^>]*?)(?:/>|>)', xml):
        found = re.search(r's="(\d+)"', m.group(2))
        out[m.group(1)] = found.group(1) if found else "0"
    return out


def merges_of(xml: str) -> list[str]:
    return re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', xml)


def compare(expected: str, got: str, sst_xml: str, formulas_expected=None,
            formulas_got=None, deviations=None, tol: float = 1e-6) -> bool:
    problems: list[str] = []
    formulas_expected = dict(formulas_expected or {})
    formulas_got = formulas_got or {}

    exp_cells, got_cells = cell_map(expected, sst_xml), cell_map(got, sst_xml)
    # Cells the sheet deliberately computes differently from the reference: the
    # total row's achievement and decline rate are ratios of the totals, while
    # the reference drags the row sum into them.
    for ref, (value, formula) in (deviations or {}).items():
        exp_cells[ref] = (value, formula)
        formulas_expected[ref] = formula
    for ref in sorted(set(exp_cells) | set(got_cells),
                      key=lambda r: (int(re.search(r"\d+", r).group()),
                                     dr.col_index(r.rstrip("0123456789")))):
        e, g = exp_cells.get(ref), got_cells.get(ref)
        if e is None:
            problems.append(f"{ref}: unexpected value {g!r}")
            continue
        if g is None:
            problems.append(f"{ref}: missing (expected {e!r})")
            continue
        want = formulas_expected.get(ref, e[1])
        have = formulas_got.get(ref, g[1])
        if want != have:
            problems.append(f"{ref}: formula {have!r} != expected {want!r}")
        try:
            if abs(float(e[0]) - float(g[0])) > tol * max(1.0, abs(float(e[0]))):
                problems.append(f"{ref}: value {g[0]!r} != expected {e[0]!r}")
        except (TypeError, ValueError):
            if e[0] != g[0]:
                problems.append(f"{ref}: value {g[0]!r} != expected {e[0]!r}")

    exp_styles, got_styles = styles_map(expected), styles_map(got)
    diff = [f"{ref} {got_styles.get(ref)} != {exp_styles[ref]}"
            for ref in sorted(exp_styles)
            if got_styles.get(ref) != exp_styles[ref]]
    if diff:
        problems.append(f"{len(diff)} style mismatch(es): " + "; ".join(diff[:12]))

    if set(merges_of(expected)) != set(merges_of(got)):
        only_expected = sorted(set(merges_of(expected)) - set(merges_of(got)))
        only_got = sorted(set(merges_of(got)) - set(merges_of(expected)))
        problems.append(f"merges differ: only in expected {only_expected}, "
                        f"only in generated {only_got}")

    exp_dims = re.search(r'<dimension ref="([^"]+)"', expected).group(1)
    got_dims = re.search(r'<dimension ref="([^"]+)"', got).group(1)
    if exp_dims != got_dims:
        problems.append(f"dimension {got_dims} != {exp_dims}")

    exp_rows = re.findall(r'<row r="(\d+)"([^>]*?)(?:/>|>)', expected)
    got_rows = re.findall(r'<row r="(\d+)"([^>]*?)(?:/>|>)', got)
    for (er, ea), (gr, ga) in zip(exp_rows, got_rows):
        if er != gr or ea != ga:
            problems.append(f"row {gr} attrs {ga!r} != expected {er} {ea!r}")

    if problems:
        print(f"\n{len(problems)} difference(s):")
        for line in problems[:60]:
            print("  -", line)
        return False
    print("\nidentical: values, formulas, styles, merges, dimension, row attrs")
    return True


if __name__ == "__main__":
    raise SystemExit(main())
