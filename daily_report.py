"""Daily "Successful Financial & Decline Transaction" compiled report sheet.

The compiled workbook (``September_Successful_Financial_and_Decline_Transaction_Report_28.xlsx``)
holds one sheet per day.  Each sheet carries four side-by-side blocks, a summary
footer and eight embedded OLE/EMF logos that ``openpyxl`` silently drops.

This module adds a day to such a workbook:

  * the eight source summaries are parsed and folded into the four blocks,
  * the new sheet is cloned from the donor day at the **OOXML** level, so styles,
    merges, column widths, formulas and the logo objects survive untouched,
  * the footer is recomputed from the block totals; the values that cannot be
    derived from the eight files (monthly plans, IPS RTP, incident notes) are
    supplied by the caller once per month.
"""

from __future__ import annotations

import difflib
import io
import re
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any, Iterable, Mapping, Sequence

# ---------------------------------------------------------------------------
# Source workbooks - one uploader each
# ---------------------------------------------------------------------------
SOURCE_KEYS: tuple[str, ...] = (
    "iss_success",
    "acq_success",
    "iss_decline",
    "acq_decline",
    "ips_success",
    "ips_decline",
    "qr_success",
    "qr_decline",
)

#: What the file is usually called, kept so the regression harness and the
#: sample data stay in step.  The screen never asks for these names: each
#: upload is recognised by its column headers, not its file name.
SOURCE_LABELS: dict[str, str] = {
    "iss_success": "Iss_Report Sucess.xlsx",
    "acq_success": "Acq_Report Sucess.xlsx",
    "iss_decline": "Iss_Report Decline.xlsx",
    "acq_decline": "Acq_Report Decline.xlsx",
    "ips_success": "IPS success for source and destination.xlsx",
    "ips_decline": "IPS Declined for source and destination.xlsx",
    "qr_success": "QR success for source and destination.xlsx",
    "qr_decline": "QR Declined for source and destination.xlsx",
}

#: The report each upload has to contain, in the words the screen shows.
SOURCE_REPORT_NAMES: dict[str, str] = {
    "iss_success": "Issuer card - successful transactions",
    "acq_success": "Acquirer card - successful transactions",
    "iss_decline": "Issuer card - declined transactions",
    "acq_decline": "Acquirer card - declined transactions",
    "ips_success": "IPS - successful interbank transfers",
    "ips_decline": "IPS - declined interbank transfers",
    "qr_success": "QR - successful interbank payments",
    "qr_decline": "QR - declined interbank payments",
}

#: Grouped the way the screen presents the uploads, each with the columns that
#: identify the report.  A file is only read as the kind of report it is if
#: these column headers are present.
SOURCE_GROUPS: tuple[tuple[str, tuple[str, ...], str], ...] = (
    ("Card reports",
     ("iss_success", "acq_success", "iss_decline", "acq_decline"),
     "ISS_BANKS or ACQ_BANKS, CASH_WITHDRAWAL, AMOUNT_CW, BALANCE_INQUIRY, "
     "PURCHASE, AMOUNT_POS, STA_REQUEST, EPG, EPG-C"),
    ("IPS and QR interbank summaries",
     ("ips_success", "ips_decline", "qr_success", "qr_decline"),
     "BANK_NAME, ISSUER_TXN_COUNT, ISSUER_TOTAL_AMOUNT, ACQUIRER_TXN_COUNT, "
     "ACQUIRER_TOTAL_AMOUNT"),
)

#: Institutions that are technical noise rather than reportable banks.
IGNORED_BANKS = {"systeminstitution"}


# ---------------------------------------------------------------------------
# Sheet geometry (1-based column indexes, as in the reference workbook)
# ---------------------------------------------------------------------------
@dataclass(frozen=True)
class BlockSpec:
    """One of the four side-by-side blocks of a daily sheet."""

    key: str
    bank_col: int
    first_col: int
    last_col: int
    fields: tuple[str, ...]


_INTEROP_FIELDS = (
    "ips_iss_cnt", "ips_iss_amt", "ips_acq_cnt", "ips_acq_amt",
    "qr_iss_cnt", "qr_iss_amt", "qr_acq_cnt", "qr_acq_amt",
)

_CARD_FIELDS = (
    "cw", "cw_amt", "acq_cw", "acq_cw_amt",
    "bal", "acq_bal",
    "pur", "pur_amt", "acq_pur", "acq_pur_amt",
    "sta", "acq_sta",
    "epg", "epg_c", "acq_epg", "acq_epg_c",
)

BLOCKS: tuple[BlockSpec, ...] = (
    BlockSpec("success_interbank", 1, 2, 9, _INTEROP_FIELDS),      # A:I
    BlockSpec("success_card", 11, 12, 27, _CARD_FIELDS),          # K:AA
    BlockSpec("decline_interbank", 29, 30, 37, _INTEROP_FIELDS),  # AC:AK
    BlockSpec("decline_card", 39, 40, 55, _CARD_FIELDS),          # AM:BC
)

BLOCK_BY_KEY = {b.key: b for b in BLOCKS}

FIRST_DATA_ROW = 5
_LAST_CONTENT_COL = 55          # BC


@dataclass(frozen=True)
class FooterRow:
    """One summary line of the footer, expressed against the block totals.

    ``{t1}``-``{t4}`` stand for the four block total-row numbers.
    """

    label: str
    success: tuple[str, ...]
    success_value: tuple[str, ...] = ()
    declined: tuple[str, ...] = ()
    declined_value: tuple[str, ...] = ()
    rtp: bool = False


FOOTER_ROWS: tuple[FooterRow, ...] = (
    FooterRow("Cash Withdrawal", ("L{t2}",), ("M{t2}",), ("AN{t4}",), ("AO{t4}",)),
    FooterRow("Balance Inquiry & Mini Statement", ("P{t2}", "V{t2}"), (),
              ("AR{t4}", "AX{t4}")),
    FooterRow("POS Purchase", ("R{t2}",), ("S{t2}",), ("AT{t4}",), ("AU{t4}",)),
    FooterRow("IPS P2P", ("B{t1}",), ("C{t1}",), ("AD{t3}",), ("AE{t3}",)),
    FooterRow("NPG-card and NPG-online",
              ("X{t2}", "Y{t2}", "Z{t2}", "AA{t2}"), (),
              ("AZ{t4}", "BA{t4}", "BB{t4}", "BC{t4}")),
    FooterRow("IPS QR", ("F{t1}",), ("G{t1}",), ("AH{t3}",), ("AI{t3}",)),
    FooterRow("IPS RTP", (), (), (), (), rtp=True),
)

#: Footer lines aggregated by the "Total interbank" row (NPG and RTP excluded).
FOOTER_TOTAL_INCLUDED = (0, 1, 2, 3, 5)

#: Total-row cells that are ratios of the row's own figures rather than sums of
#: the lines above, as ``(cell, numerator, denominator)``.  The achievement is
#: the total taken over the total plan, the decline rate over the total
#: combined.  The reference workbook drags one ``B57+B58+...`` across the whole
#: total row, so it *adds* these two percentages together; recomputing them is
#: the one place this sheet deliberately departs from the reference.
TOTAL_RATIO_CELLS: tuple[tuple[str, str, str], ...] = (
    ("E", "C", "B"),
    ("I", "F", "H"),
)

#: What a non-numeric monthly plan looks like in the reference workbook.
PLAN_PLACEHOLDER = "                             -  "


# ---------------------------------------------------------------------------
# Bank-name normalisation
# ---------------------------------------------------------------------------
def normalize_bank(name: Any) -> str:
    """Fold a bank name to a comparison key (case/punctuation insensitive)."""
    return re.sub(r"[^a-z0-9]", "", ("" if name is None else str(name)).lower())


class BankMatcher:
    """Maps incoming file spellings onto the donor sheet's bank names.

    The donor workbook is the naming authority, so its bank rows are reused
    verbatim and incoming spellings are normalised onto them.  Only banks that
    are genuinely new get appended.
    """

    def __init__(self, donor_names: Iterable[str]) -> None:
        self._exact: dict[str, str] = {}
        for name in donor_names:
            key = normalize_bank(name)
            if key and key not in self._exact:
                self._exact[key] = name
        self.ignored: list[str] = []
        self.new_banks: list[str] = []

    def resolve_all(self, names: Iterable[str]) -> dict[str, str | None]:
        """``{incoming name: donor name}``; ``None`` for ignored institutions.

        Must be called before any values are accumulated: a name can only be
        folded onto the donor list once every spelling has been seen.
        """
        mapping: dict[str, str | None] = {}
        for name in names:
            key = normalize_bank(name)
            if not key or key in IGNORED_BANKS:
                if name and name not in self.ignored:
                    self.ignored.append(name)
                mapping[name] = None
                continue
            hit = self._exact.get(key)
            if hit is None:
                close = difflib.get_close_matches(key, list(self._exact), n=1, cutoff=0.94)
                if close:
                    self._exact[key] = hit = self._exact[close[0]]
                else:
                    hit = name
                    if name not in self.new_banks:
                        self.new_banks.append(name)
            mapping[name] = hit
        return mapping


# ---------------------------------------------------------------------------
# Parsing the eight source workbooks
# ---------------------------------------------------------------------------
def _number(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, bool):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).replace(",", "").strip()
    if not text:
        return 0.0
    try:
        return float(text)
    except ValueError:
        return 0.0


def _sheet_rows(data: bytes) -> list[list[Any]]:
    import openpyxl

    book = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
    try:
        return [list(row) for row in book.active.iter_rows(values_only=True)]
    finally:
        book.close()


def _header_index(header: Sequence[Any], wanted: str) -> int | None:
    target = normalize_bank(wanted)
    for i, cell in enumerate(header):
        if normalize_bank(cell) == target:
            return i
    return None


def parse_interop(data: bytes) -> list[tuple[str, list[float]]]:
    """Parse an ``... success/declined for source and destination`` workbook.

    Returns ``(bank, [issuer count, issuer amount, acquirer count, acquirer amount])``.
    """
    rows = _sheet_rows(data)
    if not rows:
        return []
    header = rows[0]
    names = ("BANK_NAME", "ISSUER_TXN_COUNT", "ISSUER_TOTAL_AMOUNT",
             "ACQUIRER_TXN_COUNT", "ACQUIRER_TOTAL_AMOUNT")
    idx = {name: _header_index(header, name) for name in names}
    missing = [n for n, v in idx.items() if v is None]
    if missing:
        raise ValueError(f"missing column(s): {', '.join(missing)}")
    out: list[tuple[str, list[float]]] = []
    for row in rows[1:]:
        bank = row[idx["BANK_NAME"]]
        if bank is None or not str(bank).strip():
            continue
        out.append((str(bank).strip(), [_number(row[idx[n]]) for n in names[1:]]))
    return out


_CARD_COLUMN_ORDER = (
    "CASH_WITHDRAWAL", "AMOUNT_CW", "BALANCE_INQUIRY", "PURCHASE",
    "AMOUNT_POS", "STA_REQUEST", "EPG", "EPG-C",
)


def parse_card(data: bytes) -> list[tuple[str, list[float]]]:
    """Parse an ``Iss_Report``/``Acq_Report`` success or decline workbook.

    Returns ``(bank, [cw, cw amount, balance inquiry, purchase, purchase
    amount, mini statement, EPG, EPG-C])``.
    """
    rows = _sheet_rows(data)
    if not rows:
        return []
    header = rows[0]
    bank_at = _header_index(header, "ISS_BANKS")
    if bank_at is None:
        bank_at = _header_index(header, "ACQ_BANKS")
    if bank_at is None:
        bank_at = 1
    idx = {name: _header_index(header, name) for name in _CARD_COLUMN_ORDER}
    missing = [n for n, v in idx.items() if v is None]
    if missing:
        raise ValueError(f"missing column(s): {', '.join(missing)}")
    out: list[tuple[str, list[float]]] = []
    for row in rows[1:]:
        bank = row[bank_at]
        if bank is None or not str(bank).strip():
            continue
        out.append((str(bank).strip(),
                    [_number(row[idx[n]]) for n in _CARD_COLUMN_ORDER]))
    return out


# ---------------------------------------------------------------------------
# Monthly (non-derivable) footer values
# ---------------------------------------------------------------------------
@dataclass
class MonthlyFooter:
    """Everything on the footer that cannot be derived from the eight files."""

    plans: list[Any] = field(default_factory=list)
    rtp_count: float = 0.0
    rtp_value: float = 0.0
    card_note: str = ""
    interop_note: str = ""

    def with_defaults(self) -> "MonthlyFooter":
        """Pad ``plans`` so it lines up with the footer rows."""
        plans = list(self.plans[:len(FOOTER_ROWS)])
        plans += [PLAN_PLACEHOLDER] * (len(FOOTER_ROWS) - len(plans))
        self.plans = plans
        return self


# ---------------------------------------------------------------------------
# OOXML helpers
# ---------------------------------------------------------------------------
def col_letter(index: int) -> str:
    """1-based column index -> spreadsheet column letters."""
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def col_index(letters: str) -> int:
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch) - 64)
    return value


def _escape(text: Any) -> str:
    return str(text).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _num(value: Any) -> str:
    """Render a number the way the reference workbook stores it."""
    number = float(value)
    rounded = round(number)
    if abs(number - rounded) < 1e-9:
        return str(int(rounded))
    return repr(number)


_SI_RE = re.compile(r"<si>(.*?)</si>", re.S)
_T_RE = re.compile(r"<t[^>]*>(.*?)</t>", re.S)
_V_RE = re.compile(r"<v>(.*?)</v>", re.S)


def _si_text(body: str) -> str:
    return "".join(_T_RE.findall(body))


class SharedStrings:
    """Append-only view over ``xl/sharedStrings.xml``."""

    def __init__(self, xml: str) -> None:
        self._xml = xml
        self._items = [_si_text(b) for b in _SI_RE.findall(xml)]
        self._index: dict[str, int] = {}
        for i, text in enumerate(self._items):
            self._index.setdefault(text, i)
        self._added: list[str] = []
        self._base_count = int(re.search(r'\scount="(\d+)"', xml).group(1))
        self._unique = int(re.search(r'\suniqueCount="(\d+)"', xml).group(1))

    def index(self, text: Any) -> int:
        key = str(text)
        found = self._index.get(key)
        if found is not None:
            return found
        pos = self._unique + len(self._added)
        self._index[key] = pos
        self._added.append(key)
        return pos

    def text(self, index: int) -> str:
        if index < self._unique:
            return self._items[index]
        offset = index - self._unique
        return self._added[offset] if offset < len(self._added) else ""

    def render(self) -> bytes:
        if not self._added:
            return self._xml.encode("utf-8")
        items = "".join(
            f'<si><t xml:space="preserve">{_escape(t)}</t></si>' for t in self._added
        )
        unique = self._unique + len(self._added)
        count = self._base_count + len(self._added)
        head = re.sub(r'\scount="\d+"', f' count="{count}"', self._xml, count=1)
        head = re.sub(r'\suniqueCount="\d+"', f' uniqueCount="{unique}"', head, count=1)
        return (head.replace("</sst>", items + "</sst>")).encode("utf-8")


_ROW_RE = re.compile(r'<row r="(\d+)"([^>]*?)(?:/>|>(.*?)</row>)', re.S)
_CELL_RE = re.compile(r'<c r="([A-Z]+)(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
_STYLE_RE = re.compile(r's="(\d+)"')


def _parse_rows(xml: str) -> dict[int, tuple[str, str]]:
    return {int(m.group(1)): (m.group(2) or "", m.group(3) or "")
            for m in _ROW_RE.finditer(xml)}


def _majority(values: Iterable[str]) -> str | None:
    tally: dict[str, int] = {}
    for value in values:
        tally[value] = tally.get(value, 0) + 1
    if not tally:
        return None
    return max(tally, key=lambda k: (tally[k], -int(k)))


def _parse_cells(rows: Mapping[int, tuple[str, str]],
                 sst: SharedStrings) -> dict[int, dict[str, tuple[str, str | None, str | None]]]:
    """``{row: {column: (style, text, raw <v>)}}`` for every cell of a sheet."""
    cells: dict[int, dict[str, tuple[str, str | None, str | None]]] = {}
    for number, (_, body) in rows.items():
        per_row: dict[str, tuple[str, str | None, str | None]] = {}
        for m in _CELL_RE.finditer(body):
            col, attrs, inner = m.group(1), m.group(3), m.group(4) or ""
            text = _T_RE.search(inner)
            value = _V_RE.search(inner)
            content: str | None
            if text is not None:
                content = text.group(1)
            elif value is not None and 't="s"' in attrs:
                content = sst.text(int(value.group(1)))
            else:
                content = None
            found = _STYLE_RE.search(attrs)
            per_row[col] = (found.group(1) if found else "0", content,
                            value.group(1) if value is not None else None)
        cells[number] = per_row
    return cells


_F_RE = re.compile(r"<f([^>]*?)(?:/>|>(.*?)</f>)", re.S)


def _parse_formulas(rows: Mapping[int, tuple[str, str]]) -> dict[int, dict[str, str]]:
    """``{row: {column: formula}}`` with shared formulas resolved.

    Excel stores a dragged formula once and leaves ``<f t="shared" si="n"/>``
    followers behind, so the master is looked up by index.
    """
    raw: dict[int, dict[str, tuple[str, str | None]]] = {}
    masters: dict[str, str] = {}
    for number, (_, body) in rows.items():
        per_row: dict[str, tuple[str, str | None]] = {}
        for m in _CELL_RE.finditer(body):
            found = _F_RE.search(m.group(4) or "")
            if not found:
                continue
            attrs, text = found.group(1) or "", found.group(2)
            index = re.search(r'si="(\d+)"', attrs)
            key = index.group(1) if index else None
            per_row[m.group(1)] = (text or "", key)
            if text and key:
                masters[key] = text
        if per_row:
            raw[number] = per_row
    resolved: dict[int, dict[str, str]] = {}
    for number, per_row in raw.items():
        resolved[number] = {
            col: (text or masters.get(key or "", ""))
            for col, (text, key) in per_row.items()
        }
    return resolved


# ---------------------------------------------------------------------------
# Donor sheet
# ---------------------------------------------------------------------------
@dataclass
class DonorSheet:
    """Everything the generator borrows from the donor day's sheet."""

    xml: str
    rows: dict[int, tuple[str, str]]
    cells: dict[int, dict[str, tuple[str, str | None, str | None]]]
    total_rows: dict[str, int]
    data_style: dict[str, dict[str, str]]
    total_style: dict[str, dict[str, str]]
    total_sums: dict[str, set[str]]
    idle_style: dict[str, dict[str, str]]
    banks: dict[str, list[str]]
    stray: dict[int, dict[str, str]]
    filler: dict[str, str]
    footer_start: int
    date_attrs: str
    footer_attrs: dict[int, str]
    footer_style: dict[int, dict[str, str]]
    footer_texts: dict[int, dict[str, str]]
    footer_merges: list[str]
    fixed_merges: list[str]
    max_col: int

    @classmethod
    def parse(cls, xml: str, sst: SharedStrings) -> "DonorSheet":
        rows = _parse_rows(xml)
        if not rows:
            raise ValueError("donor sheet contains no rows")

        cells = _parse_cells(rows, sst)
        formulas = _parse_formulas(rows)
        last_row = max(rows)
        total_rows: dict[str, int] = {}
        for block in BLOCKS:
            letter = col_letter(block.bank_col)
            total = FIRST_DATA_ROW
            while total <= last_row:
                cell = cells.get(total, {}).get(letter)
                if cell is None or cell[1] is None:
                    break
                if str(cell[1]).strip().lower() == "total":
                    break
                total += 1
            total_rows[block.key] = total

        data_style: dict[str, dict[str, str]] = {}
        total_style: dict[str, dict[str, str]] = {}
        total_sums: dict[str, set[str]] = {}
        idle_style: dict[str, dict[str, str]] = {}
        for block in BLOCKS:
            total = total_rows[block.key]
            columns = [col_letter(c) for c in range(block.bank_col, block.last_col + 1)]
            data: dict[str, str] = {}
            totals: dict[str, str] = {}
            idle: dict[str, str] = {}
            for col in columns:
                seen = [cells[r][col][0] for r in range(FIRST_DATA_ROW, total)
                        if col in cells.get(r, {})]
                style = _majority(seen)
                if style:
                    data[col] = style
                if col in cells.get(total, {}):
                    totals[col] = cells[total][col][0]
                seen = [cells[r][col][0] for r in range(total + 1, last_row + 1)
                        if col in cells.get(r, {})]
                style = _majority(seen)
                if style:
                    idle[col] = style
            data_style[block.key] = data
            total_style[block.key] = totals
            idle_style[block.key] = idle
            # The reference only drags the SUM() across some of the columns.
            total_sums[block.key] = {
                col for col, formula in formulas.get(total, {}).items()
                if formula.strip().upper().startswith("SUM(")}

        # Background/spacer cells (J, AB, AL and the styled cells beyond the
        # blocks) are cloned verbatim: the reference only styles some of these
        # on some rows, so a per-column majority would not reproduce it.
        block_columns = {col_letter(c) for b in BLOCKS
                         for c in range(b.bank_col, b.last_col + 1)}
        stray: dict[int, dict[str, str]] = {}
        for number, per_row in cells.items():
            if not FIRST_DATA_ROW <= number <= last_row:
                continue
            extra = {c: s for c, (s, _, _) in per_row.items() if c not in block_columns}
            if extra:
                stray[number] = extra
        filler: dict[str, str] = {}
        for extra in stray.values():
            for col, style in extra.items():
                filler.setdefault(col, style)

        banks: dict[str, list[str]] = {}
        for block in BLOCKS:
            letter = col_letter(block.bank_col)
            names = []
            for r in range(FIRST_DATA_ROW, total_rows[block.key]):
                cell = cells.get(r, {}).get(letter)
                if cell is not None and cell[1] is not None and str(cell[1]).strip():
                    names.append(str(cell[1]).strip())
            banks[block.key] = names

        date_row = max(total_rows.values()) + 1
        if date_row not in rows:
            raise ValueError("donor sheet has no date row")

        footer_attrs: dict[int, str] = {}
        footer_style: dict[int, dict[str, str]] = {}
        footer_texts: dict[int, dict[str, str]] = {}
        for offset in range(0, len(FOOTER_ROWS) + 3):
            number = date_row + offset
            if number not in rows:
                continue
            footer_attrs[offset] = rows[number][0]
            footer_style[offset] = {c: s for c, (s, _, _) in cells[number].items()}
            if offset:
                footer_texts[offset] = {c: t for c, (_, t, _) in cells[number].items()
                                        if t is not None}

        footer_merges: list[str] = []
        fixed_merges: list[str] = []
        for ref in re.findall(r'<mergeCell ref="([A-Z]+\d+:[A-Z]+\d+)"/>', xml):
            start = int(re.search(r"(\d+)", ref).group(1))
            (footer_merges if start >= date_row else fixed_merges).append(ref)

        max_col = max((col_index(c) for per_row in cells.values() for c in per_row),
                      default=_LAST_CONTENT_COL)
        return cls(
            xml=xml, rows=rows, cells=cells, total_rows=total_rows,
            data_style=data_style, total_style=total_style, total_sums=total_sums,
            idle_style=idle_style,
            banks=banks, stray=stray, filler=filler, footer_start=date_row,
            date_attrs=rows[date_row][0], footer_attrs=footer_attrs,
            footer_style=footer_style, footer_texts=footer_texts,
            footer_merges=footer_merges, fixed_merges=fixed_merges, max_col=max_col,
        )

    def plan_text(self, index: int) -> str:
        """The donor's monthly-plan text for footer line ``index``."""
        return self.footer_texts.get(2 + index, {}).get("B", PLAN_PLACEHOLDER)

    def data_row_style(self, block_key: str, row: int, col: str) -> str:
        """The style a data cell gets on ``row``.

        The reference sheets carry a few per-row style artifacts, so the donor
        row wins whenever it is itself a data row of the same block; the
        per-column majority covers rows the donor used for something else.
        """
        total = self.total_rows[block_key]
        if FIRST_DATA_ROW <= row < total:
            cell = self.cells.get(row, {}).get(col)
            if cell is not None:
                return cell[0]
        return self.data_style[block_key].get(col, "0")


def _all_columns(cells: Mapping[int, Mapping[str, Any]]) -> list[str]:
    out: set[str] = set()
    for per_row in cells.values():
        out.update(per_row)
    return sorted(out, key=col_index)


# ---------------------------------------------------------------------------
# Cell emission
# ---------------------------------------------------------------------------
@dataclass
class Value:
    """A cell to write as ``formula`` with a cached numeric ``value``."""

    value: float
    formula: str


def _cell_xml(ref: str, style: str, value: Any, sst: SharedStrings | None) -> str:
    style_attr = f' s="{style}"' if style and style != "0" else ""
    if value is None:
        return f'<c r="{ref}"{style_attr}/>'
    if isinstance(value, Value):
        return (f'<c r="{ref}"{style_attr}><f>{_escape(value.formula)}</f>'
                f'<v>{_num(value.value)}</v></c>')
    if isinstance(value, str):
        if sst is None:
            return (f'<c r="{ref}"{style_attr} t="inlineStr">'
                    f'<is><t xml:space="preserve">{_escape(value)}</t></is></c>')
        return f'<c r="{ref}"{style_attr} t="s"><v>{sst.index(value)}</v></c>'
    return f'<c r="{ref}"{style_attr}><v>{_num(value)}</v></c>'


def _row_xml(number: int, attrs: str, cells: Mapping[int, tuple[str, Any]],
             sst: SharedStrings | None) -> str:
    body = "".join(_cell_xml(f"{col_letter(col)}{number}", style, value, sst)
                   for col, (style, value) in sorted(cells.items()))
    if not body:
        return f'<row r="{number}"{attrs}/>'
    return f'<row r="{number}"{attrs}>{body}</row>'


# ---------------------------------------------------------------------------
# Block assembly
# ---------------------------------------------------------------------------
_CARD_ISS_SLOTS = (0, 1, 4, 6, 7, 10, 12, 13)
_CARD_ACQ_SLOTS = (2, 3, 5, 8, 9, 11, 14, 15)
_INTERBANK_IPS_SLOTS = (0, 1, 2, 3)
_INTERBANK_QR_SLOTS = (4, 5, 6, 7)


@dataclass
class Diagnostics:
    ignored: list[str] = field(default_factory=list)
    new_banks: dict[str, list[str]] = field(default_factory=dict)
    missing_sources: list[str] = field(default_factory=list)
    skipped_rows: list[str] = field(default_factory=list)


def _accumulate(target: dict[str, list[float]], bank: str,
                values: Sequence[float], slots: Sequence[int]) -> None:
    row = target.setdefault(bank, [0.0] * 16)
    for value, slot in zip(values, slots):
        row[slot] += value


def _block_of(key: str) -> tuple[str, tuple[int, ...]]:
    """Which block a source workbook feeds, and into which slots."""
    if key.startswith(("ips_", "qr_")):
        block = "success_interbank" if key.endswith("success") else "decline_interbank"
        slots = _INTERBANK_IPS_SLOTS if key.startswith("ips_") else _INTERBANK_QR_SLOTS
    else:
        block = "success_card" if key.endswith("success") else "decline_card"
        slots = _CARD_ISS_SLOTS if key.startswith("iss_") else _CARD_ACQ_SLOTS
    return block, slots


def build_blocks(sources: Mapping[str, bytes], donor: DonorSheet,
                 diagnostics: Diagnostics | None = None
                 ) -> dict[str, list[tuple[str, list[float]]]]:
    """Fold the eight source workbooks into per-block ``(bank, values)`` rows."""
    diagnostics = diagnostics or Diagnostics()
    parsed: dict[str, list[tuple[str, list[float]]]] = {}
    for key in SOURCE_KEYS:
        data = sources.get(key)
        if not data:
            diagnostics.missing_sources.append(key)
            continue
        try:
            parsed[key] = (parse_interop(data) if key.startswith(("ips_", "qr_"))
                           else parse_card(data))
        except ValueError as exc:
            raise ValueError(f"{SOURCE_LABELS[key]}: {exc}") from exc

    raw: dict[str, dict[str, list[float]]] = {b.key: {} for b in BLOCKS}
    matchers = {b.key: BankMatcher(donor.banks[b.key]) for b in BLOCKS}
    mapping: dict[str, dict[str, str | None]] = {}

    # Names first, so that a spelling folded onto the donor is known before any
    # value is accumulated under it.
    for block in BLOCKS:
        names: list[str] = []
        for key in SOURCE_KEYS:
            if key in parsed and _block_of(key)[0] == block.key:
                names.extend(bank for bank, _ in parsed[key])
        mapping[block.key] = matchers[block.key].resolve_all(names)

    for key, rows in parsed.items():
        block_key, slots = _block_of(key)
        for bank, values in rows:
            display = mapping[block_key].get(bank)
            if display is None:
                continue
            _accumulate(raw[block_key], display, values, slots)

    out: dict[str, list[tuple[str, list[float]]]] = {}
    for block in BLOCKS:
        matcher = matchers[block.key]
        ordered = [b for b in donor.banks[block.key] if b in raw[block.key]]
        for name in matcher.new_banks:
            if name in raw[block.key] and name not in ordered:
                ordered.append(name)
        if matcher.new_banks:
            diagnostics.new_banks[block.key] = list(matcher.new_banks)
        for name in matcher.ignored:
            if name not in diagnostics.ignored:
                diagnostics.ignored.append(name)
        out[block.key] = [(name, raw[block.key][name]) for name in ordered]
    return out


# ---------------------------------------------------------------------------
# Footer
# ---------------------------------------------------------------------------
_TOTAL_KEYS = ("t1", "t2", "t3", "t4")
_TOTAL_TO_BLOCK = ("success_interbank", "success_card",
                   "decline_interbank", "decline_card")


def _footer_value(refs: Sequence[str], totals: Mapping[str, Mapping[str, float]]
                  ) -> tuple[Any, float]:
    """Return ``(cell, number)`` for a footer value column.

    A single reference is written as a literal (that is what the reference
    workbook does); several references become an ``a+b`` formula, but only
    when they add up to something.
    """
    if not refs:
        return 0.0, 0.0
    values: list[float] = []
    for ref in refs:
        col, token = re.match(r"([A-Z]+)\{(t\d)\}", ref).groups()
        block = _TOTAL_TO_BLOCK[_TOTAL_KEYS.index(token)]
        values.append(totals[block].get(col, 0.0))
    total = 0.0
    for value in values:
        total += value
    if len(values) == 1 or not total:
        return total, total
    return Value(total, "+".join(_num(v) for v in values)), total


def _build_footer_grid(totals: Mapping[str, Mapping[str, float]], footer: MonthlyFooter,
                      first_row: int) -> list[dict[str, Any]]:
    """The seven category rows, fully resolved (cells plus numbers for sums)."""
    footer = footer.with_defaults()
    grid: list[dict[str, Any]] = []
    for index, spec in enumerate(FOOTER_ROWS):
        number = first_row + index
        plan = footer.plans[index]
        plan_num = float(plan) if isinstance(plan, (int, float)) and not isinstance(plan, bool) else None
        if spec.rtp:
            success, success_num = footer.rtp_count, float(footer.rtp_count)
            value, value_num = footer.rtp_value, float(footer.rtp_value)
            declined, declined_num = 0.0, 0.0
            declined_value, declined_value_num = 0.0, 0.0
        else:
            success, success_num = _footer_value(spec.success, totals)
            value, value_num = _footer_value(spec.success_value, totals)
            declined, declined_num = _footer_value(spec.declined, totals)
            declined_value, declined_value_num = _footer_value(spec.declined_value, totals)
        combined, combined_num = (Value(declined_num + success_num,
                                       f"+F{number}+C{number}"),
                                 declined_num + success_num)
        # The reference only fills the achievement when the line has a value
        # to compare against the plan (Balance Inquiry carries a count but no
        # value, and its achievement stays empty).
        if plan_num and value_num:
            achievement, achievement_num = (Value(success_num / plan_num,
                                                  f"C{number}/B{number}"),
                                           success_num / plan_num)
        else:
            achievement, achievement_num = 0.0, 0.0
        if declined_num:
            declined_pct, pct_num = (Value(declined_num / combined_num,
                                           f"F{number}/H{number}"),
                                     declined_num / combined_num)
        else:
            declined_pct, pct_num = 0.0, 0.0
        grid.append({
            "row": number, "label": spec.label, "plan": plan, "plan_num": plan_num or 0.0,
            "C": success, "D": value, "E": achievement, "F": declined, "G": declined_value,
            "H": combined, "I": declined_pct,
            "nums": {"B": plan_num or 0.0, "C": success_num, "D": value_num,
                     "E": achievement_num, "F": declined_num, "G": declined_value_num,
                     "H": combined_num, "I": pct_num},
        })
    return grid


# ---------------------------------------------------------------------------
# New sheet XML
# ---------------------------------------------------------------------------
def _render_sheet(donor: DonorSheet,
                  blocks: Mapping[str, list[tuple[str, list[float]]]],
                  footer: MonthlyFooter, day: date,
                  sst: SharedStrings | None) -> str:
    total_rows = {key: FIRST_DATA_ROW + len(rows) for key, rows in blocks.items()}
    body_end = max(total_rows.values())

    totals: dict[str, dict[str, float]] = {}
    for block in BLOCKS:
        sums = [0.0] * len(block.fields)
        for _, values in blocks[block.key]:
            for i in range(len(block.fields)):
                sums[i] += values[i]
        totals[block.key] = {col_letter(block.first_col + i): sums[i]
                             for i in range(len(block.fields))}

    rows_xml: list[str] = []
    for number in range(1, FIRST_DATA_ROW):
        if number in donor.rows:
            rows_xml.append(f'<row r="{number}"{donor.rows[number][0]}>'
                            f'{donor.rows[number][1]}</row>')

    # The block frames do not stop at a block's own total row: the reference
    # keeps them running down to the last data row, so the donor decides which
    # columns exist on which row.
    donor_last = max(donor.total_rows.values())
    for number in range(FIRST_DATA_ROW, body_end + 1):
        src = number if number <= donor_last else donor_last
        template = donor.cells.get(src, {})
        cells: dict[int, tuple[str, Any]] = {}
        for block in BLOCKS:
            total = total_rows[block.key]
            columns = range(block.bank_col, block.last_col + 1)
            if number < total:
                for col in columns:
                    letter = col_letter(col)
                    cells[col] = (donor.data_row_style(block.key, src, letter), None)
                name, values = blocks[block.key][number - FIRST_DATA_ROW]
                cells[block.bank_col] = (
                    donor.data_row_style(block.key, src,
                                         col_letter(block.bank_col)), name)
                for i, col in enumerate(range(block.first_col, block.last_col + 1)):
                    cells[col] = (donor.data_row_style(block.key, src, col_letter(col)),
                                  values[i])
            elif number == total:
                for col in columns:
                    letter = col_letter(col)
                    cells[col] = (donor.total_style[block.key].get(letter, "0"), None)
                last = total - 1
                cells[block.bank_col] = (
                    donor.total_style[block.key][col_letter(block.bank_col)], "Total")
                summed = donor.total_sums.get(block.key, set())
                for col in range(block.first_col, block.last_col + 1):
                    letter = col_letter(col)
                    value = totals[block.key][letter]
                    if letter in summed:
                        cells[col] = (donor.total_style[block.key][letter],
                                      Value(value, f"SUM({letter}{FIRST_DATA_ROW}:{letter}{last})"))
                    else:
                        cells[col] = (donor.total_style[block.key][letter], value)
            else:
                for col in columns:
                    letter = col_letter(col)
                    if letter in template:
                        cells[col] = (template[letter][0], None)
        # Cloned exactly as in the donor when the row exists there, otherwise
        # the donor's usual pattern (covers rows added by newly appended banks).
        for letter, style in donor.stray.get(src, donor.filler).items():
            cells.setdefault(col_index(letter), (style, None))
        rows_xml.append(_row_xml(number, donor.rows.get(src, ("", ""))[0], cells, sst))

    # ── footer ──────────────────────────────────────────────────────────────
    date_row = body_end + 1
    head_row = date_row + 1
    first_footer_row = head_row + 1
    total_row = first_footer_row + len(FOOTER_ROWS)

    date_cells: dict[int, tuple[str, Any]] = {
        col_index(col): (style, None) for col, style in donor.footer_style.get(0, {}).items()
    }
    date_cells[1] = (donor.footer_style.get(0, {}).get("A", "0"), f"Date {day:%d.%m.%Y}")
    rows_xml.append(_row_xml(date_row, donor.date_attrs, date_cells, sst))

    head_text = donor.footer_texts.get(1, {})
    head_cells = {col_index(col): (style, head_text.get(col))
                  for col, style in donor.footer_style.get(1, {}).items()}
    rows_xml.append(_row_xml(head_row, donor.footer_attrs.get(1, ""), head_cells, sst))

    grid = _build_footer_grid(totals, footer, first_footer_row)
    for offset, entry in enumerate(grid, start=2):
        number = entry["row"]
        styles = donor.footer_style.get(offset, {})
        cells = {col_index(col): (style, None) for col, style in styles.items()}
        cells[1] = (styles.get("A", "0"), entry["label"])
        cells[2] = (styles.get("B", "0"), entry["plan"])
        for col, key in ((3, "C"), (4, "D"), (5, "E"), (6, "F"),
                         (7, "G"), (8, "H"), (9, "I")):
            cells[col] = (styles.get(key, "0"), entry[key])
        if offset == 2:
            notes = donor.footer_texts.get(2, {})
            card_note = footer.card_note or notes.get("J", "")
            if card_note:
                cells[10] = (styles.get("J", "0"), card_note)
            interop_note = footer.interop_note or notes.get("M", "")
            if interop_note:
                cells[13] = (styles.get("M", "0"), interop_note)
        rows_xml.append(_row_xml(number, donor.footer_attrs.get(offset, ""), cells, sst))

    total_offset = 2 + len(FOOTER_ROWS)
    total_styles = donor.footer_style.get(total_offset, {})
    picked = [first_footer_row + i for i in FOOTER_TOTAL_INCLUDED]
    cells = {col_index(col): (style, None) for col, style in total_styles.items()}
    label = (donor.footer_texts.get(total_offset, {}).get("A")
             or "Total interbank (Financial only)")
    cells[1] = (total_styles.get("A", "0"), label)

    summed: dict[str, float] = {}
    for col, key in ((2, "B"), (3, "C"), (4, "D"),
                     (6, "F"), (7, "G"), (8, "H")):
        total = 0.0
        for i in FOOTER_TOTAL_INCLUDED:
            total += grid[i]["nums"][key]
        summed[key] = total
        cells[col] = (total_styles.get(key, "0"),
                      Value(total, "+".join(f"{key}{n}" for n in picked)))

    # The achievement and decline cells are ratios of the total row's own
    # figures, not sums of the line percentages above them.
    for key, numerator, denominator in TOTAL_RATIO_CELLS:
        bottom = summed[denominator]
        cells[col_index(key)] = (
            total_styles.get(key, "0"),
            Value(summed[numerator] / bottom,
                  f"{numerator}{total_row}/{denominator}{total_row}")
            if bottom else 0.0)
    rows_xml.append(_row_xml(total_row, donor.footer_attrs.get(total_offset, ""),
                             cells, sst))

    head = donor.xml[:donor.xml.index("<sheetData>")]
    tail = donor.xml[donor.xml.index("</sheetData>") + len("</sheetData>"):]
    dimension = f"A1:{col_letter(max(donor.max_col, _LAST_CONTENT_COL))}{total_row}"
    head = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{dimension}"/>',
                  head, count=1)
    head = head.replace(' tabSelected="1"', "", 1)
    head = re.sub(r'xr:uid="\{[^}]*\}"', f'xr:uid="{{{str(uuid.uuid4()).upper()}}}"',
                  head, count=1)

    shift = date_row - donor.footer_start
    merges = list(donor.fixed_merges)
    for ref in donor.footer_merges:
        start_col, start_row, end_col, end_row = re.match(
            r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref).groups()
        merges.append(f"{start_col}{int(start_row) + shift}:"
                      f"{end_col}{int(end_row) + shift}")
    merge_xml = (f'<mergeCells count="{len(merges)}">'
                 + "".join(f'<mergeCell ref="{m}"/>' for m in merges)
                 + "</mergeCells>")
    if "<mergeCells" in tail:
        tail = re.sub(r"<mergeCells.*?</mergeCells>", merge_xml, tail, count=1, flags=re.S)
    else:  # pragma: no cover - the reference always has footer merges
        tail = tail.replace("<pageMargins", merge_xml + "<pageMargins", 1)

    return head + "<sheetData>" + "".join(rows_xml) + "</sheetData>" + tail


# ---------------------------------------------------------------------------
# Package surgery
# ---------------------------------------------------------------------------
def _sheet_order(workbook_xml: str) -> list[tuple[str, str, str]]:
    return [(m.group(1), m.group(2), m.group(3)) for m in re.finditer(
        r'<sheet name="([^"]*)"[^>]*?sheetId="(\d+)"[^>]*?r:id="([^"]+)"',
        workbook_xml)]


def _rel_target(rels_xml: str, rel_id: str) -> str | None:
    m = re.search(rf'<Relationship Id="{re.escape(rel_id)}"[^>]*?Target="([^"]+)"', rels_xml)
    return m.group(1) if m else None


def _rel_by_type(rels_xml: str, suffix: str) -> tuple[str, str] | None:
    """``(rId, Target)`` of the first relationship whose Type ends with ``suffix``."""
    m = re.search(rf'<Relationship Id="([^"]+)"[^>]*?Type="[^"]*/{re.escape(suffix)}"'
                  rf'[^>]*?Target="([^"]+)"', rels_xml)
    if m:
        return m.group(1), m.group(2)
    m = re.search(rf'<Relationship Id="([^"]+)"[^>]*?Target="([^"]+)"[^>]*?Type="[^"]*/'
                  rf'{re.escape(suffix)}"', rels_xml)
    return (m.group(1), m.group(2)) if m else None


def _part_path(target: str) -> str:
    return "xl/" + target.lstrip("/").replace("../", "")


def sheet_date(name: str) -> date | None:
    try:
        return datetime.strptime(name.strip(), "%d.%m.%Y").date()
    except ValueError:
        return None


def month_of(name: str) -> tuple[int, int] | None:
    parsed = sheet_date(name)
    return (parsed.year, parsed.month) if parsed else None


def pick_donor_sheet(sheet_names: Sequence[str], replacing: str | None = None) -> str:
    """The most recent dated sheet, ignoring the one being replaced."""
    candidates = [n for n in sheet_names if n != replacing] or list(sheet_names)
    if not candidates:
        raise ValueError("the workbook has no sheets to use as a template")
    dated = [n for n in candidates if sheet_date(n)]
    if not dated:
        raise ValueError("the workbook has no dd.mm.yyyy sheet to use as a template")
    return max(dated, key=lambda n: sheet_date(n))


def read_monthly_footer(sheet_xml: str, sst_xml: str) -> MonthlyFooter:
    """Read the manual footer values out of an existing sheet."""
    sst = SharedStrings(sst_xml)
    donor = DonorSheet.parse(sheet_xml, sst)
    start = donor.footer_start

    def raw(row: int, col: str) -> Any:
        """A cell as a string for text cells and a float for numeric ones."""
        cell = donor.cells.get(row, {}).get(col)
        if not cell:
            return None
        _, text, number = cell
        if text is not None:
            return text
        if number is not None:
            try:
                return float(number)
            except ValueError:
                return number
        return None

    plans: list[Any] = []
    for index in range(len(FOOTER_ROWS)):
        value = raw(start + 2 + index, "B")
        plans.append(donor.plan_text(index) if value is None else value)
    rtp_row = start + 2 + next(i for i, s in enumerate(FOOTER_ROWS) if s.rtp)
    return MonthlyFooter(
        plans=plans,
        rtp_count=_number(raw(rtp_row, "C")),
        rtp_value=_number(raw(rtp_row, "D")),
        card_note=raw(start + 2, "J") or "",
        interop_note=raw(start + 2, "M") or "",
    )


def _is_number(text: str) -> bool:
    try:
        float(text)
    except (TypeError, ValueError):
        return False
    return True


def build_daily_sheet(template: bytes, *, sheet_name: str, day: date,
                      sources: Mapping[str, bytes], footer: MonthlyFooter,
                      donor_name: str | None = None,
                      diagnostics: Diagnostics | None = None) -> bytes:
    """Return the template workbook with ``sheet_name`` added or replaced."""
    diagnostics = diagnostics or Diagnostics()
    with zipfile.ZipFile(io.BytesIO(template)) as zin:
        parts = {name: zin.read(name) for name in zin.namelist()}

    workbook_xml = parts["xl/workbook.xml"].decode("utf-8")
    rels_xml = parts["xl/_rels/workbook.xml.rels"].decode("utf-8")
    order = _sheet_order(workbook_xml)
    names = [n for n, _, _ in order]
    donor_name = donor_name or pick_donor_sheet(names, sheet_name)

    donor_rid = next(r for n, _, r in order if n == donor_name)
    donor_target = _rel_target(rels_xml, donor_rid)
    if not donor_target:
        raise ValueError(f"cannot resolve the donor sheet part for {donor_name!r}")
    donor_part = _part_path(donor_target)

    sst_part = "xl/sharedStrings.xml"
    sst = SharedStrings(parts.get(sst_part, "").decode("utf-8"))
    donor = DonorSheet.parse(parts[donor_part].decode("utf-8"), sst)

    blocks = build_blocks(sources, donor, diagnostics)
    new_sheet = _render_sheet(donor, blocks, footer, day, sst)

    sheet_no = _max_part(parts, "xl/worksheets/sheet") + 1
    new_sheet_part = f"xl/worksheets/sheet{sheet_no}.xml"
    drawing_no = _max_part(parts, "xl/drawings/drawing") + 1
    vml_no = _max_part(parts, "xl/drawings/vmlDrawing") + 1

    parts[new_sheet_part] = new_sheet.encode("utf-8")
    if sst._added:
        parts[sst_part] = sst.render()

    _clone_logo_parts(parts, donor_part, sheet_no, drawing_no, vml_no)

    # ── content types ───────────────────────────────────────────────────────
    types = parts["[Content_Types].xml"].decode("utf-8")
    types = re.sub(r'<Override PartName="/xl/calcChain\.xml"[^>]*/>', "", types)
    override = (f'<Override PartName="/{new_sheet_part}" ContentType="application/vnd.'
                f'openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                f'<Override PartName="/xl/drawings/drawing{drawing_no}.xml" ContentType='
                f'"application/vnd.openxmlformats-officedocument.drawing+xml"/>')
    parts["[Content_Types].xml"] = types.replace("</Types>",
                                                 override + "</Types>").encode("utf-8")
    parts.pop("xl/calcChain.xml", None)

    # ── workbook and its relationships ──────────────────────────────────────
    final_names = [n for n in names if n != sheet_name]
    replaced_at: int | None = None
    if sheet_name in names:
        for name, _, rid in order:
            if name != sheet_name:
                continue
            sheets = re.search(r"<sheets>(.*?)</sheets>", workbook_xml, re.S)
            if sheets:
                elements = re.findall(r"<sheet [^>]*?/>", sheets.group(1))
                index = next((i for i, e in enumerate(elements)
                              if re.search(rf'name="{re.escape(sheet_name)}"', e)), None)
                if index is not None:
                    replaced_at = index
            workbook_xml = re.sub(
                rf'<sheet name="{re.escape(sheet_name)}"[^>]*?/>', "", workbook_xml, count=1)
            target = _rel_target(rels_xml, rid)
            rels_xml = re.sub(rf'<Relationship Id="{re.escape(rid)}"[^>]*?/>', "",
                              rels_xml, count=1)
            if target:
                parts.pop(_part_path(target), None)
                parts.pop(f"xl/worksheets/_rels/{target.split('/')[-1]}.rels", None)

    new_rid = f"rId{_max_rid(rels_xml) + 1}"
    rels_xml = re.sub(r'<Relationship[^>]*?calcChain\.xml"[^>]*?/>', "", rels_xml)
    element = (f'<sheet name="{_escape(sheet_name)}" sheetId="{_max_sheet_id(workbook_xml) + 1}"'
               f' r:id="{new_rid}"/>')
    sheets = re.search(r"<sheets>(.*?)</sheets>", workbook_xml, re.S)
    if sheets and replaced_at is not None:
        # Replace in place so the day keeps its chronological tab position.
        elements = re.findall(r"<sheet [^>]*?/>", sheets.group(1))
        elements.insert(min(replaced_at, len(elements)), element)
        workbook_xml = (workbook_xml[:sheets.start(1)] + "".join(elements)
                        + workbook_xml[sheets.end(1):])
    else:
        workbook_xml = workbook_xml.replace("</sheets>", element + "</sheets>")
    rels_xml = rels_xml.replace(
        "</Relationships>",
        f'<Relationship Id="{new_rid}" Type="http://schemas.openxmlformats.org/'
        f'officeDocument/2006/relationships/worksheet" '
        f'Target="worksheets/sheet{sheet_no}.xml"/></Relationships>')
    parts["xl/workbook.xml"] = workbook_xml.encode("utf-8")
    parts["xl/_rels/workbook.xml.rels"] = rels_xml.encode("utf-8")
    final_names.append(sheet_name)

    if "docProps/app.xml" in parts:
        parts["docProps/app.xml"] = _update_app_xml(
            parts["docProps/app.xml"].decode("utf-8"), final_names).encode("utf-8")

    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in parts.items():
            zout.writestr(name, payload)
    return out.getvalue()


def _clone_logo_parts(parts: dict[str, bytes], donor_part: str, sheet_no: int,
                      drawing_no: int, vml_no: int) -> None:
    """Copy the donor's drawing / VML / OLE logo parts for the new sheet."""
    donor_rels_part = f"xl/worksheets/_rels/{donor_part.split('/')[-1]}.rels"
    if donor_rels_part not in parts:
        return
    donor_rels = parts[donor_rels_part].decode("utf-8")

    drawing = _rel_by_type(donor_rels, "drawing")
    vml = _rel_by_type(donor_rels, "vmlDrawing")
    if not drawing or not vml:
        return
    drawing_src, vml_src = _part_path(drawing[1]), _part_path(vml[1])
    if drawing_src not in parts or vml_src not in parts:
        return

    new_drawing = f"xl/drawings/drawing{drawing_no}.xml"
    new_vml = f"xl/drawings/vmlDrawing{vml_no}.vml"
    parts[new_drawing] = parts[drawing_src]
    parts[new_vml] = parts[vml_src]

    vml_rels = f"xl/drawings/_rels/{vml_src.split('/')[-1]}.rels"
    if vml_rels in parts:
        parts[f"xl/drawings/_rels/vmlDrawing{vml_no}.vml.rels"] = parts[vml_rels]

    # Shape ids must stay unique across the workbook.
    marker = 0x40000000 + sheet_no * 0x400
    vml_text = parts[new_vml].decode("utf-8")
    vml_text = re.sub(r'id="_x0000_s(\d+)"',
                      lambda m: f'id="_x0000_s{int(m.group(1)) + marker}"', vml_text)
    parts[new_vml] = vml_text.encode("utf-8")

    new_rels = donor_rels.replace(f'Target="{drawing[1]}"',
                                  f'Target="../drawings/drawing{drawing_no}.xml"')
    new_rels = new_rels.replace(f'Target="{vml[1]}"',
                                f'Target="../drawings/vmlDrawing{vml_no}.vml"')

    # Give the new sheet its own OLE embedding parts (this workbook keeps one
    # set per sheet) so the donor sheet can still be recalculated and re-saved.
    new_rels, added = _rewrite_ole_targets(
        new_rels, parts, _max_part(parts, "xl/embeddings/oleObject") + 1)
    parts.update(added)
    parts[f"xl/worksheets/_rels/sheet{sheet_no}.xml.rels"] = new_rels.encode("utf-8")


def _rewrite_ole_targets(rels: str, parts: Mapping[str, bytes], first: int
                         ) -> tuple[str, dict[str, bytes]]:
    """Point every OLE relationship at its own copy of the embedding part."""
    added: dict[str, bytes] = {}
    counter = [first]

    def repl(match: re.Match[str]) -> str:
        old = _part_path(match.group(2))
        if old not in parts:
            return match.group(0)
        name = f"oleObject{counter[0]}.bin"
        counter[0] += 1
        added[f"xl/embeddings/{name}"] = parts[old]
        return f'{match.group(1)}../embeddings/{name}{match.group(3)}'

    out = re.sub(r'(Target=")([^"]*oleObject\d+\.bin)(")', repl, rels)
    return out, added


def _max_part(parts: Mapping[str, bytes], prefix: str) -> int:
    best = 0
    for name in parts:
        m = re.fullmatch(re.escape(prefix) + r"(\d+)\.(?:xml|vml|bin)", name)
        if m:
            best = max(best, int(m.group(1)))
    return best


def _max_rid(rels_xml: str) -> int:
    return max((int(n) for n in re.findall(r'Id="rId(\d+)"', rels_xml)), default=0)


def _max_sheet_id(workbook_xml: str) -> int:
    return max((int(i) for i in re.findall(r'<sheet [^>]*?sheetId="(\d+)"', workbook_xml)),
               default=0)


def _update_app_xml(xml: str, names: Sequence[str]) -> str:
    xml = re.sub(
        r"(<vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>)\d+(</vt:i4>)",
        lambda m: f"{m.group(1)}{len(names)}{m.group(2)}", xml, count=1)
    titles = "".join(f"<vt:lpstr>{_escape(n)}</vt:lpstr>" for n in names)
    return re.sub(
        r'(<TitlesOfParts><vt:vector size=")\d+(" baseType="lpstr">).*?(</vt:vector></TitlesOfParts>)',
        lambda m: f"{m.group(1)}{len(names)}{m.group(2)}{titles}{m.group(3)}",
        xml, count=1, flags=re.S)
