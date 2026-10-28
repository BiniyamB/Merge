"""Product Mapping module - ported from the Laravel CSR product mapping logic.

Tracks which Financial Institutions (FIs) support each EthSwitch product,
with one status per FI-product pair. Data is persisted in a local SQLite
database (product_mapping.db) so it survives Streamlit reruns.

Statuses mirror Laravel ProductMappingController::STATUSES:
    joined, kickoff, testing, licence_not_allowed, not_started, fi_not_approached
plus ``none`` (clear / not mapped).

Category order mirrors Laravel CATEGORY_ORDER: IPS=1, Card=2, Shared Services=3.
"""

from __future__ import annotations

import io
import sqlite3
from datetime import datetime
from typing import Any

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

DB_PATH = "product_mapping.db"

# ---------------------------------------------------------------------------
# Statuses (mirror Laravel STATUSES / STATUS_SHORT)
# ---------------------------------------------------------------------------
STATUSES = {
    "joined": "Joined",
    "kickoff": "On Kickoff",
    "testing": "On Test",
    "licence_not_allowed": "Licence not allow this product",
    "not_started": "Not started",
    "fi_not_approached": "FI not approached",
}
STATUS_SHORT = {
    "joined": "J",
    "kickoff": "K",
    "testing": "T",
    "licence_not_allowed": "L",
    "not_started": "N",
    "fi_not_approached": "F",
}
MAPPABLE_STATUSES = list(STATUSES.keys())  # statuses that can be set on a cell
ALL_CELL_STATUSES = MAPPABLE_STATUSES + ["none"]  # none = clear / not mapped

# ---------------------------------------------------------------------------
# Categories (mirror Laravel CATEGORY_ORDER)
# ---------------------------------------------------------------------------
CATEGORY_ORDER = {"IPS": 1, "Card": 2, "Shared Services": 3}
PRODUCT_CATEGORIES = ["IPS", "Card", "Shared Services"]
FI_CATEGORIES = ["Bank", "MFI", "PSO", "Wallet", "Payment Provider", "Other"]

# Status cell colours (aRGB hex for openpyxl; mirror Laravel blade / Vue)
STATUS_COLOURS = {
    "joined": "FF10B981",
    "kickoff": "FF3B82F6",
    "testing": "FF8B5CF6",
    "licence_not_allowed": "FFF43F5E",
    "not_started": "FFF59E0B",
    "fi_not_approached": "FF94A3B8",
    "none": "FFCBD5E1",
}


# ---------------------------------------------------------------------------
# Database connection
# ---------------------------------------------------------------------------
def _connect() -> sqlite3.Connection:
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db() -> None:
    """Create tables if missing and seed default FIs/products when empty."""
    conn = _connect()
    try:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS financial_institutions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                category TEXT,
                notes TEXT,
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS product_catalog (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                category TEXT,
                subgroup TEXT,
                description TEXT,
                is_under_project INTEGER DEFAULT 0,
                sort_order INTEGER DEFAULT 0,
                is_active INTEGER DEFAULT 1,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE IF NOT EXISTS product_mappings (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                financial_institution_id INTEGER NOT NULL,
                product_catalog_id INTEGER NOT NULL,
                status TEXT DEFAULT 'not_started',
                notes TEXT,
                mapped_by TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(financial_institution_id, product_catalog_id),
                FOREIGN KEY(financial_institution_id)
                    REFERENCES financial_institutions(id) ON DELETE CASCADE,
                FOREIGN KEY(product_catalog_id)
                    REFERENCES product_catalog(id) ON DELETE CASCADE
            );
            CREATE INDEX IF NOT EXISTS idx_pm_fi_status
                ON product_mappings(financial_institution_id, status);
            CREATE INDEX IF NOT EXISTS idx_pm_product
                ON product_mappings(product_catalog_id);
            """
        )
        conn.commit()

        fi_count = conn.execute(
            "SELECT COUNT(*) FROM financial_institutions"
        ).fetchone()[0]
        if fi_count == 0:
            _seed_defaults(conn)
    finally:
        conn.close()


def _seed_defaults(conn: sqlite3.Connection) -> None:
    """Seed default FIs and products (mirrors Laravel ProductMappingSeeder)."""
    institutions = [
        'Abay', 'Abyssinia', 'Addis', 'Ahadu', 'Ahadu Ebirr', 'Amhara',
        'AmharaEthBirr', 'Awash', 'Berhan', 'Bunna', 'CBE', 'CBE BIRR',
        'CBO', 'Coopay-E-Birr', 'Dashen', 'Dedebit MFI', 'Enat', 'Gadaa',
        'Global', 'Goh Betoch', 'HalalPay', 'H-Cash', 'Hibret', 'Hijra',
        'Kaafi MFI', 'Kacha', 'Lion', 'Metemamen MFI', 'MPESA', 'NIB',
        'NibE-Birr', 'Nisir MFI', 'Omo', 'Oromia', 'Rammis', 'Rays MFI',
        'Sahapay', 'Shabelle', 'Sidama', 'Siinqee', 'Siinqee Wallet',
        'Siket', 'telebirr', 'Tsedey', 'Tsehay', 'Vision Fund MFI',
        'Vitabirr', 'Wegagen', 'Wegagen e-Birr', 'Yaya Wallet', 'ZamZam',
        'Zemen', 'ArifPay', 'SantipmPay', 'YgoutPay', 'StarPay', 'AddisPay',
    ]
    for name in institutions:
        conn.execute(
            "INSERT OR IGNORE INTO financial_institutions (name, category, is_active) "
            "VALUES (?, NULL, 1)",
            (name,),
        )

    products = [
        ('P2P', 'IPS', None, None, 0, 1),
        ('QR (Static)', 'IPS', 'QR', 'Static QR payment', 0, 2),
        ('QR (Dynamic)', 'IPS', 'QR', 'Dynamic QR payment', 0, 3),
        ('RTP', 'IPS', None, None, 0, 4),
        ('Alias', 'IPS', None, None, 0, 5),
        ('Bulk (Source)', 'IPS', 'Bulk', 'Bulk payment source', 0, 6),
        ('Bulk (Destination)', 'IPS', 'Bulk', 'Bulk payment destination', 0, 7),
        ('CBP', 'IPS', None, None, 0, 8),
        ('E-Mandate', 'IPS', None, None, 0, 9),
        ('PISP (Open banking)', 'IPS', None, 'Open banking - under project', 1, 10),
        ('ATM', 'Card', None, None, 0, 1),
        ('POS', 'Card', None, None, 0, 2),
        ('Cardless (Issuer)', 'Card', 'Cardless Withdrawal', 'Cardless withdrawal - issuer', 0, 3),
        ('Cardless (Acquirer)', 'Card', 'Cardless Withdrawal', 'Cardless withdrawal - acquirer', 0, 4),
        ('Contactless', 'Card', None, 'Contactless card payment', 0, 5),
        ('ATM Acquirer', 'Shared Services', 'Shared Card Switch (Hosting)', None, 0, 1),
        ('POS operator/Acquirer', 'Shared Services', 'Shared Card Switch (Hosting)', None, 0, 2),
        ('Card Perso & PIN printing', 'Shared Services', 'Issuing', None, 0, 3),
        ('NPG (By Card - Issuer)', 'Shared Services', 'NPG', 'NPG by card - issuer', 0, 4),
        ('NPG (By Card - Acquirer)', 'Shared Services', 'NPG', 'NPG by card - acquirer', 0, 5),
        ('NPG (By Account)', 'Shared Services', 'NPG', 'NPG by account', 0, 6),
        ('NPG (By Alias)', 'Shared Services', 'NPG', 'NPG by alias', 0, 7),
        ('ETHBirr', 'Shared Services', None, None, 0, 8),
        ('ETHRecon', 'Shared Services', None, 'When the product is ready, it will be included', 0, 9),
        ('Web Service Card Issuance', 'Shared Services', None, None, 0, 10),
        ('IMP', 'Shared Services', None, None, 0, 11),
        ('ETH Bridge', 'Shared Services', None, None, 0, 12),
        ('Green PIN', 'Shared Services', None, None, 0, 13),
        ('ETH Agency', 'Shared Services', None, 'Under project', 1, 14),
        ('ETH Lending', 'Shared Services', None, 'Under project', 1, 15),
    ]
    for name, category, subgroup, description, under_project, sort_order in products:
        conn.execute(
            """INSERT OR IGNORE INTO product_catalog
               (name, category, subgroup, description, is_under_project, sort_order, is_active)
               VALUES (?, ?, ?, ?, ?, ?, 1)""",
            (name, category, subgroup, description, under_project, sort_order),
        )
    conn.commit()


# ---------------------------------------------------------------------------
# Financial Institution CRUD
# ---------------------------------------------------------------------------
def get_institutions(active_only: bool = True, search: str | None = None,
                     category: str | None = None) -> list[dict]:
    conn = _connect()
    try:
        query = "SELECT * FROM financial_institutions WHERE 1=1"
        params: list[Any] = []
        if active_only:
            query += " AND is_active = 1"
        if search:
            query += " AND (name LIKE ? OR category LIKE ? OR notes LIKE ?)"
            like = f"%{search}%"
            params.extend([like, like, like])
        if category == "__none__":
            query += " AND (category IS NULL OR category = '')"
        elif category:
            query += " AND category = ?"
            params.append(category)
        query += " ORDER BY name COLLATE NOCASE"
        rows = conn.execute(query, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def add_institution(name: str, category: str | None, notes: str | None) -> int:
    conn = _connect()
    try:
        cur = conn.execute(
            "INSERT INTO financial_institutions (name, category, notes, is_active) "
            "VALUES (?, ?, ?, 1)",
            (name, category, notes),
        )
        conn.commit()
        return cur.lastrowid
    except sqlite3.IntegrityError as exc:
        raise ValueError(f"Institution '{name}' already exists.") from exc
    finally:
        conn.close()


def update_institution(fi_id: int, name: str, category: str | None,
                       notes: str | None, is_active: bool) -> None:
    conn = _connect()
    try:
        cur = conn.execute(
            "UPDATE financial_institutions SET name = ?, category = ?, notes = ?, "
            "is_active = ? WHERE id = ?",
            (name, category, notes, 1 if is_active else 0, fi_id),
        )
        conn.commit()
        if cur.rowcount == 0:
            raise ValueError("Institution not found.")
    except sqlite3.IntegrityError as exc:
        raise ValueError(f"Institution '{name}' already exists.") from exc
    finally:
        conn.close()


def delete_institution(fi_id: int) -> None:
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM financial_institutions WHERE id = ?", (fi_id,))
        conn.commit()
        if cur.rowcount == 0:
            raise ValueError("Institution not found.")
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Product Catalog CRUD
# ---------------------------------------------------------------------------
def get_products(active_only: bool = True, category: str | None = None
                 ) -> list[dict]:
    conn = _connect()
    try:
        query = "SELECT * FROM product_catalog WHERE 1=1"
        params: list[Any] = []
        if active_only:
            query += " AND is_active = 1"
        if category:
            query += " AND category = ?"
            params.append(category)
        case_parts = []
        for cat in ("IPS", "Card", "Shared Services"):
            case_parts.append(f"WHEN category = '{cat}' THEN {CATEGORY_ORDER[cat]}")
        case_sql = "CASE " + " ".join(case_parts) + " ELSE 99 END"
        query += f" ORDER BY {case_sql}, sort_order, name COLLATE NOCASE"
        rows = conn.execute(query, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def add_product(name: str, category: str, subgroup: str | None,
                description: str | None, is_under_project: bool,
                sort_order: int) -> int:
    conn = _connect()
    try:
        cur = conn.execute(
            """INSERT INTO product_catalog
               (name, category, subgroup, description, is_under_project,
                sort_order, is_active)
               VALUES (?, ?, ?, ?, ?, ?, 1)""",
            (name, category, subgroup, description,
             1 if is_under_project else 0, sort_order),
        )
        conn.commit()
        return cur.lastrowid
    except sqlite3.IntegrityError as exc:
        raise ValueError(f"Product '{name}' already exists.") from exc
    finally:
        conn.close()


def update_product(product_id: int, name: str, category: str,
                   subgroup: str | None, description: str | None,
                   is_under_project: bool, sort_order: int,
                   is_active: bool) -> None:
    conn = _connect()
    try:
        cur = conn.execute(
            """UPDATE product_catalog SET name = ?, category = ?, subgroup = ?,
               description = ?, is_under_project = ?, sort_order = ?,
               is_active = ? WHERE id = ?""",
            (name, category, subgroup, description,
             1 if is_under_project else 0, sort_order,
             1 if is_active else 0, product_id),
        )
        conn.commit()
        if cur.rowcount == 0:
            raise ValueError("Product not found.")
    except sqlite3.IntegrityError as exc:
        raise ValueError(f"Product '{name}' already exists.") from exc
    finally:
        conn.close()


def delete_product(product_id: int) -> None:
    conn = _connect()
    try:
        cur = conn.execute("DELETE FROM product_catalog WHERE id = ?", (product_id,))
        conn.commit()
        if cur.rowcount == 0:
            raise ValueError("Product not found.")
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Product Mappings
# ---------------------------------------------------------------------------
def get_mappings(fi_ids: list[int] | None = None
                 ) -> dict[int, dict[int, str]]:
    """Return {fi_id: {product_id: status}} for the given FI ids (or all)."""
    conn = _connect()
    try:
        if fi_ids:
            placeholders = ",".join("?" for _ in fi_ids)
            rows = conn.execute(
                f"SELECT financial_institution_id, product_catalog_id, status "
                f"FROM product_mappings WHERE financial_institution_id IN ({placeholders})",
                fi_ids,
            ).fetchall()
        else:
            rows = conn.execute(
                "SELECT financial_institution_id, product_catalog_id, status "
                "FROM product_mappings"
            ).fetchall()
        result: dict[int, dict[int, str]] = {}
        for r in rows:
            result.setdefault(r["financial_institution_id"], {})[r["product_catalog_id"]] = r["status"]
        return result
    finally:
        conn.close()


def toggle_mapping(fi_id: int, product_id: int, status: str,
                   mapped_by: str | None = None) -> str:
    """Set or clear a mapping.

    If status == 'none' the mapping row is deleted. Otherwise it is
    inserted/updated (updateOrCreate semantics from Laravel).
    Returns the resulting status ('none' if cleared).
    """
    if status not in ALL_CELL_STATUSES:
        raise ValueError(f"Unknown status '{status}'.")
    conn = _connect()
    try:
        if status == "none":
            conn.execute(
                "DELETE FROM product_mappings "
                "WHERE financial_institution_id = ? AND product_catalog_id = ?",
                (fi_id, product_id),
            )
            conn.commit()
            return "none"
        now = datetime.utcnow().isoformat()
        conn.execute(
            """INSERT INTO product_mappings
               (financial_institution_id, product_catalog_id, status, mapped_by,
                created_at, updated_at)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(financial_institution_id, product_catalog_id)
               DO UPDATE SET status = excluded.status,
                             mapped_by = excluded.mapped_by,
                             updated_at = excluded.updated_at""",
            (fi_id, product_id, status, mapped_by, now, now),
        )
        conn.commit()
        return status
    finally:
        conn.close()


# ---------------------------------------------------------------------------
# Stats
# ---------------------------------------------------------------------------
def get_stats() -> dict[str, Any]:
    institutions = get_institutions(active_only=False)
    products = get_products(active_only=True)
    mappings = get_mappings()

    total_fis = len(institutions)
    total_products = len(products)
    mapped_pairs = sum(len(v) for v in mappings.values())
    possible = total_fis * total_products

    status_counts = {s: 0 for s in MAPPABLE_STATUSES}
    for fi_map in mappings.values():
        for st in fi_map.values():
            if st in status_counts:
                status_counts[st] += 1

    mapped_fis = len({fi_id for fi_id, fm in mappings.items() if fm})

    # Per-category coverage
    family_coverage = []
    products_by_category: dict[str, list[dict]] = {}
    for p in products:
        products_by_category.setdefault(p["category"], []).append(p)

    for cat in PRODUCT_CATEGORIES:
        cat_products = products_by_category.get(cat, [])
        product_ids = {p["id"] for p in cat_products}
        mapped = sum(
            1 for fm in mappings.values()
            for pid in fm if pid in product_ids
        )
        possible_cat = total_fis * len(cat_products)
        coverage = round((mapped / possible_cat) * 100, 1) if possible_cat else 0
        family_coverage.append({
            "name": cat,
            "products": len(cat_products),
            "mapped": mapped,
            "possible": possible_cat,
            "coverage": coverage,
        })

    return {
        "total_fis": total_fis,
        "total_products": total_products,
        "by_status": status_counts,
        "in_progress": status_counts["kickoff"] + status_counts["testing"],
        "mapped": mapped_pairs,
        "possible": possible,
        "coverage": round((mapped_pairs / possible) * 100, 1) if possible else 0,
        "mapped_fis": mapped_fis,
        "family_coverage": family_coverage,
    }


# ---------------------------------------------------------------------------
# Excel export (mirrors Laravel export)
# ---------------------------------------------------------------------------
def export_excel(search: str | None = None, fi_category: str | None = None,
                 family: str | None = None, product: str | None = None,
                 status: str | None = None, view_mode: str | None = None
                 ) -> bytes:
    """Build an Excel-styled workbook of the product mapping matrix.

    Returns the workbook bytes (.xlsx). Layout mirrors the Laravel export:
    rows = FIs, columns = products grouped by category, cell = short status.

    All filters mirror the Streamlit matrix UI so the exported file holds
    exactly the currently-filtered view.
    """
    institutions = get_institutions(active_only=True, search=search,
                                    category=fi_category)
    all_products = get_products(active_only=True)

    # Visible products (family / product filter).
    visible_products = [p for p in all_products
                        if not family or p["category"] == family]
    if product:
        visible_products = [p for p in visible_products if p["name"] == product]

    mappings = get_mappings([fi["id"] for fi in institutions])

    # Apply status / view-mode filters to institutions, mirroring the matrix.
    filtered_fis = []
    for fi in institutions:
        fm = mappings.get(fi["id"], {})
        mapped_count = 0
        for p in visible_products:
            stt = fm.get(p["id"])
            if stt is None:
                continue
            if status in ("mapped", "unmapped"):
                if (status == "mapped") == (stt is None):
                    continue
            elif status and stt != status:
                continue
            mapped_count += 1

        if view_mode == "mapped" and mapped_count == 0:
            continue
        if view_mode == "gaps" and mapped_count >= len(visible_products):
            continue

        filtered_fis.append(fi)

    ordered = sorted(visible_products,
                     key=lambda p: (CATEGORY_ORDER.get(p["category"], 99),
                                    p["sort_order"], p["name"]))
    products_by_cat: dict[str, list[dict]] = {}
    for p in ordered:
        products_by_cat.setdefault(p["category"], []).append(p)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Mapping"

    ws["A1"] = "Product Mapping per FIs"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"Exported: {datetime.utcnow().strftime('%b %d, %Y %I:%M %p')}"
    ws["A2"].font = Font(size=9, italic=True, color="64748b")

    # Flat header row (import-compatible): FI, Category, then product names.
    header_row = 4
    data_start = header_row + 1
    col = 1
    ws.cell(row=header_row, column=col, value="FI").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=header_row, column=col).fill = PatternFill("solid", fgColor="4338CA")
    col += 1
    ws.cell(row=header_row, column=col, value="Category").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=header_row, column=col).fill = PatternFill("solid", fgColor="4338CA")
    col += 1

    product_col_start = col
    for p in ordered:
        cell = ws.cell(row=header_row, column=col, value=p["name"])
        cell.font = Font(bold=True, size=8)
        cell.fill = PatternFill("solid", fgColor="1E293B")
        cell.font = Font(bold=True, color="FFFFFF", size=8)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        col += 1
    product_col_end = col - 1

    # Category colour coding on the product header cells.
    category_colours = {"IPS": "FFE0E7FF", "Card": "FFD1FAE5", "Shared Services": "FFFEF3C7"}
    for p in ordered:
        cidx = product_col_start + ordered.index(p)
        colour = category_colours.get(p["category"], "FFF1F5F9")
        ws.cell(row=header_row, column=cidx).fill = PatternFill("solid", fgColor=colour)
        ws.cell(row=header_row, column=cidx).font = Font(color="FF1E293B", bold=True, size=8)

    r = data_start
    for fi in filtered_fis:
        ws.cell(row=r, column=1, value=fi["name"]).font = Font(bold=True)
        ws.cell(row=r, column=2, value=fi["category"] or "—")
        fm = mappings.get(fi["id"], {})
        for p in ordered:
            st = fm.get(p["id"])
            short = STATUS_SHORT.get(st, "") if st else ""
            cell = ws.cell(row=r, column=product_col_start + ordered.index(p))
            cell.value = short
            cell.alignment = Alignment(horizontal="center", vertical="center")
            if st and st in STATUS_COLOURS:
                cell.fill = PatternFill("solid", fgColor=STATUS_COLOURS[st])
                cell.font = Font(color="FFFFFF", bold=True, size=9)
        r += 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    for i in range(product_col_start, product_col_end + 1):
        ws.column_dimensions[get_column_letter(i)].width = 12
    ws.freeze_panes = f"C{data_start}"

    legend_row = r + 1
    ws.cell(row=legend_row, column=1,
            value="Legend: J = Joined, K = On Kickoff, T = On Test, "
                  "L = Licence not allow this product, N = Not started, "
                  "F = FI not approached, (blank) = Not mapped").font = Font(size=8, italic=True, color="64748b")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def download_template() -> bytes:
    """Build an import template (mirrors Laravel downloadTemplate).

    Columns: Financial Institution, Category, <product names...>, Notes.
    Includes a second sheet "Statuses" with valid status keys/labels/short codes.
    """
    institutions = get_institutions(active_only=True)
    products = get_products(active_only=True)
    mappings = get_mappings([fi["id"] for fi in institutions])

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Product Mapping"

    headers = ["Financial Institution", "Category"] + [p["name"] for p in products] + ["Notes"]
    ws.append(headers)
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1E293B")
        cell.alignment = Alignment(vertical="center", wrap_text=True)

    category_colours = {"IPS": "E0E7FF", "Card": "D1FAE5", "Shared Services": "FEF3C7"}
    for idx, p in enumerate(products):
        col = 3 + idx
        colour = category_colours.get(p["category"], "F1F5F9")
        cell = ws.cell(row=1, column=col)
        cell.fill = PatternFill("solid", fgColor=colour)
        cell.font = Font(color="1E293B", size=9)

    r = 2
    for fi in institutions:
        row_vals = [fi["name"], fi["category"] or ""]
        fm = mappings.get(fi["id"], {})
        for p in products:
            st = fm.get(p["id"])
            row_vals.append(STATUSES.get(st, "") if st else "")
        row_vals.append(fi["notes"] or "")
        ws.append(row_vals)
        r += 1

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 18
    last_col = get_column_letter(2 + len(products) + 1)
    ws.column_dimensions[last_col].width = 38
    for i in range(3, 3 + len(products)):
        ws.column_dimensions[get_column_letter(i)].width = 22
    ws.row_dimensions[1].height = 34
    ws.freeze_panes = "C2"

    ins = wb.create_sheet("Instructions")
    instructions = [
        ["Product Mapping Import"],
        ["How to use this file"],
        ['1. Do not rename, remove, or reorder the columns on the "Product Mapping" sheet.'],
        ['2. Enter a financial institution name in column A. New institutions are created automatically.'],
        ['3. For each product, choose a status from the "Statuses" sheet or leave the cell blank.'],
        ["4. A blank product cell keeps the existing mapping unchanged."],
        ['5. Enter "Not mapped" (or "none") to clear an existing mapping.'],
        ["6. Statuses are case-insensitive. Labels, status keys, and short codes are accepted."],
        ["7. Column B (Category) and the Notes column are optional."],
    ]
    for row_vals in instructions:
        ins.append(row_vals)
    ins["A1"].font = Font(bold=True, size=14)
    for row in ins.iter_rows(min_row=1, max_row=len(instructions), max_col=1):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True)
    ins.column_dimensions["A"].width = 110

    sts = wb.create_sheet("Statuses")
    sts.append(["Status key", "Status label", "Short code"])
    for c in range(1, 4):
        sts.cell(row=1, column=c).font = Font(bold=True)
    key_map = {"supported": "joined", "planned": "kickoff",
               "not mapped": "none", "unmapped": "none", "none": "none",
               "clear": "none"}
    rr = 2
    for key, label in STATUSES.items():
        sts.append([key, label, STATUS_SHORT[key]])
        rr += 1
    for alias, target in key_map.items():
        if alias not in STATUSES:
            sts.append([alias, "Not mapped / clear" if target == "none" else STATUSES[target], "-"])
            rr += 1
    sts.column_dimensions["A"].width = 26
    sts.column_dimensions["B"].width = 38
    sts.column_dimensions["C"].width = 14

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Excel import (mirrors Laravel import)
# ---------------------------------------------------------------------------
def _normalize_header(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    text = " ".join(text.split())
    return text.lower()


def import_excel(file_bytes: bytes, mapped_by: str | None = None) -> dict:
    """Import mappings from an uploaded Excel file.

    Mirrors Laravel ProductMappingController::import. Returns a summary dict.
    """
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    sheet = wb.active
    rows = list(sheet.iter_rows(values_only=True))
    if len(rows) < 2:
        raise ValueError("The uploaded file has no data rows.")

    # Detect the header row: the first row whose first cell is a financial
    # institution column (handles both template layout [row 0] and export
    # layout [title rows then header row]).
    header_idx = None
    for i, r in enumerate(rows):
        first = _normalize_header(r[0]) if r and r[0] is not None else ""
        if first in ("financial institution", "financial institution name",
                     "fi", "fi name", "institution", "institution name",
                     "bank", "bank name"):
            header_idx = i
            break
    if header_idx is None:
        raise ValueError('The "Financial Institution" column is missing. Use the downloaded template.')

    headers = [_normalize_header(h) for h in rows[header_idx]]
    data_rows = rows[header_idx + 1:]

    def find_col(aliases):
        for i, h in enumerate(headers):
            if h in aliases:
                return i
        return None

    fi_idx = find_col(["financial institution", "financial institution name", "fi", "fi name",
                       "institution", "institution name", "bank", "bank name"])
    cat_idx = find_col(["category", "fi category", "institution category", "type"])
    notes_idx = find_col(["notes", "note", "comment", "comments"])

    if fi_idx is None:
        raise ValueError('The "Financial Institution" column is missing. Use the downloaded template.')

    products = get_products(active_only=True)
    products_by_header = {}
    for p in products:
        products_by_header[_normalize_header(p["name"])] = p

    legacy_headers = {
        "qr": "QR (Static)",
        "bulk": "Bulk (Source)",
        "cardless withdrawal": "Cardless (Issuer)",
        "npg": "NPG (By Card - Issuer)",
    }
    for legacy, current in legacy_headers.items():
        ck = _normalize_header(current)
        if ck in products_by_header:
            products_by_header[legacy] = products_by_header[ck]

    product_columns = {}
    for i, h in enumerate(headers):
        if h in products_by_header:
            product_columns[i] = products_by_header[h]

    if not product_columns:
        raise ValueError("No product columns matched the product catalogue. Download the current template and try again.")

    status_lookup = {}
    for key, label in STATUSES.items():
        status_lookup[_normalize_header(key)] = key
        status_lookup[_normalize_header(label)] = key
        status_lookup[_normalize_header(STATUS_SHORT[key])] = key
    status_lookup.update({
        "supported": "joined",
        "planned": "kickoff",
        "not mapped": "none",
        "unmapped": "none",
        "none": "none",
        "clear": "none",
    })

    fis_created = mappings_created = mappings_updated = mappings_cleared = skipped = 0
    errors = []

    conn = _connect()
    try:
        for idx, row in enumerate(data_rows):
            row_number = idx + 2
            fi_name = str(row[fi_idx] or "").strip() if fi_idx < len(row) else ""
            if fi_name == "":
                skipped += 1
                continue
            if len(fi_name) > 255:
                skipped += 1
                errors.append(f"Row {row_number}: financial institution name is too long.")
                continue

            category = str(row[cat_idx] or "").strip() if cat_idx is not None and cat_idx < len(row) else ""
            notes = str(row[notes_idx] or "").strip() if notes_idx is not None and notes_idx < len(row) else ""

            institution = conn.execute(
                "SELECT * FROM financial_institutions WHERE name = ?", (fi_name,)
            ).fetchone()
            if not institution:
                institution_id = add_institution(fi_name, category or None, notes or None)
                fis_created += 1
            else:
                institution_id = institution["id"]
                updates = {}
                if category and institution["category"] != category:
                    updates["category"] = category
                if notes and institution["notes"] != notes:
                    updates["notes"] = notes
                if updates:
                    set_clause = ", ".join(f"{k} = ?" for k in updates)
                    vals = list(updates.values()) + [institution_id]
                    conn.execute(f"UPDATE financial_institutions SET {set_clause} WHERE id = ?", vals)
                    conn.commit()

            for col_index, product in product_columns.items():
                value = str(row[col_index] or "").strip() if col_index < len(row) else ""
                if value == "":
                    continue
                status = status_lookup.get(_normalize_header(value))
                if status is None:
                    skipped += 1
                    errors.append(
                        f"Row {row_number} ({fi_name}): invalid status \"{value}\" for {product['name']}."
                    )
                    continue
                if status == "none":
                    cur = conn.execute(
                        "DELETE FROM product_mappings "
                        "WHERE financial_institution_id = ? AND product_catalog_id = ?",
                        (institution_id, product["id"]),
                    )
                    mappings_cleared += cur.rowcount
                    conn.commit()
                    continue

                # Determine created vs updated before upsert
                exists = conn.execute(
                    "SELECT 1 FROM product_mappings "
                    "WHERE financial_institution_id = ? AND product_catalog_id = ?",
                    (institution_id, product["id"]),
                ).fetchone()
                now = datetime.utcnow().isoformat()
                conn.execute(
                    """INSERT INTO product_mappings
                       (financial_institution_id, product_catalog_id, status, mapped_by,
                        created_at, updated_at)
                       VALUES (?, ?, ?, ?, ?, ?)
                       ON CONFLICT(financial_institution_id, product_catalog_id)
                       DO UPDATE SET status = excluded.status,
                                     mapped_by = excluded.mapped_by,
                                     updated_at = excluded.updated_at""",
                    (institution_id, product["id"], status, mapped_by, now, now),
                )
                conn.commit()
                if exists:
                    mappings_updated += 1
                else:
                    mappings_created += 1
    finally:
        conn.close()

    summary = (
        f"Import complete: {mappings_created} mapping(s) created, "
        f"{mappings_updated} updated, {mappings_cleared} cleared, "
        f"{fis_created} new institution(s)."
    )
    if skipped:
        summary += f" {skipped} row(s)/cell(s) skipped."
    if errors:
        summary += " Warnings: " + " ".join(errors[:8])

    return {
        "summary": summary,
        "mappings_created": mappings_created,
        "mappings_updated": mappings_updated,
        "mappings_cleared": mappings_cleared,
        "fis_created": fis_created,
        "skipped": skipped,
        "errors": errors,
    }
