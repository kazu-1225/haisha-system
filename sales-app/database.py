import aiosqlite
import os
from typing import Optional

DB_PATH = os.path.join(os.path.dirname(__file__), "sales.db")

CREATE_SALES_TABLE = """
CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    invoice_no TEXT,
    upload_id INTEGER NOT NULL,
    sale_date TEXT NOT NULL,
    customer_code TEXT,
    customer_name TEXT,
    product_code TEXT,
    product_name TEXT,
    category TEXT,
    quantity REAL,
    unit_price REAL,
    amount REAL,
    FOREIGN KEY (upload_id) REFERENCES uploads(id)
);
"""

CREATE_UPLOADS_TABLE = """
CREATE TABLE IF NOT EXISTS uploads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    filename TEXT NOT NULL,
    uploaded_at TEXT NOT NULL,
    row_count INTEGER,
    column_mapping TEXT
);
"""

CREATE_INDEX_SALE_DATE = """
CREATE INDEX IF NOT EXISTS idx_sale_date ON sales(sale_date);
"""

CREATE_INDEX_CUSTOMER = """
CREATE INDEX IF NOT EXISTS idx_customer ON sales(customer_name);
"""

CREATE_INDEX_PRODUCT = """
CREATE INDEX IF NOT EXISTS idx_product ON sales(product_name);
"""

CREATE_INDEX_CATEGORY = """
CREATE INDEX IF NOT EXISTS idx_category ON sales(category);
"""


from contextlib import asynccontextmanager

@asynccontextmanager
async def get_db():
    async with aiosqlite.connect(DB_PATH) as db:
        db.row_factory = aiosqlite.Row
        await db.execute("PRAGMA journal_mode=WAL")
        await db.execute("PRAGMA foreign_keys=ON")
        yield db


async def init_db():
    async with aiosqlite.connect(DB_PATH) as db:
        await db.execute(CREATE_UPLOADS_TABLE)
        await db.execute(CREATE_SALES_TABLE)
        await db.execute(CREATE_INDEX_SALE_DATE)
        await db.execute(CREATE_INDEX_CUSTOMER)
        await db.execute(CREATE_INDEX_PRODUCT)
        await db.execute(CREATE_INDEX_CATEGORY)
        await db.commit()
        # Migration: add invoice_no if not exists
        try:
            await db.execute("ALTER TABLE sales ADD COLUMN invoice_no TEXT")
            await db.commit()
        except Exception:
            pass  # Column already exists


async def insert_upload(db: aiosqlite.Connection, filename: str, uploaded_at: str,
                         row_count: int, column_mapping: str) -> int:
    cursor = await db.execute(
        "INSERT INTO uploads (filename, uploaded_at, row_count, column_mapping) VALUES (?, ?, ?, ?)",
        (filename, uploaded_at, row_count, column_mapping)
    )
    await db.commit()
    return cursor.lastrowid


async def insert_sales_batch(db: aiosqlite.Connection, rows: list[dict]):
    await db.executemany(
        """INSERT INTO sales
           (invoice_no, upload_id, sale_date, customer_code, customer_name,
            product_code, product_name, category, quantity, unit_price, amount)
           VALUES (:invoice_no, :upload_id, :sale_date, :customer_code, :customer_name,
                   :product_code, :product_name, :category, :quantity, :unit_price, :amount)""",
        rows
    )
    await db.commit()


async def get_summary(db: aiosqlite.Connection, start: Optional[str] = None, end: Optional[str] = None) -> dict:
    where, params = _date_filter(start, end)
    row = await (await db.execute(
        f"""SELECT
                COUNT(*) as transaction_count,
                COALESCE(SUM(amount), 0) as total_amount,
                COALESCE(SUM(quantity), 0) as total_quantity,
                COUNT(DISTINCT customer_name) as customer_count,
                COUNT(DISTINCT product_name) as product_count
            FROM sales {where}""", params
    )).fetchone()
    return dict(row) if row else {}


async def get_period_aggregation(db: aiosqlite.Connection, start: Optional[str], end: Optional[str],
                                  group_by: str) -> list[dict]:
    where, params = _date_filter(start, end)
    if group_by == "day":
        date_expr = "sale_date"
    elif group_by == "week":
        date_expr = "strftime('%Y-W%W', sale_date)"
    elif group_by == "month":
        date_expr = "strftime('%Y-%m', sale_date)"
    elif group_by == "year":
        date_expr = "strftime('%Y', sale_date)"
    else:
        date_expr = "strftime('%Y-%m', sale_date)"

    rows = await (await db.execute(
        f"""SELECT
                {date_expr} as period,
                COALESCE(SUM(amount), 0) as total_amount,
                COALESCE(SUM(quantity), 0) as total_quantity,
                COUNT(*) as transaction_count,
                COUNT(DISTINCT customer_name) as customer_count
            FROM sales {where}
            GROUP BY period
            ORDER BY period ASC""", params
    )).fetchall()
    return [dict(r) for r in rows]


async def get_customer_aggregation(db: aiosqlite.Connection, start: Optional[str], end: Optional[str],
                                    customer_name: Optional[str] = None) -> list[dict]:
    where, params = _date_filter(start, end)
    if customer_name:
        connector = "AND" if where else "WHERE"
        where += f" {connector} customer_name LIKE ?"
        params.append(f"%{customer_name}%")

    rows = await (await db.execute(
        f"""SELECT
                customer_code,
                customer_name,
                COALESCE(SUM(amount), 0) as total_amount,
                COALESCE(SUM(quantity), 0) as total_quantity,
                COUNT(*) as transaction_count,
                COUNT(DISTINCT sale_date) as active_days
            FROM sales {where}
            GROUP BY customer_name, customer_code
            ORDER BY total_amount DESC""", params
    )).fetchall()
    return [dict(r) for r in rows]


async def get_product_aggregation(db: aiosqlite.Connection, start: Optional[str], end: Optional[str]) -> list[dict]:
    where, params = _date_filter(start, end)
    rows = await (await db.execute(
        f"""SELECT
                product_code,
                product_name,
                category,
                COALESCE(SUM(amount), 0) as total_amount,
                COALESCE(SUM(quantity), 0) as total_quantity,
                COUNT(*) as transaction_count,
                COUNT(DISTINCT customer_name) as customer_count
            FROM sales {where}
            GROUP BY product_name, product_code, category
            ORDER BY total_amount DESC""", params
    )).fetchall()
    return [dict(r) for r in rows]


async def get_category_aggregation(db: aiosqlite.Connection, start: Optional[str], end: Optional[str]) -> list[dict]:
    where, params = _date_filter(start, end)
    rows = await (await db.execute(
        f"""SELECT
                category,
                COALESCE(SUM(amount), 0) as total_amount,
                COALESCE(SUM(quantity), 0) as total_quantity,
                COUNT(*) as transaction_count
            FROM sales {where}
            GROUP BY category
            ORDER BY total_amount DESC""", params
    )).fetchall()
    return [dict(r) for r in rows]


async def get_customers(db: aiosqlite.Connection) -> list[dict]:
    rows = await (await db.execute(
        "SELECT DISTINCT customer_code, customer_name FROM sales ORDER BY customer_name"
    )).fetchall()
    return [dict(r) for r in rows]


async def get_products(db: aiosqlite.Connection) -> list[dict]:
    rows = await (await db.execute(
        "SELECT DISTINCT product_code, product_name, category FROM sales ORDER BY product_name"
    )).fetchall()
    return [dict(r) for r in rows]


async def get_uploads(db: aiosqlite.Connection) -> list[dict]:
    rows = await (await db.execute(
        "SELECT * FROM uploads ORDER BY uploaded_at DESC"
    )).fetchall()
    return [dict(r) for r in rows]


async def delete_upload(db: aiosqlite.Connection, upload_id: int):
    await db.execute("DELETE FROM sales WHERE upload_id = ?", (upload_id,))
    await db.execute("DELETE FROM uploads WHERE id = ?", (upload_id,))
    await db.commit()


async def get_frequency_analysis(db: aiosqlite.Connection, start: Optional[str], end: Optional[str],
                                  customer_keyword: Optional[str], product_keyword: Optional[str],
                                  group_by: str, pivot_by: str) -> list[dict]:
    where, params = _date_filter(start, end)
    if customer_keyword:
        connector = "AND" if where else "WHERE"
        where += f" {connector} customer_name LIKE ?"
        params.append(f"%{customer_keyword}%")
    if product_keyword:
        connector = "AND" if where else "WHERE"
        where += f" {connector} product_name LIKE ?"
        params.append(f"%{product_keyword}%")

    if group_by == 'day':
        date_expr = "sale_date"
    elif group_by == 'week':
        date_expr = "strftime('%Y-W%W', sale_date)"
    elif group_by == 'year':
        date_expr = "strftime('%Y', sale_date)"
    else:
        date_expr = "strftime('%Y-%m', sale_date)"

    group_col = "customer_name" if pivot_by == "customer" else "product_name"

    rows = await (await db.execute(f"""
        SELECT
            {date_expr} as period,
            {group_col} as group_label,
            COUNT(DISTINCT invoice_no) as invoice_count,
            COUNT(*) as transaction_count,
            COALESCE(SUM(amount), 0) as total_amount,
            COALESCE(SUM(quantity), 0) as total_quantity
        FROM sales {where}
        GROUP BY period, {group_col}
        ORDER BY period ASC, total_amount DESC
    """, params)).fetchall()
    return [dict(r) for r in rows]


def _date_filter(start: Optional[str], end: Optional[str]) -> tuple[str, list]:
    clauses = []
    params = []
    if start:
        clauses.append("sale_date >= ?")
        params.append(start)
    if end:
        clauses.append("sale_date <= ?")
        params.append(end)
    where = ("WHERE " + " AND ".join(clauses)) if clauses else ""
    return where, params
