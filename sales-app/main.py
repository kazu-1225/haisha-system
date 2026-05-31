import io
import json
import os
from datetime import datetime
from typing import Optional

import chardet
import pandas as pd
from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles

from database import (
    delete_upload,
    get_category_aggregation,
    get_customer_aggregation,
    get_customers,
    get_db,
    get_frequency_analysis,
    get_frequency_detail_by_invoice,
    get_invoice_lines,
    get_period_aggregation,
    get_product_aggregation,
    get_product_detail,
    get_product_search,
    get_products,
    get_summary,
    get_uploads,
    init_db,
    insert_sales_batch,
    insert_upload,
)

app = FastAPI(title="T-PLANNER 売上集計システム", version="1.0.0")

STATIC_DIR = os.path.join(os.path.dirname(__file__), "static")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")

# ---------------------------------------------------------------------------
# Column name candidates for auto-detection
# ---------------------------------------------------------------------------
COLUMN_CANDIDATES = {
    "sale_date": ["売上日", "日付", "受注日", "計上日", "date", "sale_date", "受注年月日", "売上年月日"],
    "invoice_no": ["売上NO", "売上No", "売上番号", "伝票NO", "伝票No", "invoice_no"],
    "customer_code": ["得意先コード", "顧客コード", "取引先コード", "customer_code", "得意先CD", "顧客CD"],
    "customer_name": ["得意先名", "顧客名", "取引先名", "customer_name", "得意先名称", "顧客名称"],
    "product_code": ["商品コード", "品番", "product_code", "商品CD", "品目コード"],
    "product_name": ["商品名", "品名", "製品名", "product_name", "商品名称", "品目名"],
    "category": ["カテゴリ", "分類", "商品分類", "category", "大分類", "品目分類"],
    "quantity": ["数量", "qty", "quantity", "受注数量", "出荷数量"],
    "unit_price": ["単価", "unit_price", "販売単価", "売上単価"],
    "amount": ["金額", "売上金額", "受注金額", "amount", "売上高", "合計金額", "小計"],
    "remarks": ["摘要", "備考", "摘要欄", "remarks"],
    "spec": ["規格", "仕様", "規格名", "spec"],
}


def detect_encoding(raw: bytes) -> str:
    result = chardet.detect(raw)
    enc = result.get("encoding") or "utf-8"
    # Normalize common Japanese encodings
    enc_lower = enc.lower().replace("-", "")
    if enc_lower in ("shiftjis", "shiftjisx0208", "mskanji", "csshiftjis"):
        return "shift_jis"
    if enc_lower in ("eucjp", "eucjisx0208"):
        return "euc_jp"
    return enc


def auto_map_columns(df_columns: list[str]) -> dict:
    mapping = {}
    col_lower = {c: c.strip() for c in df_columns}
    for field, candidates in COLUMN_CANDIDATES.items():
        for col in df_columns:
            stripped = col.strip()
            if stripped in candidates or stripped.lower() in [c.lower() for c in candidates]:
                mapping[field] = col
                break
    return mapping


def parse_date(val) -> Optional[str]:
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s in ("nan", "None", ""):
        return None
    # Try common formats
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y%m%d", "%Y年%m月%d日",
                "%m/%d/%Y", "%d/%m/%Y", "%Y.%m.%d"):
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    try:
        from dateutil import parser as dparser
        return dparser.parse(s).strftime("%Y-%m-%d")
    except Exception:
        return None


def safe_float(val) -> Optional[float]:
    if pd.isna(val):
        return None
    try:
        cleaned = str(val).replace(",", "").replace("¥", "").replace("￥", "").strip()
        return float(cleaned)
    except (ValueError, TypeError):
        return None


@app.on_event("startup")
async def startup():
    await init_db()


@app.get("/", response_class=HTMLResponse)
async def root():
    index_path = os.path.join(STATIC_DIR, "index.html")
    with open(index_path, "r", encoding="utf-8") as f:
        return f.read()


# ---------------------------------------------------------------------------
# Upload endpoint
# ---------------------------------------------------------------------------
@app.post("/api/upload")
async def upload_csv(
    file: UploadFile = File(...),
    column_mapping_json: Optional[str] = Form(None),
):
    if not file.filename.lower().endswith((".csv", ".tsv")):
        raise HTTPException(400, "CSVまたはTSVファイルをアップロードしてください。")

    raw = await file.read()
    encoding = detect_encoding(raw)

    try:
        text = raw.decode(encoding, errors="replace")
    except Exception:
        text = raw.decode("utf-8", errors="replace")

    sep = "\t" if file.filename.lower().endswith(".tsv") else ","
    # Try comma first, fallback to tab
    try:
        df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str, skipinitialspace=True)
        if len(df.columns) == 1:
            # Might be tab-separated saved as .csv
            df = pd.read_csv(io.StringIO(text), sep="\t", dtype=str, skipinitialspace=True)
    except Exception as e:
        raise HTTPException(400, f"CSVの解析に失敗しました: {e}")

    df.columns = [str(c).strip() for c in df.columns]

    # Determine column mapping
    if column_mapping_json:
        try:
            mapping = json.loads(column_mapping_json)
        except json.JSONDecodeError:
            mapping = auto_map_columns(list(df.columns))
    else:
        mapping = auto_map_columns(list(df.columns))

    if "sale_date" not in mapping:
        raise HTTPException(
            400,
            detail={
                "message": "売上日カラムが特定できませんでした。カラムマッピングを指定してください。",
                "columns": list(df.columns),
                "detected_mapping": mapping,
            },
        )

    rows = []
    skipped = 0
    for _, row in df.iterrows():
        sale_date = parse_date(row.get(mapping.get("sale_date", ""), pd.NA))
        if not sale_date:
            skipped += 1
            continue

        rows.append({
            "upload_id": None,  # filled after insert_upload
            "invoice_no": str(row.get(mapping.get("invoice_no", ""), "") or "").strip() or None,
            "sale_date": sale_date,
            "customer_code": str(row.get(mapping.get("customer_code", ""), "") or "").strip() or None,
            "customer_name": str(row.get(mapping.get("customer_name", ""), "") or "").strip() or None,
            "product_code": str(row.get(mapping.get("product_code", ""), "") or "").strip() or None,
            "product_name": str(row.get(mapping.get("product_name", ""), "") or "").strip() or None,
            "category": str(row.get(mapping.get("category", ""), "") or "").strip() or None,
            "quantity": safe_float(row.get(mapping.get("quantity", ""), pd.NA)),
            "unit_price": safe_float(row.get(mapping.get("unit_price", ""), pd.NA)),
            "amount": safe_float(row.get(mapping.get("amount", ""), pd.NA)),
            "remarks": str(row.get(mapping.get("remarks", ""), "") or "").strip() or None,
            "spec": str(row.get(mapping.get("spec", ""), "") or "").strip() or None,
        })

    if not rows:
        raise HTTPException(400, "有効なデータ行がありませんでした。日付カラムを確認してください。")

    uploaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    async with get_db() as db:
        upload_id = await insert_upload(db, file.filename, uploaded_at, len(rows), json.dumps(mapping, ensure_ascii=False))
        for r in rows:
            r["upload_id"] = upload_id
        await insert_sales_batch(db, rows)

    return {
        "success": True,
        "upload_id": upload_id,
        "filename": file.filename,
        "row_count": len(rows),
        "skipped": skipped,
        "column_mapping": mapping,
    }


# ---------------------------------------------------------------------------
# Preview columns (for manual mapping UI)
# ---------------------------------------------------------------------------
@app.post("/api/preview-columns")
async def preview_columns(file: UploadFile = File(...)):
    raw = await file.read()
    encoding = detect_encoding(raw)
    text = raw.decode(encoding, errors="replace")
    sep = "\t" if file.filename.lower().endswith(".tsv") else ","
    df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str, skipinitialspace=True, nrows=3)
    if len(df.columns) == 1:
        df = pd.read_csv(io.StringIO(text), sep="\t", dtype=str, skipinitialspace=True, nrows=3)
    df.columns = [str(c).strip() for c in df.columns]
    mapping = auto_map_columns(list(df.columns))
    return {
        "columns": list(df.columns),
        "detected_mapping": mapping,
        "sample_rows": df.head(3).fillna("").to_dict(orient="records"),
    }


# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
@app.get("/api/summary")
async def api_summary(start: Optional[str] = None, end: Optional[str] = None):
    async with get_db() as db:
        data = await get_summary(db, start, end)
    return data


# ---------------------------------------------------------------------------
# Aggregations
# ---------------------------------------------------------------------------
@app.get("/api/aggregation/period")
async def api_period(
    start: Optional[str] = None,
    end: Optional[str] = None,
    group_by: str = "month",
):
    if group_by not in ("day", "week", "month", "year"):
        raise HTTPException(400, "group_by は day/week/month/year のいずれかを指定してください。")
    async with get_db() as db:
        data = await get_period_aggregation(db, start, end, group_by)
    return data


@app.get("/api/aggregation/customer")
async def api_customer(
    start: Optional[str] = None,
    end: Optional[str] = None,
    customer_name: Optional[str] = None,
):
    async with get_db() as db:
        data = await get_customer_aggregation(db, start, end, customer_name)
    return data


@app.get("/api/aggregation/product")
async def api_product(
    start: Optional[str] = None,
    end: Optional[str] = None,
):
    async with get_db() as db:
        products = await get_product_aggregation(db, start, end)
        categories = await get_category_aggregation(db, start, end)
    return {"products": products, "categories": categories}


# ---------------------------------------------------------------------------
# Master data
# ---------------------------------------------------------------------------
@app.get("/api/customers")
async def api_customers():
    async with get_db() as db:
        data = await get_customers(db)
    return data


@app.get("/api/products")
async def api_products():
    async with get_db() as db:
        data = await get_products(db)
    return data


@app.get("/api/uploads")
async def api_uploads():
    async with get_db() as db:
        data = await get_uploads(db)
    return data


@app.delete("/api/uploads/{upload_id}")
async def api_delete_upload(upload_id: int):
    async with get_db() as db:
        await delete_upload(db, upload_id)
    return {"success": True, "upload_id": upload_id}


# ---------------------------------------------------------------------------
# Frequency analysis
# ---------------------------------------------------------------------------
@app.get("/api/aggregation/frequency")
async def api_frequency(
    start: Optional[str] = None,
    end: Optional[str] = None,
    customer_keyword: Optional[str] = None,
    product_keyword: Optional[str] = None,
    group_by: str = "month",
    pivot_by: str = "customer",
):
    if group_by not in ("day", "week", "month", "year"):
        raise HTTPException(400, "group_by は day/week/month/year のいずれかを指定してください。")
    if pivot_by not in ("customer", "product"):
        raise HTTPException(400, "pivot_by は customer/product のいずれかを指定してください。")
    kw_list = [k.strip() for k in (product_keyword or "").replace("　", " ").replace(",", " ").split() if k.strip()]
    async with get_db() as db:
        data = await get_frequency_analysis(db, start, end, customer_keyword, kw_list, group_by, pivot_by)
    return data


# ---------------------------------------------------------------------------
# Product search (multi-keyword, grouped)
# ---------------------------------------------------------------------------
@app.get("/api/search/product")
async def api_product_search(
    start: Optional[str] = None,
    end: Optional[str] = None,
    keywords: Optional[str] = None,
    group_by: str = "customer",
):
    kw_list = [k.strip() for k in (keywords or "").split(",") if k.strip()]
    async with get_db() as db:
        data = await get_product_search(db, start, end, kw_list, group_by)
    return data


@app.get("/api/aggregation/frequency/detail")
async def api_frequency_detail(
    start: Optional[str] = None,
    end: Optional[str] = None,
    keywords: Optional[str] = None,
    group_by: str = "customer",
    group_value: str = "",
):
    kw_list = [k.strip() for k in (keywords or "").split(",") if k.strip()]
    async with get_db() as db:
        data = await get_frequency_detail_by_invoice(db, start, end, kw_list, group_value, group_by)
    return data


@app.get("/api/invoice/{invoice_no}")
async def api_invoice_lines(invoice_no: str):
    async with get_db() as db:
        data = await get_invoice_lines(db, invoice_no)
    return data


@app.get("/api/search/product/detail")
async def api_product_detail(
    start: Optional[str] = None,
    end: Optional[str] = None,
    keywords: Optional[str] = None,
    group_by: str = "customer",
    group_value: str = "",
):
    kw_list = [k.strip() for k in (keywords or "").split(",") if k.strip()]
    async with get_db() as db:
        data = await get_product_detail(db, start, end, kw_list, group_value, group_by)
    return data
