from fastapi import FastAPI, Form, Request, UploadFile, File, Depends
from starlette.middleware.sessions import SessionMiddleware
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from sqlalchemy import create_engine, Column, Integer, String, Float, DateTime, case
from sqlalchemy.orm import sessionmaker, declarative_base
from jinja2 import Environment, FileSystemLoader
from sqlalchemy import text
import pandas as pd
import io
from datetime import datetime
from fastapi.responses import FileResponse
import pdfkit
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

from fastapi.responses import HTMLResponse
from database import Base 

from jinja2 import Environment, FileSystemLoader
import os

# ---------------- PROCESS ORDER ----------------

app = FastAPI()

# ── Session middleware ──
app.add_middleware(SessionMiddleware, secret_key="inventory-secret-key-2026")

# ── Hardcoded credentials ──
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "admin123"

# ── Auto-import inventory.xlsx on startup if products table is empty ──
@app.on_event("startup")
def auto_import_inventory():
    db = SessionLocal()
    try:
        from sqlalchemy import text as _text

        # ── 1. Import inventory.xlsx → products table ──
        count = db.execute(_text("SELECT COUNT(*) FROM products")).scalar()
        if count == 0:
            xlsx_path = os.path.join(os.path.dirname(__file__), "inventory.xlsx")
            if os.path.exists(xlsx_path):
                import pandas as pd
                df = pd.read_excel(xlsx_path, engine="openpyxl")
                df.columns = [col.strip() for col in df.columns]
                for _, row in df.iterrows():
                    qty      = int(row.get("Qty", 0) or 0)
                    rate     = float(row.get("Rate", 0) or 0)
                    discount = float(row.get("Discount", 0) or 0)
                    amount   = ((qty * rate) - (qty * rate * discount / 100)) if qty > 0 else 0.0
                    db.execute(_text("""
                        INSERT INTO products (part_no, description, hsn, gst, quantity, rate, discount, amount)
                        VALUES (:pn, :desc, :hsn, :gst, :qty, :rate, :disc, :amt)
                    """), {
                        "pn":   str(row.get("Part No", "")).strip(),
                        "desc": str(row.get("Description", "")).strip(),
                        "hsn":  str(row.get("HSN", "")).strip(),
                        "gst":  float(row.get("GST", 18) or 18),
                        "qty":  qty, "rate": rate, "disc": discount,
                        "amt":  round(amount, 2)
                    })
                db.commit()
                print(f"✅ Auto-imported inventory.xlsx ({len(df)} products)")

        # ── 2. Sync orders.csv → order_items table (always refresh on startup) ──
        csv_path = os.path.join(os.path.dirname(__file__), "orders.csv")
        if os.path.exists(csv_path):
            oi_count = db.execute(_text("SELECT COUNT(*) FROM order_items")).scalar()
            import pandas as pd
            df_csv = pd.read_csv(csv_path, dtype=str, on_bad_lines="skip")
            df_csv.columns = [col.strip() for col in df_csv.columns]
            csv_count = len(df_csv)
            # Only reimport if DB is out of sync with CSV (allows Render cold starts)
            if oi_count < csv_count * 0.9:  # reimport if DB has <90% of CSV rows
                print(f"⏳ Syncing orders.csv ({csv_count} rows) into order_items ({oi_count} in DB)...")
                db.execute(_text("DELETE FROM order_items"))
                db.commit()
                batch = []
                for _, row in df_csv.iterrows():
                    part_no = str(row.get("Part No", "") or "").strip()
                    if not part_no:
                        continue
                    batch.append({
                        "pn":   part_no,
                        "desc": str(row.get("Part Desc", "") or "").strip(),
                        "hsn":  str(row.get("HSN", "") or "").strip(),
                        "mrp":  float(row.get("MRP", 0) or 0)
                    })
                    if len(batch) >= 500:
                        for item in batch:
                            db.execute(_text("""
                                INSERT INTO order_items (part_no, description, hsn, mrp)
                                VALUES (:pn, :desc, :hsn, :mrp)
                            """), item)
                        db.commit()
                        batch = []
                for item in batch:
                    db.execute(_text("""
                        INSERT INTO order_items (part_no, description, hsn, mrp)
                        VALUES (:pn, :desc, :hsn, :mrp)
                    """), item)
                db.commit()
                print(f"✅ Synced orders.csv ({csv_count} products into order_items)")
            else:
                print(f"✅ order_items up to date ({oi_count} rows)")
        else:
            print("⚠ orders.csv not found — skipping order_items sync")

    except Exception as e:
        print(f"⚠ Auto-import skipped: {e}")
        import traceback; traceback.print_exc()
    finally:
        db.close()

# ── Auth helper ──
def get_current_user(request: Request):
    return request.session.get("user")
# Raw Jinja2 rendering — bypasses Starlette Jinja2Templates API version issues
_jinja_env = Environment(
    loader=FileSystemLoader("templates"),
    cache_size=0,
    auto_reload=True
)

def render(name: str, status_code: int = 200, **ctx) -> HTMLResponse:
    html = _jinja_env.get_template(name).render(**ctx)
    return HTMLResponse(content=html, status_code=status_code)
from database import Base, engine, SessionLocal

Base.metadata.create_all(bind=engine)

# ---------------- MODELS ----------------

class Product(Base):
    __tablename__ = "products"
    id = Column(Integer, primary_key=True)
    part_no = Column(String(100))
    description = Column(String(255))
    hsn = Column(String(50))
    gst = Column(Float)
    quantity = Column(Integer)
    rate = Column(Float)
    discount = Column(Float)
    amount = Column(Float)

class Invoice(Base):
    __tablename__ = "invoices"

    id = Column(Integer, primary_key=True)
    invoice_no = Column(String(50), unique=True)
    customer_name = Column(String(255))
    date = Column(DateTime, default=datetime.utcnow)

    total_amount = Column(Float)
    gst_amount = Column(Float)
    grand_total = Column(Float)

class InvoiceItem(Base):
    __tablename__ = "invoice_items"

    id = Column(Integer, primary_key=True)

    invoice_id = Column(Integer)  # FK (can add relationship later)

    part_no = Column(String(100))
    description = Column(String(255))

    quantity = Column(Integer)
    rate = Column(Float)
    amount = Column(Float)

    hsn = Column(String(50))

# class OrderItems(Base):
#     __tablename__ = "order_items"

#     id = Column(Integer, primary_key=True)
#     part_no = Column(String(100))
#     description = Column(String(255))
#     rate = Column(Float)
#     hsn = Column(String(50))

class OrderItems(Base):
    __tablename__ = "order_items"

    id = Column(Integer, primary_key=True)
    part_no = Column(String(100))
    description = Column(String(255))
    hsn = Column(String(50))
    mrp = Column(Float)

class Quotation(Base):
    __tablename__ = "quotations"
    id = Column(Integer, primary_key=True)
    quotation_no = Column(String(50))
    customer_name = Column(String(255), default="")
    date = Column(DateTime, default=datetime.utcnow)
    total_amount = Column(Float, default=0.0)
    grand_total = Column(Float, default=0.0)

class QuotationItem(Base):
    __tablename__ = "quotation_items"
    id = Column(Integer, primary_key=True)
    quotation_id = Column(Integer)
    part_no = Column(String(100))
    description = Column(String(255))
    rate = Column(Float)
    qty = Column(Integer)
    discount = Column(Float)
    amount = Column(Float)
    hsn = Column(String(50), default="")

Base.metadata.create_all(bind=engine)

def format_inr(amount):
    s = f"{amount:,.2f}"
    parts = s.split(".")
    integer = parts[0]
    decimal = parts[1]

    if len(integer) > 3:
        last3 = integer[-3:]
        rest = integer[:-3]
        rest = ",".join([rest[max(i-2,0):i] for i in range(len(rest), 0, -2)][::-1])
        integer = rest + "," + last3

    return f"₹ {integer}.{decimal}"

# ---------------- HOME ----------------

# @app.get("/", response_class=HTMLResponse)
# def home(request: Request):
#     db = SessionLocal()
#     products = db.query(Product).all()

#     products = sorted(products, key=lambda x: x.quantity > 0)

#     # 🔥 ADD THIS
#     total_value = sum([(p.amount or 0) for p in products])
#     formatted_total = format_inr(total_value)

#     db.close()

#     return templates.TemplateResponse(
#         "index.html",
#         context={
#             "request": request,
#             "products": products,
#             "total_value": round(total_value, 2)
#         }
#     )

# ── Login routes ──
@app.get("/login", response_class=HTMLResponse)
def login_page(request: Request):
    if request.session.get("user"):
        return RedirectResponse("/", status_code=302)
    return render("login.html", error=None)

@app.post("/login", response_class=HTMLResponse)
def login(request: Request, username: str = Form(...), password: str = Form(...)):
    if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
        request.session["user"] = username
        return RedirectResponse("/", status_code=303)
    return render("login.html", status_code=401, error="Invalid username or password. Please try again.")

@app.get("/logout")
def logout(request: Request):
    request.session.clear()
    return RedirectResponse("/login", status_code=302)

# ── Protected home route ──
@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    if not request.session.get("user"):
        return RedirectResponse("/login", status_code=302)

    db = SessionLocal()

    products = db.query(Product).order_by(
        case(
            (Product.quantity == 0, 0),
            (Product.quantity < 5, 1),
            else_=2
        )
    ).all()
    total_value = sum(
    (p.quantity * p.rate) - ((p.quantity * p.rate) * (p.discount / 100))
    if p.quantity > 0 else 0
    for p in products
)

    db.close()

    return render("index.html", products=products, total_value=round(total_value, 2))
# ---------------- PRICE MASTER UPLOAD ----------------


from fastapi.responses import StreamingResponse
import pandas as pd
import io

@app.get("/export_excel")
def export_excel(start_date: str, end_date: str):

    db = SessionLocal()

    query = text("""
        SELECT 
            invoice_no,
            customer_name,
            total_amount,
            gst_amount,
            grand_total,
            date
        FROM invoices
        WHERE DATE(date) BETWEEN :start AND :end
    """)

    result = db.execute(query, {
        "start": start_date,
        "end": end_date
    }).fetchall()

    db.close()

    df = pd.DataFrame(result, columns=[
        "Invoice No", "Customer", "Total", "GST", "Grand Total", "Date"
    ])

    df["CGST"] = df["GST"] / 2
    df["SGST"] = df["GST"] / 2

    total_orders = len(df)
    total_sales = df["Total"].sum()
    total_cgst = df["CGST"].sum()
    total_sgst = df["SGST"].sum()
    grand_total = df["Grand Total"].sum()

    summary_df = pd.DataFrame({
        "Metric": [
            "Total Orders",
            "Total Sales",
            "Total CGST",
            "Total SGST",
            "Grand Total"
        ],
        "Value": [
            total_orders,
            total_sales,
            total_cgst,
            total_sgst,
            grand_total
        ]
    })

    file_path = "/tmp/sales.xlsx"

    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Sales Data", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    return FileResponse(file_path, filename="sales.xlsx")

    # -------------------------
    # WRITE TO EXCEL
    # -------------------------
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Sales Data")
        summary_df.to_excel(writer, index=False, sheet_name="Summary")

    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=sales_report.xlsx"}
    )
@app.post("/upload_excel")
def upload_excel(file: UploadFile = File(...)):
    if file.filename.endswith(".csv"):
        df = pd.read_csv(file.file)
    else:
        df = pd.read_excel(file.file, engine="openpyxl")
    df.columns = [col.strip().upper() for col in df.columns]

    db = SessionLocal()
    db.query(OrderItems).delete()

    for _, row in df.iterrows():
        part_no = str(row.get("PART NO", "")).strip()
        if not part_no:
            continue

        db.add(OrderItems(
            part_no=part_no,
            description=row.get("PART DESC", ""),
            mrp=float(str(row.get("MRP", 0)).replace(",", "")),  # ✅ FIX
            hsn=row.get("HSN", "")
        ))

    db.commit()
    db.close()

    return {"message": "Price master uploaded"}


# ---------------- ADD PRODUCT ----------------
@app.post("/add_stock/{id}")
def add_stock(id: int, qty: int = Form(...)):
    db = SessionLocal()

    product = db.query(Product).filter(Product.id == id).first()

    if not product:
        db.close()
        return {"error": "Product not found"}

    # ✅ Add stock
    product.quantity += qty

    if product.quantity <= 0:
        product.amount = 0
    else:
        product.amount = (product.quantity * product.rate) - (
            (product.quantity * product.rate) * (product.discount / 100)
        )

    db.commit()
    db.close()

    return RedirectResponse("/", status_code=303)

@app.post("/issue/{id}")
def issue_stock(id: int, qty: int = Form(...)):
    db = SessionLocal()

    product = db.query(Product).filter(Product.id == id).first()

    if not product:
        db.close()
        return {"error": "Product not found"}

    if product.quantity < qty:
        db.close()
        return {"error": "Not enough stock"}

    # ✅ STEP 1: Reduce stock
    product.quantity -= qty

    # ✅ STEP 2: Recalculate amount
    if product.quantity <= 0:
        product.quantity = 0
        product.amount = 0
    else:
        product.amount = (product.quantity * product.rate) - (
            (product.quantity * product.rate) * (product.discount / 100)
        )

    # ✅ STEP 3: Save changes
    db.commit()
    db.refresh(product)
    db.close()

    return RedirectResponse("/", status_code=303)
@app.post("/add")
def add_product(
    part_no: str = Form(...),
    description: str = Form(...),
    hsn: str = Form(...),
    gst: float = Form(...),
    quantity: int = Form(...),
    rate: float = Form(...),
    discount: float = Form(0)
):
    db = SessionLocal()

    price = db.query(OrderItems).filter(OrderItems.part_no == part_no).first()
    if price and price.mrp > 0:
        rate = price.mrp

    existing = db.query(Product).filter(Product.part_no == part_no).first()

    if existing:
        existing.quantity += quantity

        # ✅ update rate if valid
        if rate > 0:
            existing.rate = rate

        # ✅ calculate amount
        if existing.quantity <= 0:
            existing.quantity = 0
            existing.amount = 0
        else:
            existing.amount = (existing.quantity * existing.rate) - (
                (existing.quantity * existing.rate) * (existing.discount / 100)
            )

    else:
        # ✅ new product
        if quantity <= 0:
            amount = 0
        else:
            amount = (quantity * rate) - (
                (quantity * rate) * (discount / 100)
            )

        new_product = Product(
            part_no=part_no,
            description=description,
            hsn=hsn,
            gst=gst,
            quantity=quantity,
            rate=rate,
            discount=discount,
            amount=amount
        )

        db.add(new_product)

    # ✅ SINGLE COMMIT (VERY IMPORTANT)
    db.commit()
    db.close()

    return RedirectResponse("/", status_code=303)

# ---------------- GET PRICE ----------------

@app.get("/get_price/{part_no}")
def get_price(part_no: str):
    db = SessionLocal()

    price = db.query(OrderItems).filter(
        OrderItems.part_no == part_no
    ).first()

    db.close()

    if not price:
        return {"rate": 0, "hsn": "", "description": ""}

    return {
        "rate": price.mrp,   # ✅ FIXED
        "hsn": price.hsn,
        "description": price.description
    }


@app.get("/upload_page", response_class=HTMLResponse)
def upload_page():
    return """
    <html>
    <body>
        <h2>Upload Order Excel</h2>
        <form action="/upload_order" method="post" enctype="multipart/form-data">
            <input type="file" name="file" />
            <button type="submit">Upload</button>
        </form>
    </body>
    </html>
    """

@app.post("/process_order")
def process_order(file: UploadFile = File(...)):
    df = pd.read_excel(file.file)
    df.columns = [col.strip().upper() for col in df.columns]


    db = SessionLocal()
    result = []

    for _, row in df.iterrows():
        part_no = str(row.get("PART NO", "")).strip()
        qty = int(row.get("QTY", 0))

        price = db.query(OrderItems).filter(OrderItems.part_no == part_no).first()
        stock = db.query(Product).filter(Product.part_no == part_no).first()

        result.append({
            "part_no": part_no,
            "description": price.description if price else "",
            "qty": qty,
            "rate": price.mrp if price else 0,
            "available": stock.quantity if stock else 0,
            "discount": 0
        })

    db.close()
    return result
@app.get("/get_rate/{part_no}")
def get_rate(part_no: str):
    db = SessionLocal()

    # Step 1: Check order_items (price master)
    item = db.execute(
        text("SELECT mrp, description, hsn FROM order_items WHERE part_no = :p"),
        {"p": part_no}
    ).fetchone()

    rate        = float(item.mrp)   if item and item.mrp         else 0.0
    description = item.description  if item and item.description  else ""
    hsn         = str(item.hsn)     if item and item.hsn          else ""
    stock       = 0

    # Step 2: Fallback to products table if rate is 0 or not found
    product = db.query(Product).filter(Product.part_no == part_no).first()
    if product:
        stock = int(product.quantity or 0)
        if rate == 0:
            rate        = float(product.rate or 0)
            description = description or (product.description or "")
            hsn         = hsn         or (product.hsn         or "")

    db.close()
    return {"rate": rate, "description": description, "hsn": hsn, "stock": stock}


# ---------------- PRICE MASTER (orders.csv) ----------------

@app.get("/reimport_orders")
def reimport_orders():
    """Admin: reload all products from orders.csv into order_items"""
    db = SessionLocal()
    try:
        csv_path = os.path.join(os.path.dirname(__file__), "orders.csv")
        if not os.path.exists(csv_path):
            return {"error": "orders.csv not found"}
        df = pd.read_csv(csv_path, dtype=str, on_bad_lines="skip")
        df.columns = [col.strip() for col in df.columns]
        db.execute(text("DELETE FROM order_items"))
        batch, inserted = [], 0
        for _, row in df.iterrows():
            part_no = str(row.get("Part No", "") or "").strip()
            if not part_no:
                continue
            batch.append({
                "pn":   part_no,
                "desc": str(row.get("Part Desc", "") or "").strip(),
                "hsn":  str(row.get("HSN", "") or "").strip(),
                "mrp":  float(row.get("MRP", 0) or 0)
            })
            if len(batch) >= 500:
                for item in batch:
                    db.execute(text("INSERT INTO order_items (part_no, description, hsn, mrp) VALUES (:pn, :desc, :hsn, :mrp)"), item)
                db.commit()
                inserted += len(batch)
                batch = []
        for item in batch:
            db.execute(text("INSERT INTO order_items (part_no, description, hsn, mrp) VALUES (:pn, :desc, :hsn, :mrp)"), item)
        db.commit()
        inserted += len(batch)
        return {"ok": True, "imported": inserted}
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

@app.get("/search_parts")
def search_parts(q: str = ""):
    """Autocomplete: search order_items by part_no or description"""
    db = SessionLocal()
    q = q.strip().upper()
    results = db.execute(text("""
        SELECT part_no, description, mrp, hsn
        FROM order_items
        WHERE UPPER(part_no) LIKE :q OR UPPER(description) LIKE :q
        LIMIT 20
    """), {"q": f"%{q}%"}).fetchall()
    db.close()
    return [{"part_no": r.part_no, "description": r.description,
             "rate": float(r.mrp or 0), "hsn": str(r.hsn or "")} for r in results]

@app.get("/price_master")
def price_master_list(page: int = 1, q: str = ""):
    """List all products in order_items (price master) with pagination"""
    db = SessionLocal()
    per_page = 100
    offset = (page - 1) * per_page
    q_clean = f"%{q.strip().upper()}%"
    rows = db.execute(text("""
        SELECT id, part_no, description, hsn, mrp
        FROM order_items
        WHERE UPPER(part_no) LIKE :q OR UPPER(description) LIKE :q
        ORDER BY part_no
        LIMIT :lim OFFSET :off
    """), {"q": q_clean, "lim": per_page, "off": offset}).fetchall()
    total = db.execute(text("""
        SELECT COUNT(*) FROM order_items
        WHERE UPPER(part_no) LIKE :q OR UPPER(description) LIKE :q
    """), {"q": q_clean}).scalar()
    db.close()
    return {"total": total, "page": page, "per_page": per_page,
            "items": [{"id": r.id, "part_no": r.part_no, "description": r.description,
                       "hsn": r.hsn, "mrp": float(r.mrp or 0)} for r in rows]}

@app.post("/price_master/add")
async def price_master_add(request: Request):
    data = await request.json()
    db = SessionLocal()
    existing = db.query(OrderItems).filter(OrderItems.part_no == data["part_no"].strip()).first()
    if existing:
        db.close()
        return {"error": "Part No already exists"}
    item = OrderItems(
        part_no=data["part_no"].strip(),
        description=data.get("description", "").strip(),
        hsn=data.get("hsn", "").strip(),
        mrp=float(data.get("mrp", 0))
    )
    db.add(item)
    db.commit()
    db.close()
    return {"ok": True}

@app.post("/price_master/update/{item_id}")
async def price_master_update(item_id: int, request: Request):
    data = await request.json()
    db = SessionLocal()
    item = db.query(OrderItems).filter(OrderItems.id == item_id).first()
    if item:
        item.description = data.get("description", item.description)
        item.hsn = data.get("hsn", item.hsn)
        item.mrp = float(data.get("mrp", item.mrp))
        db.commit()
    db.close()
    return {"ok": True}

@app.post("/price_master/delete/{item_id}")
def price_master_delete(item_id: int):
    db = SessionLocal()
    db.query(OrderItems).filter(OrderItems.id == item_id).delete()
    db.commit()
    db.close()
    return {"ok": True}

# ---------------- QUOTATION (no stock change) ----------------

def generate_quotation_no(db):
    import calendar
    today = datetime.now()
    year_month = today.strftime("%Y-%m")
    first_day = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_num = calendar.monthrange(today.year, today.month)[1]
    last_day = today.replace(day=last_day_num, hour=23, minute=59, second=59, microsecond=999999)
    count = db.execute(text("""
        SELECT COUNT(*) FROM quotations WHERE date >= :fd AND date <= :ld
    """), {"fd": first_day, "ld": last_day}).scalar()
    return f"QT-{year_month}-{str(count+1).zfill(3)}"

@app.post("/save_quotation")
async def save_quotation(request: Request):
    """Save quotation to DB — NO stock change, NO sales report impact"""
    data = await request.json()
    db = SessionLocal()
    rows = data.get("rows", [])
    customer_name = data.get("customer_name", "")
    if not rows:
        return {"error": "No products"}
    total_amount = sum(r["rate"] * r["qty"] * (1 - r.get("discount", 0)/100) for r in rows)
    grand_total = total_amount * 1.18
    q_no = generate_quotation_no(db)
    quot = Quotation(
        quotation_no=q_no,
        customer_name=customer_name,
        date=datetime.now(),
        total_amount=round(total_amount, 2),
        grand_total=round(grand_total, 2)
    )
    db.add(quot)
    db.flush()
    for r in rows:
        sub = r["rate"] * r["qty"] * (1 - r.get("discount", 0)/100)
        db.add(QuotationItem(
            quotation_id=quot.id,
            part_no=r.get("part_no", ""),
            description=r.get("description", ""),
            rate=r["rate"],
            qty=r["qty"],
            discount=r.get("discount", 0),
            amount=round(sub, 2),
            hsn=r.get("hsn", "")
        ))
    db.commit()
    db.close()
    return {"ok": True, "quotation_no": q_no}

@app.get("/quotations")
def list_quotations():
    db = SessionLocal()
    rows = db.execute(text("""
        SELECT id, quotation_no, customer_name, date, grand_total
        FROM quotations ORDER BY id DESC LIMIT 50
    """)).fetchall()
    db.close()
    return [{"id": r.id, "quotation_no": r.quotation_no,
             "customer_name": r.customer_name,
             "date": str(r.date)[:10],
             "grand_total": float(r.grand_total or 0)} for r in rows]

@app.post("/download_quotation_pdf")
async def download_quotation_pdf(request: Request):
    """Generate quotation PDF HTML — NO stock change"""
    data = await request.json()
    rows = data.get("rows", [])
    customer_name = data.get("customer_name", "")
    items, total = [], 0
    for r in rows:
        sub = r["rate"] * r["qty"]
        disc = sub * (r.get("discount", 0) / 100)
        after = sub - disc
        gst_amt = after * 0.18
        final = after + gst_amt
        total += final
        items.append({
            "part_no": r.get("part_no", ""), "description": r.get("description", ""),
            "hsn": r.get("hsn", ""), "rate": r["rate"], "qty": r["qty"],
            "discount": r.get("discount", 0), "gst": 18,
            "amount": round(after, 2), "gst_amount": round(gst_amt, 2), "final": round(final, 2)
        })
    date_str = datetime.now().strftime("%d-%b-%Y")
    html = f"""<!DOCTYPE html><html><head><meta charset='UTF-8'>
<title>Quotation</title>
<style>
  body{{font-family:Arial,sans-serif;margin:30px;font-size:12px;}}
  h2{{text-align:center;color:#1e40af;}} .subtitle{{text-align:center;color:#64748b;margin-bottom:20px;}}
  table{{width:100%;border-collapse:collapse;margin-top:10px;}}
  th{{background:#1e40af;color:#fff;padding:8px;text-align:left;}}
  td{{padding:6px 8px;border-bottom:1px solid #e2e8f0;}}
  tr:nth-child(even){{background:#f8fafc;}}
  .total-row{{font-weight:bold;background:#e0e7ff!important;}}
  .info{{display:flex;justify-content:space-between;margin-bottom:12px;}}
  .badge{{background:#dbeafe;color:#1e40af;padding:2px 8px;border-radius:4px;font-size:11px;}}
  @media print{{button{{display:none}}}}
</style></head><body>
<h2>QUOTATION</h2>
<div class='subtitle'>This is a quotation only — not a tax invoice</div>
<div class='info'>
  <div><b>To:</b> {customer_name or '—'}</div>
  <div><b>Date:</b> {date_str} &nbsp; <span class='badge'>QUOTATION</span></div>
</div>
<table>
<tr><th>#</th><th>Part No</th><th>Description</th><th>HSN</th><th>Rate(₹)</th><th>Qty</th><th>Disc%</th><th>Amount(₹)</th><th>GST(₹)</th><th>Total(₹)</th></tr>
{''.join(f"<tr><td>{i+1}</td><td>{it['part_no']}</td><td>{it['description']}</td><td>{it['hsn']}</td><td>{it['rate']:.2f}</td><td>{it['qty']}</td><td>{it['discount']:.1f}%</td><td>{it['amount']:,.2f}</td><td>{it['gst_amount']:,.2f}</td><td>{it['final']:,.2f}</td></tr>" for i,it in enumerate(items))}
<tr class='total-row'><td colspan='9' style='text-align:right'>Grand Total (incl. 18% GST)</td><td>₹ {total:,.2f}</td></tr>
</table>
<div style='margin-top:20px;font-size:11px;color:#64748b;'>
  <b>Note:</b> Rates subject to change. GST @18% included. Valid for 30 days from date of quotation.
</div>
<br><button onclick='window.print()' style='background:#1e40af;color:#fff;border:none;padding:8px 20px;border-radius:6px;cursor:pointer;'>🖨️ Print / Save PDF</button>
</body></html>"""
    return HTMLResponse(content=html)

# ----------------------------------------------------------------



def generate_invoice_no(db):
    import calendar
    today = datetime.now()

    year_month = today.strftime("%Y-%m")   # 2026-06

    # Python-side date range — works with both SQLite and PostgreSQL
    first_day = today.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
    last_day_num = calendar.monthrange(today.year, today.month)[1]
    last_day = today.replace(day=last_day_num, hour=23, minute=59, second=59, microsecond=999999)

    count = db.execute(text("""
        SELECT COUNT(*) 
        FROM invoices 
        WHERE date >= :from_date AND date <= :to_date
    """), {"from_date": first_day, "to_date": last_day}).scalar()

    serial = str(count + 1).zfill(3)   # 001, 002, 003
    return f"{year_month}-{serial}"
@app.post("/download_pdf")
async def download_pdf(request: Request):

    payload = await request.json()
    # Support both flat list (old) and {rows:[], customer_name:""} (new)
    if isinstance(payload, list):
        data = payload
        customer_name = "Walk-In Customer"
    else:
        data = payload.get("rows", payload) if isinstance(payload.get("rows"), list) else payload
        customer_name = payload.get("customer_name", "").strip() or "Walk-In Customer"
        if isinstance(data, dict):  # still not right, fall back
            data = [payload] if "part_no" in payload else []
    db = SessionLocal()
    from collections import defaultdict

    try:
        # ===============================
        # ✅ STEP 1: VALIDATE STOCK
        # ===============================
        qty_map = defaultdict(float)

        for row in data:
            qty_map[row["part_no"]] += float(row["qty"])

        for part_no, total_qty in qty_map.items():
            product = db.query(Product).filter(Product.part_no == part_no).first()

            if not product or product.quantity < total_qty:
                raise ValueError(f"Not enough stock for {part_no}")

        # ===============================
        # ✅ STEP 2: PROCESS + REDUCE STOCK
        # ===============================
        items = []
        subtotal = 0

        for row in data:

            db_item = db.execute(
                text("SELECT description, mrp, hsn FROM order_items WHERE part_no=:p"),
                {"p": row["part_no"]}
            ).fetchone()

            desc = db_item.description if db_item and db_item.description else row.get("description", "")
            hsn  = str(db_item.hsn)    if db_item and db_item.hsn         else row.get("hsn", "")

            # Rate: orders.csv MRP first, then UI-entered rate, then product purchase rate
            rate = float(db_item.mrp) if db_item and db_item.mrp and float(db_item.mrp) > 0 else float(row.get("rate", 0))
            if rate == 0:
                prod_fallback = db.query(Product).filter(Product.part_no == row["part_no"]).first()
                if prod_fallback:
                    rate = float(prod_fallback.rate or 0)
                    desc = desc or (prod_fallback.description or "")
                    hsn  = hsn  or (prod_fallback.hsn         or "")

            qty  = float(row.get("qty", 1))
            disc = float(row.get("discount", 0))

            # 🔥 UPDATE STOCK
            product = db.query(Product).filter(
                Product.part_no == row["part_no"]
            ).first()

            if product:
                product.quantity -= qty

                if product.quantity <= 0:
                    product.quantity = 0
                    product.amount = 0
                else:
                    product.amount = (product.quantity * product.rate) - (
                        (product.quantity * product.rate) * (product.discount / 100)
        )

            # 🔥 CALCULATIONS
            base = qty * rate
            discount_amt = base * (disc / 100)
            taxable = base - discount_amt

            subtotal += taxable

            items.append({
                "part_no": row["part_no"],
                "description": desc,
                "hsn": hsn,
                "qty": qty,
                "rate": rate,
                "discount": disc,
                "taxable": round(taxable, 2)
            })

        # ===============================
        # ✅ STEP 3: GST
        # ===============================
        cgst = round(subtotal * 0.09, 2)
        sgst = round(subtotal * 0.09, 2)
        total = round(subtotal + cgst + sgst, 2)

        # ===============================
        # ✅ STEP 4: SAVE INVOICE
        # ===============================
        invoice_no = generate_invoice_no(db)

        # Use PostgreSQL-compatible insert with RETURNING, fallback to lastrowid
        from database import engine as _engine
        is_postgres = "postgresql" in str(_engine.url)

        if is_postgres:
            result = db.execute(text("""
                INSERT INTO invoices (
                    invoice_no, customer_name, date,
                    total_amount, gst_amount, grand_total
                ) VALUES (
                    :inv, :cust, :date,
                    :total, :gst, :grand
                ) RETURNING id
            """), {
                "inv": invoice_no, "cust": customer_name,
                "date": datetime.utcnow(),
                "total": round(subtotal, 2),
                "gst": cgst + sgst, "grand": total
            })
            invoice_id = result.fetchone()[0]
        else:
            result = db.execute(text("""
                INSERT INTO invoices (
                    invoice_no, customer_name, date,
                    total_amount, gst_amount, grand_total
                ) VALUES (
                    :inv, :cust, :date,
                    :total, :gst, :grand
                )
            """), {
                "inv": invoice_no, "cust": customer_name,
                "date": datetime.utcnow(),
                "total": round(subtotal, 2),
                "gst": cgst + sgst, "grand": total
            })
            invoice_id = result.lastrowid

        for it in items:
            db.execute(text("""
                INSERT INTO invoice_items 
                (invoice_id, part_no, description, quantity, rate, amount, hsn)
                VALUES (:iid, :p, :d, :q, :r, :amt, :hsn)
            """), {
                "iid": invoice_id,
                "p": it["part_no"],
                "d": it["description"],
                "q": it["qty"],
                "r": it["rate"],
                "amt": it["taxable"],
                "hsn": it["hsn"]
            })

        db.commit()

    except Exception as e:
        db.rollback()
        db.close()
        raise e

    db.close()

    # ===============================
    # ✅ STEP 5: GENERATE PDF
    # ===============================
    env = Environment(loader=FileSystemLoader("templates"))
    template = env.get_template("quotation_pdf.html")

    html = template.render(
        invoice_no=invoice_no,
        items=items,
        subtotal=round(subtotal, 2),
        cgst=cgst,
        sgst=sgst,
        total=total,
        date=datetime.now().strftime("%d-%b-%Y"),
        buyer_name=customer_name,
        buyer_address=""
    )

    # config = pdfkit.configuration(
    #     wkhtmltopdf="/usr/local/bin/wkhtmltopdf"
    # )

    # options = {
    #     "enable-local-file-access": ""
    # }

    # pdf_file = "quotation.pdf"

    # pdfkit.from_string(html, pdf_file, configuration=config, options=options)

    # return FileResponse(pdf_file, media_type="application/pdf", filename="quotation.pdf")
    from fastapi.responses import HTMLResponse

    return HTMLResponse(content=html)

from fastapi import Query
from datetime import datetime

@app.get("/sales_summary")
def sales_summary(
    start_date: str = Query(...),
    end_date: str = Query(...)
):
    db = SessionLocal()

    query = text("""
        SELECT 
            COUNT(*) as total_orders,
            SUM(total_amount) as total_sales,
            SUM(gst_amount) as total_gst,
            SUM(grand_total) as grand_total
        FROM invoices
        WHERE DATE(date) BETWEEN :start AND :end
    """)

    result = db.execute(query, {
        "start": start_date,
        "end": end_date
    }).fetchone()

    db.close()

    return {
        "orders": result.total_orders or 0,
        "sales": float(result.total_sales or 0),
        "gst": float(result.total_gst or 0),
        "grand_total": float(result.grand_total or 0)
    }


@app.get("/import_pdf_data")
def import_pdf_data():
    db = SessionLocal()

    # 🔥 call your existing function
    products = extract_products("products.pdf")  

    for p in products:
        db.add(Product(
            part_no=p.get("part_no"),
            description=p.get("description"),
            hsn=p.get("hsn"),
            gst=float(p.get("gst", 18)),
            quantity=int(p.get("qty", 0)),
            rate=float(p.get("rate", 0)),
            discount=float(p.get("discount", 0)),
            amount=float(p.get("amount", 0))
        ))

    db.commit()
    db.close()

    return {"message": "PDF imported successfully"}
@app.get("/sales")
def get_sales(start: str, end: str):
    db = SessionLocal()

    rows = db.execute(text("""
        SELECT * FROM invoices
        WHERE DATE(created_at) BETWEEN :s AND :e
        ORDER BY created_at DESC
    """), {
        "s": start,
        "e": end
    }).fetchall()

    total_sales = sum([r.total for r in rows])

    db.close()

    return {
        "total_sales": total_sales,
        "count": len(rows),
        "data": [dict(r._mapping) for r in rows]
    }

@app.post("/view_quotation", response_class=HTMLResponse)
async def view_quotation(request: Request):

    data = await request.json()

    items = []
    total = 0

    for item in data:
        subtotal = item["rate"] * item["qty"]
        discount = subtotal * (item["discount"] / 100)
        after = subtotal - discount
        gst = after * 0.18
        final = after + gst

        total += final

    
        db = SessionLocal()

     
        price = db.query(OrderItems).filter(
            OrderItems.part_no == item["part_no"]
        ).first()

        description = price.description if price else ""
        hsn = price.hsn if price else ""
        
        items.append({
            "part_no": item["part_no"],
            "description": description,
            "hsn": hsn,
            "qty": item["qty"],
            "rate": item["rate"],
            "discount": item["discount"],
            "final": round(final, 2)
        })

    return render(
        "quotation.html",
        items=items,
        total=round(total, 2),
        date=datetime.now().strftime("%d-%b-%Y"),
        buyer_name="Thakur Infraprojects Private Limited",
        buyer_address="""Plot No. 265/01, Om Sadanika, Panvel
Uran Road Uran Naka, Panvel, Raigad
GSTIN: 27AACCT6451F1ZC"""
    )
@app.get("/build_quotation")
def build_quotation():
    db = SessionLocal()

    rows = db.execute(text("SELECT * FROM order_items"))
    result = []

    for r in rows:
        stock = db.query(Product).filter(
            Product.part_no == r.part_no
        ).first()

        result.append({
            "part_no": r.part_no,
            "description": r.description,
            "rate": float(r.mrp),   # ✅ FIX
            "stock": stock.quantity if stock else 0,
            "qty": stock.quantity if stock else 0,
            "discount": 0
        })

    db.close()
    return result

@app.post("/upload_order")
def upload_order(file: UploadFile = File(...)):

    if file.filename.endswith(".csv"):
        df = pd.read_csv(file.file)
    else:
        df = pd.read_excel(file.file, engine="openpyxl")

    df.columns = [col.strip().upper() for col in df.columns]

    db = SessionLocal()

    db.execute(text("DELETE FROM order_items"))

    df = df.fillna("")

    for _, row in df.iterrows():

        part_no = str(row.get("PART NO", "")).strip()
        if not part_no:
            continue

        mrp_val = float(str(row.get("MRP", 0)).replace(",", "") or 0)

        db.execute(text("""
            INSERT INTO order_items (
                part_no, description, mrp, hsn
            ) VALUES (
                :part_no, :description, :mrp, :hsn
            )
        """), {
            "part_no": part_no,
            "description": row.get("PART DESC", ""),
            "mrp": mrp_val,
            "hsn": row.get("HSN", "")
        })

    db.commit()
    db.close()

    return {"message": "Full order data saved ✅"}
@app.get("/get_order_items")
def get_order_items():
    db = SessionLocal()

    rows = db.execute(text("SELECT * FROM order_items")).fetchall()
    result = []

    for r in rows:

        stock = db.query(Product).filter(
            Product.part_no == r.part_no
        ).first()

        result.append({
            "part_no": r.part_no,
            "description": r.description,
            "rate": r.mrp,
            "stock": stock.quantity if stock else 0,
            "qty": stock.quantity if stock else 0,   # ✅ real qty
            "discount": 0,
            "status": "OUT OF STOCK" if (not stock or stock.quantity == 0) else "AVAILABLE"
        })

    # ✅ SORT AFTER LOOP
    result = sorted(result, key=lambda x: x["stock"] != 0)

    db.close()
    return result

import os

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(
        "main:app",
        host="0.0.0.0",
        port=int(os.environ.get("PORT", 8000))
    )


   
