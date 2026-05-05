from fastapi import FastAPI, Form, Request, UploadFile, File
from fastapi.responses import HTMLResponse, RedirectResponse, StreamingResponse
from sqlalchemy import create_engine, Column, Integer, String, Float
from sqlalchemy.orm import sessionmaker, declarative_base
from fastapi.templating import Jinja2Templates
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
# ---------------- PROCESS ORDER ----------------

app = FastAPI()
templates = Jinja2Templates(directory="templates")
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

@app.get("/", response_class=HTMLResponse)
def home(request: Request):
    db = SessionLocal()

    products = db.query(Product).all()
    total_value = sum(p.amount or 0 for p in products)

    db.close()

    html = templates.get_template("index.html").render(
        request=request,
        products=products,
        total_value=round(total_value, 2)
    )

    return HTMLResponse(content=html)
# ---------------- PRICE MASTER UPLOAD ----------------


from fastapi.responses import StreamingResponse
import pandas as pd
import io

@app.get("/export_excel")
def export_excel(start_date: str, end_date: str):

    db = SessionLocal()

    query = text("""
        SELECT invoice_no, customer_name, total, cgst, sgst, created_at
        FROM invoices
        WHERE DATE(created_at) BETWEEN :start AND :end
    """)

    result = db.execute(query, {
        "start": start_date,
        "end": end_date
    }).fetchall()

    db.close()

    # -------------------------
    # DATAFRAME (Sales Data)
    # -------------------------
    df = pd.DataFrame(result, columns=[
        "Invoice No", "Customer", "Total", "CGST", "SGST", "Date"
    ])

    # -------------------------
    # SUMMARY CALCULATION
    # -------------------------
    total_orders = len(df)
    total_sales = df["Total"].sum()
    total_cgst = df["CGST"].sum()
    total_sgst = df["SGST"].sum()
    grand_total = total_sales + total_cgst + total_sgst
    # ✅ ADD TOTAL ROW BEFORE EXPORT
    df.loc[len(df)] = [
        "",
        "GRAND TOTAL",
        grand_total,
        "",
        "",
        ""
    ]

    summary_df = pd.DataFrame({
    "Metric": [
        "Total Orders",
        "Total Sales (Before GST)",
        "Total CGST",
        "Total SGST",
        "Grand Total (With GST)"
    ],
    "Value": [
        total_orders,
        total_sales,
        total_cgst,
        total_sgst,
        grand_total
    ]
})

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

    # ✅ Recalculate amount
    subtotal = product.quantity * product.rate
    discount_amt = subtotal * (product.discount / 100)
    product.amount = subtotal - discount_amt

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
    
    if product.quantity <= 0:
        product.quantity = 0
        product.amount = 0 

    # ✅ Reduce stock
    product.quantity -= qty

    # ✅ Recalculate amount
    subtotal = product.quantity * product.rate
    discount_amt = subtotal * (product.discount / 100)
    product.amount = subtotal - discount_amt

    db.commit()
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
    if price:
        rate = price.mrp

    existing = db.query(Product).filter(Product.part_no == part_no).first()

    if existing:
        existing.quantity += quantity

        subtotal = existing.quantity * existing.rate
        discount_amt = subtotal * (existing.discount / 100)
        existing.amount = subtotal - discount_amt
    else:

        subtotal = quantity * rate
        discount_amt = subtotal * (discount / 100)
        amount = subtotal - discount_amt
        db.add(Product(
            part_no=part_no,
            description=description,
            hsn=hsn,
            gst=gst,
            quantity=quantity,
            rate=rate,
            discount=discount,
            amount=amount
        ))

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

    item = db.execute(
        text("SELECT mrp FROM order_items WHERE part_no = :p"),
        {"p": part_no}
    ).fetchone()

    rate = float(item.mrp) if item else 0

    db.close()
    return {"rate": rate}

# ---------------- DOWNLOAD QUOTATION ----------------
from datetime import datetime
from sqlalchemy import text


def generate_invoice_no(db):
    today = datetime.now()

    year_month = today.strftime("%Y-%m")   # 2026-05

    # Count how many invoices already exist this month
    count = db.execute(
        text("""
            SELECT COUNT(*) 
            FROM invoices 
            WHERE DATE_FORMAT(created_at, '%Y-%m') = :ym
        """),
        {"ym": year_month}
    ).scalar()

    serial = str(count + 1).zfill(3)   # 001, 002, 003

    return f"{year_month}-{serial}"
@app.post("/download_pdf")
async def download_pdf(request: Request):

    data = await request.json()
    db = SessionLocal()

    from collections import defaultdict
    from sqlalchemy.exc import SQLAlchemyError

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

            desc = db_item.description if db_item else ""
            hsn = db_item.hsn if db_item else ""
            rate = db_item.mrp if db_item else float(row["rate"])

            qty = float(row["qty"])
            disc = float(row["discount"])

            # 🔥 UPDATE STOCK
            product = db.query(Product).filter(
                Product.part_no == row["part_no"]
            ).first()

            if product:
                product.quantity -= qty

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

        result = db.execute(text("""
            INSERT INTO invoices (invoice_no, customer_name, total, cgst, sgst)
            VALUES (:inv, :cust, :total, :cgst, :sgst)
        """), {
            "inv": invoice_no,
            "cust": "Thakur Infraprojects Private Limited",
            "total": total,
            "cgst": cgst,
            "sgst": sgst
        })

        invoice_id = result.lastrowid

        for it in items:
            db.execute(text("""
                INSERT INTO invoice_items 
                (invoice_id, part_no, description, qty, rate, discount, amount)
                VALUES (:iid, :p, :d, :q, :r, :disc, :amt)
            """), {
                "iid": invoice_id,
                "p": it["part_no"],
                "d": it["description"],
                "q": it["qty"],
                "r": it["rate"],
                "disc": it["discount"],
                "amt": it["taxable"]
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
        invoice_no=invoice_no,   # 🔥 IMPORTANT FIX
        items=items,
        subtotal=round(subtotal, 2),
        cgst=cgst,
        sgst=sgst,
        total=total,
        date=datetime.now().strftime("%d-%b-%Y"),
        buyer_name="Thakur Infraprojects Private Limited",
        buyer_address="""Plot No. 265/01, Om Sadanika, Panvel
Uran Road Uran Naka, Panvel, Raigad
GSTIN: 27AACCT6451F1ZC"""
    )

    config = pdfkit.configuration(
        wkhtmltopdf="/usr/local/bin/wkhtmltopdf"
    )

    options = {
        "enable-local-file-access": ""
    }

    pdf_file = "quotation.pdf"

    #pdfkit.from_string(html, pdf_file, configuration=config, options=options)

    #return FileResponse(pdf_file, media_type="application/pdf", filename="quotation.pdf")
    return {"message": "PDF generation disabled on server"}

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
            SUM(total) as total_sales,
            SUM(cgst) as total_cgst,
            SUM(sgst) as total_sgst
        FROM invoices
        WHERE DATE(created_at) BETWEEN :start AND :end
    """)

    result = db.execute(query, {
        "start": start_date,
        "end": end_date
    }).fetchone()

    db.close()

    return {
        "orders": result.total_orders or 0,
        "sales": float(result.total_sales or 0),
        "cgst": float(result.total_cgst or 0),
        "sgst": float(result.total_sgst or 0)
    }
import_pdf = None
from import_pdf import extract_products  # assuming your function name



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

    return templates.TemplateResponse("quotation.html", {
        "request": request,
        "items": items,
        "total": round(total, 2),
        "date": datetime.now().strftime("%d-%b-%Y"),
        "buyer_name": "Thakur Infraprojects Private Limited",
        "buyer_address": """Plot No. 265/01, Om Sadanika, Panvel
Uran Road Uran Naka, Panvel, Raigad
GSTIN: 27AACCT6451F1ZC"""
    })
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


   
