import os
import traceback
from sqlalchemy import text
from database import SessionLocal
import google.generativeai as genai

# Configure Gemini
api_key = os.environ.get("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)

def get_inventory_stats():
    """Get overall inventory statistics including total unique products, number of in-stock items,
    number of out-of-stock items, and total value of the inventory (taxable).
    """
    db = SessionLocal()
    try:
        total_products = db.execute(text("SELECT COUNT(*) FROM products")).scalar() or 0
        in_stock = db.execute(text("SELECT COUNT(*) FROM products WHERE quantity > 0")).scalar() or 0
        out_of_stock = db.execute(text("SELECT COUNT(*) FROM products WHERE quantity = 0")).scalar() or 0
        
        # Calculate total value in SQL
        rows = db.execute(text("SELECT quantity, rate, discount FROM products WHERE quantity > 0")).fetchall()
        total_value = sum((r[0] * r[1]) - ((r[0] * r[1]) * (r[2] / 100)) for r in rows)
        
        return {
            "total_unique_products": total_products,
            "in_stock_products": in_stock,
            "out_of_stock_products": out_of_stock,
            "total_inventory_value_inr": round(total_value, 2)
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def search_products(query: str):
    """Search for products in the inventory using a search query (matches part number, description, or store location).
    Returns up to 10 matching products with location information.
    """
    db = SessionLocal()
    try:
        q_clean = f"%{query.strip().upper()}%"
        sql = """
            SELECT part_no, description, hsn, gst, quantity, rate, discount, amount, vendor_name, location 
            FROM products 
            WHERE UPPER(part_no) LIKE :q OR UPPER(description) LIKE :q OR UPPER(location) LIKE :q 
            LIMIT 10
        """
        rows = db.execute(text(sql), {"q": q_clean}).fetchall()
        result = []
        for r in rows:
            result.append({
                "part_no": r[0],
                "description": r[1],
                "hsn": r[2],
                "gst_percent": r[3],
                "quantity": r[4],
                "rate": r[5],
                "discount_percent": r[6],
                "amount": r[7],
                "vendor_name": r[8],
                "store_location": r[9] or "Unassigned"
            })
        return result
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def get_product_details(part_no: str):
    """Retrieve detailed information for a specific product using its unique part number.
    Returns details including quantity, rate, discount, vendor details, and store location.
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT part_no, description, hsn, gst, quantity, rate, discount, amount, 
                   vendor_name, vendor_address, vendor_mobile, vendor_gstin, vendor_email, location
            FROM products 
            WHERE UPPER(part_no) = :p
        """
        r = db.execute(text(sql), {"p": part_no.strip().upper()}).fetchone()
        if not r:
            return {"error": f"Product with part number '{part_no}' not found."}
        return {
            "part_no": r[0],
            "description": r[1],
            "hsn": r[2],
            "gst_percent": r[3],
            "quantity": r[4],
            "rate": r[5],
            "discount_percent": r[6],
            "amount": r[7],
            "vendor_name": r[8],
            "vendor_address": r[9],
            "vendor_mobile": r[10],
            "vendor_gstin": r[11],
            "vendor_email": r[12],
            "store_location": r[13] or "Unassigned"
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def get_low_stock_products(threshold: int = 5):
    """List products whose stock quantity is less than or equal to the threshold (default 5).
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT part_no, description, quantity, rate, vendor_name, location 
            FROM products 
            WHERE quantity <= :t 
            ORDER BY quantity ASC 
            LIMIT 20
        """
        rows = db.execute(text(sql), {"t": threshold}).fetchall()
        result = []
        for r in rows:
            result.append({
                "part_no": r[0],
                "description": r[1],
                "quantity": r[2],
                "rate": r[3],
                "vendor_name": r[4],
                "store_location": r[5] or "Unassigned"
            })
        return result
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def get_sales_report(start_date: str, end_date: str, customer_name: str = None, vendor_name: str = None):
    """Retrieve sales statistics and detailed transactions between start_date and end_date (YYYY-MM-DD).
    Optionally filter by customer_name or vendor_name.
    """
    db = SessionLocal()
    try:
        sql_parts = [
            "FROM invoice_items ii",
            "JOIN invoices i ON ii.invoice_id = i.id",
            "LEFT JOIN products p ON ii.part_no = p.part_no",
            "WHERE DATE(i.date) BETWEEN :start AND :end"
        ]
        params = {"start": start_date, "end": end_date}
        
        if customer_name:
            sql_parts.append("AND LOWER(i.customer_name) LIKE LOWER(:customer)")
            params["customer"] = f"%{customer_name.strip()}%"
            
        if vendor_name:
            sql_parts.append("AND p.vendor_name = :vendor")
            params["vendor"] = vendor_name.strip()

        query = f"""
            SELECT 
                i.invoice_no,
                i.customer_name,
                ii.part_no,
                ii.description,
                ii.quantity,
                ii.rate,
                ii.amount as taxable_amount,
                COALESCE(ii.purchase_rate, 0.0) as purchase_rate,
                i.date
            {" ".join(sql_parts)}
            ORDER BY i.date DESC
        """
        result = db.execute(text(query), params).fetchall()
        
        rows = []
        for r in result:
            taxable = float(r[6] or 0.0)
            gst = taxable * 0.18
            grand = taxable + gst
            cgst = gst / 2
            sgst = gst / 2
            qty = float(r[4] or 0.0)
            prate = float(r[7] or 0.0)
            pcost = prate * qty
            profit = taxable - pcost
            
            rows.append({
                "invoice_no": r[0],
                "customer": r[1],
                "part_no": r[2],
                "description": r[3],
                "qty": qty,
                "selling_rate": r[5],
                "taxable_amount": round(taxable, 2),
                "cgst": round(cgst, 2),
                "sgst": round(sgst, 2),
                "grand_total": round(grand, 2),
                "purchase_rate": prate,
                "total_purchase_cost": round(pcost, 2),
                "profit": round(profit, 2),
                "date": r[8].strftime("%Y-%m-%d") if r[8] else ""
            })
            
        total_orders = len(set(r["invoice_no"] for r in rows)) if rows else 0
        total_sales = sum(r["taxable_amount"] for r in rows) if rows else 0.0
        total_cgst = sum(r["cgst"] for r in rows) if rows else 0.0
        total_sgst = sum(r["sgst"] for r in rows) if rows else 0.0
        grand_total = sum(r["grand_total"] for r in rows) if rows else 0.0
        total_cost = sum(r["total_purchase_cost"] for r in rows) if rows else 0.0
        total_profit = sum(r["profit"] for r in rows) if rows else 0.0
        
        return {
            "summary": {
                "total_orders": total_orders,
                "total_sales_taxable": round(total_sales, 2),
                "total_cgst": round(total_cgst, 2),
                "total_sgst": round(total_sgst, 2),
                "grand_total": round(grand_total, 2),
                "total_purchase_cost": round(total_cost, 2),
                "total_profit": round(total_profit, 2)
            },
            "transactions": rows[:15] # Return first 15 rows for context
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def get_purchase_report(start_date: str, end_date: str, vendor_name: str = None):
    """Retrieve purchase statistics and detailed transactions between start_date and end_date (YYYY-MM-DD).
    Optionally filter by vendor_name.
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT id, vendor_name, part_no, description, quantity, rate, discount, amount, date 
            FROM purchases 
            WHERE DATE(date) BETWEEN :start AND :end
        """
        params = {"start": start_date, "end": end_date}
        if vendor_name:
            sql += " AND LOWER(vendor_name) LIKE LOWER(:vendor)"
            params["vendor"] = f"%{vendor_name.strip()}%"
        sql += " ORDER BY date DESC"
        
        rows = db.execute(text(sql), params).fetchall()
        result = []
        for r in rows:
            result.append({
                "id": r[0],
                "vendor_name": r[1],
                "part_no": r[2],
                "description": r[3],
                "quantity": r[4],
                "rate": r[5],
                "discount_percent": r[6],
                "amount": r[7],
                "date": r[8].strftime("%Y-%m-%d") if r[8] else ""
            })
        
        total_spend = sum(r["amount"] for r in result)
        return {
            "summary": {
                "total_purchases_count": len(result),
                "total_spend": round(total_spend, 2)
            },
            "transactions": result[:15]
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def get_vendors():
    """Retrieve a list of all unique vendor names in the system.
    """
    db = SessionLocal()
    try:
        rows = db.execute(text("SELECT DISTINCT name FROM vendors ORDER BY name")).fetchall()
        return [r[0] for r in rows if r[0]]
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def query_chatbot(message: str, history: list = None):
    """Start/continue a chat with Google Gemini and automatic tool calls to solve inventory user query.
    """
    api_key = os.environ.get("GEMINI_API_KEY")
    if not api_key:
        return (
            "⚠️ **Gemini API key is not configured.**\n\n"
            "To activate the chatbot, please configure the `GEMINI_API_KEY` environment variable in your `.env` file.\n"
            "Example:\n`GEMINI_API_KEY=your_actual_api_key_here`"
        )
        
    try:
        genai.configure(api_key=api_key)
        
        tools = [
            get_inventory_stats,
            search_products,
            get_product_details,
            get_low_stock_products,
            get_sales_report,
            get_purchase_report,
            get_vendors
        ]
        
        model = genai.GenerativeModel(
            model_name='gemini-2.5-flash',
            tools=tools,
            system_instruction=(
                "You are InvenPro Assistant, an intelligent chatbot integrated into the InvenPro Inventory Management System. "
                "You help warehouse administrators manage and analyze inventory, stock, sales, and purchases. "
                "You have access to database tools that retrieve live inventory and sales data. "
                "Always call the appropriate tool when the user asks for stock statistics, product lookups, low stock, sales reports, or purchase reports. "
                "If the tool returns an error, explain it politely. "
                "Format numbers nicely in INR (e.g. ₹ 12,34,567.89 or using the Indian currency system when appropriate). "
                "Be brief, precise, helpful, and polite. Use markdown tables to present tabular data when helpful."
            )
        )
        
        contents = []
        if history:
            for item in history:
                role = "user" if item.get("role") == "user" else "model"
                text_content = item.get("text", "")
                if text_content.strip():
                    contents.append({
                        "role": role,
                        "parts": [text_content]
                    })
                
        chat = model.start_chat(history=contents, enable_automatic_function_calling=True)
        response = chat.send_message(message)
        return response.text
    except Exception as e:
        traceback.print_exc()
        return f"❌ **Error running Gemini Chatbot:** {str(e)}"
