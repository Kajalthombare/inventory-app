import os
import traceback
from sqlalchemy import text
from database import SessionLocal
import google.generativeai as genai

# Configure Gemini
api_key = os.environ.get("GEMINI_API_KEY")
if api_key:
    genai.configure(api_key=api_key)

def get_inventory_stats(store_name: str = "mahindra"):
    """Get overall inventory statistics for the specified store including total unique products,
    number of in-stock items, number of out-of-stock items, and total value of inventory (taxable).
    """
    db = SessionLocal()
    try:
        total_products = db.execute(text("SELECT COUNT(*) FROM products WHERE store = :store"), {"store": store_name}).scalar() or 0
        in_stock = db.execute(text("SELECT COUNT(*) FROM products WHERE store = :store AND quantity > 0"), {"store": store_name}).scalar() or 0
        out_of_stock = db.execute(text("SELECT COUNT(*) FROM products WHERE store = :store AND quantity = 0"), {"store": store_name}).scalar() or 0
        
        # Calculate total value in SQL
        rows = db.execute(text("SELECT quantity, rate, discount FROM products WHERE store = :store AND quantity > 0"), {"store": store_name}).fetchall()
        total_value = sum((r[0] * r[1]) - ((r[0] * r[1]) * (r[2] / 100)) for r in rows)
        
        return {
            "store_id": store_name,
            "total_unique_products": total_products,
            "in_stock_products": in_stock,
            "out_of_stock_products": out_of_stock,
            "total_inventory_value_inr": round(total_value, 2)
        }
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def search_products(query: str, store_name: str = "mahindra"):
    """Search for products in the specified store using a search query (matches part number, description, or store location).
    Returns up to 10 matching products with location information.
    """
    db = SessionLocal()
    try:
        q_clean = f"%{query.strip().upper()}%"
        sql = """
            SELECT part_no, description, hsn, gst, quantity, rate, discount, amount, vendor_name, location 
            FROM products 
            WHERE store = :store AND (UPPER(part_no) LIKE :q OR UPPER(description) LIKE :q OR UPPER(location) LIKE :q)
            LIMIT 10
        """
        rows = db.execute(text(sql), {"q": q_clean, "store": store_name}).fetchall()
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

def get_product_details(part_no: str, store_name: str = "mahindra"):
    """Retrieve detailed information for a specific product in the specified store using its unique part number.
    Returns details including quantity, rate, discount, vendor details, and store location.
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT part_no, description, hsn, gst, quantity, rate, discount, amount, 
                   vendor_name, vendor_address, vendor_mobile, vendor_gstin, vendor_email, location
            FROM products 
            WHERE store = :store AND UPPER(part_no) = :p
        """
        r = db.execute(text(sql), {"p": part_no.strip().upper(), "store": store_name}).fetchone()
        if not r:
            return {"error": f"Product with part number '{part_no}' not found in store '{store_name}'."}
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

def get_low_stock_products(threshold: int = 5, store_name: str = "mahindra"):
    """List products in the specified store whose stock quantity is less than or equal to the threshold (default 5).
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT part_no, description, quantity, rate, vendor_name, location 
            FROM products 
            WHERE store = :store AND quantity <= :t 
            ORDER BY quantity ASC 
            LIMIT 20
        """
        rows = db.execute(text(sql), {"t": threshold, "store": store_name}).fetchall()
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

def get_sales_report(start_date: str, end_date: str, customer_name: str = None, vendor_name: str = None, store_name: str = "mahindra"):
    """Retrieve sales statistics and detailed transactions for specified store between start_date and end_date (YYYY-MM-DD).
    Optionally filter by customer_name or vendor_name.
    """
    db = SessionLocal()
    try:
        sql_parts = [
            "FROM invoice_items ii",
            "JOIN invoices i ON ii.invoice_id = i.id",
            "LEFT JOIN products p ON (ii.part_no = p.part_no AND p.store = :store)",
            "WHERE i.store = :store AND DATE(i.date) BETWEEN :start AND :end"
        ]
        params = {"start": start_date, "end": end_date, "store": store_name}
        
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

def get_purchase_report(start_date: str, end_date: str, vendor_name: str = None, store_name: str = "mahindra"):
    """Retrieve purchase statistics and detailed transactions for specified store between start_date and end_date (YYYY-MM-DD).
    Optionally filter by vendor_name.
    """
    db = SessionLocal()
    try:
        sql = """
            SELECT id, vendor_name, part_no, description, quantity, rate, discount, amount, date 
            FROM purchases 
            WHERE store = :store AND DATE(date) BETWEEN :start AND :end
        """
        params = {"start": start_date, "end": end_date, "store": store_name}
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

def get_vendors(store_name: str = "mahindra"):
    """Retrieve a list of all unique vendor names in the specified store.
    """
    db = SessionLocal()
    try:
        rows = db.execute(text("SELECT DISTINCT name FROM vendors WHERE store = :store ORDER BY name"), {"store": store_name}).fetchall()
        return [r[0] for r in rows if r[0]]
    except Exception as e:
        return {"error": str(e)}
    finally:
        db.close()

def query_chatbot(message: str, history: list = None, store_name: str = "mahindra"):
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
        
        # Bind store_name parameter to tool wrappers
        tools = [
            lambda: get_inventory_stats(store_name=store_name),
            lambda query: search_products(query, store_name=store_name),
            lambda part_no: get_product_details(part_no, store_name=store_name),
            lambda threshold=5: get_low_stock_products(threshold, store_name=store_name),
            lambda start_date, end_date, customer_name=None, vendor_name=None: get_sales_report(start_date, end_date, customer_name, vendor_name, store_name=store_name),
            lambda start_date, end_date, vendor_name=None: get_purchase_report(start_date, end_date, vendor_name, store_name=store_name),
            lambda: get_vendors(store_name=store_name)
        ]
        
        # Set proper function names for docstring inspection
        tools[0].__name__ = "get_inventory_stats"
        tools[0].__doc__ = get_inventory_stats.__doc__
        tools[1].__name__ = "search_products"
        tools[1].__doc__ = search_products.__doc__
        tools[2].__name__ = "get_product_details"
        tools[2].__doc__ = get_product_details.__doc__
        tools[3].__name__ = "get_low_stock_products"
        tools[3].__doc__ = get_low_stock_products.__doc__
        tools[4].__name__ = "get_sales_report"
        tools[4].__doc__ = get_sales_report.__doc__
        tools[5].__name__ = "get_purchase_report"
        tools[5].__doc__ = get_purchase_report.__doc__
        tools[6].__name__ = "get_vendors"
        tools[6].__doc__ = get_vendors.__doc__

        store_title = "Mahindra Pro Spares" if store_name == "mahindra" else "Divya Automobiles"

        model = genai.GenerativeModel(
            model_name='gemini-2.5-flash',
            tools=tools,
            system_instruction=(
                f"You are InvenPro Assistant, an intelligent chatbot integrated into the InvenPro Inventory Management System. "
                f"You are currently assisting for the store: '{store_title}'. "
                f"You help warehouse administrators manage and analyze inventory, stock, sales, and purchases for '{store_title}'. "
                f"You have access to database tools that retrieve live inventory and sales data for '{store_title}'. "
                f"Always call the appropriate tool when the user asks for stock statistics, product lookups, low stock, sales reports, or purchase reports. "
                f"If the tool returns an error or no data, explain it politely. "
                f"Format numbers nicely in INR (e.g. ₹ 12,34,567.89 or using the Indian currency system when appropriate). "
                f"Be brief, precise, helpful, and polite. Use markdown tables to present tabular data when helpful."
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
