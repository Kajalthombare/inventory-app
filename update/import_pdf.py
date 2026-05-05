import pdfplumber
import re
import pandas as pd
from sqlalchemy import create_engine

engine = create_engine("mysql+pymysql://root:1234@localhost/inventory")

data = []

def parse_number(x):
    return float(x.replace(",", ""))

with pdfplumber.open("products.pdf") as pdf:
    for page in pdf.pages:
        text = page.extract_text() or ""
        lines = text.split("\n")

        i = 0
        while i < len(lines):
            line = lines[i]

            # skip headers
            if "PART NO" in line or "Description" in line or "Amount" in line:
                i += 1
                continue

            # match main row (without amount)
            match = re.search(
                r'^\d+\s+([A-Z0-9]+)\s+(.*?)\s+(\d{7,8})\s+(\d+)\s*%\s+(\d+)\s+NOS\s+([\d,]+\.\d+)\s+NOS\s+(\d+)\s*%',
                line
            )

            if match:
                part_no = match.group(1)
                description = match.group(2)
                hsn = match.group(3)
                gst = float(match.group(4))
                quantity = int(match.group(5))
                rate = parse_number(match.group(6))
                discount = float(match.group(7))

                # 🔥 get amount from NEXT LINE
                amount = 0
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if re.match(r'^[\d,]+\.\d+$', next_line):
                        amount = parse_number(next_line)
                        i += 1  # skip next line

                data.append({
                    "part_no": part_no,
                    "description": description.strip(),
                    "hsn": hsn,
                    "gst": gst,
                    "quantity": quantity,
                    "rate": rate,
                    "discount": discount,
                    "amount": amount
                })

            else:
                print("Skipping:", line)

            i += 1

# Convert
df = pd.DataFrame(data)

print(df)
print("Total rows:", len(df))

# Insert
df.to_sql("products", con=engine, if_exists="append", index=False)

print("✅ IMPORT SUCCESS 🚀")