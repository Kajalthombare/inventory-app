import pdfplumber
import re

def parse_number(x):
    return float(x.replace(",", ""))

def extract_products(pdf_path):
    data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            lines = text.split("\n")

            i = 0
            while i < len(lines):
                line = lines[i]

                if "PART NO" in line or "Description" in line:
                    i += 1
                    continue

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

                    amount = 0
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        if re.match(r'^[\d,]+\.\d+$', next_line):
                            amount = parse_number(next_line)
                            i += 1

                    data.append({
                        "part_no": part_no,
                        "description": description.strip(),
                        "hsn": hsn,
                        "gst": gst,
                        "qty": quantity,
                        "rate": rate,
                        "discount": discount,
                        "amount": amount
                    })

                i += 1

    return data