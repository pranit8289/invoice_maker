import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import os

# Output folder
output_folder = "Invoices"
os.makedirs(output_folder, exist_ok=True)

def format_address(address, max_length=50):
    if not isinstance(address, str):
        return [""]
    words = address.split()
    lines, line = [], ""
    for word in words:
        if len(line + " " + word) > max_length:
            lines.append(line)
            line = word
        else:
            line += " " + word if line else word
    if line:
        lines.append(line)
    return lines

bank_details = """
Cheque to be drawn in the favor of URBANLEAFSPACE LLP

Online Transfer Details:
Acc Name : URBANLEAF SPACE LLP
Account Number: 539705000135
IFSC Code: ICIC0005397
Bank: ICICI
Branch: Pancard Club Road, Baner Pune
"""

def generate_invoice(data):
    invoice_number = str(data["Invoice Number"])
    month_year = data['Invoice Date'].strftime('%b_%Y')
    buyer_name_safe = str(data['Buyer Name']).replace(" ", "_")[:30]
    pdf_filename = f"{output_folder}/Invoice_{invoice_number}_{month_year}_{buyer_name_safe}.pdf"

    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4

    # Logo
    try:
        logo = ImageReader(logo_path)
        c.drawImage(logo, 50, height - 80, width=100, height=50)
    except:
        print("⚠️ Logo not found.")

    new_invoice_number = f"URBN/AI/25-26/{invoice_number}"

    # Header
    c.setFont("Helvetica-Bold", 18)
    c.drawString(230, height - 60, "INVOICE")
    c.setFont("Helvetica", 10)
    c.drawString(400, height - 100, f"Date: {data['Invoice Date'].strftime('%d-%b-%Y')}")
    c.drawString(400, height - 120, f"Invoice No: {new_invoice_number}")

    # Seller Info
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 140, "From:")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 160, data["Seller Name"])
    y = height - 180
    for line in format_address(data["Seller Address"]):
        c.drawString(50, y, line)
        y -= 20
    c.drawString(50, y, f"GSTIN: {data['Seller GSTIN']}")

    # Buyer Info
    c.setFont("Helvetica-Bold", 12)
    c.drawString(300, height - 140, "Billed To:")
    c.setFont("Helvetica", 10)
    c.drawString(300, height - 160, data["Buyer Name"])
    y = height - 180
    for line in format_address(data["Buyer Address"]):
        c.drawString(300, y, line)
        y -= 20
    c.drawString(300, y, f"GSTIN: {data['Buyer GSTIN']}")

    # Table
    table_data = [["Item", "Quantity", "Unit Price", "Total Amount"]]
    base_amount = data["Base Amount"]
    table_data.append([data["Item Name"], data["Quantity"], f"{data['Unit Price']:.2f}", f"{base_amount:.2f}"])

    has_extra = False
    if data.get("Extra Qty", 0) > 0:
        extra_label = data.get("Extra Label Name", "Extra")
        extra_total = data["Extra Qty"] * data["Extra Price"]
        table_data.append([extra_label, data["Extra Qty"], f"{data['Extra Price']:.2f}", f"{extra_total:.2f}"])
        base_amount += extra_total
        has_extra = True
    

    discount = data.get("Discount", 0)
    if discount:
        table_data.append(["Base Price", "", "", f"{base_amount:.2f}"])
    has_discount = discount > 0
    if has_discount:
        table_data.append(["Discount", "1", f"{discount:.2f}", f"-{discount:.2f}"])
        base_amount -= discount

    table_data.append(["Final Base Price", "", "", f"{base_amount:.2f}"])

    # Taxes
    grand_total = base_amount
    for label in ["CGST", "SGST", "IGST"]:
        percent = data.get(label, 0)
        if percent > 0:
            amount = (base_amount * percent) / 100
            table_data.append([label, "1", f"{percent}%", f"{amount:.2f}"])
            grand_total += amount

    table_data.append(["Grand Total", "", "", f"{grand_total:.2f}"])

    # Draw table
    table = Table(table_data, colWidths=[200, 80, 100, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0a0a6c")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
    ]))

    # Dynamic offset
    extra_rows = int(has_extra) + int(has_discount)
    y_offset = 360 + (extra_rows * 20)
    table.wrapOn(c, width, height)
    table.drawOn(c, 50, height - y_offset)

    # Payment Method - UPDATED
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - (y_offset + 30), "Payment Method:")
    c.setFont("Helvetica", 10)
    c.drawString(200, height - (y_offset + 30), "UPI / Bank Transfer")  # ✅ no 'Cash'

    # Bank Details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - (y_offset + 45), "Bank Account Details:")
    c.setFont("Helvetica", 10)
    y = height - (y_offset + 60)
    for line in bank_details.strip().splitlines():
        c.drawString(50, y, line.strip())
        y -= 15

    # Footer
    c.setFont("Helvetica-Bold", 15)
    c.drawString(200, 70, "Thank you for beliving in us!")

    c.showPage()
    c.save()
    print(f"✅ Invoice generated: {pdf_filename}")

# Main
if __name__ == "__main__":
    logo_path = "Kodesk-Logo.jpg"
    excel_file = "invoice_data_May_2025.xlsx"

    df = pd.read_excel(excel_file)
    print(f"Loaded {len(df)} rows from Excel.")
    for _, row in df.iterrows():
        generate_invoice(row)
    print("✅ All invoices saved in the 'Invoices' folder.")
