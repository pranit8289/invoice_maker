import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import qrcode
import os

# Create 'Invoices' folder if not exists
output_folder = "Invoices"
if not os.path.exists(output_folder):
    os.makedirs(output_folder)

# Function to split long addresses
def format_address(address, max_length=50):
    words = address.split()
    lines = []
    current_line = ""
    for word in words:
        if len(current_line) + len(word) + 1 > max_length:
            lines.append(current_line)
            current_line = word
        else:
            current_line += " " + word if current_line else word
    lines.append(current_line)
    return lines

# Function to generate invoice PDF
def generate_invoice(data):
    invoice_number = str(data["Invoice Number"])
    month_year = data['Invoice Date'].strftime('%b_%Y')
    com_name = str(data["Buyer Name"])
    pdf_filename = f"{output_folder}/Invoice_{invoice_number}_{month_year}_{com_name}.pdf"
    
    c = canvas.Canvas(pdf_filename, pagesize=A4)
    width, height = A4
    
    # Add Company Logo
    try:
        logo = ImageReader(logo_path)
        c.drawImage(logo, 30, height - 100, width=150, height=75,)
    except:
        print("Logo not found. Skipping logo placement.")

    new_invoice_number = f"URBN/AI/25-26/{invoice_number}"

    # Invoice Header
    c.setFont("Helvetica-Bold", 28)
    c.drawString(230, height - 120, "INVOICE")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(400, height - 140, f"Date: {data['Invoice Date'].strftime('%d-%b-%Y')}")
    c.drawString(400, height - 160, f"Invoice No: {new_invoice_number}")
    
    # Seller & Buyer Details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 180, "From:")
    c.setFont("Helvetica", 10)
    c.drawString(50, height - 200, f"{data['Seller Name']}")
    y_position = height - 220
    for line in format_address(data['Seller Address']):
        c.drawString(50, y_position, line)
        y_position -= 20
    c.drawString(50, y_position, f"GSTIN: {data['Seller GSTIN']}")

    c.setFont("Helvetica-Bold", 12)
    c.drawString(300, height - 180, "Billed To:")
    c.setFont("Helvetica", 10)
    c.drawString(300, height - 200, f"{data['Buyer Name']}")
    y_position = height - 220
    # for line in format_address(data['Buyer Address']):
    #     c.drawString(300, y_position, line)
    #     y_position -= 20
    c.drawString(300, y_position, f"GSTIN: {data['Buyer GSTIN']}")
    y_position -= 80

    # Table Header
    table_data = [["Item", "Quantity", "Unit Price", "Total Amount"]]
    
    # Adding Items
    table_data.append([
        data['Item Name'], data['Quantity'], f"{data['Unit Price']:.2f}", f"{data['Base Amount']:.2f}"
    ])
    
    # GST Calculation
    gst_rate = data.get('GST Rate', 18)  # Default to 18% if not specified
    gst_amount = (data['Total Amount'] * gst_rate) / 100
    
    cgst = data.get('CGST', 0)
    sgst = data.get('SGST', 0)
    igst = data.get('IGST', 0)
    extras = data.get('Extra', 0)
    discount = data.get('Discount', 0)

    if extras > 0:
        extra_Qty = data["Extra Qty"]
        table_data.append([data["Label Name"], f"{extra_Qty}", f"{data['Extra Price']}", "{:.2f}".format(extras)])

    table_data.append(["Base Price", "", "", "{:.2f}".format(extras + data['Base Amount'])])
    
    if discount:
        table_data.append(["Discount", "1", f"{data['Discount']}", "{:.2f}".format(extras + data['Discount'])])

    table_data.append(["Final Base Price", "", "", "{:.2f}".format(extras + data['Total Amount'])])

    
  
    if cgst > 0:
        table_data.append(["CGST", "1", f"{cgst}%", "{:.2f}".format((data['Total Amount'] * cgst) / 100)])
    if sgst > 0:
        table_data.append(["SGST", "1", f"{sgst}%", "{:.2f}".format((data['Total Amount'] * sgst) / 100)])
    if igst > 0:
        table_data.append(["IGST", "1", f"{igst}%", "{:.2f}".format((data['Total Amount'] * igst) / 100)])
    
    # table_data.append(["GST", "1", f"{gst_rate}%", f"{gst_amount:.2f}"])  
    table_data.append(["Grand Total", "", "", f"{data['Grand Total']:.2f}"])

    
    # Draw Table
    table = Table(table_data, colWidths=[200, 80, 100, 100])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), "#0a0a6c"),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold', 12),  # Bold font for last row (Grand Total)
    ]))
    table.wrapOn(c, width, height)
    if extras > 0:
        table.drawOn(c, 50, height - 410)
    elif discount>0:
        table.drawOn(c, 50, height - 430)
    elif extras and discount>0:
        table.drawOn(c, 50, height - 450)
    else:
        table.drawOn(c, 50, height - 390)

    # Payment Method
    c.setFont("Helvetica-Bold", 12)
    if discount>0:
        table.drawOn(c, 50, height - 430)
    else:
        c.drawString(50, height - 440, "Payment Method: Bank Transfer")

    # Bank Details
    c.setFont("Helvetica-Bold", 12)
    c.drawString(50, height - 455, "Bank Account Details:")
    
    c.setFont("Helvetica", 10)
    y_position = height - 465
    for line in bank_details.split("\n"):
        c.drawString(38, y_position, line)
        y_position -= 10

    # Footer - Thank You Note
    c.setFont("Helvetica-Bold", 15)
    c.drawString(200, 100, "Thank you for beliving in us!")

    c.showPage()
    c.save()
    print(f"Invoice {new_invoice_number} generated successfully.")


if __name__ == "__main__":

    excel_file = "invoice_data_May_2025.xlsx"
    df = pd.read_excel(excel_file)

    # Company Logo Path
    logo_path = "Kodesk-Logo.jpg"

    bank_details = """
    Cheque to be drawn in the favor of URBANLEAFSPACE LLP

    Online Transfer Details:
    Acc Name : URBANLEAF SPACE LLP
    Account Number: 539705000135
    IFSC Code: ICIC0005397
    Bank: ICICI
    Branch: Pancard Club Road, Baner Pune
    """

    # Process each row in Excel file
    for index, row in df.iterrows():
        generate_invoice(row)

    print(f"All invoices saved in the '{output_folder}' folder.")
    
