import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:

    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_no, date = filename.split("-")

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt="Invoice No." + f"{invoice_no}", ln=1)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt="Date: " + f"{date}", ln=1)

    # Add a header
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    columns_title = df.columns
    columns_title = [item.replace("_", " ").title() for item in columns_title]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=str(columns_title[0]), border=1)
    pdf.cell(w=70, h=8, txt=str(columns_title[1]), border=1)
    pdf.cell(w=35, h=8, txt=str(columns_title[2]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns_title[3]), border=1)
    pdf.cell(w=30, h=8, txt=str(columns_title[4]), border=1, ln=1)

    # Add rows to the table

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=70, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=70, h=8, txt="", border=1)
    pdf.cell(w=35, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(total_sum), border=1, ln=1)

    # Add sum total price
    pdf.set_font(family="Times", style="B", size=10)
    pdf.cell(w=30, h=8, txt=f"The Total Price is {total_sum}", ln=1)

    # Add Company Stamp
    pdf.set_font(family="Times", style="B", size=15)
    pdf.cell(w=25, h=8, txt="Company")
    pdf.image("Logo.png", w=10)

    pdf.output(f"PDFs/{filename}.pdf")