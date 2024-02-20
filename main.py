from fpdf import FPDF  # Import the FPDF library
import pandas as pd  # Import the pandas library
import glob  # Import the glob module
from pathlib import Path  # Import the Path class from the pathlib module

# Get a list of file paths for all Excel files in the "invoices" directory
filepaths = glob.glob("invoices/*.xlsx")

# Loop through each Excel file
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    # Create a new PDF object with A4 format
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # Add a new page to the PDF
    pdf.add_page()

    # Extract the filename without extension
    filename = Path(filepath).stem
    # Extract the invoice number from the filename
    invoice_nr, date = filename.split("-")

    # Set font for the PDF text
    pdf.set_font(family="Times", size=16, style="B")
    # Add text to the PDF indicating the invoice number
    pdf.cell(w=50, h=8, txt=f"Invoice No. {invoice_nr}", ln=1)
    # Add text to the PDF indicating the invoice number (repeated line)

    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")

    # Replace "_" with " " in column names and capitalize each word
    columns = [column.replace("_", " ").title() for column in df.columns]

    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(50, 50, 50)
    pdf.set_draw_color(150, 150, 150)

    # Manually set the width for each column cell
    widths = [30, 60, 40, 30, 30]

    # Iterate over each column name, capitalize each word, and add it to the PDF with the corresponding width
    for i, column in enumerate(columns):
        pdf.cell(widths[i], 10, txt=str(column), border=1)

    # Move to the next line after adding all column names
    pdf.ln(10)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=8)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(30, 10, txt=str(row["product_id"]), border=1)
        pdf.cell(60, 10, txt=str(row["product_name"]), border=1)
        pdf.cell(40, 10, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(30, 10, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(30, 10, txt=str(row["total_price"]), border=1, ln=1)

    total_sum = df["total_price"].sum()
    pdf.cell(30, 10, txt="", border=1)
    pdf.cell(60, 10, txt="", border=1)
    pdf.cell(40, 10, txt="", border=1)
    pdf.cell(30, 10, txt="", border=1)
    pdf.cell(30, 10, txt=str(total_sum), border=1, ln=1)

    # Set total sentence
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(40, 10, txt=f"The total price is {total_sum}", ln=1)

    # Set company name and logo
    pdf.set_font(family="Times", size=14, style="B")
    pdf.set_text_color(0, 0, 0)
    pdf.cell(30, 10, txt=f"  Exxotelis")
    pdf.image("images/button.png", w=8, h=8)

    # Output the PDF file with the filename based on the invoice number
    pdf.output(f"PDFs/{filename}.pdf")
