# This script takes in raw csv files and generates structured invoices
import pandas as pd
import glob
import openpyxl
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    filename = Path(filepath).stem
    invoice_nr = filename.split("-")[0]
    pdf.set_font(family="Times", size=15, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoices nr.{invoice_nr}")
    pdf.output(f"PDFs/{filename}.pdf")