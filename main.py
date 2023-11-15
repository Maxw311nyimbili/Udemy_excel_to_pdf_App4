import pandas as pd
from fpdf import FPDF
import glob
# glob returns all file paths that match a specific pattern, in the form of a list. It, basically, iterates through a
# folder and gets file paths with a specific pattern, then returns them as a list.

filepaths = glob.glob("Invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    filename = filepath[9:24]
    # getting information from a list without indexing
    invoice_nr, date = filename.split("-")  # gets the invoice number, and date, used for the pdf.cell section
    # create pdf object
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    # add the object to a pdf page
    pdf.add_page()
    # set the styles for what will be put on the pdf
    pdf.set_font(family="Times", size=16, style="B")
    # information to the page
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)
    pdf.cell(w=50, h=8, txt=f"Date {date}")
    pdf.output(f"PDFs/{filename}.pdf")



