import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob('invoices/*.xlsx')

for filepath in filepaths:

    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Invoice nr.{invoice_nr}', align='L', ln=1)

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=50, h=8, txt=f'Date: {date}', align='L', ln=1)

    df = pd.read_excel(filepath, sheet_name='Sheet 1')

    # Add header
    colums = df.columns
    colums = [title.replace("_", " ").title() for title in colums]
    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=colums[0], border=1)
    pdf.cell(w=60, h=8, txt=colums[1], border=1)
    pdf.cell(w=40, h=8, txt=colums[2], border=1)
    pdf.cell(w=30, h=8, txt=colums[3], border=1)
    pdf.cell(w=30, h=8, txt=colums[4], border=1, ln=1)

    # Add rows to the table
    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10)
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=8, txt=str(row['product_id']), border=1)
        pdf.cell(w=60, h=8, txt=str(row['product_name']), border=1)
        pdf.cell(w=40, h=8, txt=str(row['amount_purchased']), border=1, align='R')
        pdf.cell(w=30, h=8, txt=str(row['price_per_unit']), border=1, align='R')
        pdf.cell(w=30, h=8, txt=str(row['total_price']), border=1, ln=1, align='R')
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(sum(df['total_price'])), border=1, align='R', ln=1)

    # Add total sum sentence
    pdf.set_font(family='Times', size=14, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=f"Total price is {str(sum(df['total_price']))}", ln=1)

    # Add company name and logo
    pdf.set_font(family='Times', size=14, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=35, h=8, txt="Our Company")
    pdf.image('pythonhow.png', w=10)

    pdf.output(f'PDFs/{filename}.pdf')

