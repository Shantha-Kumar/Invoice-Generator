import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob(pathname='exel/*.xlsx')
print(filepaths)
for filepath in filepaths:
    df = pd.read_excel(filepath, sheet_name='Sheet 1')
    columns = [item.replace('_', " ").title() for item in df.columns]
    pdf = FPDF(orientation='p', unit='mm', format='A4')
    pdf.add_page()

    filename = Path(filepath).stem

    print(filename)
    invoice_nr, date = filename.split('-')

    pdf.set_font(family='Times', size=16, style='B')
    pdf.cell(w=30, h=8, txt=f'Invoice No.{invoice_nr}', ln=1)

    pdf.set_font(family='Times', size=15, style='B')
    pdf.cell(w=30, h=8, txt=f'Date {date}', ln=1)

    pdf.set_font(family='Times', size=10, style='B')
    pdf.cell(w=30, h=7, txt=columns[0], border=1)
    pdf.cell(w=70, h=7, txt=columns[1], border=1)
    pdf.cell(w=35, h=7, txt=columns[2], border=1)
    pdf.cell(w=30, h=7, txt=columns[3], border=1)
    pdf.cell(w=30, h=7, txt=columns[4], ln=1, border=1)

    pdf.set_font(family='Times', size=14)
    for index, row in df.iterrows():
        pdf.cell(w=30, h=7, txt=str(row['product_id']), border=1)
        pdf.cell(w=70, h=7, txt=str(row['product_name']), border=1)
        pdf.cell(w=35, h=7, txt=str(row['amount_purchased']), border=1)
        pdf.cell(w=30, h=7, txt=str(row['price_per_unit']), border=1)
        pdf.cell(w=30, h=7, txt=str(row['total_price']), ln=1, border=1)

    total_sum = df['total_price'].sum()
    pdf.cell(w=30, h=7, txt="", border=1)
    pdf.cell(w=70, h=7, txt="", border=1)
    pdf.cell(w=35, h=7, txt="", border=1)
    pdf.cell(w=30, h=7, txt="", border=1)
    pdf.cell(w=30, h=7, txt=str(total_sum), ln=1, border=1)

    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=30, h=8, txt=f'The total sum is {total_sum}', ln=1)

    pdf.set_font(family='Times', size=12, style='B''I')
    pdf.cell(w=30, h=8, txt='Shantha Corp')
    pdf.image('thumb.png', w=5, h=8)

    pdf.output(f"PDFs/{invoice_nr}.pdf")
