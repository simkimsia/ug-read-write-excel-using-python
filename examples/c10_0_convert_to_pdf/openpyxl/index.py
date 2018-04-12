import subprocess


from openpyxl import load_workbook
from pdfrw import PdfWriter


def convert_to_pdf(xlsx_path, pdf_path):

    workbook = load_workbook(xlsx_path)
    worksheet = workbook.active

    pw = PdfWriter(pdf_path)

    ws_range = worksheet.iter_rows('A1:H13')
    for row in ws_range:
        s = ''
        for cell in row:
            if cell.value is None:
                s += ' ' * 11
            else:
                s += str(cell.value).rjust(10) + ' '
        pw.writeLine(s)
    pw.savePage()
    pw.close()
