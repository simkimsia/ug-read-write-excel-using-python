from openpyxl.styles.borders import Border, Side


def set_thin_border(workbook):
    ws = workbook.active
    a1 = ws['A1']
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))

    a1.border = thin_border

    return a1
