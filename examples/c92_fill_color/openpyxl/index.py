from openpyxl.styles import PatternFill, GradientFill


def set_fill_color_green(workbook):
    # read
    # http://openpyxl.readthedocs.io/en/stable/api/openpyxl.styles.fills.html
    ws = workbook.active
    a1 = ws['A1']
    # 2 different fill types
    fill = PatternFill("solid", fgColor="DDDDDD")
    fill = GradientFill(stop=("000000", "FFFFFF"))
    fill = PatternFill(
        fill_type=None,
        start_color='FFFFFFFF',
        end_color='FF000000')
    a1.fill = fill

    return a1
