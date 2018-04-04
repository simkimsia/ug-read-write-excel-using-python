from openpyxl.styles import Font, colors


def set_font_color_red(workbook):
    ws = workbook.active
    a1 = ws['A1']
    ft = Font(color=colors.RED)
    a1.font = ft

    return a1


def set_font_size(workbook, size):
    ws = workbook.active
    a1 = ws['A1']
    ft = Font(size=size)
    a1.font = ft

    return a1
