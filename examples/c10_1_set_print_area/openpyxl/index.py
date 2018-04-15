from openpyxl import load_workbook


def set_print_area(xlsx_path, range_in_string, sheet_name=None):
    workbook = load_workbook(xlsx_path)
    worksheet = workbook.active
    if sheet_name is not None:
        worksheet = workbook[sheet_name]

    worksheet.print_area = range_in_string
    workbook.save(xlsx_path)
