from openpyxl import load_workbook


def list_all_sheet_names(file_path='../3_sheets.xlsx'):
    wb = load_workbook(filename=file_path)
    return wb.sheetnames


def has_sheet(file_path, looking_for_sheet_name):
    sheet_names = list_all_sheet_names(file_path)
    return looking_for_sheet_name in sheet_names


def add_new_sheet(file_path, new_sheet_name):
    wb = load_workbook(filename=file_path)
    wb.create_sheet(new_sheet_name)
    return wb.save(file_path)


def delete_sheet(file_path, sheet_name):
    book = Workbook()
    book.save(file_path)


def order_sheets_alphabetically(file_path, order="asc"):
    book = Workbook()
    book.save(file_path)


def rename_sheet(file_path, sheet_name, new_sheet_name):
    book = Workbook()
    book.save(file_path)
