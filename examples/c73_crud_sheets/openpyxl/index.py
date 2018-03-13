from openpyxl import Workbook


def list_all_sheet_names(file_path='../3_sheets.xlsx'):
    book = Workbook()
    book.save(file_path)


def add_new_sheet(file_path, new_sheet_name):
    book = Workbook()
    book.save(file_path)


def delete_sheet(file_path, sheet_name):
    book = Workbook()
    book.save(file_path)


def order_sheets_alphabetically(file_path, order="asc"):
    book = Workbook()
    book.save(file_path)


def rename_sheet(file_path, sheet_name, new_sheet_name):
    book = Workbook()
    book.save(file_path)
