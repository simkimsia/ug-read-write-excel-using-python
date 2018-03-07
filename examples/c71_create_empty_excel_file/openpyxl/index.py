from openpyxl import Workbook


def save_empty_file(save_as_path = 'c71_openpyxl_empty.xlsx'):
    book = Workbook()
    book.save(save_as_path)
