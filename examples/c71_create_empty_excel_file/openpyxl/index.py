from openpyxl import Workbook


def save_empty_file():
    book = Workbook()
    save_as_path = 'empty.xlsx'
    book.save(save_as_path)


save_empty_file()
