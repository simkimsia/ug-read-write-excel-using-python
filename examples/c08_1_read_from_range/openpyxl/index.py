from openpyxl import load_workbook


def read_from_range(file_path, sheet_name, start_cell, end_cell):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    return ws[start_cell:end_cell]
