from pyexcel import load_workbook, Workbook
import re


def get_row_col_stats(start_cell, end_cell):
    # separate the number and letters from start_cell
    start_row_in_letter = extract_out_letters(start_cell)
    start_col = extract_out_numbers(start_cell)
    # separate the number and letters from end_cell
    end_row_in_letter = extract_out_letters(end_cell)
    end_col = extract_out_numbers(end_cell)


def read_from_range(file_path, sheet_name, start_cell, end_cell):
    row_col_stats = get_row_col_stats(start_cell, end_cell)
    wb = load_workbook(file_path, use_iterators=True)
    ws = wb[sheet_name]
    return ws[start_cell:end_cell]
