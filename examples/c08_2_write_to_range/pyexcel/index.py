from pyexcel import get_book
from examples.c08_0_convert_indices_coordinates.openpyxl \
    import index as convert_index


def write_to_range(file_path, sheet_name, start_cell, end_cell, data):
    book = get_book(file_name=file_path)
    # data is expected to be 2 dimensional list
    start_cell_index = convert_index.coordinate_to_index(start_cell, True)

    start_col = start_cell_index[0]
    start_row = start_cell_index[1]
    sheet = book[sheet_name]
    sheet.paste([start_row, start_col], rows=data)
    return book.save_as(file_path)
