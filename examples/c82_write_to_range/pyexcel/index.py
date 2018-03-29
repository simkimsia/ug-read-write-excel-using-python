from pyexcel import get_book
from examples.c80_convert_indices_coordinates.openpyxl \
    import index as convert_index


def write_to_range(file_path, sheet_name, start_cell, end_cell, data):
    book = get_book(file_name=file_path)
    # data is expected to be 2 dimensional list
    start_cell_index = convert_index.coordinate_to_index(start_cell, True)
    end_cell_index = convert_index.coordinate_to_index(end_cell, True)

    start_col = start_cell_index[0]
    start_row = start_cell_index[1]
    end_col = end_cell_index[0]
    end_row = end_cell_index[1]
    count_row = 0
    sheet = book[sheet_name]
    for row in range(start_row, end_row+1):
        count_col = 0
        for col in range(start_col, end_col+1):
            print(row)
            print(col)
            print(data[count_row][count_col])
            sheet[row, col] = data[count_row][count_col]
            count_col += 1
        count_row += 1
    return book.save_as(file_path)
