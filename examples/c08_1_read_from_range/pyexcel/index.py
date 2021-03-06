from pyexcel import get_sheet
from examples.c08_0_convert_indices_coordinates.openpyxl \
    import index as convert_index


def read_from_range(file_path, sheet_name, start_cell, end_cell):
    start_cell_index = convert_index.coordinate_to_index(start_cell)
    end_cell_index = convert_index.coordinate_to_index(end_cell)
    start_col = start_cell_index[0]
    print(start_col)
    start_row = start_cell_index[1]
    end_col = end_cell_index[0]
    end_row = end_cell_index[1]
    row_limit = end_row - start_row + 1
    col_limit = end_col - start_col + 1
    return get_sheet(
        file_name=file_path,
        start_row=start_row, row_limit=row_limit,
        start_column=start_col, column_limit=col_limit
    )
