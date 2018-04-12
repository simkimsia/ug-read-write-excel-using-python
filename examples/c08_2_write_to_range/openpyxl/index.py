from openpyxl import load_workbook
from examples.c08_0_convert_indices_coordinates.openpyxl \
    import index as convert_index


def write_to_range(file_path, sheet_name, start_cell, end_cell, data):
    # data is expected to be 2 dimensional list
    start_cell_index = convert_index.coordinate_to_index(start_cell, False)
    end_cell_index = convert_index.coordinate_to_index(end_cell, False)
    start_col = start_cell_index[0]
    start_row = start_cell_index[1]
    end_col = end_cell_index[0]
    end_row = end_cell_index[1]
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    count_row = 0
    for i in range(start_row, end_row+1):
        count_col = 0
        for j in range(start_col, end_col+1):
            ws.cell(row=i, column=j).value = data[count_row][count_col]
            count_col += 1
        count_row += 1
    return wb.save(file_path)
