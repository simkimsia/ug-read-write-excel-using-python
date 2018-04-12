from openpyxl.utils import \
    coordinate_from_string, column_index_from_string, \
    get_column_letter


def coordinate_to_index(coordinate_string, zero_based=True):
    # returns ('A',4) for example
    xy = coordinate_from_string(coordinate_string)
    # openpyxl will assume column A returns 1
    col = column_index_from_string(xy[0])
    row = xy[1]
    if zero_based:
        col = col - 1
        row = row - 1
    return (col, row)


def index_to_coordinate(index_tuple, zero_based=True):
    col = index_tuple[0]
    row = index_tuple[1]
    if zero_based:
        col = col + 1
        row = row + 1
    col_letters = get_column_letter(col)
    return str(col_letters) + str(row)
