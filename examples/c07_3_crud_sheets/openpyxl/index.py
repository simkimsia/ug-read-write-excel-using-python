from openpyxl import load_workbook, Workbook
from examples.c07_2_rename_validate_sheet_name.custom \
    import index as rename_index


def list_all_sheet_names(file_path='../3_sheets.xlsx'):
    wb = load_workbook(filename=file_path)
    return wb.sheetnames


def has_sheet(file_path, looking_for_sheet_name):
    sheet_names = list_all_sheet_names(file_path)
    return looking_for_sheet_name in sheet_names


def read_sheet_data(file_path, sheet_name):
    if not has_sheet(file_path, sheet_name):
        raise Exception
    wb = load_workbook(file_path)
    sheet = wb[sheet_name]
    data = []
    cols, rows = sheet.max_column, sheet.max_row
    data = [[0 for x in range(cols)] for y in range(rows)]
    for row_index, row in enumerate(
        sheet.iter_rows(
            max_col=sheet.max_column,
            max_row=sheet.max_row
        )
    ):
        for col_index, cell in enumerate(row):
            data[row_index][col_index] = cell.value
    return data


def write_sheet_data(workbook, sheet_name, data):
    if sheet_name in workbook.sheetnames:
        raise Exception
    workbook.create_sheet(sheet_name)
    sheet = workbook[sheet_name]
    for row_index in range(len(data)):
        for col_index in range(len(data[row_index])):
            sheet.cell(
                row=row_index + 1,
                column=col_index + 1
            ).value = data[row_index][col_index]
    return workbook


def add_new_sheet(file_path, new_sheet_name):
    if (has_sheet(file_path, new_sheet_name)):
        raise Exception
    wb = load_workbook(filename=file_path)
    wb.create_sheet(new_sheet_name)
    return wb.save(file_path)


def delete_sheet(file_path, sheet_name):
    if (not has_sheet(file_path, sheet_name)):
        raise Exception
    wb = load_workbook(filename=file_path)
    del wb[sheet_name]
    return wb.save(file_path)


# This is not easy to do even with other libraries like xlsxwriter
# as it involves copying
# data from one sheet in one book to another book
# this is also why I recommend pyexcel
def order_sheets_alphabetically(
    file_path,
    dest_file_path,
    reverse=False
):
    wb = load_workbook(filename=file_path)
    sheet_names = wb.sheetnames
    sheet_names.sort(reverse=reverse)
    # create empty workbook first
    destBook = Workbook()
    for sheet_name in sheet_names:
        data = read_sheet_data(file_path, sheet_name)
        destBook = write_sheet_data(destBook, sheet_name, data)
    # finally remove unnecessary sheet
    del destBook['Sheet']
    destBook.save(dest_file_path)
    return True


# Again this is not easy to do even with other libraries like xlsxwriter
# as it involves copying
# data from one sheet in one book to another book
# this is also why I recommend pyexcel
def rename_sheet(file_path, dest_file_path, sheet_name, new_sheet_name):
    if (not has_sheet(file_path, sheet_name)):
        raise Exception
    if (has_sheet(file_path, new_sheet_name)):
        raise Exception
    new_sheet_name = rename_index.sanitise_sheet_name(new_sheet_name)
    sheet_names = list_all_sheet_names(file_path)
    destBook = Workbook()
    for original_sheet_name in sheet_names:
        saved_as = original_sheet_name
        if original_sheet_name == sheet_name:
            saved_as = new_sheet_name
        data = read_sheet_data(file_path, sheet_name)
        destBook = write_sheet_data(destBook, saved_as, data)
    # finally remove unnecessary sheet
    del destBook['Sheet']
    destBook.save(dest_file_path)
    return True
