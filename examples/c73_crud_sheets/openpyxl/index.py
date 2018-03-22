from openpyxl import load_workbook, Workbook


def list_all_sheet_names(file_path='../3_sheets.xlsx'):
    wb = load_workbook(filename=file_path)
    return wb.sheetnames


def has_sheet(file_path, looking_for_sheet_name):
    sheet_names = list_all_sheet_names(file_path)
    return looking_for_sheet_name in sheet_names


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
    wb = load_workbook(file_path, use_iterators=True)
    sheet = wb.worksheets[0]
    row_count = sheet.max_row
    column_count = sheet.max_column
    return True


# Again this is not easy to do even with other libraries like xlsxwriter
# as it involves copying
# data from one sheet in one book to another book
# this is also why I recommend pyexcel
def rename_sheet(file_path, sheet_name, new_sheet_name):
    return True
