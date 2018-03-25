import pyexcel
from pyexcel import Sheet
from collections import OrderedDict
from examples.c72_rename_validate_sheet_name.custom \
    import index as rename_index


def list_all_sheet_names(file_path='../3_sheets.xlsx'):
    book = pyexcel.get_book(file_name=file_path)
    return book.sheet_names()


def has_sheet(file_path, looking_for_sheet_name):
    sheet_names = list_all_sheet_names(file_path)
    return looking_for_sheet_name in sheet_names


def read_sheet_data(file_path, sheet_name):
    if not has_sheet(file_path, sheet_name):
        raise Exception
    book = pyexcel.get_book(file_name=file_path)
    sheet = book[sheet_name]
    # return sheet.content
    return sheet.to_array()


def write_sheet_data(workbook_data, sheet_name, data):
    if sheet_name in workbook_data:
        raise Exception
    # workbook_data is expected to be OrderedDict
    workbook_data.update({sheet_name: data})
    return workbook_data


def add_new_sheet(file_path, new_sheet_name):
    if (has_sheet(file_path, new_sheet_name)):
        raise Exception
    book = pyexcel.get_book(file_name=file_path)
    book += Sheet(name=new_sheet_name)
    return book.save_as(file_path)


def delete_sheet(file_path, sheet_name):
    if (not has_sheet(file_path, sheet_name)):
        raise Exception
    book = pyexcel.get_book(file_name=file_path)
    del book[sheet_name]
    return book.save_as(file_path)


def order_sheets_alphabetically(
    file_path,
    dest_file_path,
    reverse=False
):
    book = pyexcel.get_book(file_name=file_path)
    sheet_names = book.sheet_names()
    sheet_names.sort(reverse=reverse)
    data = OrderedDict()
    for sheet_name in sheet_names:
        data.update({sheet_name: book[sheet_name]})
    return pyexcel.save_book_as(
        bookdict=data, dest_file_name=dest_file_path)


def rename_sheet(file_path, dest_file_path, sheet_name, new_sheet_name):
    if (not has_sheet(file_path, sheet_name)):
        raise Exception
    if (has_sheet(file_path, new_sheet_name)):
        raise Exception
    new_sheet_name = rename_index.sanitise_sheet_name(new_sheet_name)
    sheet_names = list_all_sheet_names(file_path)
    book = pyexcel.get_book(file_name=file_path)
    data = OrderedDict()
    for original_sheet_name in sheet_names:
        saved_as = original_sheet_name
        if original_sheet_name == sheet_name:
            saved_as = new_sheet_name
        data.update({saved_as: book[original_sheet_name]})
    return pyexcel.save_book_as(
        bookdict=data, dest_file_name=dest_file_path
    )
