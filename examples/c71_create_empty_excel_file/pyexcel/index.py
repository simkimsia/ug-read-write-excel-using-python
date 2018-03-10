import pyexcel


def save_empty_file(save_as_path='c71_pyexcel_empty.xlsx'):
    book = pyexcel.Book()
    book.save_as(save_as_path)
