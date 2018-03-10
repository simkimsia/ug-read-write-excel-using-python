import pyexcel


def save_empty_file(save_as_path='c71_pyexcel_empty.xlsx'):
    '''
    This works for xlsx only
    '''
    book = pyexcel.Book()
    book.save_as(save_as_path)


def save_empty_file_another_way(save_as_path='c71_pyexcel_empty.xlsx'):
    '''
    This works for both xlsx and xls
    '''
    data = []
    pyexcel.save_as(array=data, dest_file_name=save_as_path)
