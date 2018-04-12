from pathlib import Path


def read_file_format(file_path='test_file.xlsx'):
    '''
    returns .xlsx or .xls
    '''
    return Path(file_path).suffix
