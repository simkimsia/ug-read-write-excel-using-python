from pathlib import Path
from examples.c71_create_empty_excel_file.openpyxl import index
import unittest


class TestOpenPyXLEmptyFile(unittest.TestCase):
    # def setUp(self):
    #     dir_path = Path('../examples/c71_create_empty_excel_file/openpyxl/')
    #     for f in dir_path.glob('*.xlsx'):
    #         f.unlink()

    # def tearDown(self):
    #     dir_path = Path('../examples/c71_create_empty_excel_file/openpyxl/')
    #     for f in dir_path.glob('*.xlsx'):
    #         f.unlink()

    def test(self):
        # self.assertFalse(
        #     Path(
        #         '../examples/c71_create_empty_excel_file/openpyxl/empty.xlsx'
        #     ).is_file())
        index.save_empty_file()
        # self.assertTrue(
        #     Path(
        #         '../examples/c71_create_empty_excel_file/openpyxl/empty.xlsx'
        #     ).is_file())
