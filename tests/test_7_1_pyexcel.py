from pathlib import Path
from examples.c71_create_empty_excel_file.pyexcel import index
import unittest
import os


class TestPyExcelEmptyFile(unittest.TestCase):
    def setUp(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def tearDown(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def test_save_using_workbook_object(self):
        # python35 cannot write to Path, so use os.path
        save_path = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)),
            'c71_pyexcel_empty.xlsx')
        self.assertFalse(
            Path(save_path).is_file())
        index.save_empty_file(save_path)
        self.assertTrue(
            Path(save_path).is_file())

    def test_save_using_pyexcel_library_directly(self):
        # python35 cannot write to Path, so use os.path
        save_path = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)),
            'c71_pyexcel_empty.xlsx')
        self.assertFalse(
            Path(save_path).is_file())
        index.save_empty_file_another_way(save_path)
        self.assertTrue(
            Path(save_path).is_file())
