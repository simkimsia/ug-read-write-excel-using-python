from pathlib import Path
from examples.c71_create_empty_excel_file.openpyxl import index
import unittest


class TestOpenPyXLEmptyFile(unittest.TestCase):
    def setUp(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def tearDown(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def test(self):
        save_path = Path(
            Path(__file__).resolve().parent,
            'c71_openpyxl_empty.xlsx')
        self.assertFalse(
            save_path.is_file())
        index.save_empty_file(save_path)
        self.assertTrue(
            save_path.is_file())
