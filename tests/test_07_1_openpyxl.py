from pathlib import Path
from examples.c07_1_create_empty_excel_file.openpyxl import index
import os
from base_test_cases import ExcelTest


class TestOpenPyXLEmptyFile(ExcelTest):

    def test(self):
        # python35 cannot write to Path, so use os.path
        save_path = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)),
            'c07_1_openpyxl_empty.xlsx')
        self.assertFalse(
            Path(save_path).is_file())
        index.save_empty_file(save_path)
        self.assertTrue(
            Path(save_path).is_file())
