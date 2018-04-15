from pathlib import Path
from examples.c07_1_create_empty_excel_file.pyexcel import index as create
from examples.c07_0_read_file_format.pathlib import index
import os
from base_test_cases import ExcelTest


class TestPathLibReadFileFormat(ExcelTest):

    def test_xlsx_can_be_read(self):
        # python35 cannot write to Path, so use os.path
        save_path_xlsx = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)),
            'c07_0_read_xlsx.xlsx')
        create.save_empty_file(save_path_xlsx)
        self.assertTrue(
            Path(save_path_xlsx).is_file())
        self.assertEqual(
            '.xlsx',
            index.read_file_format(save_path_xlsx))

    def test_xls_can_be_read(self):
        # python35 cannot write to Path, so use os.path
        save_path_xls = os.path.join(
            os.path.dirname(
                os.path.abspath(__file__)),
            'c07_0_openpyxl_read_xls.xls')
        create.save_empty_file_another_way(save_path_xls)
        self.assertTrue(
            Path(save_path_xls).is_file())
        self.assertEqual(
            '.xls',
            index.read_file_format(save_path_xls))
