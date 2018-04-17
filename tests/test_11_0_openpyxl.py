from pathlib import Path
from examples.c07_1_create_empty_excel_file.openpyxl \
    import index as create_index
from examples.c11_0_hidden_rows_columns.openpyxl import index
from base_test_cases import ExcelTest
import os


class TestOpenPyXLHideRowsCols(ExcelTest):

    def test_hide_row(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        write_xlsx_path = os.path.join(
            tests_path, 'hide_row13.xlsx')
        create_index.save_empty_file(write_xlsx_path)
        self.assertTrue(
            Path(write_xlsx_path).is_file())

        # ensure no hidden rows at first
        hidden_row_dimensions = index.get_hidden_row_dimensions(
            write_xlsx_path)
        expected_hidden_row_dimensions = []
        self.assertEqual(hidden_row_dimensions, expected_hidden_row_dimensions)

        index.hide_row(write_xlsx_path, [1, 3])

        # now check if rows are hidden correctly
        hidden_row_dimensions = index.get_hidden_row_dimensions(
            write_xlsx_path)
        expected_hidden_row_dimensions = [1, 3]
        self.assertEqual(hidden_row_dimensions, expected_hidden_row_dimensions)

    def test_get_hidden_row_dimensions(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        project_root = os.path.dirname(
            tests_path)
        examples_path = os.path.join(project_root, 'examples')
        sample_xlsx_path = os.path.join(
            examples_path, 'c11_0_hidden_rows_columns', 'hideCG36.xlsx')
        self.assertTrue(
            Path(sample_xlsx_path).is_file())

        hidden_row_dimensions = index.get_hidden_row_dimensions(
            sample_xlsx_path)
        expected_hidden_row_dimensions = [3, 6]
        self.assertEqual(hidden_row_dimensions, expected_hidden_row_dimensions)

    def test_hide_col(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        write_xlsx_path = os.path.join(
            tests_path, 'hide_colAD.xlsx')
        create_index.save_empty_file(write_xlsx_path)
        self.assertTrue(
            Path(write_xlsx_path).is_file())

        # ensure no hidden cols at first
        hidden_col_dimensions = index.get_hidden_col_dimensions(
            write_xlsx_path)
        expected_hidden_col_dimensions = []
        self.assertEqual(hidden_col_dimensions, expected_hidden_col_dimensions)

        index.hide_col(write_xlsx_path, ['A', 'D'])

        # now check if cols are hidden correctly
        hidden_col_dimensions = index.get_hidden_col_dimensions(
            write_xlsx_path)
        expected_hidden_col_dimensions = ['A', 'D']
        self.assertEqual(hidden_col_dimensions, expected_hidden_col_dimensions)

    def test_get_hidden_col_dimensions(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        project_root = os.path.dirname(
            tests_path)
        examples_path = os.path.join(project_root, 'examples')
        sample_xlsx_path = os.path.join(
            examples_path, 'c11_0_hidden_rows_columns', 'hideCG36.xlsx')
        self.assertTrue(
            Path(sample_xlsx_path).is_file())

        hidden_col_dimensions = index.get_hidden_col_dimensions(
            sample_xlsx_path)
        expected_hidden_col_dimensions = ['C', 'G']
        self.assertEqual(hidden_col_dimensions, expected_hidden_col_dimensions)
