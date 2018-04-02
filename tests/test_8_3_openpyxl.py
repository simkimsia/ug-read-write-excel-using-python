from pathlib import Path
from examples.c83_freeze_panes.openpyxl import index
from examples.c71_create_empty_excel_file.openpyxl \
    import index as create_index
from base_test_cases import ExcelTest
import os


class TestOpenPyXLFreezePanes(ExcelTest):

    def test_has_freeze_panes(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        project_root = os.path.dirname(
            tests_path)
        examples_path = os.path.join(project_root, 'examples')
        sample_xlsx_path = os.path.join(
            examples_path, 'c83_freeze_panes', 'freeze_panes.xlsx')
        self.assertTrue(
            Path(sample_xlsx_path).is_file())

        has = index.has_freeze_pane(
            sample_xlsx_path, 'freeze_first_row')
        self.assertTrue(has)

        has_not = index.has_freeze_pane(
            sample_xlsx_path, 'empty')
        self.assertFalse(has_not)

    def test_freeze_panes(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        write_xlsx_path = os.path.join(
            tests_path, 'output_freeze.xlsx')
        create_index.save_empty_file(write_xlsx_path)
        self.assertTrue(
            Path(write_xlsx_path).is_file())

        has = index.has_freeze_pane(
            write_xlsx_path, 'Sheet')
        self.assertFalse(has)

        index.freeze_pane(write_xlsx_path, 'Sheet', 'C7')

        has = index.has_freeze_pane(
            write_xlsx_path, 'Sheet')
        self.assertTrue(has)

        index.unfreeze_pane(write_xlsx_path, 'Sheet')

        has = index.has_freeze_pane(
            write_xlsx_path, 'Sheet')
        self.assertFalse(has)
