from pathlib import Path
from examples.c10_0_convert_to_pdf.openpyxl import index
from base_test_cases import ExcelTest
import os


class TestOpenPyXLReadRange(ExcelTest):

    def test_convert_to_pdf(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        project_root = os.path.dirname(
            tests_path)
        examples_path = os.path.join(project_root, 'examples')
        sample_xlsx_path = os.path.join(
            examples_path, 'c08_1_read_from_range', '9cell_range.xlsx')
        sample_pdf_path = os.path.join(
            examples_path, 'c08_1_read_from_range', '9cell_range.pdf')
        index.convert_to_pdf(sample_xlsx_path, sample_pdf_path)
        self.assertTrue(True)
