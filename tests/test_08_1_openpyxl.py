from pathlib import Path
from examples.c08_1_read_from_range.openpyxl import index
from base_test_cases import ExcelTest
import os


class TestOpenPyXLReadRange(ExcelTest):

    def test_read_from_range(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        project_root = os.path.dirname(
            tests_path)
        examples_path = os.path.join(project_root, 'examples')
        sample_xlsx_path = os.path.join(
            examples_path, 'c08_1_read_from_range', '9cell_range.xlsx')
        self.assertTrue(
            Path(sample_xlsx_path).is_file())
        range = index.read_from_range(sample_xlsx_path, 'range', 'A1', 'C3')
        expected_data = [
            'a1', 'b1', 'c1',
            'a2', 'b2', 'c2',
            'a3', 'b3', 'c3',
        ]
        counter = 0
        for row in range:
            for cell in row:
                self.assertEqual(cell.value, expected_data[counter])
                counter += 1
