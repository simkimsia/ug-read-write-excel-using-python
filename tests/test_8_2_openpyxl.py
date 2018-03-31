from pathlib import Path
from examples.c82_write_to_range.openpyxl import index
from examples.c81_read_from_range.openpyxl \
    import index as read_from_range_index
from examples.c71_create_empty_excel_file.openpyxl \
    import index as create_index
import unittest
import os
import sys
import pytest


class TestOpenPyXLWriteRange(unittest.TestCase):

    def setUp(self):
        if pytest.config.getoption("--skip-updown"):
            return
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def tearDown(self):
        if pytest.config.getoption("--skip-updown"):
            return
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def test_write_to_range(self):
        # python35 cannot write to Path, so use os.path
        tests_path = os.path.dirname(
            os.path.abspath(__file__))
        write_xlsx_path = os.path.join(
            tests_path, 'write_to_range.xlsx')
        create_index.save_empty_file(write_xlsx_path)
        self.assertTrue(
            Path(write_xlsx_path).is_file())

        write_data = [
            ['a1', 'b1', 'c1', ],
            ['a2', 'b2', 'c2', ],
            ['a3', 'b3', 'c3', ],
        ]
        index.write_to_range(
            write_xlsx_path, 'Sheet', 'B2', 'D4', write_data)
        range = read_from_range_index.read_from_range(
            write_xlsx_path, 'Sheet', 'B2', 'D4')
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