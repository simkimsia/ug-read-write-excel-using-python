from pathlib import Path
from examples.c80_convert_indices_coordinates.openpyxl import index
from examples.c71_create_empty_excel_file.openpyxl \
    import index as create_index
import unittest
import os


class TestOpenPyXLIndicesCoordinates(unittest.TestCase):
    def setUp(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def tearDown(self):
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xlsx'):
            f.unlink()

    def test_coordinates_to_indices(self):
        # python35 cannot write to Path, so use os.path
        coordinate_string = 'A4'
        zero_based = True
        expected_indices = (0, 3)
        indices = index.coordinate_to_index(coordinate_string, zero_based)
        self.assertEqual(indices, expected_indices)
