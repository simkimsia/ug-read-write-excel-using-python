from pathlib import Path
from examples.c80_convert_indices_coordinates.openpyxl import index
import unittest


class TestOpenPyXLIndicesCoordinates(unittest.TestCase):

    def test_coordinates_to_indices(self):
        # python35 cannot write to Path, so use os.path
        coordinate_string = 'A4'
        zero_based = True
        expected_indices = (0, 3)
        indices = index.coordinate_to_index(coordinate_string, zero_based)
        self.assertEqual(indices, expected_indices)

        zero_based = False
        expected_indices = (1, 4)
        indices = index.coordinate_to_index(coordinate_string, zero_based)
        self.assertEqual(indices, expected_indices)

        coordinate_string = 'AB24'
        zero_based = True
        expected_indices = (27, 23)
        indices = index.coordinate_to_index(coordinate_string, zero_based)
        self.assertEqual(indices, expected_indices)

        zero_based = False
        expected_indices = (28, 24)
        indices = index.coordinate_to_index(coordinate_string, zero_based)
        self.assertEqual(indices, expected_indices)
