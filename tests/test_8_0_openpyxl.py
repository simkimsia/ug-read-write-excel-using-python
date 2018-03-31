from examples.c80_convert_indices_coordinates.openpyxl import index
from base_test_cases import ExcelTest


class TestOpenPyXLIndicesCoordinates(ExcelTest):

    def test_coordinates_to_indices(self):
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

    def test_indices_to_coordinates(self):
        indices = (1, 4)
        zero_based = False
        expected_coordinate_string = 'A4'
        coordinate_string = index.index_to_coordinate(indices, zero_based)
        self.assertEqual(coordinate_string, expected_coordinate_string)

        zero_based = True
        expected_coordinate_string = 'B5'
        coordinate_string = index.index_to_coordinate(indices, zero_based)
        self.assertEqual(coordinate_string, expected_coordinate_string)

        indices = (28, 24)
        zero_based = False
        expected_coordinate_string = 'AB24'
        coordinate_string = index.index_to_coordinate(indices, zero_based)
        self.assertEqual(coordinate_string, expected_coordinate_string)

        zero_based = True
        expected_coordinate_string = 'AC25'
        coordinate_string = index.index_to_coordinate(indices, zero_based)
        self.assertEqual(coordinate_string, expected_coordinate_string)
