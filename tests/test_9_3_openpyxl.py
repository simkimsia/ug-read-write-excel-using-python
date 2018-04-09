from openpyxl.styles import Border, Side
from examples.c93_border.openpyxl import index
from openpyxl import Workbook
from base_test_cases import ExcelTest


class TestOpenPyXLBorder(ExcelTest):

    def test_border(self):
        wb = Workbook()
        ws = wb.active
        a1 = ws['A1']
        self.assertEqual(a1.border, Border())
        a1 = index.set_thin_border(wb)
        thin_border = Border(left=Side(style='thin'),
                             right=Side(style='thin'),
                             top=Side(style='thin'),
                             bottom=Side(style='thin'))
        self.assertEqual(a1.border, thin_border)
