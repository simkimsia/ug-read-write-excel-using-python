from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference


def make_bar_chart(file_path):
    wb = load_workbook(file_path)
    ws = wb.active
    for i in range(10):
        ws.append([i])
    values = Reference(ws, min_col=1, min_row=1, max_col=1, max_row=10)
    chart = BarChart()
    chart.add_data(values)
    ws.add_chart(chart, "E15")
    wb.save(file_path)
