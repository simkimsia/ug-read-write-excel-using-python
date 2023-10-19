from openpyxl import load_workbook


def has_freeze_pane(file_path, sheet_name):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    return ws.freeze_panes != 'A1' and ws.freeze_panes is not None


def freeze_pane(file_path, sheet_name, cell_or_string):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    ws.freeze_panes = cell_or_string
    return wb.save(file_path)


def unfreeze_pane(file_path, sheet_name):
    wb = load_workbook(file_path)
    ws = wb[sheet_name]
    ws.freeze_panes = None
    return wb.save(file_path)
