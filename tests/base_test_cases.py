from pathlib import Path
import unittest
import pytest


class ExcelTest(unittest.TestCase):
    def setUp(self):
        if pytest.config.getoption("--skip-updown"):
            return
        if pytest.config.getoption("--skip-up"):
            return
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xls*'):
            f.unlink()

    def tearDown(self):
        if pytest.config.getoption("--skip-updown"):
            return
        if pytest.config.getoption("--skip-down"):
            return
        dir_path = Path(__file__).resolve().parent
        for f in dir_path.glob('*.xls*'):
            f.unlink()
