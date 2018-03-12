from examples.c72_rename_validate_sheet_name.custom import index
import unittest


class TestCustomValidateSheetName(unittest.TestCase):
    def test_validate(self):
        expect_false = '#abc'
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = "%abc"
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = "&abc"
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = "/abc"
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = "*abc"
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = "?abc"
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = '\\ asdasd'
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_false = '1234567890123456789012345678901234'
        self.assertFalse(index.validate_sheet_name(expect_false))

        expect_true = '123456789012345678901234567890123'
        self.assertTrue(index.validate_sheet_name(expect_true))

    def test_sanitise(self):
        original_name = '1234567890123456789012345678901234'
        expected = '123456789012345678901234567890123'
        self.assertEqual(index.sanitise_sheet_name(original_name), expected)

        original_name = '123?456%789/012*345#678\\901&234567'
        expected = '123456789-012345678-901234567'
        self.assertEqual(index.sanitise_sheet_name(original_name), expected)
