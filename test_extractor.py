import math
import unittest
from extractor import extract, is_number, is_range, mid_range, get_value
from openpyxl import Workbook
import os


class TestExtractor(unittest.TestCase):

    def setUp(self):
        # Create dummy Excel files for testing
        self.source_file = "test_source.xlsx"
        self.destination_file = "test_destination.xlsx"

        wb_source = Workbook()
        ws_source = wb_source.active
        ws_source.title = "Input"
        ws_source["D6"] = 42
        wb_source.save(self.source_file)

        wb_destination = Workbook()
        ws_destination = wb_destination.active
        ws_destination.title = "Summary by Type"
        wb_destination.save(self.destination_file)

    def tearDown(self):
        # Clean up test files
        if os.path.exists(self.source_file):
            os.remove(self.source_file)
        if os.path.exists(self.destination_file):
            os.remove(self.destination_file)
        if os.path.exists(self.destination_file + ".new"):
            os.remove(self.destination_file + ".new")

    def test_is_number(self):
        self.assertTrue(is_number("123"))
        self.assertTrue(is_number("-123.45"))
        self.assertTrue(is_number("+123.45"))
        self.assertTrue(is_number("nan"))
        self.assertFalse(is_number("abc"))
        self.assertFalse(is_number("123-456"))

    def test_is_range(self):
        self.assertTrue(is_range("123 - 456"))
        self.assertTrue(is_range("-123.45 - 678.90"))
        self.assertFalse(is_range("123"))
        self.assertFalse(is_range("abc - def"))

    def test_mid_range(self):
        self.assertEqual(mid_range("10 - 20"), 15.0)
        self.assertEqual(mid_range("-10 - 10"), 0.0)
        self.assertEqual(mid_range("5.5-7.5"), 6.5)
        self.assertEqual(mid_range("-10-10"), 0.0)
        self.assertEqual(mid_range("8 - -6"), 1.0)
        self.assertEqual(mid_range("-8 - -6"), -7.0)
        self.assertTrue(math.isnan(mid_range("abc - def")))
        self.assertTrue(math.isnan(mid_range("boo")))

    def test_get_value(self):
        self.assertEqual(get_value("123"), 123.0)
        self.assertEqual(get_value("10 - 20"), 15.0)
        self.assertTrue(math.isnan(get_value("abc")))


if __name__ == "__main__":
    unittest.main()
