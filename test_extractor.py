import math
import unittest
import extractor as ex
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
        self.assertTrue(ex.is_number("123"))
        self.assertTrue(ex.is_number("-123.45"))
        self.assertTrue(ex.is_number("+123.45"))
        self.assertTrue(ex.is_number("nan"))
        self.assertFalse(ex.is_number("abc"))
        self.assertFalse(ex.is_number("123-456"))

    def test_is_range(self):
        self.assertTrue(ex.is_range("123 - 456"))
        self.assertTrue(ex.is_range("-123.45 - 678.90"))
        self.assertFalse(ex.is_range("123"))
        self.assertFalse(ex.is_range("abc - def"))

    def test_convert_value(self):
        self.assertEqual(ex.convert_value("123"), 123.0)
        self.assertEqual(ex.convert_value("10 - 20"), 15.0)
        self.assertTrue(math.isnan(ex.convert_value("abc")))

    def test_convert_range(self):
        self.assertEqual(ex.convert_range("10 - 20"), 15.0)
        self.assertEqual(ex.convert_range("-10 - 10"), 0.0)
        self.assertEqual(ex.convert_range("5.5-7.5"), 6.5)
        self.assertEqual(ex.convert_range("-10-10"), 0.0)
        self.assertEqual(ex.convert_range("8 - -6"), 1.0)
        self.assertEqual(ex.convert_range("-8 - -6"), -7.0)
        self.assertTrue(math.isnan(ex.convert_range("abc - def")))
        self.assertTrue(math.isnan(ex.convert_range("boo")))

    def test_convert_column(self):
        self.assertEqual(ex.convert_column("A"), 1)
        self.assertEqual(ex.convert_column("B"), 2)
        self.assertEqual(ex.convert_column("Z"), 26)
        self.assertEqual(ex.convert_column("AA"), 27)
        self.assertEqual(ex.convert_column("AB"), 28)
        self.assertEqual(ex.convert_column("AZ"), 52)
        self.assertEqual(ex.convert_column("BA"), 53)
        self.assertEqual(ex.convert_column("CG"), 85)
        self.assertEqual(ex.convert_column("ABC"), 0)

    def test_convert_na(self):
        self.assertEqual(ex.convert_na(None), ex.N_A)
        self.assertEqual(ex.convert_na(math.nan), ex.N_A)
        self.assertEqual(ex.convert_na("text"), "text")
        self.assertEqual(ex.convert_na(1.23), 1.23)

    def test_may_be_convert(self):
        self.assertEqual(ex.may_be_convert(None), None)
        self.assertEqual(ex.may_be_convert(1.23), 1.23)
        self.assertEqual(ex.may_be_convert("1.23"), 1.23)
        self.assertEqual(ex.may_be_convert("2 - 4"), 3.0)


if __name__ == "__main__":
    unittest.main()
