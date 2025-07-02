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

    def test_new_file_name(self):
        self.assertEqual(ex.new_file_name("file.ext"), "file.org.ext")
        self.assertEqual(ex.new_file_name("/path/file.ext"), "/path/file.org.ext")
        self.assertEqual(ex.new_file_name("file"), "file.org")


if __name__ == "__main__":
    unittest.main()
