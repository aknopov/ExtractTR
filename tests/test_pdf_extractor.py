import unittest
from unittest.mock import Mock, patch
import fitz # PyMuPDF
import pdf_extractor as pdf
import xcl_extractor as xcl


def identity_fun(arg):
    return arg


# pylint: disable=protected-access
class TestPdfExtractor(unittest.TestCase):

    def setUp(self):
        self.doc = fitz.open()
        page = self.doc.new_page(width=595, height=842)
        page.insert_text((50, 50), "Client:")
        page.insert_text((110, 50), "Client Name")
        page.insert_text((200, 50), "Borehole/Sample No.:")
        page.insert_text((350, 50), "ENG-PQ24-15 / SC40")
        page.insert_text((200, 70), "Sample Depth (m):")
        page.insert_text((350, 70), "56.9 - 57.6m")
        page.insert_text((200, 90), "Liquid Limit:")
        page.insert_text((350, 90), "39")

        # wb_destination = Workbook()
        # ws_destination = wb_destination.active
        # ws_destination.title = "Summary by Type"
        # wb_destination.save(self.destination_file) #UC ???


    def tearDown(self):
        self.doc.close()
        # Clean up test files
        # if os.path.exists(self.source_file):
        #     os.remove(self.source_file)
        # if os.path.exists(self.destination_file):
        #     os.remove(self.destination_file)


    def test_validate_mappings(self):
        for mapping in pdf.MAPPINGS:
            print(mapping)
            self.assertTrue(mapping["key"].islower())
            self.assertTrue(mapping["out"].isalpha())
            self.assertGreaterEqual(mapping["num"], len(mapping["offset"]))


    def test_merge_vals(self):
        self.assertEqual("ENG-PQ24-15_SC40", pdf._merge_vals(["ENG-PQ24-15", "/", "SC40"]))
        self.assertEqual(3.0, pdf._merge_vals(["2", "-", "4"]))
        self.assertEqual(3.0, pdf._merge_vals(["2", "-", "4m"]))
        self.assertEqual(2.0, pdf._merge_vals(["2", "something", "4"]))
        self.assertEqual("ab", pdf._merge_vals(["a", "b"]))


    def test_copy_one_value_basic(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        # Test mapping
        mapping = { "page": 0, "key": "client:", "num": 1, "out": "B", "offset": [0]}

        with patch('xcl_extractor.cnv.convert_na', side_effect=identity_fun), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):

            pdf._copy_one_value(self.doc, excl_wb, 10, mapping)

        mock_output_sheet.cell.assert_called_once_with(
            row=11,  # last_row + offset + 1 = 10 + 0 + 1
            column=2,  # column B
            value="Client" #UC "Client Name"!
        )


    def test_copy_one_value_with_merge(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        # Test mapping
        mapping = { "page": 0, "key": "borehole/sample no.:", "num": 3, "out": "B", "offset": [0]}

        # Mock converter functions
        with patch('xcl_extractor.cnv.convert_na', side_effect=identity_fun), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):

            pdf._copy_one_value(self.doc, excl_wb, 10, mapping)

        mock_output_sheet.cell.assert_called_once_with(
            row=11,  # last_row + offset + 1 = 10 + 0 + 1
            column=2,  # column B
            value="ENG-PQ24-15_SC40"
        )


    #UC def test_extract_file(self): #UC just for testing
    #     pdf.extract_file(self.source_file, self.destination_file)

    #     wb_in = load_workbook(self.source_file, data_only=True)
