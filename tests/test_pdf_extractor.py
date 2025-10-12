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
        page.insert_text((110, 50), "Client Name in Few Words")
        page.insert_text((200, 70), "Borehole/Sample No.:")
        page.insert_text((350, 70), "ENG-PQ24-15 / SC40")
        page.insert_text((200, 90), "Sample Type:")
        page.insert_text((350, 90), "Intact")
        page.insert_text((200, 110), "Sample Depth (m):")
        page.insert_text((350, 110), "56.9 - 57.6m")
        page.insert_text((200, 130), "Liquid Limit:")
        page.insert_text((350, 130), "39")
        #UC self.doc.save("tests/test_source.pdf")


    def tearDown(self):
        self.doc.close()


    def test_validate_mappings(self):
        for mapping in pdf.MAPPINGS:
            print(mapping)
            self.assertTrue(mapping["key"].islower())
            self.assertTrue(mapping["out"].isalpha())
            if "offset" in mapping:
                self.assertEqual(mapping["num"], len(mapping["offset"]))
            if "till" in mapping:
                self.assertEqual(mapping["num"], 1)


    def test_merge_vals(self):
        self.assertEqual("ENG-PQ24-15_SC40", pdf._merge_vals(["ENG-PQ24-15", "/", "SC40"]))
        self.assertEqual(3.0, pdf._merge_vals(["2", "-", "4"]))
        self.assertEqual(3.0, pdf._merge_vals(["2", "-", "4m"]))
        self.assertEqual(2.0, pdf._merge_vals(["2", "something", "4"]))
        self.assertEqual("a b", pdf._merge_vals(["a", "b"]))


    def test_copy_one_value_basic(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        # Test mapping
        mapping = {"page": 0, "key": "liquid limit:", "num": 1, "out": "I"}

        with patch('xcl_extractor.cnv.convert_na', side_effect=identity_fun), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):

            pdf._copy_one_value(self.doc, excl_wb, 10, mapping)

        mock_output_sheet.cell.assert_called_once_with(
            row=11,  # last_row + 1
            column=2,  # column B
            value="39"
        )


    def test_copy_one_value_with_merge(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        # Mock converter functions
        with patch('xcl_extractor.cnv.convert_na', side_effect=identity_fun), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):
            # 1
            mapping = {"page": 0, "key": "borehole/sample no.:", "till": "sample type:",  "num": 1, "out": "B"}
            pdf._copy_one_value(self.doc, excl_wb, 10, mapping)
            mock_output_sheet.cell.assert_called_once_with(
                row=11,  # last_row + offset + 1 = 10 + 0 + 1
                column=2,  # column B
                value="ENG-PQ24-15_SC40"
            )
            # 2
            mock_output_sheet.reset_mock()
            mapping = {"page": 0, "key": "client:", "till": "borehole/sample no.:", "num": 1, "out": "B"}
            pdf._copy_one_value(self.doc, excl_wb, 11, mapping)
            mock_output_sheet.cell.assert_called_once_with(
                row=12,
                column=2,
                value="Client Name in Few Words"
            )


    def test_copy_one_value_for_range(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        mapping = {"page": 0, "key": "sample depth (m):", "till": "liquid limit:", "num": 1, "out": "H"}

        with patch('xcl_extractor.cnv.convert_na', side_effect=identity_fun), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=8):

            pdf._copy_one_value(self.doc, excl_wb, 3, mapping)

        mock_output_sheet.cell.assert_called_once_with(
            row=4,  # last_row + 1
            column=8,  # column B
            value=57.25
        )

    # UC
    # def test_extract_file(self): #UC just for testing
    #     pdf.extract_file("tests/sample1.pdf", "tests/output.xlsx")
