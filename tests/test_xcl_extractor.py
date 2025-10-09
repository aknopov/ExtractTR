import unittest
from unittest.mock import Mock, patch
import converter as cnv
import xcl_extractor as xcl
import math


# pylint: disable=protected-access
class TestExtractor(unittest.TestCase):

    def test_print_mappings(self):
        xcl._print_mappings()


    def test_new_file_name(self):
        self.assertEqual(xcl._new_file_name("file.ext"), "file.org.ext")
        self.assertEqual(xcl._new_file_name("/path/file.ext"), "/path/file.org.ext")
        self.assertEqual(xcl._new_file_name("file"), "file.org")


    def test_max_row(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        mock_output_sheet.max_row = 123
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        self.assertEqual(xcl.max_row(excl_wb), 123)


    def test_insert_one_value(self):
        wb_out = Mock()
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet
        excl_wb = xcl.ExcelWorkbook(wb_out)

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=42.5), \
             patch('xcl_extractor.cnv.convert_na', return_value=42.5), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):

            xcl.insert_one_value("foo", excl_wb, 10, "B")

            # Verify the cell was written to the correct location
            mock_output_sheet.cell.assert_called_once_with(
                row=10,  # last_row
                column=2,  # column B
                value=42.5
            )


    def test_copy_one_value_basic(self):
        """Test basic value copying without additional parameters"""
        # Create mock workbooks
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cell
        mock_sheet = Mock()
        mock_cell = Mock()
        mock_cell.value = 42.5
        mock_sheet.__getitem__ = Mock(return_value=mock_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "out": "B",
            "offset": 0
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=42.5), \
             patch('xcl_extractor.cnv.convert_na', return_value=42.5), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=2):

            xcl._copy_one_value(wb_in, wb_out, 10, mapping)

            # Verify the cell was written to the correct location
            mock_output_sheet.cell.assert_called_once_with(
                row=11,  # last_row + offset + 1 = 10 + 0 + 1
                column=2,  # column B
                value=42.5
            )

    def test_copy_one_value_with_in2(self):
        """Test value copying with two input cells (in1 + in2)"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = 10.0
        mock_cell2 = Mock()
        mock_cell2.value = 20.0
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_cell2)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with in2
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "in2": "D7",
            "out": "C",
            "offset": 1
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', side_effect=[10.0, 20.0]), \
             patch('xcl_extractor.cnv.convert_na', return_value=15.0), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=3):

            xcl._copy_one_value(wb_in, wb_out, 5, mapping)

            # Verify the average value was written
            mock_output_sheet.cell.assert_called_once_with(
                row=7,  # last_row + offset + 1 = 5 + 1 + 1
                column=3,  # column C
                value=15.0
            )

    def test_copy_one_value_with_conditional(self):
        """Test value copying with conditional logic"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = 100.0
        mock_condition_cell = Mock()
        mock_condition_cell.value = "CU"  # This matches the condition
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_condition_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with conditional
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "out": "D",
            "offset": 2,
            "if": "L11",
            "is": "CU"
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=100.0), \
             patch('xcl_extractor.cnv.convert_na', return_value=100.0), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=4):

            xcl._copy_one_value(wb_in, wb_out, 8, mapping)

            # Verify the cell was written when condition is met
            mock_output_sheet.cell.assert_called_once_with(
                row=11,  # last_row + offset + 1 = 8 + 2 + 1
                column=4,  # column D
                value=100.0
            )

    def test_copy_one_value_conditional_false(self):
        """Test value copying when conditional logic is false"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = 100.0
        mock_condition_cell = Mock()
        mock_condition_cell.value = "CD"  # This doesn't match the condition
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_condition_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with conditional
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "out": "D",
            "offset": 0,
            "if": "L11",
            "is": "CU"
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=100.0), \
             patch('xcl_extractor.cnv.convert_na', return_value=100.0), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=4):

            xcl._copy_one_value(wb_in, wb_out, 10, mapping)

            # Verify no cell was written when condition is not met
            mock_output_sheet.cell.assert_not_called()

    def test_copy_one_value_conditional_no_input(self):
        """Test value copying when conditional logic is false"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = None
        mock_condition_cell = Mock()
        mock_condition_cell.value = "CU"  # This matches the condition
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_condition_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with conditional
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "out": "D",
            "offset": 0,
            "if": "L11",
            "is": "CU"
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=None), \
             patch('xcl_extractor.cnv.convert_na', return_value=cnv.N_A), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=4):

            xcl._copy_one_value(wb_in, wb_out, 10, mapping)

            # Verify the cell was written when condition is met
            mock_output_sheet.cell.assert_called_once_with(
                row=11,
                column=4,
                value=cnv.N_A
            )

    def test_copy_one_value_with_nan_in2(self):
        """Test value copying when in2 contains NaN"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = 15.0
        mock_cell2 = Mock()
        mock_cell2.value = float('nan')
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_cell2)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with in2 containing NaN
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "in2": "D7",
            "out": "E",
            "offset": 0
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', side_effect=[15.0, float('nan')]), \
             patch('xcl_extractor.cnv.convert_na', return_value=15.0), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=5):

            xcl._copy_one_value(wb_in, wb_out, 12, mapping)

            # Verify only the first value was used (no averaging due to NaN)
            mock_output_sheet.cell.assert_called_once_with(
                row=13,  # last_row + offset + 1 = 12 + 0 + 1
                column=5,  # column E
                value=15.0
            )

    def test_copy_one_value_with_both_nan(self):
        """Test value copying when both in1 and in2 contain NaN"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cells
        mock_sheet = Mock()
        mock_cell1 = Mock()
        mock_cell1.value = float('nan')
        mock_cell2 = Mock()
        mock_cell2.value = float('nan')
        mock_sheet.__getitem__ = Mock(side_effect=lambda x: mock_cell1 if x == "D6" else mock_cell2)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping with both cells containing NaN
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "in2": "D7",
            "out": "F",
            "offset": 1
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', side_effect=[float('nan'), float('nan')]), \
             patch('xcl_extractor.cnv.convert_na', return_value="N/A"), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=6):

            xcl._copy_one_value(wb_in, wb_out, 15, mapping)

            # Verify the N/A value was written
            mock_output_sheet.cell.assert_called_once_with(
                row=17,  # last_row + offset + 1 = 15 + 1 + 1
                column=6,  # column F
                value="N/A"
            )

    def test_copy_one_value_string_input(self):
        """Test value copying with string input that gets converted"""
        wb_in = Mock()
        wb_out = Mock()

        # Mock the input sheet and cell
        mock_sheet = Mock()
        mock_cell = Mock()
        mock_cell.value = "25.5"
        mock_sheet.__getitem__ = Mock(return_value=mock_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        # Mock the output sheet
        mock_output_sheet = Mock()
        wb_out.active = mock_output_sheet

        # Test mapping
        mapping = {
            "sheet": "Input",
            "in1": "D6",
            "out": "G",
            "offset": 3
        }

        # Mock converter functions
        with patch('xcl_extractor.cnv.may_be_convert', return_value=25.5), \
             patch('xcl_extractor.cnv.convert_na', return_value=25.5), \
             patch('xcl_extractor.cnv.col_name_to_idx', return_value=7):

            xcl._copy_one_value(wb_in, wb_out, 20, mapping)

            # Verify the converted value was written
            mock_output_sheet.cell.assert_called_once_with(
                row=24,  # last_row + offset + 1 = 20 + 3 + 1
                column=7,  # column G
                value=25.5
            )

    def test_get_cell_value_success(self):
        """Test successful cell value extraction"""
        wb_in = Mock()
        mock_sheet = Mock()
        mock_cell = Mock()
        mock_cell.value = 42.5

        # Mock the sheet access
        mock_sheet.__getitem__ = Mock(return_value=mock_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        result = xcl._get_cell_value(wb_in, "Input", "D6")

        self.assertEqual(result, 42.5)
        wb_in.__getitem__.assert_called_once_with("Input")
        mock_sheet.__getitem__.assert_called_once_with("D6")

    def test_get_cell_value_string(self):
        """Test cell value extraction with string value"""
        wb_in = Mock()
        mock_sheet = Mock()
        mock_cell = Mock()
        mock_cell.value = "Test String"

        mock_sheet.__getitem__ = Mock(return_value=mock_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        result = xcl._get_cell_value(wb_in, "Report", "L6")

        self.assertEqual(result, "Test String")

    def test_get_cell_value_none(self):
        """Test cell value extraction with None value"""
        wb_in = Mock()
        mock_sheet = Mock()
        mock_cell = Mock()
        mock_cell.value = None

        mock_sheet.__getitem__ = Mock(return_value=mock_cell)
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        result = xcl._get_cell_value(wb_in, "Input", "D8")

        self.assertIsNone(result)

    def test_get_cell_value_key_error(self):
        """Test cell value extraction when sheet or cell doesn't exist"""
        wb_in = Mock()

        # Mock KeyError when accessing sheet
        wb_in.__getitem__ = Mock(side_effect=KeyError("Sheet 'NonExistent' not found"))

        with patch('xcl_extractor.log.warning') as mock_log:
            result = xcl._get_cell_value(wb_in, "NonExistent", "A1")

            self.assertTrue(math.isnan(result))
            mock_log.assert_called_once()

    def test_get_cell_value_cell_key_error(self):
        """Test cell value extraction when cell doesn't exist in sheet"""
        wb_in = Mock()
        mock_sheet = Mock()

        # Mock KeyError when accessing cell
        mock_sheet.__getitem__ = Mock(side_effect=KeyError("Cell 'Z999' not found"))
        wb_in.__getitem__ = Mock(return_value=mock_sheet)

        with patch('xcl_extractor.log.warning') as mock_log:
            result = xcl._get_cell_value(wb_in, "Input", "Z999")

            self.assertTrue(math.isnan(result))
            mock_log.assert_called_once()

    def test_is_number(self):
        self.assertTrue(xcl._is_number(3.14159))
        self.assertTrue(xcl._is_number(42))

        self.assertFalse(xcl._is_number(math.nan))
        self.assertFalse(xcl._is_number(None))
        self.assertFalse(xcl._is_number("Hello"))
        self.assertFalse(xcl._is_number(True))
        self.assertTrue(xcl._is_number(float('inf')))  # Infinity is not None and not NaN


if __name__ == "__main__":
    unittest.main()
