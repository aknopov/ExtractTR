import math
import unittest
import converter as cnv


class TestConverter(unittest.TestCase):

    def setUp(self):
        pass

    def tearDown(self):
        pass


    def test_is_number(self):
        self.assertTrue(cnv.is_number("123"))
        self.assertTrue(cnv.is_number("-123.45"))
        self.assertTrue(cnv.is_number("+123.45"))
        self.assertTrue(cnv.is_number("nan"))
        self.assertFalse(cnv.is_number("abc"))
        self.assertFalse(cnv.is_number("123-456"))

    def test_is_range(self):
        self.assertTrue(cnv.is_range("123 - 456"))
        self.assertTrue(cnv.is_range("-123.45 - 678.90"))
        self.assertFalse(cnv.is_range("123"))
        self.assertFalse(cnv.is_range("abc - def"))

    def test_convert_value(self):
        self.assertEqual(cnv.convert_value("123"), 123.0)
        self.assertEqual(cnv.convert_value("10 - 20"), 15.0)
        self.assertTrue(math.isnan(cnv.convert_value("abc")))

    def test_convert_range(self):
        self.assertEqual(cnv.convert_range("10 - 20"), 15.0)
        self.assertEqual(cnv.convert_range("-10 - 10"), 0.0)
        self.assertEqual(cnv.convert_range("5.5-7.5"), 6.5)
        self.assertEqual(cnv.convert_range("-10-10"), 0.0)
        self.assertEqual(cnv.convert_range("8 - -6"), 1.0)
        self.assertEqual(cnv.convert_range("-8 - -6"), -7.0)
        self.assertTrue(math.isnan(cnv.convert_range("abc - def")))
        self.assertTrue(math.isnan(cnv.convert_range("boo")))

    def test_col_name_to_idx(self):
        self.assertEqual(cnv.col_name_to_idx("A"), 1)
        self.assertEqual(cnv.col_name_to_idx("B"), 2)
        self.assertEqual(cnv.col_name_to_idx("Z"), 26)
        self.assertEqual(cnv.col_name_to_idx("AA"), 27)
        self.assertEqual(cnv.col_name_to_idx("AB"), 28)
        self.assertEqual(cnv.col_name_to_idx("AZ"), 52)
        self.assertEqual(cnv.col_name_to_idx("BA"), 53)
        self.assertEqual(cnv.col_name_to_idx("CG"), 85)
        self.assertEqual(cnv.col_name_to_idx("ABC"), 0)

    def test_col_indes_to_name(self):
        self.assertEqual(cnv.col_idx_to_name(1), "A")
        self.assertEqual(cnv.col_idx_to_name(26), "Z")
        self.assertEqual(cnv.col_idx_to_name(27), "AA")
        self.assertEqual(cnv.col_idx_to_name(28), "AB")
        self.assertEqual(cnv.col_idx_to_name(52), "AZ")
        self.assertEqual(cnv.col_idx_to_name(53), "BA")
        self.assertEqual(cnv.col_idx_to_name(78), "BZ")
        self.assertEqual(cnv.col_idx_to_name(79), "CA")
        self.assertEqual(cnv.col_idx_to_name(85), "CG")
        self.assertEqual(cnv.col_idx_to_name(26**2-1), None)

    def test_convert_na(self):
        self.assertEqual(cnv.convert_na(None), cnv.N_A)
        self.assertEqual(cnv.convert_na('#NUM!'), cnv.N_A)
        self.assertEqual(cnv.convert_na(math.nan), cnv.N_A)
        self.assertEqual(cnv.convert_na("text"), "text")
        self.assertEqual(cnv.convert_na(1.23), 1.23)

    def test_may_be_convert(self):
        self.assertEqual(cnv.may_be_convert(None), None)
        self.assertEqual(cnv.may_be_convert(1.23), 1.23)
        self.assertEqual(cnv.may_be_convert("1.23"), 1.23)
        self.assertEqual(cnv.may_be_convert("2 - 4"), 3.0)
