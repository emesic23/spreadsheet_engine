"""
Caltech CS130 - Winter 2022

Testing general cell functions. Checks overwriting
cell types with other cell types
"""

import unittest

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
from sheets._utils import split_sheet_cell, col_to_num, get_cell_location   # pylint: disable=E0401
from sheets._utils import num_to_col, check_cell_inbounds, get_target_area  # pylint: disable=E0401
from sheets._utils import add_quotes_sheet_name  # pylint: disable=E0401


MAX_COL_NUM = 26 * 26 * 26 * 26 + 26 * 26 * 26 + 26 * 26 + 26
MAX_ROW_NUM = 9999


class TestGeneralCell(unittest.TestCase):
    """
    Testing general cell functions. Checks overwriting
    cell types with other cell types
    """

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def testColToNum(self):
        """Test column to number and make sure proper value errors are raised"""
        self.assertTrue(col_to_num("A") == 1)
        self.assertTrue(col_to_num("a") == 1)
        self.assertTrue(col_to_num("Z") == 26)
        self.assertTrue(col_to_num("z") == 26)
        self.assertTrue(col_to_num("ZZ") == 26 * 26 + 26)
        self.assertTrue(col_to_num("Zz") == 26 * 26 + 26)
        self.assertTrue(col_to_num("zz") == 26 * 26 + 26)

        self.assertRaises(ValueError, col_to_num, "ZZZZZ")
        self.assertRaises(ValueError, col_to_num, "ZZZZA")

    def testNumToCol(self):
        """Test number to column and make sure proper value errors are raised"""
        # Test A through Z
        for ii in range(0, 26):
            self.assertTrue(num_to_col(ii + 1) == chr(ii + 97))

        # Test AA through AZ
        for ii in range(0, 26):
            self.assertTrue(num_to_col(26 + ii + 1) == "a" + str(chr(ii + 97)))

        # Test BA through BZ
        for ii in range(0, 26):
            self.assertTrue(num_to_col(26 * 2 + ii + 1) == "b" + str(chr(ii + 97)))

        self.assertTrue(num_to_col(MAX_COL_NUM) == "zzzz")

        self.assertRaises(ValueError, num_to_col, MAX_COL_NUM + 1)
        self.assertRaises(ValueError, num_to_col, 0)

    def testGetCellLocation(self):
        """Test get cell location and make sure proper value errors are raised"""

        self.assertTrue(get_cell_location("A$1") == (1, 1, False, True))
        self.assertTrue(get_cell_location("a23") == (1, 23, False, False))
        self.assertTrue(get_cell_location("$Z5") == (26, 5, True, False))
        self.assertTrue(get_cell_location("$z$44") == (26, 44, True, True))
        self.assertTrue(get_cell_location("ZZ55") == (26 * 26 + 26, 55, False, False))
        self.assertTrue(get_cell_location("Zz23") == (26 * 26 + 26, 23, False, False))
        self.assertTrue(get_cell_location("zz42") == (26 * 26 + 26, 42, False, False))

        self.assertRaises(ValueError, get_cell_location, "ZZZZZ1")
        self.assertRaises(ValueError, get_cell_location, "ZZZZA1")

        self.assertRaises(ValueError, get_cell_location, "A10000")
        self.assertRaises(ValueError, get_cell_location, "A99999")

    def testSplitSheetCell(self):
        """Test splitting sheet name and cell location"""

        aN, aC = split_sheet_cell("'My Sheet'!A2")
        bN, bC = split_sheet_cell("Sheet1!CDS212")
        cN, cC = split_sheet_cell("'My Sheet'!A2", remove_quotes=False)

        self.assertEqual(aN, "My Sheet")
        self.assertEqual(cN, "'My Sheet'")
        self.assertEqual(aC, "A2")
        self.assertEqual(bN, "Sheet1")
        self.assertEqual(bC, "CDS212")
        self.assertEqual(cC, "A2")

    def testAddQuotesSheetName(self):
        """
        Test function that adds quotes to sheet names with special characters
        """
        self.assertEqual("'Sheet 1'", add_quotes_sheet_name("Sheet 1"))
        self.assertEqual("'Sheet@1'", add_quotes_sheet_name("Sheet@1"))
        self.assertEqual("'Shee!1'", add_quotes_sheet_name("Shee!1"))
        self.assertEqual("Sheet2", add_quotes_sheet_name("Sheet2"))

    def testCheckCellInbounds(self):
        """Test cell (int, int) inbounds"""
        self.assertTrue(check_cell_inbounds(1, 1))
        self.assertTrue(check_cell_inbounds(MAX_COL_NUM, 1))
        self.assertTrue(check_cell_inbounds(1, MAX_ROW_NUM))

        self.assertFalse(check_cell_inbounds(0, 0))
        self.assertFalse(check_cell_inbounds(1, 0))
        self.assertFalse(check_cell_inbounds(0, 1))
        self.assertFalse(check_cell_inbounds(MAX_COL_NUM + 1, 1))
        self.assertFalse(check_cell_inbounds(1, MAX_ROW_NUM + 1))
        self.assertFalse(check_cell_inbounds(MAX_COL_NUM + 1, MAX_ROW_NUM + 1))

    def testGetTargetArea(self):
        """
        Test method that lists all cells within a start and an end location
        inclusive. Start and end can be any corners of the rectangle.
        """

        tmp_row = [(4, 1), (4, 2), (4, 3), (4, 4)]
        tmp_col = [(1, 4), (2, 4), (3, 4), (4, 4)]

        tmp_rect = []
        for col in range(2, 6):
            for row in range(4, 6):
                tmp_rect.append((col, row))

        # Single cell
        self.assertEqual(([(col_to_num('A'), 2)], 1, 1), get_target_area('A2', 'A2'))
        self.assertEqual(([(col_to_num('C'), 12)], 1, 1), get_target_area('C12', 'C12'))

        # Row
        self.assertEqual((tmp_row, 1, 4), get_target_area('d1', 'D4'))
        self.assertEqual((tmp_row, 1, 4), get_target_area('d4', 'D1'))

        # Column
        self.assertEqual((tmp_col, 4, 1), get_target_area('A4', 'D4'))
        self.assertEqual((tmp_col, 4, 1), get_target_area('D4', 'A4'))

        # Rectangle
        self.assertEqual((tmp_rect, 4, 2), get_target_area('B4', 'E5'))
        self.assertEqual((tmp_rect, 4, 2), get_target_area('e4', 'B5'))
        self.assertEqual((tmp_rect, 4, 2), get_target_area('b5', 'e4'))
        self.assertEqual((tmp_rect, 4, 2), get_target_area('E5', 'B4'))

        # Check that value errors returned for cells out of bounds
        self.assertRaises(ValueError, get_target_area, 'ZZZZZ1', 'A1')
        self.assertRaises(ValueError, get_target_area, 'A1', 'ZZZZZ1')
        self.assertRaises(ValueError, get_target_area, 'ZZZ20391', 'A1')
        self.assertRaises(ValueError, get_target_area, 'A1', 'ZZZ20391')


if __name__ == '__main__':
    unittest.main()
