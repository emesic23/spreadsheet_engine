"""
Caltech CS130 - Winter 2022

Testing of the common sheet methods like create, delete, and exetent.
Also includes case insensitivity checks
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestSheetCreating(unittest.TestCase):
    """
    Testing of the common sheet methods like create, delete, and exetent.
    Also includes case insensitivity checks
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        # Create workbook and two sheets
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        self.wb.del_sheet(self.name_b)
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet2

    def tearDown(self):
        pass

    def testCreating(self):
        """Check sheet number and names"""
        sheetNames = self.wb.list_sheets()

        self.assertTrue(self.wb.num_sheets() == 2)
        self.assertTrue(len(sheetNames) == 2)
        self.assertTrue(sheetNames[0] == 'Sheet1')
        self.assertTrue(sheetNames[1] == 'Sheet2')

    def testDeleteAll(self):
        """Test deleting all sheets"""
        sheetNames = self.wb.list_sheets().copy()
        for sheetName in sheetNames:
            try:
                self.wb.del_sheet(sheetName)
            except KeyError as e:
                self.fail("Raised unexpected key error " + str(e))
        self.assertTrue(self.wb.num_sheets() == 0)

    def testExtent(self):
        """Test only addition of cells changing extent"""
        self.wb.set_cell_contents(self.name_a, 'a1', '42')
        self.wb.set_cell_contents(self.name_a, 'b8', '3')

        # Test deletion of a cell changing extent
        self.wb.set_cell_contents(self.name_c, 'c10', '2')
        self.wb.set_cell_contents(self.name_c, 'c6', '3')
        self.wb.set_cell_contents(self.name_c, 'c10', None)

        self.assertEqual((2, 8), self.wb.get_sheet_extent(self.name_a))
        self.assertEqual((3, 6), self.wb.get_sheet_extent(self.name_c))

        # Test deleting all cells
        self.wb.set_cell_contents(self.name_c, 'c6', '')
        self.assertEqual((0, 0), self.wb.get_sheet_extent(self.name_c))

        self.wb.set_cell_contents(self.name_a, 'b8', '')
        self.assertEqual((1, 1), self.wb.get_sheet_extent(self.name_a))

        self.wb.set_cell_contents(self.name_a, 'a1', None)
        self.assertEqual((0, 0), self.wb.get_sheet_extent(self.name_a))

    def testExtentLimits(self):
        """Test [0, 0] and [1 + 4 * 26, 9999] limits"""
        self.wb.set_cell_contents(self.name_c, 'ZZZZ9999', '42')

        # Lower
        self.assertTrue((0, 0) == self.wb.get_sheet_extent(self.name_a))

        # Upper
        self.assertTrue((475254, 9999) == self.wb.get_sheet_extent(self.name_c))

    def testExtentShrinks(self):
        """Test shrinking the extent"""
        self.assertEqual((0, 0), self.wb.get_sheet_extent(self.name_a))
        self.wb.set_cell_contents(self.name_a, 'e1', '1')
        self.wb.set_cell_contents(self.name_a, 'f6', '1')
        self.assertEqual((6, 6), self.wb.get_sheet_extent(self.name_a))
        self.wb.set_cell_contents(self.name_a, 'e1', '')
        self.assertEqual((6, 6), self.wb.get_sheet_extent(self.name_a))
        self.wb.set_cell_contents(self.name_a, 'f6', '')
        self.assertEqual((0, 0), self.wb.get_sheet_extent(self.name_a))

        self.wb.set_cell_contents(self.name_a, 'e1', '1')
        self.wb.set_cell_contents(self.name_a, 'f6', '1')
        self.wb.set_cell_contents(self.name_a, 'f6', '')
        self.assertEqual((5, 1), self.wb.get_sheet_extent(self.name_a))
        self.wb.set_cell_contents(self.name_a, 'f1', '1')
        self.assertEqual((6, 1), self.wb.get_sheet_extent(self.name_a))

    def testCaseInsensitivity(self):
        """Test inserting out of bounds or to invalid cell name"""
        self.wb.set_cell_contents(self.name_a, 'aB39', '42')
        self.wb.set_cell_contents(self.name_a, 'bb8', '42')
        self.assertTrue(decimal.Decimal('42') == self.wb.get_cell_value(self.name_a, 'aB39'))
        self.assertTrue(decimal.Decimal('42') == self.wb.get_cell_value(self.name_a, 'bb8'))

    def testInvalidCells(self):
        """Test inserting out of bounds or to invalid cell name"""
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'ZZZZZ99999', '42')
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'ZZZZZ-99999', '42')
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'A99A9', '42')

    def testKeyErrors(self):
        """Make sure that key errors are generated appropriately"""
        self.assertRaises(KeyError, self.wb.get_sheet_extent, "Sheet Doesnt Exist")
        self.assertRaises(KeyError, self.wb.set_cell_contents, "Sheet Doesnt Exist", 'a1', '1')
        self.assertRaises(KeyError, self.wb.get_cell_contents, "Sheet Doesnt Exist", 'a1')
        self.assertRaises(KeyError, self.wb.del_sheet, "Sheet Doesnt Exist")


if __name__ == '__main__':
    unittest.main()
