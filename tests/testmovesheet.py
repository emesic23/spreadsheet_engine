"""
Caltech CS130 - Winter 2022

Testing general cell functions. Checks overwriting
cell types with other cell types
"""

import unittest

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestGeneralCell(unittest.TestCase):
    """
    Testing general cell functions. Checks overwriting
    cell types with other cell types
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet("My Sheet")  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet("My! Sheet")  # Sheet3

    def tearDown(self):
        pass

    def testBasicMove(self):
        """Test basic move"""
        self.wb.move_sheet(self.name_a, self.index_c)

        sheet_names = self.wb.list_sheets()
        self.assertTrue(sheet_names[0] == "My Sheet")
        self.assertTrue(sheet_names[1] == "My! Sheet")
        self.assertTrue(sheet_names[2] == "Sheet1")

    def testInvalidMove(self):
        """Test invalid move"""
        self.assertRaises(IndexError, self.wb.move_sheet, self.name_a, len(self.wb.list_sheets()))
        self.assertRaises(IndexError, self.wb.move_sheet, self.name_a, -1)
        self.assertRaises(KeyError, self.wb.move_sheet, "SheetDoesNotExist", 1)


if __name__ == '__main__':
    unittest.main()
