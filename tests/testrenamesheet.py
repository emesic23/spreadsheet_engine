"""
Caltech CS130 - Winter 2022

Tests to ensure that sheetnames are being generated, stored
deleted, and remade correctly. Also includes case insensitivity
checks
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestRenameSheet(unittest.TestCase):
    """
    Tests to ensure that sheetnames are being generated, stored
    deleted, and remade correctly. Also includes case insensitivity
    checks
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()
        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet('My Sheet')  # MySheet
        (self.index_d, self.name_d) = self.wb.new_sheet('.?!,: ;!@#$%^&*()-_')  # .?!,:;!@#$%^&*()-_
        (self.index_e, self.name_e) = self.wb.new_sheet('M!y !Sheet 99')  # MySheet

    def tearDown(self):
        pass

    def testSheetNamesValidUserSpecified(self):
        """Test various valid user specified sheet names"""
        self.wb.rename_sheet('Sheet1', 'mY sHeEt 2')  # My Sheet
        # (index_b, name_b) = wb.new_sheet('mY sHeEt 2')  # mY sHeEt 2
        self.wb.rename_sheet('.?!,: ;!@#$%^&*()-_', '1?!,  :;!@#')  # .?!,:;!@#$%^&*()-_
        # Already exists in other case
        self.wb.rename_sheet('My Sheet', 'my sheet')
        # Sheet1 doesn't exist anymore so should be Sheet1
        self.wb.new_sheet()  # Sheet1

        sheetNames = self.wb.list_sheets()

        self.assertTrue(sheetNames[0] == 'mY sHeEt 2')
        self.assertTrue(sheetNames[1] == 'Sheet2')
        self.assertTrue(sheetNames[2] == 'my sheet')
        self.assertTrue(sheetNames[3] == '1?!,  :;!@#')
        self.assertTrue(sheetNames[4] == 'M!y !Sheet 99')
        self.assertTrue(sheetNames[5] == 'Sheet1')

    def testSheetNamesInvalidUserSpecified(self):
        """Test error raising when user renames sheet poorly"""

        # No empty strings
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', '')
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', '   ')

        # No start/end with white space
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', ' name')
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', 'name ')

        # No quotes
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', '\"my sheet\"')
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', '\'my sheet\'')

        # Empty string
        self.assertRaises(ValueError, self.wb.rename_sheet, 'My Sheet', "")

    def testFormulaUpdatesOnRename(self):
        """Test that formulas in cells update after sheet name is changed"""

        self.wb.set_cell_contents('Sheet1', 'a1', '3.0')
        self.wb.set_cell_contents('Sheet2', 'a1', '=Sheet1!a1')
        self.wb.set_cell_contents('Sheet2', 'b1', '=Sheet1DoesntExist!a1 + testingSheet1!a1 + Sheet1!a1')
        self.wb.set_cell_contents('Sheet2', 'c1', '=Sheet1!a1 + Sheet1!a1 + Sheet1!a1')
        self.wb.set_cell_contents('My Sheet', 'a1', "='.?!,: ;!@#$%^&*()-_'!a1")

        self.wb.rename_sheet('Sheet1', 'Renamed')
        self.assertTrue(self.wb.get_cell_contents('Sheet2', 'a1') == '=Renamed!a1')
        self.assertTrue(self.wb.get_cell_contents('Sheet2', 'b1') == '=Sheet1DoesntExist!a1+testingSheet1!a1+Renamed!a1')
        self.assertTrue(self.wb.get_cell_contents('Sheet2', 'c1') == '=Renamed!a1+Renamed!a1+Renamed!a1')
        self.assertTrue(self.wb.get_cell_value('Sheet2', 'a1') == decimal.Decimal(3.0))
        self.assertTrue(self.wb.get_cell_value('Sheet2', 'c1') == decimal.Decimal(9.0))

    def testDiamondRename(self):
        """Test rename with a diamond"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + d1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=3')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'a1'))

        self.wb.rename_sheet(self.name_a, "Sheet10")

        self.wb.set_cell_contents("Sheet10", 'c1', '=10')
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'c1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'b1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'd1'))
        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value("Sheet10", 'a1'))

    def testRenameResurrect(self):
        """Test resurrect with rename"""
        self.wb.set_cell_contents('Sheet1', 'a1', "='Sheet !100'!a1")
        self.wb.set_cell_contents('Sheet1', 'b1', '=Sheet300!a1')
        self.wb.set_cell_contents('Sheet1', 'c1', "='M!y !Sheet 99'!a1")
        self.wb.set_cell_contents(self.name_c, 'a1', '=1')
        self.wb.set_cell_contents(self.name_e, 'a1', '=1')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.wb.rename_sheet(self.name_c, "Sheet !100")
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value("Sheet1", 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value("Sheet1", 'a1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.BAD_REFERENCE)
        self.wb.rename_sheet("Sheet !100", 'Sheet300')
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value("Sheet1", 'b1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value("Sheet1", 'a1'))


if __name__ == '__main__':
    unittest.main()
