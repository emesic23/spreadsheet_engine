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


class TestCopySheet(unittest.TestCase):
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
        (self.index_e, self.name_e) = self.wb.new_sheet('My! Sheet!')  # MySheet

    def tearDown(self):
        pass

    def testBasicCopy(self):
        """Test various valid user specified sheet names"""
        self.wb.set_cell_contents('Sheet1', 'a1', '3.0')
        self.wb.set_cell_contents('Sheet1', 'b1', '=a1 + 9')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1') == decimal.Decimal(12))

        (_, namecpy) = self.wb.copy_sheet(self.name_a)  # My Sheet

        self.assertTrue(namecpy == 'Sheet1_1')
        self.assertTrue(self.wb.get_cell_value(namecpy, 'b1') == decimal.Decimal(12))
        self.wb.set_cell_contents('Sheet1_1', 'a1', '=6')
        self.assertTrue(self.wb.get_cell_value(namecpy, 'b1') == decimal.Decimal(15))

    def testCopyResurrect(self):
        """Test resurrecting a sheet using copy"""
        self.wb.set_cell_contents('Sheet1', 'a1', "=Sheet1_1!a1")

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        (_, namecpy) = self.wb.copy_sheet(self.name_a)  # My Sheet

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(namecpy, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents('Sheet1_1', 'a1', '=3')
        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1') == decimal.Decimal(3))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1') == decimal.Decimal(3))

    def testCopyWithQuotesResurrect(self):
        """Test resurrecting a sheet that has quotes in its name using copy"""
        self.wb.set_cell_contents('My Sheet', 'a1', "='My Sheet_1'!a1")

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_c, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        (_, namecpy) = self.wb.copy_sheet(self.name_c)  # My Sheet

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_c, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(namecpy, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents('My Sheet_1', 'a1', '=3')

        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1') == decimal.Decimal(3))
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'a1') == decimal.Decimal(3))

    def testCopyWithExclamationResurrect(self):
        """Test resurrecting a sheet that has quotes and exclamations in its name using copy"""
        self.wb.set_cell_contents('My! Sheet!', 'a1', "='My! Sheet!_1'!a1")

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_e, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_e, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        (_, namecpy) = self.wb.copy_sheet(self.name_e)  # My! Sheet!

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_e, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_e, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(namecpy, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents('My! Sheet!_1', 'a1', '=3')

        self.assertTrue(self.wb.get_cell_value(namecpy, 'a1') == decimal.Decimal(3))
        self.assertTrue(self.wb.get_cell_value(self.name_e, 'a1') == decimal.Decimal(3))


if __name__ == '__main__':
    unittest.main()
