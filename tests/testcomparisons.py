"""
Caltech CS130 - Winter 2022

Testing advanced formulas that result in numbers
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestNumberFormulas(unittest.TestCase):
    """Testing advanced formulas that result in numbers"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""

        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testEqualsAndNotEquals(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 = a2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a1 == a2')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a1 <> a2')
        self.wb.set_cell_contents(self.name_a, 'a6', '=a1 != a2')
        self.wb.set_cell_contents(self.name_a, 'b1', '-5')
        self.wb.set_cell_contents(self.name_a, 'b2', '-5')
        self.wb.set_cell_contents(self.name_a, 'b3', '=b1 = b2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=b1 == b2')
        self.wb.set_cell_contents(self.name_a, 'b5', '=b1 <> b2')
        self.wb.set_cell_contents(self.name_a, 'b6', '=b1 != b2')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a6'))

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b6'))

    def testGreaterThanAndLessThan(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 > a2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a1 >= a2')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a1 < a2')
        self.wb.set_cell_contents(self.name_a, 'a6', '=a1 <= a2')
        self.wb.set_cell_contents(self.name_a, 'b1', '-5')
        self.wb.set_cell_contents(self.name_a, 'b2', '-5')
        self.wb.set_cell_contents(self.name_a, 'b3', '=b1 > b2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=b1 >= b2')
        self.wb.set_cell_contents(self.name_a, 'b5', '=b1 < b2')
        self.wb.set_cell_contents(self.name_a, 'b6', '=b1 <= b2')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a6'))

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b6'))

    def testComparisonPrecedence(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'type')
        self.wb.set_cell_contents(self.name_a, 'a2', 'type')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 < a2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a1 < a2 & "type"')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))

    def testEmptyComparisonWithString(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'type')
        self.wb.set_cell_contents(self.name_a, 'a2', '')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 = b2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a2 = b2')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a1 > b2')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a5'))

    def testEmptyComparisonWithNumber(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '123')
        self.wb.set_cell_contents(self.name_a, 'a2', '0')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 = b2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a2 = b2')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a1 > b2')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a5'))

    def testEmptyComparisonWithBool(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', 'false')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 = b2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a2 = b2')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a1 > b2')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a5'))

    def testDifferentTypeComparisons(self):
        """Test comparisons between different types"""
        self.wb.set_cell_contents('Sheet1', 'a1', '123')
        self.wb.set_cell_contents('Sheet1', 'a2', 'abc')
        self.wb.set_cell_contents('Sheet1', 'a3', 'false')
        self.wb.set_cell_contents('Sheet1', 'a4', '=a2 > a1')
        self.wb.set_cell_contents('Sheet1', 'a5', '=a2 > a3')
        self.wb.set_cell_contents('Sheet1', 'a6', '=a1 > a3')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a6'))

    def testComparisonDifferentSheets(self):
        """Test comparisons outside the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a2', '4')

        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!a1 > Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=Sheet1!a1 < Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=Sheet1!a1 == Sheet2!a2')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))

    def testEdgeCases(self):
        """Test various edge cases caught by previous smoke tests"""
        self.wb.set_cell_contents(self.name_a, 'a1', '= 4 < 5')
        self.wb.set_cell_contents(self.name_a, 'b1', '=-5+a1')

        self.wb.set_cell_contents(self.name_b, 'a1', '= 4 > 5')
        self.wb.set_cell_contents(self.name_b, 'b1', '=-5+a1')

        self.assertTrue(decimal.Decimal('-5') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('-5') == self.wb.get_cell_value(self.name_b, 'b1'))


if __name__ == '__main__':
    unittest.main()
