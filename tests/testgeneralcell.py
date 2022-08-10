"""
Caltech CS130 - Winter 2022

Testing general cell functions. Checks overwriting
cell types with other cell types
"""


import unittest
import decimal

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
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testUpdatingValuesNumberandString(self):
        """Test updating of numbers to strings and vice versa"""

        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a1', '3')
        self.wb.set_cell_contents(self.name_c, 'a1', '6')

        self.wb.set_cell_contents(self.name_a, 'b1', 'apple')
        self.wb.set_cell_contents(self.name_b, 'b1', 'banana')
        self.wb.set_cell_contents(self.name_c, 'b1', 'peach')

        # Updates
        self.wb.set_cell_contents(self.name_a, 'a1', 'hello')
        self.wb.set_cell_contents(self.name_b, 'a1', '   hi')
        self.wb.set_cell_contents(self.name_c, 'a1', 'hey  ')

        self.wb.set_cell_contents(self.name_a, 'b1', '5')
        self.wb.set_cell_contents(self.name_b, 'b1', '   12')
        self.wb.set_cell_contents(self.name_c, 'b1', '3  ')

        # Checks
        self.assertTrue('hello' == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue('hi' == self.wb.get_cell_value(self.name_b, 'a1'))
        self.assertTrue('hey' == self.wb.get_cell_value(self.name_c, 'a1'))

        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_c, 'b1'))


if __name__ == '__main__':
    unittest.main()
