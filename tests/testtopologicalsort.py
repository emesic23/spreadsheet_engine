"""
Caltech CS130 - Winter 2022

Testing topological sort
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestTopologicalSort(unittest.TestCase):
    """Testing topological sort"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testDiamond(self):
        """Test a small diamond"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + d1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=3')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'a1'))

        self.wb.set_cell_contents(self.name_a, 'c1', '=10')

        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testLargerDiamond(self):
        """Test a large diamond"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + c1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=d1 + e1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=e1 + f1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=g1 + h1')
        self.wb.set_cell_contents(self.name_a, 'e1', '=h1 + i1')
        self.wb.set_cell_contents(self.name_a, 'f1', '=i1 + j1')
        self.wb.set_cell_contents(self.name_a, 'g1', '=k1')
        self.wb.set_cell_contents(self.name_a, 'h1', '=k1 + l1')
        self.wb.set_cell_contents(self.name_a, 'i1', '=l1 + m1')
        self.wb.set_cell_contents(self.name_a, 'j1', '=m1')
        self.wb.set_cell_contents(self.name_a, 'k1', '=n1')
        self.wb.set_cell_contents(self.name_a, 'l1', '=n1 + o1')
        self.wb.set_cell_contents(self.name_a, 'm1', '=o1')
        self.wb.set_cell_contents(self.name_a, 'n1', '=p1')
        self.wb.set_cell_contents(self.name_a, 'o1', '=p1')
        self.wb.set_cell_contents(self.name_a, 'p1', '=1')

        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testTop(self):
        """Test a basic sort"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + c1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=1')

        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'a1'))

        self.wb.set_cell_contents(self.name_a, 'c1', '=2')

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('4') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testTopDiffSheets(self):
        """Test a basic sort over multiple sheets"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=sheet2!b1 + sheet3!c1')
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet3!c1')
        self.wb.set_cell_contents(self.name_c, 'c1', '=1')

        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_c, 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'a1'))

        self.wb.set_cell_contents(self.name_c, 'c1', '=2')

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_c, 'c1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('4') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testDependencies(self):
        """Test daisy chained dependencies"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'e1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'f1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'g1', '=h1')
        self.wb.set_cell_contents(self.name_a, 'h1', '=1')

        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'f1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'g1'))

        self.wb.set_cell_contents(self.name_a, 'h1', '=2')

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'f1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'g1'))


if __name__ == '__main__':
    unittest.main()
