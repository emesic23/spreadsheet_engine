"""
Caltech CS130 - Winter 2022

Test basic number functionality
"""

import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestBasicNumbers(unittest.TestCase):
    """Test basic number functionality"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testEmptyCellIsNone(self):
        """Test empty cell"""
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'z30'))

    def testBlankCellReferenceZero(self):
        """Test blank cells are the same as zero"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=12')
        self.wb.set_cell_contents(self.name_a, 'a2', '= a1 + z30')
        self.wb.set_cell_contents(self.name_a, 'a3', '= z29 + z30')

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'a3'))

    def testValues(self):
        """Test normal insertion of numbers and retrieve in different order"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_a, 'a2', '  4')
        self.wb.set_cell_contents(self.name_a, 'b2', '2')

        self.wb.set_cell_contents(self.name_b, 'a1', '2.0000')
        self.wb.set_cell_contents(self.name_b, 'a2', '21')
        self.wb.set_cell_contents(self.name_b, 'b2', '3')

        self.wb.set_cell_contents(self.name_c, 'a1', '2')
        self.wb.set_cell_contents(self.name_c, 'a2', '21')
        self.wb.set_cell_contents(self.name_c, 'b2', '3')

        self.assertTrue(decimal.Decimal('4') == self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'b2'))

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b2'))
        self.assertTrue(decimal.Decimal('21') == self.wb.get_cell_value(self.name_b, 'a2'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_b, 'a1'))

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_c, 'a1'))
        self.assertTrue(decimal.Decimal('21') == self.wb.get_cell_value(self.name_c, 'a2'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_c, 'b2'))

    def testContents(self):
        """Test insertion with spaces for value and contents"""
        self.wb.set_cell_contents(self.name_a, 'a1', '  12')
        self.wb.set_cell_contents(self.name_a, 'a2', ' 4')
        self.wb.set_cell_contents(self.name_a, 'b2', '2  ')

        self.wb.set_cell_contents(self.name_b, 'a1', '2 ')
        self.wb.set_cell_contents(self.name_b, 'a2', ' 21 ')
        self.wb.set_cell_contents(self.name_b, 'b2', '  3  ')

        # Test empty cells
        self.assertIsNone(self.wb.get_cell_contents(self.name_c, 'b8'))

        self.assertEqual('4', self.wb.get_cell_contents(self.name_a, 'a2'))
        self.assertEqual('12', self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertEqual('2', self.wb.get_cell_contents(self.name_a, 'b2'))

        self.assertEqual('3', self.wb.get_cell_contents(self.name_b, 'b2'))
        self.assertEqual('21', self.wb.get_cell_contents(self.name_b, 'a2'))
        self.assertEqual('2', self.wb.get_cell_contents(self.name_b, 'a1'))

    def testUpdatingValues(self):
        """Test updating of numbers"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_a, 'a2', '  4')
        self.wb.set_cell_contents(self.name_a, 'b2', '2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=$b$4')
        self.wb.set_cell_contents(self.name_a, 'b4', '6')

        self.wb.set_cell_contents(self.name_b, 'a1', '2')
        self.wb.set_cell_contents(self.name_b, 'a2', '21')
        self.wb.set_cell_contents(self.name_b, 'b2', '3')
        self.wb.set_cell_contents(self.name_b, 'b3', '=b$4')
        self.wb.set_cell_contents(self.name_b, 'b4', '7')

        self.wb.set_cell_contents(self.name_c, 'a1', '2')
        self.wb.set_cell_contents(self.name_c, 'a2', '21')
        self.wb.set_cell_contents(self.name_c, 'b2', '3')
        self.wb.set_cell_contents(self.name_c, 'b3', '=b$4')
        self.wb.set_cell_contents(self.name_c, 'b4', '8')

        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertEqual(decimal.Decimal('7'), self.wb.get_cell_value(self.name_b, 'b3'))
        self.assertEqual(decimal.Decimal('8'), self.wb.get_cell_value(self.name_c, 'b3'))

        # Updates
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_b, 'a2', '9')
        self.wb.set_cell_contents(self.name_c, 'b2', '2')
        self.wb.set_cell_contents(self.name_a, 'b4', '3')
        self.wb.set_cell_contents(self.name_b, 'b4', '4')
        self.wb.set_cell_contents(self.name_c, 'b4', '5')

        self.assertTrue(decimal.Decimal('4') == self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'a1'))  # Updated
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertEqual(decimal.Decimal('3'), self.wb.get_cell_value(self.name_a, 'b3'))  # Updated

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b2'))
        self.assertTrue(decimal.Decimal('9') == self.wb.get_cell_value(self.name_b, 'a2'))  # Updated
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_b, 'a1'))
        self.assertEqual(decimal.Decimal('4'), self.wb.get_cell_value(self.name_b, 'b3'))  # Updated

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_c, 'a1'))
        self.assertTrue(decimal.Decimal('21') == self.wb.get_cell_value(self.name_c, 'a2'))
        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_c, 'b2'))  # Updated
        self.assertEqual(decimal.Decimal('5'), self.wb.get_cell_value(self.name_c, 'b3'))  # Updated

    def testReferenceWithinSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'a2'))

    def testUpdatingReferenceWithinSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')
        self.wb.set_cell_contents(self.name_a, 'a1', ' 42')

        self.assertTrue(decimal.Decimal('42') == self.wb.get_cell_value(self.name_a, 'a2'))

    def testReferenceOutsideSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_b, 'a2', '=Sheet1!a1')

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_b, 'a2'))

    def testUpdatingReferenceOutsideSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', ' 12')
        self.wb.set_cell_contents(self.name_b, 'a2', '=Sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'a1', ' 42')

        self.assertTrue(decimal.Decimal('42') == self.wb.get_cell_value(self.name_b, 'a2'))


if __name__ == '__main__':
    unittest.main()
