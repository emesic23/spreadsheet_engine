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

    def testFormulaSameSheet(self):
        """Test formulas within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a4+a5')
        self.wb.set_cell_contents(self.name_a, 'a4', '-5')
        self.wb.set_cell_contents(self.name_a, 'a5', '-5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=a1+a2*a3')

        self.assertTrue(decimal.Decimal('-28') == self.wb.get_cell_value(self.name_a, 'b1'))

        self.wb.set_cell_contents(self.name_a, 'a5', '4')

        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b1'))

    def testNumberWithParseableString(self):
        """Test number plus a string that can be parsed as a number"""
        self.wb.set_cell_contents('Sheet1', 'a1', '12')
        self.wb.set_cell_contents('Sheet1', 'a2', "'14")
        self.wb.set_cell_contents('Sheet1', 'a3', '=a1+a2')

        self.assertEqual(decimal.Decimal('12'), self.wb.get_cell_value('Sheet1', 'a1'))
        self.assertEqual("14", self.wb.get_cell_value('Sheet1', 'a2'))
        self.assertEqual(decimal.Decimal('26'), self.wb.get_cell_value('Sheet1', 'a3'))

    def testNumberWithNonparseableString(self):
        """Test number plus a string that cannot be parsed as a number"""
        self.wb.set_cell_contents('Sheet1', 'a1', '12')
        self.wb.set_cell_contents('Sheet1', 'a2', "'14a")
        self.wb.set_cell_contents('Sheet1', 'a3', '=a1+a2')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3').get_type() == sheets.CellErrorType.TYPE_ERROR)

    def testAdditionWeirdSpacing(self):
        """Test addition with weird spacing"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1 + 5')

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'b1'))

        self.wb.set_cell_contents(self.name_a, 'a1', ' 1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1 + 5')

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'b1'))

        self.wb.set_cell_contents(self.name_a, 'a1', '    1      ')
        self.wb.set_cell_contents(self.name_a, 'b1', '   =     a1           + 5       ')

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'b1'))

    def testAdditionSameSheet(self):
        """Test addition within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-10')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1+a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=a1 +a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=a1+ a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=a1 + a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= a1 + a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = a1 + a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = a1 + 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= a1 + 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= a1 + 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 + a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1+a2+a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=a1 +a2   + a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = a1 + a2 + a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(a1 +a2   + a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = a1 + (a2 + a3)')

        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('13') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('-10') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testAdditionDifferentSheets(self):
        """Test addition outside the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a2', '4')
        self.wb.set_cell_contents(self.name_c, 'a3', '-10')

        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!a1+Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=Sheet1!a1 +Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=Sheet1!a1+ Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=Sheet1!a1 + Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= Sheet1!a1 + Sheet2!a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = Sheet1!a1 + Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = Sheet1!a1 + 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= Sheet1!a1 + 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= Sheet1!a1 + 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 + Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=Sheet1!a1+Sheet2!a2+Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=Sheet1!a1 +Sheet2!a2   + Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = Sheet1!a1 + Sheet2!a2 + Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(Sheet1!a1 +Sheet2!a2   + Sheet3!a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = Sheet1!a1 + (Sheet2!a2 + Sheet3!a3)')

        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('16') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('13') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('-10') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testSubtractionSameSheet(self):
        """Test subtraction within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-10')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1-a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=a1 -a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=a1- a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=a1 - a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= a1 - a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = a1 - a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = a1 - 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= a1 - 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= a1 - 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 - a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1-a2-a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=a1 -a2   - a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = a1 - a2 - a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(a1 -a2   - a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = a1 - (a2 - a3)')

        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('11') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-2') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testSubtractionDifferentSheets(self):
        """Tests for subtraction outside of the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a2', '4')
        self.wb.set_cell_contents(self.name_c, 'a3', '-10')

        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!a1-Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=Sheet1!a1 -Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=Sheet1!a1- Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=Sheet1!a1 - Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= Sheet1!a1 - Sheet2!a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = Sheet1!a1 - Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = Sheet1!a1 - 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= Sheet1!a1 - 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= Sheet1!a1 - 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 - Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=Sheet1!a1-Sheet2!a2-Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=Sheet1!a1 -Sheet2!a2   - Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = Sheet1!a1 - Sheet2!a2 - Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(Sheet1!a1 -Sheet2!a2   - Sheet3!a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = Sheet1!a1 - (Sheet2!a2 - Sheet3!a3)')

        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('8') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('11') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('18') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-2') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testMultiplicationSameSheet(self):
        """Test multiplication within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-10')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1*a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=a1 *a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=a1* a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=a1 * a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= a1 * a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = a1 * a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = a1 * 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= a1 * 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= a1 * 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 * a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1*a2*a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=a1 *a2   * a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = a1 * a2 * a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(a1 *a2   * a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = a1 * (a2 * a3)')

        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testMultiplicationDifferentSheets(self):
        """Tests for multiplication outside of the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a2', '4')
        self.wb.set_cell_contents(self.name_c, 'a3', '-10')

        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!a1*Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=Sheet1!a1 *Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=Sheet1!a1* Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=Sheet1!a1 * Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= Sheet1!a1 * Sheet2!a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = Sheet1!a1 * Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = Sheet1!a1 * 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= Sheet1!a1 * 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= Sheet1!a1 * 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 * Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=Sheet1!a1*Sheet2!a2*Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=Sheet1!a1 *Sheet2!a2   * Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = Sheet1!a1 * Sheet2!a2 * Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(Sheet1!a1 *Sheet2!a2   * Sheet3!a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = Sheet1!a1 * (Sheet2!a2 * Sheet3!a3)')

        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('48') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-480') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testDivisionSameSheet(self):
        """Test division within the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'a2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-10')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1/a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=a1 /a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=a1/ a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=a1 / a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= a1 / a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = a1 / a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = a1 / 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= a1 / 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= a1 / 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 / a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1/a2/a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=a1 /a2   / a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = a1 / a2 / a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(a1 /a2   / a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = a1 / (a2 / a3)')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c2').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-30') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testDivisionDifferentSheets(self):
        """Tests for division outside the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_b, 'a2', '4')
        self.wb.set_cell_contents(self.name_c, 'a3', '-10')

        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!a1/Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b2', '=Sheet1!a1 /Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b3', '=Sheet1!a1/ Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b4', '=Sheet1!a1 / Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'b5', '= Sheet1!a1 / Sheet2!a2    ')
        self.wb.set_cell_contents(self.name_a, 'b6', '   = Sheet1!a1 / Sheet2!a2')
        self.wb.set_cell_contents(self.name_a, 'c1', '   = Sheet1!a1 / 0')
        self.wb.set_cell_contents(self.name_a, 'c2', '= Sheet1!a1 / 0')
        self.wb.set_cell_contents(self.name_a, 'c3', '= Sheet1!a1 / 1')
        self.wb.set_cell_contents(self.name_a, 'c4', '= 0 / Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd1', '=Sheet1!a1/Sheet2!a2/Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd2', '=Sheet1!a1 /Sheet2!a2   / Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'd3', '   = Sheet1!a1 / Sheet2!a2 / Sheet3!a3')
        self.wb.set_cell_contents(self.name_a, 'e1', '=(Sheet1!a1 /Sheet2!a2   / Sheet3!a3 )')
        self.wb.set_cell_contents(self.name_a, 'e2', '   = Sheet1!a1 / (Sheet2!a2 / Sheet3!a3)')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b6'))

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c2').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(decimal.Decimal('12') == self.wb.get_cell_value(self.name_a, 'c3'))
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'c4'))

        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'd3'))

        self.assertTrue(decimal.Decimal('-0.3') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('-30') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testDivisionWithRounding(self):
        """Tests division with numbers getting rounded"""
        self.wb.set_cell_contents(self.name_a, 'a1', '3')
        self.wb.set_cell_contents(self.name_a, 'b1', '10')
        self.wb.set_cell_contents(self.name_a, 'c1', '=3/a1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1/3')

        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'd1'))

    def testDivisionWithoutRounding(self):
        """Tests division with numbers getting rounded"""
        self.wb.set_cell_contents(self.name_a, 'a1', '3.0')
        self.wb.set_cell_contents(self.name_a, 'b1', '10.0')
        self.wb.set_cell_contents(self.name_a, 'c1', '=3.0/a1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=a1/3.0')
        self.wb.set_cell_contents(self.name_a, 'e1', '=a1/b1')
        self.wb.set_cell_contents(self.name_a, 'e2', '=b1/a1')

        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('1') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('0.3') == self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertTrue(decimal.Decimal('3.333333333333333333333333333') == self.wb.get_cell_value(self.name_a, 'e2'))

    def testEdgeCases(self):
        """Test various edge cases caught by previous smoke tests"""
        self.wb.set_cell_contents(self.name_a, 'a1', '45')
        self.wb.set_cell_contents(self.name_a, 'b1', '=-5+a1')
        self.wb.set_cell_contents(self.name_a, 'c1', '-5')
        self.wb.set_cell_contents(self.name_a, 'd1', '=-5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=5 + 4')
        self.wb.set_cell_contents(self.name_a, 'b2', '=5 * 4')
        self.wb.set_cell_contents(self.name_a, 'c2', '="Yes, " & "please"')
        self.wb.set_cell_contents(self.name_a, 'd2', '=-d1')

        self.assertTrue(decimal.Decimal('40') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('-5') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('-5') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('9') == self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertTrue("Yes, please" == self.wb.get_cell_value(self.name_a, 'c2'))


if __name__ == '__main__':
    unittest.main()
