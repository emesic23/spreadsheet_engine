"""
Caltech CS130 - Winter 2022

Testing the move cell function.
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestMoveCell(unittest.TestCase):
    """
    Testing move cell function.
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet("Sheet! 3")  # Sheet 3

    def tearDown(self):
        pass

    def testMoveSingleCell(self):
        """Test moving a single cell and overwriting a cell that is already there"""

        # Test number
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.set_cell_contents(self.name_a, 'b1', '20')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'b1')

        # Test string
        self.wb.set_cell_contents(self.name_a, 'C1', 'abc')
        self.wb.move_cells(self.name_a, 'c1', 'C1', 'd1')

        # Test double quotes
        self.wb.set_cell_contents(self.name_a, 'z1', '"abc"')
        self.wb.move_cells(self.name_a, 'z1', 'z1', 'z2')

        # Assert number
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('12'), self.wb.get_cell_value(self.name_a, 'b1'))

        # Assert string
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'd1'))

        # Assert double quotes
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'z1'))
        self.assertEqual('"abc"', self.wb.get_cell_contents(self.name_a, 'z2'))
        self.assertEqual('"abc"', self.wb.get_cell_value(self.name_a, 'z2'))

    def testMoveSingleCellFormula(self):
        """Test moving a single cell formula"""

        self.wb.set_cell_contents(self.name_a, 'a1', '=a2')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'd5')

        self.wb.set_cell_contents(self.name_a, 'aa1', '=bb1')
        self.wb.move_cells(self.name_a, 'aa1', 'aa1', 'aa2')

        self.wb.set_cell_contents(self.name_a, 'd372', '=78+e373  + d373+f373')
        self.wb.move_cells(self.name_a, 'd372', 'd372', 'f4')

        # Assert moved
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual('=d6', self.wb.get_cell_contents(self.name_a, 'd5'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'd5'))

        # Assert moved
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'aa1'))
        self.assertEqual('=bb2', self.wb.get_cell_contents(self.name_a, 'aa2'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'aa2'))

        # Assert moved
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd372'))
        self.assertEqual('=78+g5+f5+h5', self.wb.get_cell_contents(self.name_a, 'f4'))
        self.assertEqual(decimal.Decimal('78'), self.wb.get_cell_value(self.name_a, 'f4'))

    def testMoveSingleCellFormulaWithAbsRefs(self):
        """Test moving a single cell formula with absolute references"""

        self.wb.set_cell_contents(self.name_a, 'a2', '42')
        self.wb.set_cell_contents(self.name_a, 'a1', '=$A$2')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'd5')

        self.wb.set_cell_contents(self.name_b, 'a2', '42')
        self.wb.set_cell_contents(self.name_b, 'a1', '=a$2')
        self.wb.move_cells(self.name_b, 'a1', 'a1', 'd5')

        self.wb.set_cell_contents(self.name_a, 'aa1', '=$bb5')
        self.wb.move_cells(self.name_a, 'aa1', 'aa1', 'aa2')

        self.wb.set_cell_contents(self.name_a, 'e373', '10')
        self.wb.set_cell_contents(self.name_a, 'd373', '11')
        self.wb.set_cell_contents(self.name_a, 'f373', '-10')
        self.wb.set_cell_contents(self.name_a, 'f5', '-10')
        self.wb.set_cell_contents(self.name_a, 'd372', '=78+$e$373  + d$373+$f373')
        self.wb.move_cells(self.name_a, 'd372', 'd372', 'f4')

        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual('=$a$2', self.wb.get_cell_contents(self.name_a, 'd5'))
        self.assertEqual(decimal.Decimal('42'), self.wb.get_cell_value(self.name_a, 'd5'))

        self.assertIsNone(self.wb.get_cell_contents(self.name_b, 'a1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'a1'))
        self.assertEqual('=d$2', self.wb.get_cell_contents(self.name_b, 'd5'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_b, 'd5'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'aa1'))
        self.assertEqual('=$bb6', self.wb.get_cell_contents(self.name_a, 'aa2'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'aa2'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'd372'))
        self.assertEqual('=78+$e$373+f$373+$f5', self.wb.get_cell_contents(self.name_a, 'f4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd372'))
        self.assertEqual(decimal.Decimal('68'), self.wb.get_cell_value(self.name_a, 'f4'))

    def testMoveSingleCellFormulaMultiSheet(self):
        """Test moving a single cell formula between sheets"""

        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet2!a2')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'd5', self.name_b)

        self.wb.set_cell_contents(self.name_b, 'aa1', '=Sheet1!bb1')
        self.wb.move_cells(self.name_b, 'aa1', 'aa1', 'aa2', self.name_a)

        self.wb.set_cell_contents(self.name_a, 'd372', "=78+Sheet2!e373  + 'Sheet! 3'!d373+Sheet2!f373")
        self.wb.move_cells(self.name_a, 'd372', 'd372', 'f4')

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertEqual('=Sheet2!d6', self.wb.get_cell_contents(self.name_b, 'd5'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_b, 'd5'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_b, 'aa1'))
        self.assertEqual('=Sheet1!bb2', self.wb.get_cell_contents(self.name_a, 'aa2'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'aa2'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'd372'))
        self.assertEqual("=78+Sheet2!g5+'Sheet! 3'!f5+Sheet2!h5", self.wb.get_cell_contents(self.name_a, 'f4'))
        self.assertEqual(decimal.Decimal('78'), self.wb.get_cell_value(self.name_a, 'f4'))

    def testMoveSingleCellFormulaMultiSheetWithAbsRefs(self):
        """Test moving a single cell formula between sheets"""

        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet2!$a2')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'd5', self.name_b)

        self.wb.set_cell_contents(self.name_b, 'aa1', '=Sheet1!bb$1')
        self.wb.move_cells(self.name_b, 'aa1', 'aa1', 'aa2', self.name_a)

        self.wb.set_cell_contents(self.name_a, 'd372', "=78+Sheet2!$e$373  + 'Sheet! 3'!d$373+Sheet2!$f373")
        self.wb.move_cells(self.name_a, 'd372', 'd372', 'f4')

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertEqual('=Sheet2!$a6', self.wb.get_cell_contents(self.name_b, 'd5'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_b, 'd5'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_b, 'aa1'))
        self.assertEqual('=Sheet1!bb$1', self.wb.get_cell_contents(self.name_a, 'aa2'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'aa2'))

        # Assert copied
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'd372'))
        self.assertEqual("=78+Sheet2!$e$373+'Sheet! 3'!f$373+Sheet2!$f5", self.wb.get_cell_contents(self.name_a, 'f4'))
        self.assertEqual(decimal.Decimal('78'), self.wb.get_cell_value(self.name_a, 'f4'))

    def testMoveSingleCellMultiSheet(self):
        """Test moving a single cell between sheets"""

        # Test number
        self.wb.set_cell_contents(self.name_a, 'a1', '12')
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'b1', self.name_b)

        # Test string
        self.wb.set_cell_contents(self.name_a, 'C1', 'abc')
        self.wb.move_cells(self.name_a, 'c1', 'C1', 'd1', self.name_c)

        # Assert number
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('12'), self.wb.get_cell_value(self.name_b, 'b1'))

        # Assert string
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_c, 'd1'))

    def testMoveErrorCell(self):
        """Test moving an error"""

        # Test error
        self.wb.set_cell_contents(self.name_a, 'f1', '#DIV/0!')
        self.wb.move_cells(self.name_a, 'f1', 'f1', 'f2')

        # Assert error
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'f1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'f2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'f2').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)

    def testMoveBlankCells(self):
        """Test moving a blank cells"""

        # Test blank
        self.wb.move_cells(self.name_a, 'a1', 'a1', 'b1')  # Single cell
        self.wb.move_cells(self.name_a, 'd1', 'e4', 'pp1')  # Rectangle

        # Assert blank
        # Single cell
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'b1'))

        # Rectangle
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'e1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'e2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'e3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'e4'))

        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pp1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pp2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pp3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pp4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pq1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pq2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pq3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'pq4'))

    def testMoveRowCol(self):  # pylint: disable=R0915
        """Test moving rows and columns and cells around them"""

        # row = ['b2', 'c2', 'd2', 'e2', 'f2']  # Row on Sheet1
        # col = ['b2', 'b3', 'b4', 'b5', 'b6']  # Col on Sheet2

        # Set up row
        self.wb.set_cell_contents(self.name_a, 'b2', 'moved from b2')
        self.wb.set_cell_contents(self.name_a, 'c2', 'moved from c2')
        self.wb.set_cell_contents(self.name_a, 'd2', 'moved from d2')
        self.wb.set_cell_contents(self.name_a, 'e2', 'moved from e2')
        self.wb.set_cell_contents(self.name_a, 'f2', 'moved from f2')

        # Move row
        self.wb.move_cells(self.name_a, 'b2', 'f2', 'c5', self.name_c)

        # Set up col
        self.wb.set_cell_contents(self.name_b, 'b2', 'moved from b2')
        self.wb.set_cell_contents(self.name_b, 'b3', 'moved from b3')
        self.wb.set_cell_contents(self.name_b, 'b4', 'moved from b4')
        self.wb.set_cell_contents(self.name_b, 'b5', 'moved from b5')
        self.wb.set_cell_contents(self.name_b, 'b6', 'moved from b6')

        # Move col
        self.wb.move_cells(self.name_b, 'b2', 'b6', 'aa12', self.name_c)

        # Check moved row
        self.assertEqual('moved from b2', self.wb.get_cell_value(self.name_c, 'c5'))
        self.assertEqual('moved from c2', self.wb.get_cell_value(self.name_c, 'd5'))
        self.assertEqual('moved from d2', self.wb.get_cell_value(self.name_c, 'e5'))
        self.assertEqual('moved from e2', self.wb.get_cell_value(self.name_c, 'f5'))
        self.assertEqual('moved from f2', self.wb.get_cell_value(self.name_c, 'g5'))

        # Check cells around moved row
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'b4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'b5'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'b6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'c4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'c6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'd4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'd6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'e4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'e6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'f4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'f6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'g4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'g6'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'h4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'h5'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'h6'))

        # Check original cells row
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'c2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'd2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'e2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'f2'))

        # Check moved col
        self.assertEqual('moved from b2', self.wb.get_cell_value(self.name_c, 'aa12'))
        self.assertEqual('moved from b3', self.wb.get_cell_value(self.name_c, 'aa13'))
        self.assertEqual('moved from b4', self.wb.get_cell_value(self.name_c, 'aa14'))
        self.assertEqual('moved from b5', self.wb.get_cell_value(self.name_c, 'aa15'))
        self.assertEqual('moved from b6', self.wb.get_cell_value(self.name_c, 'aa16'))

        # Check cells around moved col
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z11'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'aa11'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab11'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z12'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab12'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z13'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab13'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z14'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab14'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z15'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab15'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z16'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab16'))

        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'z17'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'aa17'))
        self.assertIsNone(self.wb.get_cell_value(self.name_c, 'ab17'))

        # Check original cell col
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'b2'))
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'b3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'b4'))
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'b5'))
        self.assertIsNone(self.wb.get_cell_value(self.name_b, 'b6'))

    def testMoveRect(self):
        """Test moving rows and columns and cells around them"""
        # Set up rectangle
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'a2', '2')
        self.wb.set_cell_contents(self.name_a, 'a3', '3')
        self.wb.set_cell_contents(self.name_a, 'a4', '4')
        self.wb.set_cell_contents(self.name_a, 'a5', '5')
        self.wb.set_cell_contents(self.name_a, 'b1', '6')
        self.wb.set_cell_contents(self.name_a, 'b2', '7')
        self.wb.set_cell_contents(self.name_a, 'b3', '8')
        self.wb.set_cell_contents(self.name_a, 'b4', '9')
        self.wb.set_cell_contents(self.name_a, 'b5', '10')
        self.wb.set_cell_contents(self.name_a, 'c1', '11')
        self.wb.set_cell_contents(self.name_a, 'c2', '12')
        self.wb.set_cell_contents(self.name_a, 'c3', '13')
        self.wb.set_cell_contents(self.name_a, 'c4', '14')
        self.wb.set_cell_contents(self.name_a, 'c5', '15')
        self.wb.set_cell_contents(self.name_a, 'd1', '16')
        self.wb.set_cell_contents(self.name_a, 'd2', '17')
        self.wb.set_cell_contents(self.name_a, 'd3', '18')
        self.wb.set_cell_contents(self.name_a, 'd4', '19')
        self.wb.set_cell_contents(self.name_a, 'd5', '20')

        self.wb.move_cells(self.name_a, 'b2', 'c4', 'ab2')

        # Checked moved cells
        self.assertEqual('7', self.wb.get_cell_contents(self.name_a, 'ab2'))
        self.assertEqual('8', self.wb.get_cell_contents(self.name_a, 'ab3'))
        self.assertEqual('9', self.wb.get_cell_contents(self.name_a, 'ab4'))
        self.assertEqual('12', self.wb.get_cell_contents(self.name_a, 'ac2'))
        self.assertEqual('13', self.wb.get_cell_contents(self.name_a, 'ac3'))
        self.assertEqual('14', self.wb.get_cell_contents(self.name_a, 'ac4'))

        # Check unmoved and from cells
        self.assertEqual('1', self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertEqual('2', self.wb.get_cell_contents(self.name_a, 'a2'))
        self.assertEqual('3', self.wb.get_cell_contents(self.name_a, 'a3'))
        self.assertEqual('4', self.wb.get_cell_contents(self.name_a, 'a4'))
        self.assertEqual('5', self.wb.get_cell_contents(self.name_a, 'a5'))
        self.assertEqual('6', self.wb.get_cell_contents(self.name_a, 'b1'))
        self.assertEqual('10', self.wb.get_cell_contents(self.name_a, 'b5'))
        self.assertEqual('11', self.wb.get_cell_contents(self.name_a, 'c1'))
        self.assertEqual('15', self.wb.get_cell_contents(self.name_a, 'c5'))
        self.assertEqual('16', self.wb.get_cell_contents(self.name_a, 'd1'))
        self.assertEqual('17', self.wb.get_cell_contents(self.name_a, 'd2'))
        self.assertEqual('18', self.wb.get_cell_contents(self.name_a, 'd3'))
        self.assertEqual('19', self.wb.get_cell_contents(self.name_a, 'd4'))
        self.assertEqual('20', self.wb.get_cell_contents(self.name_a, 'd5'))

        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'b2'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'b3'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'b4'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'c2'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'c3'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'c4'))

    def testMoveReferencedCell(self):
        """Test moving a cell that is referenced by another"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=a2+a3')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a4+a5')
        self.wb.set_cell_contents(self.name_a, 'a3', '1')
        self.wb.set_cell_contents(self.name_a, 'a4', '2')
        self.wb.set_cell_contents(self.name_a, 'a5', '6')

        self.assertEqual(decimal.Decimal('9'), self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('8'), self.wb.get_cell_value(self.name_a, 'a2'))
        self.wb.move_cells(self.name_a, 'a4', 'a4', 'aa1')

        # Check references
        self.assertEqual(decimal.Decimal('7'), self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'a2'))

        # Check others
        self.assertEqual(decimal.Decimal('1'), self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'a5'))

        # Check to
        self.assertEqual(decimal.Decimal('2'), self.wb.get_cell_value(self.name_a, 'aa1'))

    def testMoveIntoReferencedCell(self):
        """Test moving into cell that is referenced by another"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=a2+a3')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a4+a5')
        self.wb.set_cell_contents(self.name_a, 'a3', '1')
        self.wb.set_cell_contents(self.name_a, 'a4', '2')
        self.wb.set_cell_contents(self.name_a, 'a5', '6')
        self.wb.set_cell_contents(self.name_a, 'a6', '6')

        self.assertEqual(decimal.Decimal('9'), self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('8'), self.wb.get_cell_value(self.name_a, 'a2'))

        self.wb.move_cells(self.name_a, 'a6', 'a6', 'a4')

        # Check references
        self.assertEqual(decimal.Decimal('13'), self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('12'), self.wb.get_cell_value(self.name_a, 'a2'))

        # Check others
        self.assertEqual(decimal.Decimal('1'), self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a6'))

    def testMoveIntoReferencedCellCauseError(self):
        """Test moving into cell that is referenced by another to cause an error"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=a2+a3')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a4+a5')
        self.wb.set_cell_contents(self.name_a, 'a3', '1')
        self.wb.set_cell_contents(self.name_a, 'a4', '2')
        self.wb.set_cell_contents(self.name_a, 'a5', '6')
        self.wb.set_cell_contents(self.name_a, 'a6', 'abc')

        self.assertEqual(decimal.Decimal('9'), self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(decimal.Decimal('8'), self.wb.get_cell_value(self.name_a, 'a2'))

        self.wb.move_cells(self.name_a, 'a6', 'a6', 'a4')

        # Check references
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2').get_type() == sheets.CellErrorType.TYPE_ERROR)

        # Check others
        self.assertEqual(decimal.Decimal('1'), self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a6'))

    def testSheetNameDontExist(self):
        """
        Test that key errors are raised if from or to sheet do not exist
        """

        self.assertRaises(KeyError, self.wb.move_cells, 'SheetDoesntExist', 'a1', 'a1', 'c1')
        self.assertRaises(KeyError, self.wb.move_cells, self.name_a, 'a1', 'a1', 'c1', 'SheetDoesntExist')
        self.assertRaises(KeyError, self.wb.move_cells, 'SheetDoesntExist', 'a1', 'a1', 'c1', 'SheetDoesntExist')

    def testCellOutofBounds(self):
        """
        Test that value errors are raised if any cell is out of bounds
        """

        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'ZZZZZ1', 'a1', 'c1')
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'a38292', 'c1')
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'a38292', 'ZZZZZ9999')
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'a38292', '312cs1')

    def testNoMovementWhenTargetAreaOutOfBounds(self):
        """
        Tests that no cells are moved when a target area would have out of
        bounds cells when moved to the new location and the call raises
        a ValueError
        """

        self.wb.set_cell_contents(self.name_a, 'a1', 'no moving')
        self.wb.set_cell_contents(self.name_a, 'a2', 'no moving')
        self.wb.set_cell_contents(self.name_a, 'a3', 'no moving')
        self.wb.set_cell_contents(self.name_a, 'b1', 'no moving')
        self.wb.set_cell_contents(self.name_a, 'b2', 'no moving')
        self.wb.set_cell_contents(self.name_a, 'b3', 'no moving')

        # Assert that trying to move raises a value error
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'b3', 'ZZZZ9999')
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'b3', 'A9999')
        self.assertRaises(ValueError, self.wb.move_cells, self.name_a, 'a1', 'b3', 'ZZZZ99')

        # Assert that the cells are unchanged
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertEqual('no moving', self.wb.get_cell_value(self.name_a, 'b3'))

        # Spot check the target areas above to make sure no cells were changed
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'ZZZZ9999'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'A9999'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'ZZZZ99'))

    def testMoveWithOverlap(self):
        """
        Test target area overlapping with original area
        """

        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'a2', '2')
        self.wb.set_cell_contents(self.name_a, 'a3', '3')
        self.wb.set_cell_contents(self.name_a, 'b1', '4')
        self.wb.set_cell_contents(self.name_a, 'b2', '5')
        self.wb.set_cell_contents(self.name_a, 'b3', '6')

        self.wb.move_cells(self.name_a, 'a1', 'b3', 'a2')

        self.assertEqual(decimal.Decimal('1'), self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertEqual(decimal.Decimal('2'), self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual(decimal.Decimal('3'), self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual(decimal.Decimal('4'), self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertEqual(decimal.Decimal('5'), self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertEqual(decimal.Decimal('6'), self.wb.get_cell_value(self.name_a, 'b4'))

        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'b1'))

    def testMoveSoRefsOutOfBounds(self):
        """Test moving a cell so that the references in the formulas are out of bounds"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=a7+a2')
        self.wb.set_cell_contents(self.name_a, 'a2', '=z2')
        self.wb.set_cell_contents(self.name_b, 'd6', '=a1+z30')
        self.wb.set_cell_contents(self.name_b, 'd7', '=z30+a1')

        self.wb.move_cells(self.name_a, 'a1', 'a2', 'ZZZY9996')
        self.wb.move_cells(self.name_b, 'd6', 'd6', 'a6')
        self.wb.move_cells(self.name_b, 'd7', 'd7', 'd1')

        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_b, 'd6'))
        self.assertIsNone(self.wb.get_cell_contents(self.name_b, 'd7'))

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'ZZZY9996').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'ZZZY9996'), '=#REF!+zzzy9997')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'ZZZY9997').get_type(),
                         sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'ZZZY9997'), '=#REF!')

        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a6').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_b, 'a6'), '=#REF!+w30')
        self.assertEqual(self.wb.get_cell_contents(self.name_b, 'd1'), '=z24+#REF!')


if __name__ == '__main__':
    unittest.main()
