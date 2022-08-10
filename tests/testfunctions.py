# pylint: skip-file
# TODO fix pylint in this file

"""
Caltech CS130 - Winter 2022

Test basic number functionality
"""

import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestFunctions(unittest.TestCase):
    """Test basic boolean functionality"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testAnd(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=AND(a1, a2, a3, a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=AND(a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=AND(a1, a4)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=AND()')
        self.wb.set_cell_contents(self.name_a, 'b5', '=AND(a1)')
        self.wb.set_cell_contents(self.name_a, 'b6', '=AND(1, "TRUE")')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b6'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b6'))

    def testAndMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_c, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=AND(Sheet1!a1, Sheet2!a2, Sheet3!a3, Sheet1!a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=AND(Sheet2!a2, Sheet3!a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=AND(Sheet1!a1, Sheet1!a4)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))

    def testOr(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=OR(a1, a2, a3, a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=OR(a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=OR(a1, a4)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=OR()')
        self.wb.set_cell_contents(self.name_a, 'b5', '=OR(a1)')
        self.wb.set_cell_contents(self.name_a, 'b6', '=OR(0, "TRUE")')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b6'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b6'))

    def testOrMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_c, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=OR(Sheet1!a1, Sheet2!a2, Sheet3!a3, Sheet1!a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=OR(Sheet2!a2, Sheet3!a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=OR(Sheet1!a1, Sheet1!a4)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))

    def testNot(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 > 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=NOT(a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=NOT(a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=NOT()')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3').get_type() == sheets.CellErrorType.TYPE_ERROR)

    def testNotMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 > 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=NOT(Sheet1!a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=NOT(Sheet2!a2)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2'))

    def testXor(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=XOR(a1, a2, a3, a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=XOR(a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=XOR(a1, a4)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=XOR(a1, a2)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=XOR()')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5').get_type() == sheets.CellErrorType.TYPE_ERROR)

    def testXorMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_c, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '=4 <> 5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=XOR(Sheet1!a1, Sheet2!a2, Sheet3!a3, Sheet1!a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=XOR(Sheet2!a2, Sheet3!a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=XOR(Sheet1!a1, Sheet1!a4)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=XOR(Sheet1!a1, Sheet2!a2)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4'))

    def testExact(self):
        self.wb.set_cell_contents(self.name_a, 'a1', 'testing')
        self.wb.set_cell_contents(self.name_a, 'a2', 'test')

        self.wb.set_cell_contents(self.name_a, 'b1', '=EXACT(a1, "testing")')
        self.wb.set_cell_contents(self.name_a, 'b2', '=EXACT(a1, a2)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))

    def testExactMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', 'testing')
        self.wb.set_cell_contents(self.name_b, 'a2', 'test')

        self.wb.set_cell_contents(self.name_a, 'b1', '=EXACT(Sheet1!a1, "testing")')
        self.wb.set_cell_contents(self.name_a, 'b2', '=EXACT(Sheet1!a1, Sheet2!a2)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))

    def testIf(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_a, 'a4', '5')
        self.wb.set_cell_contents(self.name_a, 'a5', '7')

        self.wb.set_cell_contents(self.name_a, 'b1', '=IF(a1, a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=IF(a2, 6, a5)')

        self.assertTrue(decimal.Decimal(5) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(7) == self.wb.get_cell_value(self.name_a, 'b2'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=4 < 5')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b2'))

    def testIfMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=4 < 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 > 5')
        self.wb.set_cell_contents(self.name_a, 'a3', '=4 = 5')
        self.wb.set_cell_contents(self.name_b, 'a4', '5')
        self.wb.set_cell_contents(self.name_a, 'a5', '7')

        self.wb.set_cell_contents(self.name_a, 'b1', '=IF(Sheet1!a1, Sheet2!a4)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=IF(Sheet2!a2, 6, Sheet1!a5)')

        self.assertTrue(decimal.Decimal(5) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(7) == self.wb.get_cell_value(self.name_a, 'b2'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=4 > 5')
        self.wb.set_cell_contents(self.name_b, 'a2', '=4 < 5')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b2'))

    def testIfError(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_a, 'a2', '#VALUE!')

        self.wb.set_cell_contents(self.name_a, 'b1', '=IFERROR(a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=IFERROR(a2, 6)')

        self.assertTrue("" == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b2'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=4')
        self.wb.set_cell_contents(self.name_a, 'a2', '=5')

        self.assertTrue(decimal.Decimal(4) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(5) == self.wb.get_cell_value(self.name_a, 'b2'))

    def testIfErrorMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_b, 'a2', '#VALUE!')

        self.wb.set_cell_contents(self.name_a, 'b1', '=IFERROR(Sheet1!a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=IFERROR(Sheet2!a2, 6)')

        self.assertTrue("" == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b2'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=4')
        self.wb.set_cell_contents(self.name_b, 'a2', '=5')

        self.assertTrue(decimal.Decimal(4) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal(5) == self.wb.get_cell_value(self.name_a, 'b2'))

    def testChoose(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '5')
        self.wb.set_cell_contents(self.name_a, 'a2', '6')
        self.wb.set_cell_contents(self.name_a, 'a3', '2')

        self.wb.set_cell_contents(self.name_a, 'b1', '=CHOOSE(a3, a1, a2)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=CHOOSE(0, a1, a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=CHOOSE(4, a1, a2)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=CHOOSE(a3, 12, 24)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=CHOOSE(a3, 1, a1:a3)')

        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(decimal.Decimal(24) == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b5'), [[decimal.Decimal(5)], [decimal.Decimal(6)], [decimal.Decimal(2)]])

    def testChooseMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '5')
        self.wb.set_cell_contents(self.name_b, 'a2', '6')
        self.wb.set_cell_contents(self.name_c, 'a3', '2')

        self.wb.set_cell_contents(self.name_a, 'b1', '=CHOOSE(Sheet3!a3, Sheet1!a1, Sheet2!a2)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=CHOOSE(0, Sheet1!a1, Sheet2!a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=CHOOSE(4, Sheet1!a1, Sheet2!a2)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=CHOOSE(Sheet3!a3, 12, 24)')

        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(decimal.Decimal(24) == self.wb.get_cell_value(self.name_a, 'b4'))

    # TODO Make sure that when a cell is deleted it's value is set to None
    def testIsBlank(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '')
        self.wb.set_cell_contents(self.name_a, 'a2', 'False')
        self.wb.set_cell_contents(self.name_a, 'a3', '0')
        self.wb.set_cell_contents(self.name_a, 'a4', '""')

        self.wb.set_cell_contents(self.name_a, 'b1', '=ISBLANK(a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=ISBLANK(a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=ISBLANK(a3)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=ISBLANK(a4)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b4'))

    def testIsBlankMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '')
        self.wb.set_cell_contents(self.name_a, 'a2', 'False')
        self.wb.set_cell_contents(self.name_a, 'a3', '0')
        self.wb.set_cell_contents(self.name_a, 'a4', '""')

        self.wb.set_cell_contents(self.name_a, 'b1', '=ISBLANK(a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=ISBLANK(a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=ISBLANK(a3)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=ISBLANK(a4)')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b3'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b4'))

    def testIsError(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_a, 'a2', '#VALUE!')
        self.wb.set_cell_contents(self.name_a, 'a3', '=b1+')
        self.wb.set_cell_contents(self.name_a, 'a4', '=ISERROR(b4)')
        self.wb.set_cell_contents(self.name_a, 'a5', '=4')
        self.wb.set_cell_contents(self.name_a, 'a6', '=5')

        self.wb.set_cell_contents(self.name_a, 'b1', '=ISERROR(a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=ISERROR(a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=ISERROR(a3)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=ISERROR(a4)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=ISERROR(a5)')
        self.wb.set_cell_contents(self.name_a, 'b6', '=ISERROR(a6)')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertFalse(self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertFalse(self.wb.get_cell_value(self.name_a, 'b6'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=4')
        self.wb.set_cell_contents(self.name_a, 'a2', '=5')

        self.assertFalse(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertFalse(self.wb.get_cell_value(self.name_a, 'b2'))

    def testIsErrorMultiSheet(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_b, 'a2', '#VALUE!')
        self.wb.set_cell_contents(self.name_c, 'a3', '=b1+')
        self.wb.set_cell_contents(self.name_a, 'a4', '=ISERROR(Sheet1!b4)')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a4'))

        self.wb.set_cell_contents(self.name_a, 'b1', '=ISERROR(Sheet1!a1)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=ISERROR(Sheet2!a2)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=ISERROR(Sheet3!a3)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=ISERROR(Sheet1!a4)')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents(self.name_a, 'a1', '=4')
        self.wb.set_cell_contents(self.name_b, 'a2', '=5')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'b2'))

    def testVersion(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '=VERSION()')
        self.wb.set_cell_contents(self.name_a, 'a2', '=EXACT(a2, "1.4.0")')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2'))

    def testIndirect(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '5')
        self.wb.set_cell_contents(self.name_a, 'a2', '=INDIRECT("Sheet1!a3")')
        self.wb.set_cell_contents(self.name_a, 'a3', '=INDIRECT("a2")')
        self.wb.set_cell_contents(self.name_a, 'a4', '=INDIRECT("!2")')
        self.wb.set_cell_contents(self.name_a, 'a5', '=INDIRECT("a1")')

        self.assertTrue(decimal.Decimal(5) == self.wb.get_cell_value(self.name_a, 'a5'))

        self.wb.set_cell_contents(self.name_a, 'a1', '6')
        self.assertTrue(decimal.Decimal(6) == self.wb.get_cell_value(self.name_a, 'a5'))

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a4').get_type() == sheets.CellErrorType.BAD_REFERENCE)

    # # TODO Test in functions testing as well as number to bool
    # def testImplicitConversionsStringToBool(self):
    #     """Test implicit conversions from booleans to string"""
    #     self.wb.set_cell_contents(self.name_a, 'a1', '="true"')
    #     self.wb.set_cell_contents(self.name_a, 'a3', '="false"')
    #     self.wb.set_cell_contents(self.name_a, 'a1', 'nope')
    #     self.wb.set_cell_contents(self.name_a, 'a3', '=a1 & "test"')
    #     self.wb.set_cell_contents(self.name_a, 'a4', '=a2 & "test"')

    #     self.assertTrue("TRUEtest" == self.wb.get_cell_value(self.name_a, 'a3'))
    #     self.assertTrue("FALSEtest" == self.wb.get_cell_value(self.name_a, 'a4'))

    #     self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
    #     self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)

    def testMin(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'a2', '13')
        self.wb.set_cell_contents(self.name_a, 'a3', '69')
        self.wb.set_cell_contents(self.name_a, 'a4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'a5', '"70"')
        self.wb.set_cell_contents(self.name_a, 'a6', '"Hello!"')

        self.wb.set_cell_contents(self.name_a, 'b1', '=MIN(a1, a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=MIN(a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=MIN(a1:a5)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=MIN(a6)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=MIN(a1:a6)')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal(1))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal(69))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal(1))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), sheets.CellError))

    def testMax(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'a2', '13')
        self.wb.set_cell_contents(self.name_a, 'a3', '69')
        self.wb.set_cell_contents(self.name_a, 'a4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'a5', '"70"')
        self.wb.set_cell_contents(self.name_a, 'a6', '"Hello!"')

        self.wb.set_cell_contents(self.name_a, 'b1', '=MAX(a1, a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=MAX(a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=MAX(a1:a5)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=MAX(a6)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=MAX(a1:a6)')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal(69))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal(69))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal(70))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), sheets.CellError))

    def testSum(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'a2', '13')
        self.wb.set_cell_contents(self.name_a, 'a3', '69')
        self.wb.set_cell_contents(self.name_a, 'a4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'a5', '"70"')
        self.wb.set_cell_contents(self.name_a, 'a6', '"Hello!"')

        self.wb.set_cell_contents(self.name_a, 'b1', '=SUM(a1, a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=SUM(a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=SUM(a1:a5)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=SUM(a6)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=SUM(a1:a6)')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal(83))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal(69))
        self.assertAlmostEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal(157.2))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), sheets.CellError))

    def testAverage(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '2')
        self.wb.set_cell_contents(self.name_a, 'a2', '13')
        self.wb.set_cell_contents(self.name_a, 'a3', '69')
        self.wb.set_cell_contents(self.name_a, 'a4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'a5', '"70"')
        self.wb.set_cell_contents(self.name_a, 'a6', '"Hello!"')

        self.wb.set_cell_contents(self.name_a, 'b1', '=AVERAGE(a1, a2, a3)')
        self.wb.set_cell_contents(self.name_a, 'b2', '=AVERAGE(a3)')
        self.wb.set_cell_contents(self.name_a, 'b3', '=AVERAGE(a1:a5)')
        self.wb.set_cell_contents(self.name_a, 'b4', '=AVERAGE(a6)')
        self.wb.set_cell_contents(self.name_a, 'b5', '=AVERAGE(a1:a6)')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal(28))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal(69))
        self.assertAlmostEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal((2 + 13 + 69 + 4.20 + 70) / 5))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b4').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b4'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b5').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b5'), sheets.CellError))

    def testHLookup(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '2')
        self.wb.set_cell_contents(self.name_a, 'b2', '13')
        self.wb.set_cell_contents(self.name_a, 'c3', '69')
        self.wb.set_cell_contents(self.name_a, 'd4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'e1', 'Hello')
        self.wb.set_cell_contents(self.name_a, 'e5', 'HelloWorld')

        self.wb.set_cell_contents(self.name_a, 'a2', '102')
        self.wb.set_cell_contents(self.name_a, 'a3', '103')
        self.wb.set_cell_contents(self.name_a, 'a4', '104')
        self.wb.set_cell_contents(self.name_a, 'a5', '105')

        self.wb.set_cell_contents(self.name_b, 'a1', '=HLOOKUP("Hello", Sheet1!a1:e5, 5)')
        self.wb.set_cell_contents(self.name_b, 'b1', '=HLOOKUP("Hello", Sheet1!a6:e15, 5)')
        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a1'), 'HelloWorld')
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))

    def testVLookup(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '2')
        self.wb.set_cell_contents(self.name_a, 'b2', '13')
        self.wb.set_cell_contents(self.name_a, 'c3', '69')
        self.wb.set_cell_contents(self.name_a, 'd4', '4.20')
        self.wb.set_cell_contents(self.name_a, 'e5', '70')

        self.wb.set_cell_contents(self.name_a, 'a2', '102')
        self.wb.set_cell_contents(self.name_a, 'a3', '103')
        self.wb.set_cell_contents(self.name_a, 'a4', '104')
        self.wb.set_cell_contents(self.name_a, 'a5', '105')
        self.wb.set_cell_contents(self.name_a, 'a5', '69420')

        self.wb.set_cell_contents(self.name_b, 'a1', '=VLOOKUP(69420, Sheet1!a1:e5, 5)')
        self.wb.set_cell_contents(self.name_b, 'b1', '=VLOOKUP(69420, Sheet1!a6:e15, 5)')
        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a1'), decimal.Decimal(70))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))

    def testFunctionRangeUpdate(self):
        self.wb.set_cell_contents(self.name_a, 'a1', '2')
        self.wb.set_cell_contents(self.name_a, 'b2', '3')
        self.wb.set_cell_contents(self.name_a, 'c3', '4')
        self.wb.set_cell_contents(self.name_a, 'd4', '5')
        self.wb.set_cell_contents(self.name_a, 'e5', '6')

        self.wb.set_cell_contents(self.name_a, 'h1', '=SUM(A1:E5)')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'h1'), decimal.Decimal(20))
        self.wb.set_cell_contents(self.name_a, 'c3', '20')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'h1'), decimal.Decimal(36))
        self.wb.set_cell_contents(self.name_a, 'a1', 'HELLO')
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'h1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'h1'), sheets.CellError))


if __name__ == '__main__':
    unittest.main()
