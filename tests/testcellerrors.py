"""
Caltech CS130 - Winter 2022

Testing various cell errors
"""

import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestCellErrors(unittest.TestCase):
    """Testing various cell errors"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testDivideByZeroAndTypeError(self):
        """Cell errors testing"""
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_a, 'a2', '#VALUE!')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1+a2')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a2'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3').get_type() == sheets.CellErrorType.TYPE_ERROR)

        self.wb.set_cell_contents(self.name_a, 'a2', '5')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3').get_type() == sheets.CellErrorType.DIVIDE_BY_ZERO)

    def testPropogateErrors(self):
        """Test that errors propogate"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=a1')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents(self.name_a, 'b1', '=5')

        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_a, 'c1'))

    def testBadReferenceFromSheet(self):
        """Cell errors testing"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet10!a1')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.wb.new_sheet("Sheet10")
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testBadReferenceDelSheet(self):
        """Test bad references when a sheet is deleted"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet2!a1')
        self.wb.set_cell_contents(self.name_b, 'a1', '=Sheet1!b1')
        self.assertTrue(decimal.Decimal('0') == self.wb.get_cell_value(self.name_a, 'a1'))
        self.wb.del_sheet(self.name_b)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

    def testOutOfBoundsError(self):
        """Test out of bounds errors"""
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'ZZZZZ99999', '42')
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'ZZZZZ-99999', '42')
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_c, 'A99A9', '42')

        self.wb.set_cell_contents(self.name_a, 'a1', '=helloworld1234123')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.wb.set_cell_contents(self.name_a, 'a1', '=aaaaa10000')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.wb.set_cell_contents(self.name_a, 'a1', '=a9998121251')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

    def testNumOpTypeErrors(self):
        """Test common cell errors with operations between string and numbers"""
        self.wb.set_cell_contents(self.name_a, 'a1', '="hello" + 1')
        self.wb.set_cell_contents(self.name_a, 'b1', '="hello" - 1')
        self.wb.set_cell_contents(self.name_a, 'c1', '="hello" * 1')
        self.wb.set_cell_contents(self.name_a, 'd1', '="hello" / 1')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'd1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'd1').get_type() == sheets.CellErrorType.TYPE_ERROR)

        self.wb.set_cell_contents(self.name_a, 'a1', '=1 + "hello"')
        self.wb.set_cell_contents(self.name_a, 'b1', '=1 - "hello"')
        self.wb.set_cell_contents(self.name_a, 'c1', '=1 * "hello"')
        self.wb.set_cell_contents(self.name_a, 'd1', '=1 / "hello"')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.TYPE_ERROR)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'd1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'd1').get_type() == sheets.CellErrorType.TYPE_ERROR)

    def testOverwriteError(self):
        """Test deleting a cell with an error"""
        self.wb.set_cell_contents(self.name_a, 'a1', '#DIV/0!')
        self.wb.set_cell_contents(self.name_a, 'a1', '')
        self.wb.set_cell_contents(self.name_a, 'a2', '#VALUE!')
        self.wb.set_cell_contents(self.name_a, 'a2', '')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1+a2')

        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertEqual(decimal.Decimal('0'), self.wb.get_cell_value(self.name_a, 'a3'))


if __name__ == '__main__':
    unittest.main()
