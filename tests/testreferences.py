"""
Caltech CS130 - Winter 2022

Testing for references. Basic testing to ensure that references are
working as expected and auto-updating. Also includes case
insensitivity tests
"""


import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestReferences(unittest.TestCase):
    """
    Testing for references. Basic testing to ensure that references are
    working as expected and auto-updating. Also includes case
    insensitivity tests
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet("sheet1")  # sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet("Sheet2")  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet("sheet 3")  # sheet 3
        (self.index_d, self.name_d) = self.wb.new_sheet("s!!4")  # s!!4

        self.wb.set_cell_contents(self.name_a, 'a1', '3')
        self.wb.set_cell_contents(self.name_a, 'a2', '5')
        self.wb.set_cell_contents(self.name_a, 'a3', '1')
        self.wb.set_cell_contents(self.name_a, 'a4', '0')

        self.wb.set_cell_contents(self.name_c, 'a1', '3')
        self.wb.set_cell_contents(self.name_c, 'a2', '5')

        self.wb.set_cell_contents(self.name_d, 'a1', '3')
        self.wb.set_cell_contents(self.name_d, 'a2', '5')

    def tearDown(self):
        pass

    def testReferencesWithinSheet(self):
        """Test references and formula parsing within a sheet"""
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1')
        self.wb.set_cell_contents(self.name_a, 'b2', '=   a1')
        self.wb.set_cell_contents(self.name_a, 'b3', '    =a1')
        self.wb.set_cell_contents(self.name_a, 'b4', '=a1    ')
        self.wb.set_cell_contents(self.name_a, 'b5', '= a1')
        self.wb.set_cell_contents(self.name_a, 'b6', ' = a1')
        self.wb.set_cell_contents(self.name_a, 'b7', '= a1  ')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b6'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b7'))

    def testUpdatedValue(self):
        """Test deleting and resurrecting a sheet leading to new value"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'a1', '4')

        self.assertTrue(decimal.Decimal('4') == self.wb.get_cell_value(self.name_b, 'b1'))

    def testCircularReferencesWithinSheet(self):
        """Test creating a circular reference within a sheet"""
        self.wb.set_cell_contents(self.name_a, 'b1', '=a1')
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

        self.wb.set_cell_contents(self.name_a, 'c1', '=c1')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

    def testReferencesOutsideSheet(self):
        """Test references and formula parsing outside of a single sheet"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.set_cell_contents(self.name_b, 'b2', '=   Sheet1!a1')
        self.wb.set_cell_contents(self.name_b, 'b3', '    =SHEET1!a1')
        self.wb.set_cell_contents(self.name_b, 'b4', '=Sheet1!a1    ')
        self.wb.set_cell_contents(self.name_b, 'b5', '= Sheet1!a1')
        self.wb.set_cell_contents(self.name_b, 'b6', ' = Sheet1!a1')
        self.wb.set_cell_contents(self.name_b, 'b7', '= Sheet1!a1  ')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b2'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b3'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b4'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b5'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b6'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b7'))

    def testCircularReferencesOutsideSheet(self):
        """Test creating a circular reference within outside a single sheet"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet2!b1')

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

    def testDeleteReferencedSheet(self):
        """Test deleting a sheet leading to bad references"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.del_sheet(self.name_a)

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

    def testDeleteAndResurrectReferencedSheet(self):
        """Test deleting and resurrecting a sheet leading to new value"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.del_sheet(self.name_a)
        self.wb.new_sheet(self.name_a)
        self.wb.set_cell_contents(self.name_a, 'a1', '2')

        self.assertTrue(decimal.Decimal('2') == self.wb.get_cell_value(self.name_b, 'b1'))

    def testLargerCircularReference(self):
        """Test creating a large circular reference"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=Sheet2!b1')
        self.wb.set_cell_contents(self.name_b, 'd1', '=sheet1!c1')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'd1'))

        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet2!d1')
        print("OPEBNUOFBEUIOAWBFUOAWEBF", self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'c1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'd1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'd1').get_type() == sheets.CellErrorType.CIRCULAR_REFERENCE)

    def testBadReferences(self):
        """Check bad references"""
        self.wb.set_cell_contents(self.name_b, 'b1', '=sheet100!a1')  # Sheet does not exist
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_b, 'ZZZZZ9999', '=sheet1!a1')  # Out of range
        self.assertRaises(ValueError, self.wb.set_cell_contents, self.name_b, 'ZZZZ99999', '=sheet1!a1')  # Out of range

        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'b1'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'b1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

    def testQuotedReferences(self):
        """Test that references that need to be quoted (or ones that don't that are quoted)"""
        # still work
        self.wb.set_cell_contents(self.name_b, 'a1', '=\'sheet1\'!a1')
        self.wb.set_cell_contents(self.name_b, 'a2', '=\'sheet1\'!a2')

        self.wb.set_cell_contents(self.name_b, 'b1', '=\'Sheet 3\'!a1')
        self.wb.set_cell_contents(self.name_b, 'b2', '=\'sheet 3\'!a2')

        self.wb.set_cell_contents(self.name_b, 'c1', '=\'s!!4\'!a1')
        self.wb.set_cell_contents(self.name_b, 'c2', '=\'s!!4\'!a2')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'a1'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_b, 'a2'))

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'b1'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_b, 'b2'))

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_b, 'c1'))
        self.assertTrue(decimal.Decimal('5') == self.wb.get_cell_value(self.name_b, 'c2'))


if __name__ == '__main__':
    unittest.main()
