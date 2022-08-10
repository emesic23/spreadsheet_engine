"""
Caltech CS130 - Winter 2022

Test basic number functionality
"""

import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestBasicNumbers(unittest.TestCase):
    """Test basic boolean functionality"""

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

    def testValues(self):
        """Test basic boolean values"""
        self.wb.set_cell_contents(self.name_a, 'a1', '   trUe')
        self.wb.set_cell_contents(self.name_a, 'a2', '=TRue')
        self.wb.set_cell_contents(self.name_b, 'a3', 'faLSE   ')
        self.wb.set_cell_contents(self.name_c, 'a4', '=FalSe')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a1'), bool))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a2'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_b, 'a3'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_b, 'a3'), bool))
        self.assertTrue(not self.wb.get_cell_value(self.name_c, 'a4'))
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_c, 'a4'), bool))

    def testContents(self):
        """Test basic boolean values"""
        self.wb.set_cell_contents(self.name_a, 'a1', '   trUe')
        self.wb.set_cell_contents(self.name_a, 'a2', '=TRue')
        self.wb.set_cell_contents(self.name_b, 'a3', 'faLSE   ')
        self.wb.set_cell_contents(self.name_c, 'a4', '=FalSe')

        self.assertEqual('trUe', self.wb.get_cell_contents(self.name_a, 'a1'))
        self.assertEqual('=TRue', self.wb.get_cell_contents(self.name_a, 'a2'))
        self.assertEqual('faLSE', self.wb.get_cell_contents(self.name_b, 'a3'))
        self.assertEqual('=FalSe', self.wb.get_cell_contents(self.name_c, 'a4'))

    def testUpdatingValues(self):
        """Test updating of numbers"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', 'false')

        self.wb.set_cell_contents(self.name_b, 'a1', 'true')
        self.wb.set_cell_contents(self.name_b, 'a2', 'false')

        self.wb.set_cell_contents(self.name_c, 'a1', 'true')
        self.wb.set_cell_contents(self.name_c, 'a2', 'false')

        # Updates
        self.wb.set_cell_contents(self.name_a, 'a1', 'false')
        self.wb.set_cell_contents(self.name_b, 'a2', 'true')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a2'))  # Updated

        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a1'))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a2'))  # Updated

        self.assertTrue(self.wb.get_cell_value(self.name_c, 'a1'))
        self.assertTrue(not self.wb.get_cell_value(self.name_c, 'a2'))

    def testReferenceWithinSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')

        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a2'))

    def testUpdatingReferenceWithinSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')
        self.wb.set_cell_contents(self.name_a, 'a1', 'false')

        self.assertTrue(not self.wb.get_cell_value(self.name_a, 'a2'))

    def testReferenceOutsideSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_b, 'a2', '=Sheet1!a1')

        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a2'))

    def testUpdatingReferenceOutsideSheet(self):
        """Test reference with a sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_b, 'a2', '=Sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'a1', 'false')

        self.assertTrue(not self.wb.get_cell_value(self.name_b, 'a2'))

    def testImplicitConversionsBoolToNumber(self):
        """Test implicit conversions from booleans to numbers"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', 'false')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 + 1')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a2 + 1')

        self.assertTrue(decimal.Decimal(2) == self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue(decimal.Decimal(1) == self.wb.get_cell_value(self.name_a, 'a4'))

    def testImplicitConversionsBoolToString(self):
        """Test implicit conversions from booleans to string"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'true')
        self.wb.set_cell_contents(self.name_a, 'a2', 'false')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1 & "test"')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a2 & "test"')

        self.assertTrue("TRUEtest" == self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue("FALSEtest" == self.wb.get_cell_value(self.name_a, 'a4'))

    # # TODO Test in functions testing as well as number to bool  # pylint: disable=W0511
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


if __name__ == '__main__':
    unittest.main()
