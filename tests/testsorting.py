"""
Caltech CS130 - Winter 2022

Test sorting functionality
"""

import unittest
import decimal

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestSorting(unittest.TestCase):
    """Test sorting functionality"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet("Sheet 3")  # Sheet3

    def tearDown(self):
        pass

    def testErrorRaisingSheetName(self):
        """Test that function raises key error for invalid sheet name"""
        self.assertRaises(KeyError, self.wb.sort_region, "I dont exist", 'a1', 'A3', [1])

    def testErrorInvalidSourceArea(self):
        """Test that function raises value error for invalid source region"""
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'aaaaaa1', 'A3', [1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'a1000000', 'A3', [1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'a1', 'aaaaaA3', [1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'a1', 'A3000000', [1])

    def testErrorSortCols(self):
        """Test that invalid sort columns result in ValueErrors"""
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [0])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [-2])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [1, 1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [1, -1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [1, 2])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'a3', [2])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'd3', [2, 1, 3, 4, 10])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'd3', [1, 2, 1])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'd3', [-2, 2])
        self.assertRaises(ValueError, self.wb.sort_region, self.name_a, 'A1', 'd3', [-2, -1, -3, -4, -10])

    def testSingleColSortAscending(self):
        """Test ascending sort on three values using single column"""
        # Already sorted
        self.wb.set_cell_contents(self.name_a, 'a1', 'apple')
        self.wb.set_cell_contents(self.name_a, 'a2', 'banana')
        self.wb.set_cell_contents(self.name_a, 'a3', 'orange')
        self.wb.sort_region(self.name_a, 'a1', 'a3', [1])

        # Mixed
        self.wb.set_cell_contents(self.name_a, 'b2', 'apple')
        self.wb.set_cell_contents(self.name_a, 'b3', 'orange')
        self.wb.set_cell_contents(self.name_a, 'b4', 'banana')
        self.wb.sort_region(self.name_a, 'B2', 'b4', [1])

        # # Backwards
        self.wb.set_cell_contents(self.name_a, 'c10', 'orange')
        self.wb.set_cell_contents(self.name_a, 'c11', 'banana')
        self.wb.set_cell_contents(self.name_a, 'c12', 'apple')
        self.wb.sort_region(self.name_a, 'c10', 'c12', [1])

        # Already sorted column
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), 'orange')

        # Mixed
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), 'orange')

        # Backwards
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c10'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c11'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c12'), 'orange')

    def testSingleColSortDescending(self):
        """Test descending sort on three values using single column"""
        # Already sorted
        self.wb.set_cell_contents(self.name_a, 'a1', 'apple')
        self.wb.set_cell_contents(self.name_a, 'a2', 'banana')
        self.wb.set_cell_contents(self.name_a, 'a3', 'orange')
        self.wb.sort_region(self.name_a, 'a1', 'a3', [-1])

        # Mixed
        self.wb.set_cell_contents(self.name_a, 'b2', 'apple')
        self.wb.set_cell_contents(self.name_a, 'b3', 'orange')
        self.wb.set_cell_contents(self.name_a, 'b4', 'banana')
        self.wb.sort_region(self.name_a, 'B2', 'b4', [-1])

        # # Backwards
        self.wb.set_cell_contents(self.name_a, 'c10', 'orange')
        self.wb.set_cell_contents(self.name_a, 'c11', 'banana')
        self.wb.set_cell_contents(self.name_a, 'c12', 'apple')
        self.wb.sort_region(self.name_a, 'c10', 'c12', [-1])

        # Already sorted column
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), 'orange')

        # Mixed
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), 'orange')

        # Backwards
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c12'), 'apple')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c11'), 'banana')
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c10'), 'orange')

    def testMultiColSortAscending(self):
        """Test ascending sort on values using two columns"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '3')
        self.wb.set_cell_contents(self.name_a, 'a2', '-50')
        self.wb.set_cell_contents(self.name_a, 'b2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-50')
        self.wb.set_cell_contents(self.name_a, 'b3', '-2')
        self.wb.set_cell_contents(self.name_a, 'a4', '30')
        self.wb.set_cell_contents(self.name_a, 'b4', '23')

        self.wb.sort_region(self.name_a, 'a1', 'b4', [1, 2])

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), decimal.Decimal('30'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal('-2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal('3'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), decimal.Decimal('23'))

    def testMultiColSortDescending(self):
        """Test ascending sort on values using two columns with second col descending"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '3')
        self.wb.set_cell_contents(self.name_a, 'a2', '-50')
        self.wb.set_cell_contents(self.name_a, 'b2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-50')
        self.wb.set_cell_contents(self.name_a, 'b3', '-2')
        self.wb.set_cell_contents(self.name_a, 'a4', '30')
        self.wb.set_cell_contents(self.name_a, 'b4', '23')

        self.wb.sort_region(self.name_a, 'a1', 'b4', [1, -2])

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), decimal.Decimal('30'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal('-2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal('3'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), decimal.Decimal('23'))

    def testMultiColSortDescending2(self):
        """Test descendng sort on values using two columns with both col descending"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '3')
        self.wb.set_cell_contents(self.name_a, 'a2', '-50')
        self.wb.set_cell_contents(self.name_a, 'b2', '4')
        self.wb.set_cell_contents(self.name_a, 'a3', '-50')
        self.wb.set_cell_contents(self.name_a, 'b3', '-2')
        self.wb.set_cell_contents(self.name_a, 'a4', '30')
        self.wb.set_cell_contents(self.name_a, 'b4', '23')
        self.wb.set_cell_contents(self.name_a, 'a5', '8')
        self.wb.set_cell_contents(self.name_a, 'a6', '12')

        self.wb.sort_region(self.name_a, 'a1', 'b6', [-2, -1])

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal('23'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal('3'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), decimal.Decimal('-2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b5'), None)
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b6'), None)

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), decimal.Decimal('30'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a5'), decimal.Decimal('12'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a6'), decimal.Decimal('8'))

    def testSingleColSortDescendingEntriesInOrder(self):
        """
        Test descendng sort on values using one columns to see if cells that can't
        be differentiated appear in the right order
        """
        self.wb.set_cell_contents(self.name_c, 'a1', '1')
        self.wb.set_cell_contents(self.name_c, 'b1', '3')
        self.wb.set_cell_contents(self.name_c, 'a2', '-50')
        self.wb.set_cell_contents(self.name_c, 'b2', '4')
        self.wb.set_cell_contents(self.name_c, 'a3', '-50')
        self.wb.set_cell_contents(self.name_c, 'b3', '-2')
        self.wb.set_cell_contents(self.name_c, 'a4', '30')
        self.wb.set_cell_contents(self.name_c, 'b4', '23')
        self.wb.set_cell_contents(self.name_c, 'a5', '8')
        self.wb.set_cell_contents(self.name_c, 'a6', '12')

        self.wb.sort_region(self.name_c, 'a1', 'b6', [-2])

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b1'), decimal.Decimal('23'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b2'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b3'), decimal.Decimal('3'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b4'), decimal.Decimal('-2'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b5'), None)
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b6'), None)

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a1'), decimal.Decimal('30'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a2'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a3'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a4'), decimal.Decimal('-50'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a5'), decimal.Decimal('8'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a6'), decimal.Decimal('12'))

    def testFunctionSorting(self):
        """Test ascending sort on three formulas using single column"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '2')
        self.wb.set_cell_contents(self.name_a, 'c1', '=a1+b1+a1')
        self.wb.set_cell_contents(self.name_a, 'a2', '2')
        self.wb.set_cell_contents(self.name_a, 'b2', '2')
        self.wb.set_cell_contents(self.name_a, 'c2', '=a2+b2')
        self.wb.set_cell_contents(self.name_a, 'a3', '6')
        self.wb.set_cell_contents(self.name_a, 'b3', '2')
        self.wb.set_cell_contents(self.name_a, 'c3', '=a3 - b3')
        self.wb.set_cell_contents(self.name_a, 'a3', '6')
        self.wb.set_cell_contents(self.name_a, 'b3', '2')
        self.wb.set_cell_contents(self.name_a, 'c3', '=a3 - b3 + 10')
        self.wb.set_cell_contents(self.name_a, 'a4', '320')
        self.wb.set_cell_contents(self.name_a, 'b4', '2')
        self.wb.set_cell_contents(self.name_a, 'c4', '=a1 - b4')

        self.wb.sort_region(self.name_a, 'a2', 'c4', [-1])

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c1'), decimal.Decimal('4'))

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), decimal.Decimal('320'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal('2'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c2').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c2'), '=#REF!-b2')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), decimal.Decimal('6'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c3'), decimal.Decimal('14'))
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c3'), '=a3-b3+10')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c4'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c4'), '=a4+b4')

    def testFunctionSortingOnFunctionValues(self):
        """Test ascending sort on three formulas using single column"""
        self.wb.set_cell_contents(self.name_a, 'a1', '1')
        self.wb.set_cell_contents(self.name_a, 'b1', '2')
        self.wb.set_cell_contents(self.name_a, 'c1', '=a1+b1+a1')
        self.wb.set_cell_contents(self.name_a, 'a2', '2')
        self.wb.set_cell_contents(self.name_a, 'b2', '2')
        self.wb.set_cell_contents(self.name_a, 'c2', '=a2+b2')
        self.wb.set_cell_contents(self.name_a, 'a3', '6')
        self.wb.set_cell_contents(self.name_a, 'b3', '2')
        self.wb.set_cell_contents(self.name_a, 'c3', '=a3 - b3')
        self.wb.set_cell_contents(self.name_a, 'a3', '6')
        self.wb.set_cell_contents(self.name_a, 'b3', '2')
        self.wb.set_cell_contents(self.name_a, 'c3', '=a3 - b3 + 10')
        self.wb.set_cell_contents(self.name_a, 'a4', '320')
        self.wb.set_cell_contents(self.name_a, 'b4', '222')
        self.wb.set_cell_contents(self.name_a, 'c4', '=a1 + b4')

        self.wb.sort_region(self.name_a, 'a2', 'c4', [-3])

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a1'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b1'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c1'), decimal.Decimal('4'))

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a2'), decimal.Decimal('320'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b2'), decimal.Decimal('222'))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'c2').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c2'), '=#REF!+b2')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a3'), decimal.Decimal('6'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b3'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c3'), decimal.Decimal('14'))
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c3'), '=a3-b3+10')

        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'b4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'c4'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_contents(self.name_a, 'c4'), '=a4+b4')

    def testFunctionSortingOnFunctionValuesWithSheetName(self):
        """Test ascending sort on three formulas using single column"""
        self.wb.set_cell_contents(self.name_c, 'a1', '1')
        self.wb.set_cell_contents(self.name_c, 'b1', '2')
        self.wb.set_cell_contents(self.name_c, 'c1', "='Sheet 3'!a1+'Sheet 3'!b1+'Sheet 3'!a1")
        self.wb.set_cell_contents(self.name_c, 'a2', '2')
        self.wb.set_cell_contents(self.name_c, 'b2', '2')
        self.wb.set_cell_contents(self.name_c, 'c2', "='Sheet 3'!a2+'Sheet 3'!b2")
        self.wb.set_cell_contents(self.name_c, 'a3', '6')
        self.wb.set_cell_contents(self.name_c, 'b3', '2')
        self.wb.set_cell_contents(self.name_c, 'c3', "='Sheet 3'!a3 - 'Sheet 3'!b3 + 10")
        self.wb.set_cell_contents(self.name_c, 'a4', '320')
        self.wb.set_cell_contents(self.name_c, 'b4', '222')
        self.wb.set_cell_contents(self.name_c, 'c4', "='Sheet 3'!a1 + 'Sheet 3'!b4")

        self.wb.sort_region(self.name_c, 'a2', 'c4', [-3])

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a1'), decimal.Decimal('1'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b1'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'c1'), decimal.Decimal('4'))

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a2'), decimal.Decimal('320'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b2'), decimal.Decimal('222'))
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'c2').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_c, 'c2'), "='Sheet 3'!#REF!+'Sheet 3'!b2")

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a3'), decimal.Decimal('6'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b3'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'c3'), decimal.Decimal('14'))
        self.assertEqual(self.wb.get_cell_contents(self.name_c, 'c3'), "='Sheet 3'!a3-'Sheet 3'!b3+10")

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'a4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'b4'), decimal.Decimal('2'))
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'c4'), decimal.Decimal('4'))
        self.assertEqual(self.wb.get_cell_contents(self.name_c, 'c4'), "='Sheet 3'!a4+'Sheet 3'!b4")

    def testSortingErrors(self):
        """Test ascending sort on three formulas using single column"""
        self.wb.set_cell_contents(self.name_c, 'd23', '')
        self.wb.set_cell_contents(self.name_c, 'd24', '=#REF!')
        self.wb.set_cell_contents(self.name_c, 'd25', None)
        self.wb.set_cell_contents(self.name_c, 'd26', '30')
        self.wb.set_cell_contents(self.name_c, 'd27', "")
        self.wb.set_cell_contents(self.name_c, 'd28', '=#ERROR!')

        self.wb.sort_region(self.name_c, 'd23', 'd28', [1])

        self.assertEqual(self.wb.get_cell_value(self.name_c, 'd23'), None)
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'd24'), None)
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'd25'), None)
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'd26').get_type()
                        == sheets.CellErrorType.PARSE_ERROR)
        self.assertEqual(self.wb.get_cell_contents(self.name_c, 'd26'), "=#ERROR!")
        self.assertTrue(self.wb.get_cell_value(self.name_c, 'd27').get_type()
                        == sheets.CellErrorType.BAD_REFERENCE)
        self.assertEqual(self.wb.get_cell_contents(self.name_c, 'd27'), "=#REF!")
        self.assertEqual(self.wb.get_cell_value(self.name_c, 'd28'), decimal.Decimal('30'))


if __name__ == '__main__':
    unittest.main()
