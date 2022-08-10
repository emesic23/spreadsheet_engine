"""
Caltech CS130 - Winter 2022

Testing basic string functionality and parsing
"""

import unittest

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestBasicStrings(unittest.TestCase):
    """Testing basic string functionality and parsing"""

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testValues(self):
        """Test values for parsing"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_a, 'a2', '1234a')
        self.wb.set_cell_contents(self.name_a, 'a3', '="    abc"')
        self.wb.set_cell_contents(self.name_a, 'a4', '"    abc"')
        self.wb.set_cell_contents(self.name_a, 'a5', "ab \'ab\' ab")
        self.wb.set_cell_contents(self.name_a, 'a6', '   abc')
        self.wb.set_cell_contents(self.name_a, 'a7', '   abc   ')
        self.wb.set_cell_contents(self.name_a, 'a8', 'abc   ')

        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual('1234a', self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertEqual('    abc', self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual('"    abc"', self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual("ab \'ab\' ab", self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a6'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a7'))
        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a8'))

    def testSingleQuoteStrings(self):
        """Test strings that start with one quote"""
        self.wb.set_cell_contents(self.name_a, 'a1', "'abc")
        self.wb.set_cell_contents(self.name_a, 'b1', "'   abc")
        self.wb.set_cell_contents(self.name_a, 'a2', "'135.2")
        self.wb.set_cell_contents(self.name_a, 'b2', "'   135.2")
        self.wb.set_cell_contents(self.name_a, 'a3', "'=A1+B1")
        self.wb.set_cell_contents(self.name_a, 'b3', "'   =A1+B1")
        self.wb.set_cell_contents(self.name_a, 'a4', "'#REF!")
        self.wb.set_cell_contents(self.name_a, 'a5', "  '#REF!")
        self.wb.set_cell_contents(self.name_a, 'b4', "'   #REF!")
        self.wb.set_cell_contents(self.name_a, 'b5', "''")
        self.wb.set_cell_contents(self.name_a, 'b6', "'   '")

        self.assertEqual('abc', self.wb.get_cell_value(self.name_a, 'a1'))

        self.assertEqual('   abc', self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertEqual("'   abc", self.wb.get_cell_contents(self.name_a, 'b1'))

        self.assertEqual('135.2', self.wb.get_cell_value(self.name_a, 'a2'))
        self.assertEqual('   135.2', self.wb.get_cell_value(self.name_a, 'b2'))
        self.assertEqual('=A1+B1', self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertEqual('   =A1+B1', self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertEqual('#REF!', self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual('#REF!', self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertEqual('   #REF!', self.wb.get_cell_value(self.name_a, 'b4'))
        self.assertEqual("'", self.wb.get_cell_value(self.name_a, 'b5'))
        self.assertEqual("   '", self.wb.get_cell_value(self.name_a, 'b6'))

    def testOverwrite(self):
        """Test overwriting string with other string"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_a, 'a1', 'cba')

        self.assertTrue('cba' == self.wb.get_cell_value(self.name_a, 'a1'))

    def testReference(self):
        """Test references"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')

        self.assertTrue('abc' == self.wb.get_cell_value(self.name_a, 'a2'))

    def testUpdatingReferenceSameSheet(self):
        """Test references in the same sheet"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_a, 'a2', '=a1')
        self.wb.set_cell_contents(self.name_a, 'a1', 'defg')

        self.assertTrue('defg' == self.wb.get_cell_value(self.name_a, 'a2'))

    def testUpdatingReferenceDifferentSheet(self):
        """Test references in different sheets"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_b, 'a2', '=Sheet1!a1')
        self.wb.set_cell_contents(self.name_a, 'a1', 'defg')

        self.assertTrue('defg' == self.wb.get_cell_value(self.name_b, 'a2'))

    def testEmptyCell(self):
        """Test empty cell"""
        self.assertIsNone(self.wb.get_cell_value(self.name_a, 'a1'))

    def testStringConcat(self):
        """Test string concatenation"""
        self.wb.set_cell_contents(self.name_a, 'a1', 'abc')
        self.wb.set_cell_contents(self.name_a, 'a2', 'def')
        self.wb.set_cell_contents(self.name_a, 'a3', '=a1&a2')
        self.wb.set_cell_contents(self.name_a, 'a4', '=a1&a20')
        self.wb.set_cell_contents(self.name_a, 'a5', '=a15&a20')
        self.wb.set_cell_contents(self.name_a, 'a6', '="ab"&a20')
        self.wb.set_cell_contents(self.name_a, 'a7', '=a20&"ab"')

        self.wb.set_cell_contents(self.name_a, 'b1', '   abc')
        self.wb.set_cell_contents(self.name_a, 'b2', '   def   ')
        self.wb.set_cell_contents(self.name_a, 'b3', '=b1&b2')

        self.assertTrue('abcdef' == self.wb.get_cell_value(self.name_a, 'a3'))
        self.assertTrue('abcdef' == self.wb.get_cell_value(self.name_a, 'b3'))
        self.assertTrue('abc' == self.wb.get_cell_value(self.name_a, 'a4'))
        self.assertEqual('=a15&a20', self.wb.get_cell_contents(self.name_a, 'a5'))
        self.assertEqual('', self.wb.get_cell_value(self.name_a, 'a5'))
        self.assertEqual('ab', self.wb.get_cell_value(self.name_a, 'a6'))
        self.assertEqual('ab', self.wb.get_cell_value(self.name_a, 'a7'))

    def testNumericalAndStringConcat(self):
        """Test numerical and string concatenation"""
        self.wb.set_cell_contents(self.name_a, 'a1', '=(1 + 1) & "hello"')
        self.wb.set_cell_contents(self.name_a, 'a4', '=1.000 & " is one"')

        self.wb.set_cell_contents(self.name_b, 'a1', '=(1-1) & " is zero"')
        self.wb.set_cell_contents(self.name_b, 'a2', '=(1+1) & " is two"')
        self.wb.set_cell_contents(self.name_b, 'a3', '="three is " &(3+3-3)')

        self.assertTrue('2hello' == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertEqual(self.wb.get_cell_value(self.name_a, 'a4'), "1 is one")

        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a1'), "0 is zero")
        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a2'), "2 is two")
        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a3'), "three is 3")

    def testUnparseableStrings(self):
        """Test unparseable strings"""
        self.wb.set_cell_contents(self.name_a, 'a3', '=String')
        self.assertTrue(isinstance(self.wb.get_cell_value(self.name_a, 'a3'), sheets.CellError))
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'a3').get_type() == sheets.CellErrorType.PARSE_ERROR)

    def testQuotedString(self):
        """Test strings in quotes"""
        self.wb.set_cell_contents(self.name_a, 'a1', '"Hello"')
        self.wb.set_cell_contents(self.name_a, 'a2', '"       Hello   "')

        self.assertTrue('"Hello"' == self.wb.get_cell_value(self.name_a, 'a1'))
        self.assertTrue('"       Hello   "' == self.wb.get_cell_value(self.name_a, 'a2'))


if __name__ == '__main__':
    unittest.main()
