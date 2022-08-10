"""
Caltech CS130 - Winter 2022

Tests to ensure that sheetnames are being generated, stored
deleted, and remade correctly. Also includes case insensitivity
checks
"""


import unittest

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestSheetName(unittest.TestCase):
    """
    Tests to ensure that sheetnames are being generated, stored
    deleted, and remade correctly. Also includes case insensitivity
    checks
    """

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def testSheetNamesGenerated(self):
        """Test name generation without deletion"""
        wb = sheets.Workbook()

        (_, name_a) = wb.new_sheet()  # Sheet1
        (_, name_b) = wb.new_sheet()  # Sheet2
        (_, name_c) = wb.new_sheet("sheet3")  # sheet3
        (_, name_d) = wb.new_sheet('MySheet')  # MySheet
        (_, name_e) = wb.new_sheet()  # Sheet4

        sheetNames = wb.list_sheets()

        self.assertTrue(len(sheetNames) == 5)

        self.assertTrue(sheetNames[0] == 'Sheet1')
        self.assertTrue(name_a == 'Sheet1')

        self.assertTrue(sheetNames[1] == 'Sheet2')
        self.assertTrue(name_b == 'Sheet2')

        self.assertTrue(sheetNames[2] == 'sheet3')
        self.assertTrue(name_c == 'sheet3')

        self.assertTrue(sheetNames[3] == 'MySheet')
        self.assertTrue(name_d == 'MySheet')

        self.assertTrue(sheetNames[4] == 'Sheet4')
        self.assertTrue(name_e == 'Sheet4')

    def testSheetNamesGeneratedNamesWithDelete(self):
        """Test generated sheet names with user specified and delete"""
        wb = sheets.Workbook()

        (_, name_a) = wb.new_sheet()  # Sheet1
        (_, name_b) = wb.new_sheet()  # Sheet2 - deleted
        (_, name_c) = wb.new_sheet()  # Sheet3
        (_, name_d) = wb.new_sheet('MySheet')  # MySheet
        (_, name_e) = wb.new_sheet()  # Sheet4 - deleted
        wb.del_sheet(name_e)
        (_, name_f) = wb.new_sheet()  # Sheet4
        wb.del_sheet(name_b)
        (_, name_g) = wb.new_sheet()  # Sheet2

        sheetNames = wb.list_sheets()

        self.assertTrue(len(sheetNames) == 5)

        self.assertTrue(sheetNames[0] == 'Sheet1')
        self.assertTrue(name_a == 'Sheet1')

        self.assertTrue(name_b == 'Sheet2')

        self.assertTrue(sheetNames[1] == 'Sheet3')
        self.assertTrue(name_c == 'Sheet3')

        self.assertTrue(sheetNames[2] == 'MySheet')
        self.assertTrue(name_d == 'MySheet')

        self.assertTrue(sheetNames[3] == 'Sheet4')
        self.assertTrue(name_e == 'Sheet4')
        self.assertTrue(name_f == 'Sheet4')

        self.assertTrue(sheetNames[4] == 'Sheet2')
        self.assertTrue(name_g == 'Sheet2')

    def testSheetNamesValidUserSpecified(self):
        """Test various valid user specified sheet names"""
        wb = sheets.Workbook()

        (_, name_a) = wb.new_sheet('My Sheet')  # My Sheet
        (_, name_b) = wb.new_sheet('mY sHeEt 2')  # mY sHeEt 2
        (_, name_c) = wb.new_sheet('.?!,:;!@#$%^&*()-_')  # .?!,:;!@#$%^&*()-_

        sheetNames = wb.list_sheets()

        self.assertTrue(len(sheetNames) == 3)

        self.assertTrue(sheetNames[0] == 'My Sheet')
        self.assertTrue(name_a == 'My Sheet')

        self.assertTrue(sheetNames[1] == 'mY sHeEt 2')
        self.assertTrue(name_b == 'mY sHeEt 2')

        self.assertTrue(sheetNames[2] == '.?!,:;!@#$%^&*()-_')
        self.assertTrue(name_c == '.?!,:;!@#$%^&*()-_')

    def testSheetNamesInvalidUserSpecified(self):
        """Test error raising when user names sheet poorly"""
        wb = sheets.Workbook()

        # Already exists in other case
        wb.new_sheet('My Sheet')
        self.assertRaises(ValueError, wb.new_sheet, 'my sheet')

        # No empty strings
        self.assertRaises(ValueError, wb.new_sheet, '')
        self.assertRaises(ValueError, wb.new_sheet, '   ')

        # No start/end with white space
        self.assertRaises(ValueError, wb.new_sheet, ' name')
        self.assertRaises(ValueError, wb.new_sheet, 'name ')

        # No quotes
        self.assertRaises(ValueError, wb.new_sheet, '\"my sheet\"')
        self.assertRaises(ValueError, wb.new_sheet, '\'my sheet\'')

        # Empty string
        self.assertRaises(ValueError, wb.new_sheet, "")


if __name__ == '__main__':
    unittest.main()
