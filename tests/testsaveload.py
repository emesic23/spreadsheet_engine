# pylint: skip-file
# TODO fix pylint in this file

"""
Caltech CS130 - Winter 2022

Testing for references. Basic testing to ensure that references are
working as expected and auto-updating. Also includes case
insensitivity tests
"""


import unittest
import decimal
import json

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


class TestSaveLoad(unittest.TestCase):
    """
    Testing for references. Basic testing to ensure that references are
    working as expected and auto-updating. Also includes case
    insensitivity tests
    """

    def setUp(self):
        pass

    def tearDown(self):
        pass

    def testBasicLoad(self):
        """Test references and formula parsing within a sheet"""
        f = open("tests/jsons/ref/BasicLoad.json", 'r')
        self.wb = sheets.Workbook.load_workbook(f)

        self.assertEqual("'123", self.wb.get_cell_contents("Sheet1", 'a1'))
        self.assertEqual("123", self.wb.get_cell_value("Sheet1", 'a1'))
        self.assertEqual(decimal.Decimal('5.3'), self.wb.get_cell_value("Sheet1", 'b1'))
        self.assertEqual(decimal.Decimal('651.9'), self.wb.get_cell_value("Sheet1", 'c1'))
        self.assertEqual(decimal.Decimal('20'), self.wb.get_cell_value("Sheet2", 'A1'))
        self.assertEqual(decimal.Decimal('369'), self.wb.get_cell_value("Sheet2", 'b1'))

        f.close()

    def testBasicSave(self):
        f = open("tests/jsons/ref/BasicLoad.json", 'r')
        f2 = open("tests/jsons/out/BasicSaveOutput.json", 'w')
        self.wb = sheets.Workbook.load_workbook(f)
        self.wb.save_workbook(f2)
        f.close()
        f2.close()
        f = open("tests/jsons/ref/BasicLoad.json", 'r')
        f2 = open("tests/jsons/out/BasicSaveOutput.json", 'r')
        json1 = json.load(f)
        json2 = json.load(f2)
        self.assertEqual(json1, json2)
        f.close()
        f2.close()

    # TODO Not sure how they make the json to do this  # pylint: disable=W0511
    # def testTypeErrors(self):
    #     f = open("tests/jsons/ref/SheetDict.json", 'r')
    #     self.assertRaises(TypeError, sheets.Workbook.load_workbook(f))


if __name__ == '__main__':
    unittest.main()
