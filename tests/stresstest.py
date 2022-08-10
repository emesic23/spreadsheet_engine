"""
Caltech CS130 - Winter 2022

Stress test functions for performance analysis
"""

import cProfile
import os
import unittest
import decimal

# from pyinstrument import Profiler
from pstats import Stats

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401


class TestStressTest(unittest.TestCase):
    """
    Stress test functions for performance analysis
    """
    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        # Set up dumps directory
        is_exist = os.path.exists('dumps')

        if not is_exist:
            os.makedirs('dumps')

        # Set up workbook
        self.wb = sheets.Workbook()
        self.profiler = cProfile.Profile()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet("Sheet 3")  # "Sheet 3"

    def tearDown(self):
        pass

    def testLargerDiamond(self):
        """Stress test a large diamond over single workbook"""
        filename = "dumps/testLargerDiamond"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + c1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=d1 + e1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=e1 + f1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=g1 + h1')
        self.wb.set_cell_contents(self.name_a, 'e1', '=h1 + i1')
        self.wb.set_cell_contents(self.name_a, 'f1', '=i1 + j1')
        self.wb.set_cell_contents(self.name_a, 'g1', '=k1')
        self.wb.set_cell_contents(self.name_a, 'h1', '=k1 + l1')
        self.wb.set_cell_contents(self.name_a, 'i1', '=l1 + m1')
        self.wb.set_cell_contents(self.name_a, 'j1', '=m1')
        self.wb.set_cell_contents(self.name_a, 'k1', '=n1')
        self.wb.set_cell_contents(self.name_a, 'l1', '=n1 + o1')
        self.wb.set_cell_contents(self.name_a, 'm1', '=o1')
        self.wb.set_cell_contents(self.name_a, 'n1', '=p1')
        self.wb.set_cell_contents(self.name_a, 'o1', '=p1')
        self.profiler.enable()

        for ii in range(200):
            self.wb.set_cell_contents(self.name_a, 'p1', str(ii))
        self.wb.set_cell_contents(self.name_a, 'p1', '1')

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value(self.name_a, 'a1'))

    def testLargerDiamondMultiSheet(self):
        """Stress test a large diamond over many workbooks"""
        filename = "dumps/testLargerDiamondMultiSheet"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        self.wb.set_cell_contents(self.name_a, 'a1', '=Sheet1!b1 + Sheet1!c1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=Sheet1!d1 + Sheet2!e1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=Sheet2!e1 + Sheet2!f1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=Sheet2!g1 + Sheet2!h1')
        self.wb.set_cell_contents(self.name_b, 'e1', '=Sheet2!h1 + Sheet2!i1')
        self.wb.set_cell_contents(self.name_b, 'f1', '=Sheet2!i1 + Sheet2!j1')
        self.wb.set_cell_contents(self.name_b, 'g1', "='Sheet 3'!k1")
        self.wb.set_cell_contents(self.name_b, 'h1', "='Sheet 3'!k1 + 'Sheet 3'!l1")
        self.wb.set_cell_contents(self.name_b, 'i1', "='Sheet 3'!l1 + 'Sheet 3'!m1")
        self.wb.set_cell_contents(self.name_b, 'j1', "='Sheet 3'!m1")
        self.wb.set_cell_contents(self.name_c, 'k1', "='Sheet 3'!n1")
        self.wb.set_cell_contents(self.name_c, 'l1', "='Sheet 3'!n1 + 'Sheet 3'!o1")
        self.wb.set_cell_contents(self.name_c, 'm1', "='Sheet 3'!o1")
        self.wb.set_cell_contents(self.name_c, 'n1', "='Sheet 3'!p1")
        self.wb.set_cell_contents(self.name_c, 'o1', "='Sheet 3'!p1")
        self.profiler.enable()

        for ii in range(200):
            self.wb.set_cell_contents(self.name_c, 'p1', str(ii))
        self.wb.set_cell_contents(self.name_c, 'p1', '1')

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

        self.assertEqual(decimal.Decimal('20'), self.wb.get_cell_value(self.name_a, 'a1'))

    def testDiamondRename(self):
        """Test a diamond over renamed sheet"""

        filename = "dumps/testDiamondRename"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        # self.profiler.enable()
        self.wb.set_cell_contents(self.name_a, 'a1', '=b1 + d1')
        self.wb.set_cell_contents(self.name_a, 'b1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'd1', '=c1')
        self.wb.set_cell_contents(self.name_a, 'c1', '=3')

        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'c1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'b1'))
        self.assertTrue(decimal.Decimal('3') == self.wb.get_cell_value(self.name_a, 'd1'))
        self.assertTrue(decimal.Decimal('6') == self.wb.get_cell_value(self.name_a, 'a1'))

        self.profiler.enable()
        for ii in range(200):
            self.wb.rename_sheet(self.name_a, "Sheet" + str(ii) + "ada")
            self.name_a = "Sheet" + str(ii) + "ada"
        self.wb.rename_sheet(self.name_a, "Sheet10")

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)
        self.wb.set_cell_contents("Sheet10", 'c1', '=10')

        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'c1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'b1'))
        self.assertTrue(decimal.Decimal('10') == self.wb.get_cell_value("Sheet10", 'd1'))
        self.assertTrue(decimal.Decimal('20') == self.wb.get_cell_value("Sheet10", 'a1'))

    def testCopyResurrect(self):
        """Test a diamond over many copied sheets"""

        filename = "dumps/testCopyResurrect"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        copies = 100
        self.wb.set_cell_contents(self.name_a, 'a1', '=1')
        self.wb.set_cell_contents(
            self.name_b, 'a1', '=' + " + ".join([f'Sheet1_{i}!a1' for i in range(1, copies + 1)]))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.profiler.enable()
        for _ in range(copies):
            self.wb.copy_sheet(self.name_a)

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a1'),
                         decimal.Decimal(copies))

    def testNewSheetResurrect(self):
        """Test a diamond over many new sheets"""

        filename = "dumps/testNewSheetResurrect"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        copies = 100
        self.wb.set_cell_contents(self.name_a, 'a1', '=1')
        self.wb.set_cell_contents(
            self.name_b, 'a1', '=' + " + ".join([f'Sheet1_{i}!a1' for i in range(1, copies + 1)]))
        self.assertTrue(self.wb.get_cell_value(self.name_b, 'a1').get_type() == sheets.CellErrorType.BAD_REFERENCE)

        self.profiler.enable()
        for ii in range(1, copies + 1):
            name = f'Sheet1_{ii}'
            self.wb.new_sheet(name)
        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

        self.assertEqual(self.wb.get_cell_value(self.name_b, 'a1'),
                         decimal.Decimal(0))

    def testLargeLoad(self):
        """Test loading a large area"""

        filename = "dumps/testLargeLoad"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoad.json", 'r', encoding="utf8") as file:
            self.profiler.enable()

            self.wb = sheets.Workbook.load_workbook(file)

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testCopySheet(self):
        """Test copying a sheet"""

        filename = "dumps/testCopySheet"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoad.json", 'r', encoding="utf8") as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.copy_sheet("Sheet 3")

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testRenameSheet(self):
        """Test renaming a sheet"""

        filename = "dumps/testRenameSheet"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoad.json", 'r', encoding="utf8") as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.rename_sheet("Sheet 3", "New Sheet Name")

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testMoveBlankCellArea(self):
        """Test moving a large target area that is entirely blank"""

        filename = "dumps/testMoveBlankCellArea"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        self.profiler.enable()

        self.wb.move_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

    def testMoveCellArea(self):
        """Test moving a large target area that has cells"""

        filename = "dumps/testMoveCellArea"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoad.json", 'r', encoding="utf8") as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.move_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testMoveCellAreaWithRefs(self):
        """Test moving a large target area that has cells with refs"""

        filename = "dumps/testMoveCellAreaWithRefs"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoadWithRefs.json", 'r', encoding="utf8") as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.move_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testCopyBlankCellArea(self):
        """Test copying a large target area that is entirely blank"""

        filename = "dumps/testCopyBlankCellArea"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        self.profiler.enable()

        self.wb.copy_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

        profiler = Stats(self.profiler)
        profiler.strip_dirs().dump_stats(filename)

    def testCopyCellArea(self):
        """Test copying a large target area that has cells"""

        filename = "dumps/testCopyCellArea"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoad.json", 'r', encoding='utf8') as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.copy_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)

    def testCopyCellAreaWithRefs(self):
        """Test copying a large target area that has cells with refs"""

        filename = "dumps/testCopyCellAreaWithRefs"
        is_exist = os.path.exists(filename)

        if not is_exist:
            with open(filename, "w+", encoding='utf8') as file_handle:
                file_handle.close()

        with open("tests/jsons/ref/LargeLoadWithRefs.json", 'r', encoding='utf8') as file:
            self.wb = sheets.Workbook.load_workbook(file)

            self.profiler.enable()

            self.wb.copy_cells(self.name_a, 'a1', 'zz99', 'a1', self.name_b)

            profiler = Stats(self.profiler)
            profiler.strip_dirs().dump_stats(filename)


if __name__ == '__main__':
    unittest.main()
