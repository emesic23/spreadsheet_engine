"""
Caltech CS130 - Winter 2022

Testing general cell functions. Checks overwriting
cell types with other cell types
"""


import unittest
import decimal
import logging

import tests.context  # noqa: F401  # pylint: disable=E0401,W0611
import sheets  # pylint: disable=E0401,W0611


def on_cells_changed(_, changed_cells):
    """
    Function gets called when cells in workbook get changed. Prints
    changed cells
    """
    logging.info(f'Cell(s) changed:  {changed_cells}')  # pylint: disable=W1203  # Logging format is fine


class TestUpdateCellNotifications(unittest.TestCase):
    """
    Testing general cell functions. Checks overwriting
    cell types with other cell types
    """

    def setUp(self):
        """Set up tests, runs at the beginning of each case"""
        self.wb = sheets.Workbook()

        (self.index_a, self.name_a) = self.wb.new_sheet()  # Sheet1
        (self.index_b, self.name_b) = self.wb.new_sheet()  # Sheet2
        (self.index_c, self.name_c) = self.wb.new_sheet()  # Sheet3

    def tearDown(self):
        pass

    def testInsertsOnSingleSheet(self):
        """Test updating of cells on a single sheet"""
        self.wb.notify_cells_changed(on_cells_changed)

        with self.assertLogs() as captured:

            # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
            self.wb.set_cell_contents(self.name_a, "A1", "'123")

            # Generates one call to notify functions, with the argument [('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_a, "C1", "=" + self.name_a + "!A1+" + self.name_a + "!B1")

            # Generates one or more calls to notify functions, indicating that cells B1
            # and C1 have changed.  For example, there might be one call with the argument
            # [('Sheet1', 'B1'), ('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_a, "B1", "5.3")

        self.assertEqual(len(captured.records), 3)
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet1', 'C1')]")
        self.assertTrue(
            captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet1', 'B1'), ('Sheet1', 'C1')]"
            or captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet1', 'C1'), ('Sheet1', 'B1')]")

    def testInsertsOnMultipleSheets(self):
        """Test updating of cells on a single sheet"""
        self.wb.notify_cells_changed(on_cells_changed)

        with self.assertLogs() as captured:

            # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
            self.wb.set_cell_contents(self.name_a, "A1", "'123")

            # Generates one call to notify functions, with the argument [('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_b, "C1", "=" + self.name_a + "!A1+" + self.name_a + "!B1")

            # Generates one or more calls to notify functions, indicating that cells B1
            # and C1 have changed.  For example, there might be one call with the argument
            # [('Sheet1', 'B1'), ('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_a, "B1", "5.3")

        self.assertEqual(len(captured.records), 3)
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")
        self.assertTrue(
            captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet1', 'B1'), ('Sheet2', 'C1')]"
            or captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet2', 'C1'), ('Sheet1', 'B1')]")

    def testDeletion(self):
        """Test updating of cells on a single sheet"""
        self.wb.notify_cells_changed(on_cells_changed)

        with self.assertLogs() as captured:

            # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
            self.wb.set_cell_contents(self.name_a, "A1", "'123")

            # Generates one call to notify functions, with the argument [('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_b, "C1", "=" + self.name_a + "!A1+" + self.name_a + "!B1")

            # Generates one or more calls to notify functions, indicating that cells B1
            # and C1 have changed.  For example, there might be one call with the argument
            # [('Sheet1', 'B1'), ('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_a, "B1", "5.3")

            self.wb.del_sheet(self.name_a)

        self.assertEqual(len(captured.records), 4)
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")
        self.assertTrue(
            captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet1', 'B1'), ('Sheet2', 'C1')]"
            or captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet2', 'C1'), ('Sheet1', 'B1')]")
        self.assertEqual(captured.records[3].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")

    def testResurrection(self):
        """Test updating of cells on a single sheet"""
        self.wb.notify_cells_changed(on_cells_changed)

        with self.assertLogs() as captured:

            # Generates one call to notify functions, with the argument [('Sheet1', 'A1')].
            self.wb.set_cell_contents(self.name_a, "A1", "'123")

            # Generates one call to notify functions, with the argument [('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_b, "C1", "=" + self.name_a + "!A1+" + self.name_a + "!B1")

            # Generates one or more calls to notify functions, indicating that cells B1
            # and C1 have changed.  For example, there might be one call with the argument
            # [('Sheet1', 'B1'), ('Sheet1', 'C1')].
            self.wb.set_cell_contents(self.name_a, "B1", "5.3")

            self.wb.del_sheet(self.name_a)

            self.wb.new_sheet()
            self.wb.set_cell_contents(self.name_a, "A1", "'123")
            self.wb.set_cell_contents(self.name_a, "B1", "5.3")

        self.assertEqual(len(captured.records), 7)

        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")
        self.assertTrue(
            captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet1', 'B1'), ('Sheet2', 'C1')]"
            or captured.records[2].getMessage() == "Cell(s) changed:  [('Sheet2', 'C1'), ('Sheet1', 'B1')]")
        self.assertEqual(captured.records[3].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")
        self.assertEqual(captured.records[4].getMessage(), "Cell(s) changed:  [('Sheet2', 'C1')]")
        self.assertTrue(
            captured.records[5].getMessage() == "Cell(s) changed:  [('Sheet1', 'A1'), ('Sheet2', 'C1')]"
            or captured.records[5].getMessage() == "Cell(s) changed:  [('Sheet2', 'C1'), ('Sheet1', 'A1')]")
        self.assertTrue(
            captured.records[6].getMessage() == "Cell(s) changed:  [('Sheet1', 'B1'), ('Sheet2', 'C1')]"
            or captured.records[6].getMessage() == "Cell(s) changed:  [('Sheet2', 'C1'), ('Sheet1', 'B1')]")

    def testRename(self):
        """Test rename function"""
        self.wb.notify_cells_changed(on_cells_changed)
        with self.assertLogs() as captured:
            self.wb.set_cell_contents('Sheet1', 'a1', '3.0')
            self.wb.set_cell_contents('Sheet2', 'a1', '=Sheet1!a1')
            self.wb.rename_sheet('Sheet1', 'Renamed')
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet2', 'A1')]")
        self.assertEqual(captured.records[2].getMessage(), "Cell(s) changed:  [('Renamed', 'A1')]")
        self.assertEqual(captured.records[3].getMessage(), "Cell(s) changed:  [('Sheet2', 'A1')]")
        self.assertEqual(captured.records[4].getMessage(), "Cell(s) changed:  [('Sheet2', 'A1')]")

    def testResurrectionByRename(self):
        """Test resurrection by rename function"""
        self.wb.notify_cells_changed(on_cells_changed)
        with self.assertLogs() as captured:
            self.wb.set_cell_contents('Sheet1', 'a1', '=Sheet100!a1')
            self.wb.set_cell_contents(self.name_c, 'a1', '=1')
            self.wb.rename_sheet(self.name_c, 'Sheet100')
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet3', 'A1')]")
        self.assertEqual(captured.records[2].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertTrue(
            captured.records[3].getMessage() == "Cell(s) changed:  [('Sheet1', 'A1'), ('Sheet100', 'A1')]"
            or captured.records[3].getMessage() == "Cell(s) changed:  [('Sheet100', 'A1'), ('Sheet1', 'A1')]")

    def testCopy(self):
        """Test copy function"""
        self.wb.notify_cells_changed(on_cells_changed)
        self.wb.set_cell_contents('Sheet1', 'a1', '3.0')
        self.wb.set_cell_contents('Sheet1', 'b1', '=a1 + 9')
        self.assertTrue(self.wb.get_cell_value(self.name_a, 'b1') == decimal.Decimal(12))
        with self.assertLogs() as captured:
            self.wb.copy_sheet(self.name_a)  # My Sheet
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1_1', 'A1')]")
        self.assertEqual(captured.records[1].getMessage(), "Cell(s) changed:  [('Sheet1_1', 'B1')]")

    def testResurrectionByCopy(self):
        """Test resurrection by copy function"""
        self.wb.notify_cells_changed(on_cells_changed)
        self.wb.set_cell_contents('Sheet1', 'a1', '=Sheet1_1!a1')
        with self.assertLogs() as captured:
            self.wb.copy_sheet(self.name_a)
        self.assertEqual(captured.records[0].getMessage(), "Cell(s) changed:  [('Sheet1', 'A1')]")
        self.assertTrue(
            captured.records[1].getMessage() == "Cell(s) changed:  [('Sheet1_1', 'A1'), ('Sheet1', 'A1')]"
            or captured.records[1].getMessage() == "Cell(s) changed:  [('Sheet1', 'A1'), ('Sheet1_1', 'A1')]")


if __name__ == '__main__':
    unittest.main()
