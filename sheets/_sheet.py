"""
Caltech CS130 - Winter 2022

File containing the sheet class
"""

import re

from ._utils import col_to_num


class Sheet:
    """
    A sheet containing cells.

    Attributes:
       sheet_id       The id of the sheet
       name           The sheet name in original casing
       entries        A dictionary of the cells

    Methods:
       get_extent     Returns the (number of rows, number of columns) in the
                      sheet as a tuple of ints
       get_cell       Return the cell at the specified location
    """

    def __init__(self, sheet_id: str, name: str):
        self.sheet_id = sheet_id
        self.name = name
        self.entries = {}

    def get_extent(self):
        """
        Returns the (number of rows, number of columns) in the sheet as a
        tuple of ints
        """
        extent_row = 0
        extent_col = 0
        for cell_location in self.entries:
            col, row, _ = re.split(r'(\d+)', cell_location)
            col = col_to_num(col)
            col = int(col)
            row = int(row)
            extent_col = max(extent_col, col)
            extent_row = max(extent_row, row)

        return extent_col, extent_row

    def get_cell(self, location: str) -> any:
        """Return the cell at the specified location"""
        return self.entries[location]
