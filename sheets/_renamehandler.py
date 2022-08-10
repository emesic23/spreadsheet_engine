"""
Caltech CS130 - Winter 2022

File containing the reference getter class
"""

import lark
from ._utils import get_cell_location, num_to_col


# TODO Update docs # pylint: disable=W0511
class RenameHandler(lark.Visitor):
    """
    The reference getter class

    Attributes:
       references
       workbook
       sheet

    Methods:
       cell        Updates the cell value and type.
    """

    # TODO update type # pylint: disable=W0511
    def __init__(self,  # pylint: disable=R0913  # Not too many arguments
                 tree,
                 sheetchangefrom=None,
                 sheetchangeto=None,
                 cellchangefrom=None,
                 coldelta=None,
                 rowdelta=None):
        self.sheetchangefrom = sheetchangefrom
        self.sheetchangeto = sheetchangeto
        self.sheettype = None
        if self.sheetchangefrom:
            if self.sheetchangeto[0] == '"' and self.sheetchangeto[-1] == '"':
                self.sheettype = "QUOTED_SHEET_NAME"
            else:
                self.sheettype = "SHEET_NAME"
        self.cellchangefrom = cellchangefrom
        self.coldelta = coldelta
        self.rowdelta = rowdelta
        self.original_tree = tree.copy()

    # TODO update types and docstring # pylint: disable=W0511
    def cell(self, tree):  # pylint: disable=R0912  # Not too many branches
        """Process cell"""
        if self.sheetchangefrom and self.cellchangefrom:
            if (tree.children[0].type in ('SHEET_NAME', 'QUOTED_SHEET_NAME')
                    and tree.children[0].value.lower() == self.sheetchangefrom.lower()
                    and tree.children[-1].value.lower() == self.cellchangefrom.lower()):

                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)

                rowcol = tree.children[-1].value
                col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
                if not is_absolute_col:
                    col += self.coldelta
                if not is_absolute_row:
                    row += self.rowdelta

                col = num_to_col(col)
                loc_string = ''
                if is_absolute_col:
                    loc_string += "$"
                loc_string += col
                if is_absolute_row:
                    loc_string += "$"
                loc_string += str(row)
                tree.children[-1] = lark.Token('CELLREF', loc_string)

        if self.sheetchangeto:
            if tree.children[0].type == "SHEET_NAME" and tree.children[0].value.lower() == self.sheetchangefrom.lower():
                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)
            elif tree.children[0].type == "QUOTED_SHEET_NAME" and tree.children[0].value.lower() == self.sheetchangefrom.lower():
                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)

        if self.cellchangefrom:
            rowcol = tree.children[-1].value
            col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
            if not is_absolute_col:
                col += self.coldelta
            if not is_absolute_row:
                row += self.rowdelta
            col = num_to_col(col, throw_errors=False)
            loc_string = ''
            if is_absolute_col:
                loc_string += "$"
            loc_string += col
            if is_absolute_row:
                loc_string += "$"
            loc_string += str(row)
            tree.children[-1] = lark.Token('CELLREF', loc_string)

    def range(self, tree):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511,R0912,R0915
        if self.sheetchangefrom and self.cellchangefrom:
            if (tree.children[0].type in ('SHEET_NAME', 'QUOTED_SHEET_NAME')
                    and tree.children[0].value.lower() == self.sheetchangefrom.lower()
                    and tree.children[-1].value.lower() == self.cellchangefrom.lower()):

                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)

                rowcol = tree.children[-2].value
                col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
                if not is_absolute_col:
                    col += self.coldelta
                if not is_absolute_row:
                    row += self.rowdelta

                col = num_to_col(col)
                loc_string = ''
                if is_absolute_col:
                    loc_string += "$"
                loc_string += col
                if is_absolute_row:
                    loc_string += "$"
                loc_string += str(row)
                tree.children[-2] = lark.Token('CELLREF', loc_string)

                rowcol = tree.children[-1].value
                col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
                if not is_absolute_col:
                    col += self.coldelta
                if not is_absolute_row:
                    row += self.rowdelta

                col = num_to_col(col)
                loc_string = ''
                if is_absolute_col:
                    loc_string += "$"
                loc_string += col
                if is_absolute_row:
                    loc_string += "$"
                loc_string += str(row)
                tree.children[-1] = lark.Token('CELLREF', loc_string)

        if self.sheetchangeto:
            if tree.children[0].type == "SHEET_NAME" and tree.children[0].value.lower() == self.sheetchangefrom.lower():
                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)
            elif tree.children[0].type == "QUOTED_SHEET_NAME" and tree.children[0].value.lower() == self.sheetchangefrom.lower():
                tree.children[0] = lark.Token(self.sheettype, self.sheetchangeto)

        if self.cellchangefrom:
            rowcol = tree.children[-2].value
            col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
            if not is_absolute_col:
                col += self.coldelta
            if not is_absolute_row:
                row += self.rowdelta

            col = num_to_col(col)
            loc_string = ''
            if is_absolute_col:
                loc_string += "$"
            loc_string += col
            if is_absolute_row:
                loc_string += "$"
            loc_string += str(row)
            tree.children[-2] = lark.Token('CELLREF', loc_string)

            rowcol = tree.children[-1].value
            col, row, is_absolute_col, is_absolute_row = get_cell_location(rowcol)
            if not is_absolute_col:
                col += self.coldelta
            if not is_absolute_row:
                row += self.rowdelta

            col = num_to_col(col)
            loc_string = ''
            if is_absolute_col:
                loc_string += "$"
            loc_string += col
            if is_absolute_row:
                loc_string += "$"
            loc_string += str(row)
            tree.children[-1] = lark.Token('CELLREF', loc_string)
