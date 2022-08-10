"""
Caltech CS130 - Winter 2022

File containing the reference getter class
"""
import re
import lark
from ._cell import CellReference
from ._utils import cell_range


# TODO Update docs # pylint: disable=W0511
class ReferenceGetter(lark.visitors.Interpreter):
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
    def __init__(self, workbook, sheet):
        self.references = []
        self.workbook = workbook
        self.sheet = sheet

    # TODO update types and docstring # pylint: disable=W0511
    def cell(self, tree):
        """Process cell"""
        usedsheet = ""
        col, row, _ = re.split(r'(\d+)', tree.children[-1].value)
        try:
            if tree.children[0].type == "SHEET_NAME":
                usedsheet = self.workbook.get_sheet(tree.children[0].value)
            elif tree.children[0].type == "QUOTED_SHEET_NAME":
                usedsheet = self.workbook.get_sheet(
                    tree.children[0].value[1:-1])
            else:
                usedsheet = self.sheet
            self.references.append(CellReference(usedsheet.name, col, row))
        except KeyError:
            name = tree.children[0].value.strip("'")
            self.references.append(CellReference(name, col, row))

    # TODO write docstring Esmir  # pylint: disable=W0511
    def function(self, tree):
        """ TODO """
        children = self.visit_children(tree)
        if len(children) == 2:
            func, params = children
        else:
            func = children[0]
        func = func.lower()
        if func == "indirect":
            if tree.children[1].children[0].data == "string":
                try:
                    if "!" in params[0]:
                        cellref = params[0].split("!")
                        sheet = "!".join(cellref[:-1])
                        col, row, _ = re.split(r'(\d+)', cellref[-1])
                    else:
                        sheet = self.sheet.name
                        col, row, _ = re.split(r'(\d+)', params[0])
                    self.references.append(CellReference(sheet, col, row))
                except IndexError:
                    pass

    def range(self, tree):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
        children = tree.children
        sheetname = None
        if len(children) == 2:
            start, end = children
        else:
            sheetname, start, end = children

        cell_trees = []
        for row in cell_range(start, end):
            row_trees = []
            for cell in row:
                rule = lark.Token('RULE', 'cell')
                cellref = lark.Token('CELLREF', cell)

                if sheetname:
                    sheet = lark.Token('SHEET_NAME', sheetname)
                    row_trees.append(lark.Tree(rule, [sheet, cellref]))
                else:
                    row_trees.append(lark.Tree(rule, [cellref]))
            cell_trees.append(row_trees)

        for row in cell_trees:
            for tree_in_row in row:
                self.visit(tree_in_row)

    def string(self, tree):
        """
        Parse string. Truncates on either side due to quotes
        """
        return tree.children[0][1:-1]
