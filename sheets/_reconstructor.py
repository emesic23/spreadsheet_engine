"""
Caltech CS130 - Winter 2022

File containing the custom reconstructor
"""

import lark
from ._utils import get_cell_location


# TODO Update docs # pylint: disable=W0511
class ContentReconstructor(lark.visitors.Interpreter):
    """
    The Content Reconstructor Class

    Attributes:
        content
    Methods:
       cell        Updates the cell value and type.
    """

    # TODO update type # pylint: disable=W0511
    def __init__(self):
        self.content = '='

    def add_expr(self, tree):  # TODO standardize returns # pylint: disable=W0511,R1710,R0911
        """
        Parse formula with addition or subtraction
        """
        self.visit(tree.children[0])
        op = tree.children[1].value
        self.content += str(op)
        self.visit(tree.children[-1])

    def mul_expr(self, tree):  # TODO standardize returns # pylint: disable=W0511,R1710,R0911
        """
        Parse formula with addition or subtraction
        """
        self.visit(tree.children[0])
        op = tree.children[1].value
        self.content += str(op)
        self.visit(tree.children[-1])

    def number(self, tree):
        """
        Parse number
        """
        self.content += tree.children[0]

    def concat_expr(self, tree):
        """
        Parse formula with concat (&)
        """
        self.visit(tree.children[0])
        self.content += '&'
        self.visit(tree.children[-1])

    def unary_op(self, tree):
        """
        Parse formula with unary ops
        """
        self.content += '-'
        self.visit(tree.children[-1])

    def cell(self, tree):
        """
        Parse cell
        """
        # TODO merge comparisons (pylint)
        cell_location = ""
        try:
            get_cell_location(tree.children[-1].value)
            cell_location = tree.children[-1].value
        except (ValueError, IndexError):
            cell_location = "#REF!"

        if tree.children[0].type == "SHEET_NAME" or tree.children[0].type == "QUOTED_SHEET_NAME":  # pylint: disable=R1714
            self.content += tree.children[0].value + '!' + cell_location
        else:
            self.content += cell_location

    def parens(self, tree):
        """
        Parse parentheses
        """
        self.content += '('
        self.visit(tree.children[-1])
        self.content += ')'

    def string(self, tree):
        """
        Parse string. Truncates on either side due to quotes
        """
        self.content += tree.children[0]

    def error(self, tree):
        """
        Parse error
        """
        self.content += tree.children[0]

    def range(self, tree):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
        if len(tree.children) == 3:
            self.content += tree.children[0].value + '!'
        self.content += tree.children[1].value + ':'
        self.content += tree.children[2].value
