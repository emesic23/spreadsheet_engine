"""
Caltech CS130 - Winter 2022

File containing the CellType, CellErrorType, CellError, and Cell classes
"""

import enum
from typing import Optional, Tuple

from lark import Tree

from ._sheet import Sheet
from ._utils import split_sheet_cell, get_cell_location, num_to_col


class CellType(enum.Enum):
    """This enum specifies the types of cells that a spreadsheet can hold"""

    # A string value
    STRING = 1

    # A number value
    NUMBER = 2

    # An Expression
    EXPRESSION = 3

    # A Boolean
    BOOLEAN = 4


class CellErrorType(enum.Enum):
    """This enum specifies the kinds of errors that spreadsheet cells can hold."""

    # A formula doesn't parse successfully ("#ERROR!")
    PARSE_ERROR = 1

    # A cell is part of a circular reference ("#CIRCREF!")
    CIRCULAR_REFERENCE = 2

    # A cell-reference is invalid in some way ("#REF!")
    BAD_REFERENCE = 3

    # Unrecognized function name ("#NAME?")
    BAD_NAME = 4

    # A value of the wrong type was encountered during evaluation ("#VALUE!")
    TYPE_ERROR = 5

    # A divide-by-zero was encountered during evaluation ("#DIV/0!")
    DIVIDE_BY_ZERO = 6


class CellError:
    """
    This class represents an error value from user input, cell parsing, or
    evaluation.

    Attributes:
       _error_type    Type of cell
       _detail        Details of the error
       _exception     The exception raised if the CellError was generated from
                      a raised exception
    """

    def __init__(self, error_type: CellErrorType, detail: str,
                 exception: Optional[Exception] = None):
        self._error_type = error_type
        self._detail = detail
        self._exception = exception

    def get_type(self) -> CellErrorType:
        """The category of the cell error."""
        return self._error_type

    def get_detail(self) -> str:
        """More detail about the cell error."""
        return self._detail

    def get_exception(self) -> Optional[Exception]:
        """
        If the cell error was generated from a raised exception, this is the
        exception that was raised.  Otherwise this will be None.
        """
        return self._exception

    def __str__(self) -> str:
        """Return error as a string"""
        return f'ERROR[{self._error_type}, "{self._detail}"]'

    def __repr__(self) -> str:
        """Return self representation"""
        return self.__str__()


class Cell:
    """
    The cell class

    Attributes:
       sheet              The parent sheet
       value              The evaluated cell or error type if error
       contents           The contents of the cell
       type               The cell type from the CellType enum

    Methods:
       update_cell        Updates the cell value and type.
       update_value       Updates the value of the cell.
       update_type        Updates the type of the cell.
       get_value          Returns the value of the cell.
       get_type           Returns the type of the cell.
       get_contents       Returns the contents of the cell.
       get_sheet          Returns the parent sheet.
    """

    def __init__(self,  # pylint: disable=R0913  # Not too many arguments
                 sheet: Sheet,
                 contents: object = None,
                 value: object = 0,
                 cell_type: CellType = CellType.NUMBER,
                 parse_tree: Tree = None):
        self.sheet = sheet              # The parent sheet
        self.value = value              # The evaluated cell or error type if error
        self.contents = contents        # Contents of the cell
        self.cell_type = cell_type      # The type of cell or error
        self.parse_tree = parse_tree    # The cell's parse tree

    def update_cell(self, value: object, cell_type: str) -> None:
        """Updates the cell value and type"""
        self.update_value(value)
        self.update_type(cell_type)

    def update_value(self, value: object) -> None:
        """Updates the value of the cell"""
        self.value = value

    def update_type(self, cell_type: CellType):
        """Updates the type of the cell"""
        self.cell_type = cell_type

    def get_value(self) -> object:
        """Returns the value of the cell"""
        return self.value

    def get_type(self) -> str:
        """Returns the type of the cell"""
        return self.cell_type

    def get_contents(self) -> object:
        """Returns the contents of the cell"""
        return self.contents

    def get_tree(self) -> object:
        """Returns the cell's parse tree"""
        return self.parse_tree

    def get_sheet(self) -> Sheet:
        """Returns the parent sheet"""
        return self.sheet

    def copy(self, sheet):
        """Returns a new copy of the cell"""
        return Cell(sheet, self.contents, self.value, self.cell_type, self.parse_tree)


class CellReference:
    """
    The cell reference class

    Attributes:
       sheet              The parent sheet
       col                The cell's column
       row                The cell's row
       abs_col            If the reference to the cell's column is absolute
       abs_row            If the reference to the cell's row is absolute
       loc                The cell's location


    Methods:
       get_sheet          Returns the parent sheet.
       get_abs_col        Returns if the reference to the cell's column is absolute
       get_abs_row        Returns if the reference to the cell's row is absolute
       get_cell_loc       Returns cell location
    """

    def __init__(self,
                 sheet: str,
                 col: str,
                 row: str,):
        self.sheet = sheet.lower()
        self.col = col.lower()
        self.row = str(row).lower()
        self.loc = self.sheet + '!' + self.col + self.row

        sheet, cell_loc = split_sheet_cell(self.loc)

        try:
            stripped_col, stripped_row, _, _ = get_cell_location(cell_loc)
            stripped_col = num_to_col(stripped_col)
        except ValueError:
            stripped_col = "#REF!"
            stripped_row = ""

        self.loc_without_refs = str(sheet) + '!' + str(stripped_col) + str(stripped_row)

    def get_sheet(self) -> str:
        """Returns the parent sheet"""
        return self.sheet

    def get_col(self) -> str:
        """Returns the cell column"""
        return self.col

    def get_row(self) -> str:
        """Returns the cell row"""
        return self.row

    def get_rowcol(self) -> str:
        """Returns the cell rowcol"""
        return self.col + self.row

    def get_cell_loc(self) -> str:
        """Returns the cell location"""
        return self.loc

    def get_cell_loc_without_refs(self) -> str:
        """Returns the cell location"""
        return self.loc_without_refs

    def __eq__(self, __o: object) -> bool:
        return self.loc_without_refs == __o.loc_without_refs

    def __hash__(self) -> int:
        return hash(self.loc_without_refs)

    def __str__(self) -> str:
        return self.loc_without_refs


class SortingCell:
    """
    Class to hold a cell object and its original location. Used
    in SortingCellRow class to make sorting rows easier
    """
    def __init__(self,
                 cell: Cell,
                 original_loc: Tuple):
        self.cell = cell
        self.is_blank = False
        self.is_error = False
        if not cell:
            self.is_blank = True
        elif cell:
            self.is_blank = (cell.contents is None or cell.contents == '')
            if not self.is_blank:
                self.is_error = type(self.cell.value) == CellError  # pylint: disable=C0123
        self.original_col, self.original_row = original_loc
        self.original_loc = original_loc

    def __str__(self) -> str:
        if self.is_blank:
            return "Blank cell from " + str(self.original_loc)
        return str(self.cell.value) + " cell from " + str(self.original_loc)

    def __eq__(self, __o: object) -> bool:  # pylint: disable=R0911
        if self.is_blank and __o.is_blank:
            return True
        if self.is_error and __o.is_error:
            return self.cell.value.get_type() == __o.cell.value.get_type()
        if self.is_blank or __o.is_blank:
            return False
        if self.is_error or __o.is_error:
            return False
        if self.is_error and __o.is_blank:
            return False
        if self.is_blank and __o.is_error:
            return False
        return self.cell.value == __o.cell.value

    def __lt__(self, __o: object) -> bool:  # pylint: disable=R0911
        if self.is_blank and __o.is_blank:
            return False
        if self.is_error and __o.is_error:
            return self.cell.value.get_type().value < __o.cell.value.get_type().value
        if self.is_blank:
            return True
        if self.is_error:
            return not __o.is_blank
        if __o.is_blank:
            return False
        if __o.is_error:
            return False
        return self.cell.value < __o.cell.value

    def __gt__(self, __o: object) -> bool:
        return not self.__lt__(__o)


class SortingCellRow:
    """
    Class to work with row of cells stored as sorting cells
    """
    def __init__(self, sort_cols, original_row):
        self.cells = []
        self.sort_cols = sort_cols

        self.original_row = original_row
        # self.original_loc = original_loc

    def add_cell(self, cell: Cell, original_loc: Tuple):
        """
        Add a cell as a sorting cell to cells array
        """
        self.cells.append(SortingCell(cell, original_loc))

    def __str__(self) -> str:
        to_return = "Cell row containing:\n"

        for cell in self.cells:
            to_return += "\t" + str(cell) + " holding " + str(cell) + "\n"

        return to_return

    def __eq__(self, __o: object) -> bool:
        combined_bool = True
        # For each sort col, compare values
        for col in self.sort_cols:
            combined_bool = combined_bool and (self.cells[abs(col) - 1] == __o.cells[abs(col) - 1])

        return combined_bool

    def __lt__(self, __o: object) -> bool:
        # For each sort col, compare values
        for col in self.sort_cols:
            do_opposite = col < 0
            if self.cells[abs(col) - 1] != __o.cells[abs(col) - 1]:
                res = self.cells[abs(col) - 1] < __o.cells[abs(col) - 1]
                if do_opposite:
                    return not res
                return res

        return False

    def __gt__(self, __o: object) -> bool:
        # For each sort col, compare values
        return not self.__lt__(__o)
