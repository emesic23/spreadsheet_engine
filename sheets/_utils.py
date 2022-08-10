"""
Caltech CS130 - Winter 2022

File containing various utility functions

Methods:
   is_number      Checks whether a value is a number. Returns boolean.
   is_formula     Checks whether a value is a formula. Returns boolean.
   is_string      Checks whether a value is a string. Returns boolean.
   is_error       Checks whether a value is a error. Returns boolean.
   is_quoted      Checks whether a value is a quoted. Returns boolean.
   col_to_num     Converts column string to a number. Returns int.
"""

import decimal
import re
from typing import Optional


MAX_COL_NUM = 26 * 26 * 26 * 26 + 26 * 26 * 26 + 26 * 26 + 26
MAX_ROW_NUM = 9999


def is_number(val: Optional[str]) -> bool:
    """Checks whether a value is a number. Returns boolean."""
    try:
        decimal.Decimal(val)
        return True
    except (decimal.InvalidOperation, TypeError):
        return False


def is_formula(val: Optional[str]) -> bool:
    """Checks whether a value is a formula. Returns boolean."""
    try:
        return val[0] == '='
    except (IndexError, TypeError):
        return False


def is_string(val: Optional[str]) -> bool:
    """Checks whether a value is a string. Returns boolean."""
    try:
        return val[0] != '='
    except (IndexError, TypeError):
        return False


def is_error(val: Optional[str]) -> bool:
    """Checks whether a value is a error. Returns boolean."""
    try:
        return val[0] == '#'
    except (IndexError, TypeError):
        return False


def is_quoted(val: Optional[str]) -> bool:
    """Checks whether a value is a quoted. Returns boolean."""
    try:
        return val[0] == "'"
    except (IndexError, TypeError):
        return False


def col_to_num(col: str, throw_errors=True) -> int:
    """Converts column string to a number. Returns int."""
    col = col.lower()

    num = 0
    for ii in range(len(col)):
        char = col[-(ii + 1)]
        num += (ord(char) - 96) * 26**(ii)

    if throw_errors and (num > MAX_COL_NUM or num < 1):
        raise ValueError("Cell Reference Out of Bounds")

    return num


def num_to_col(num: int, throw_errors=True) -> str:
    """
    Given a column as a number, returns the string representation or raises
    a ValueError if the column is out of bounds.
    """
    if throw_errors and (num > MAX_COL_NUM or num < 1):
        raise ValueError("Number to Column Out of Bounds")

    col = ""
    while num > 0:
        num, remainder = divmod(num - 1, 26)
        col = chr(97 + remainder) + col
    return col


def split_sheet_cell(location: str, remove_quotes: bool = True) -> tuple[str, str]:
    """
    Splits a string of SheetName!CellLoc into (SheetName, CellLoc) accounting
    for the possibility of quotes around the sheetname
    """
    last_exclam = location.rfind('!')

    if remove_quotes and location[0] == "'":
        return (location[1:last_exclam - 1], location[last_exclam + 1:])

    return (location[:last_exclam], location[last_exclam + 1:])


def add_quotes_sheet_name(original_sheet_name: str) -> str:
    """
    If there are special characters in the string, add quotes around it
    for both old and new names unless there are already quotes around it.
    Return the quoted string
    """
    special = '.?!,:;!@#$%^&*()-_ '

    if original_sheet_name[0] == "'":
        return original_sheet_name

    contains_special = False
    for cc in original_sheet_name:
        if cc in special:
            contains_special = True

    if contains_special:
        original_sheet_name = "'" + original_sheet_name + "'"

    return original_sheet_name


def check_cell_inbounds(col: int, row: int) -> bool:
    """
    Checks whether cell is inbounds. Returns boolean.
    """
    return (1 <= row <= MAX_ROW_NUM) and (1 <= col <= MAX_COL_NUM)


def get_target_area(start_location: str, end_location: str) -> list[(int, int)]:
    """
    Get all cells in the rectangle indicated by the start and end
    locations (inclusive). Note that start_location and end_location
    are not necessarily the top left corner and the bottom right
    corner respectively. Returns a list of cells as tuples and the
    number of columns and rows in the target area.
    """
    # Get the start and end columns and rows as numbers
    start_col, start_row, _, _ = get_cell_location(start_location)

    end_col, end_row, _, _ = get_cell_location(end_location)

    # Columns increase left to right
    left_col = start_col
    right_col = end_col
    if left_col > end_col:
        left_col = end_col
        right_col = start_col

    # Rows increase top to bottom
    top_row = start_row
    bot_row = end_row
    if top_row > end_row:
        top_row = end_row
        bot_row = start_row

    # Iterate through range and list all cells
    target_area = []
    for ii in range(left_col, right_col + 1):
        for jj in range(top_row, bot_row + 1):
            target_area.append((ii, jj))

    return target_area, right_col + 1 - left_col, bot_row + 1 - top_row


def get_cell_location(location: str) -> tuple[int, int, bool, bool]:
    """
    Gets location of the cell from a string and checks that it is inbounds.
    Returns int, int
    """

    cell_location = location.lower()
    col, row, _ = re.split(r'(\d+)', cell_location)

    is_absolute_col = (col[0] == '$')
    if is_absolute_col:
        col = col[1:]

    is_absolute_row = (col[-1] == '$')
    if is_absolute_row:
        col = col[:-1]

    # Get column number and make sure its inbounds
    col = col_to_num(col)

    # Check row inbounds
    if int(row) > MAX_ROW_NUM or int(row) < 1:
        raise ValueError("Cell Reference Out of Bounds")

    return (col, int(row), is_absolute_col, is_absolute_row)


def to_boolean(value) -> bool:
    """Converts values to booleans"""
    if isinstance(value, str):
        value = value.lower()
        if value == 'true':
            return True

        if value == 'false':
            return False

        return ValueError("String is not a boolean!")

    if isinstance(value, decimal.Decimal):
        return value != 0

    if isinstance(value, bool):
        return value

    if value is None:
        return False

    raise TypeError("Wrong Type to convert to boolean")


def arr_to_boolean(arr) -> list:
    """Converts array of values to booleans"""
    for ii in range(len(arr)):  # pylint: disable=C0200
        arr[ii] = to_boolean(arr[ii])
    return arr


def to_number(val) -> decimal.Decimal:
    """Converts value to number"""
    if isinstance(val, str):
        val = val.strip('"')
    if val is None:
        return decimal.Decimal(0)
    return decimal.Decimal(val)


def arr_to_num(arr) -> list:  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
    ret = []
    for element in arr:
        if element is not None:
            ret.append(to_number(element))
    return ret


def organize_range(start, end):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
    (start_col, start_row, _, _) = get_cell_location(start)
    (end_col, end_row, _, _) = get_cell_location(end)

    col1, col2 = 0, 0
    if start_col < end_col:
        col1 = start_col
        col2 = end_col
    else:
        col1 = end_col
        col2 = start_col

    row1, row2 = 0, 0
    if start_row < end_row:
        row1 = start_row
        row2 = end_row
    else:
        row1 = end_row
        row2 = start_row

    return col1, row1, col2, row2


def cell_range(start, end):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
    col1, row1, col2, row2 = organize_range(start, end)

    ret = []
    for row in range(row1, row2 + 1):
        temp = []
        for col in range(col1, col2 + 1):
            temp.append((num_to_col(col) + str(row)).upper())
        ret.append(temp)

    return ret


def condense_params(params):  # TODO write docstring esmir/jon  # pylint: disable=C0116,W0511
    new_params = []
    for element in params:
        if isinstance(element, list):
            for row in element:
                for column in row:
                    new_params.append(column)
        else:
            new_params.append(element)
    return new_params
