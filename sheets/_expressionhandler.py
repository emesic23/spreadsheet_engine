"""
Caltech CS130 - Winter 2022

File containing the expression handler class
"""

from functools import reduce

import decimal
import re
import lark

from ._cell import CellType, CellError, CellErrorType
from ._utils import col_to_num, is_string, is_number, to_boolean, arr_to_boolean, cell_range, condense_params, arr_to_num, to_number
from ._lark_utils import match_errors
from ._version import version


def value_fixer(lhs, rhs, default):  # TODO fix branching  # pylint: disable=R0912,W0511
    """
    Check that neither side of the equation is None and
    sets both sides to error (accounting for priority)
    if either side is an error
    """

    if lhs is None or lhs == "":
        lhs = default
    if rhs is None or rhs == "":
        rhs = default

    lerror = False
    rerror = False

    # Check if left side is an error
    if isinstance(lhs, lark.Token):
        if lhs.type == 'ERROR_VALUE':
            lerror = True
    if isinstance(lhs, CellError):
        lerror = True

    # Check if right side is an error
    if isinstance(rhs, lark.Token):
        if rhs.type == 'ERROR_VALUE':
            rerror = True
    if isinstance(rhs, CellError):
        rerror = True

    if isinstance(lhs, bool):
        lhs = bool_value_fixer(lhs, default)
    if isinstance(rhs, bool):
        rhs = bool_value_fixer(rhs, default)

    # If both sides are errors, set error by priority
    if lerror and rerror:
        if lhs.get_type().value < rhs.get_type().value:
            rhs = lhs
        else:
            lhs = rhs
    elif lerror and (not rerror):  # If only the left is error, set right to same error
        rhs = lhs
    elif (not lerror) and rerror:  # If only the right is error, set left to same error
        lhs = rhs

    return lhs, rhs


# TODO write docstring jon/esmir   # pylint: disable=W0511
def bool_value_fixer(val, default):
    """ TODO """
    if isinstance(default, str):
        return str(val).upper()
    if is_number(default):
        return decimal.Decimal(val)
    return val


# TODO write docstring jon/esmir   # pylint: disable=W0511
def comp_value_fixer(lhs, rhs):
    """ TODO """
    if (lhs is None or lhs == "") and (is_string(rhs)):
        lhs, rhs = value_fixer(lhs, rhs, "")
    elif (lhs is None or lhs == "") and (is_number(rhs)):
        lhs, rhs = value_fixer(lhs, rhs, decimal.Decimal(0))
    elif (lhs is None or lhs == "") and isinstance(rhs, bool):
        lhs = False
    elif (lhs is None or lhs == ""):
        lhs = 0

    if (rhs is None or rhs == "") and (is_string(lhs)):
        lhs, rhs = value_fixer(lhs, rhs, "")
    elif (rhs is None or rhs == "") and (is_number(lhs)):
        lhs, rhs = value_fixer(lhs, rhs, decimal.Decimal(0))
    elif (rhs is None or rhs == "") and isinstance(lhs, bool):
        rhs = False
    elif (rhs is None or rhs == ""):
        rhs = 0

    return lhs, rhs


# TODO write docstring jon/esmir   # pylint: disable=W0511
def type_priority(val):
    """ TODO """
    if isinstance(val, bool):
        return 3
    if is_string(val):
        return 2
    return 1


# TODO write docstring jon/esmir   # pylint: disable=W0511
# TODO fix returns   # pylint: disable=W0511
def comp_helper(val1, val2):  # pylint: disable=R0911
    """ TODO """
    if not isinstance(val1, type(val2)):
        priority1 = type_priority(val1)
        priority2 = type_priority(val2)

        if priority1 > priority2:
            return 1
        if priority1 < priority2:
            return -1
        return 0

    if is_string(val1) and is_string(val2):
        if val1.lower() > val2.lower():
            return 1
        if val1.lower() < val2.lower():
            return -1
        return 0

    if val1 > val2:
        return 1

    if val1 < val2:
        return -1
    return 0


# TODO write docstring jon/esmir   # pylint: disable=W0511
def get_first_error(arr) -> CellError:
    """ TODO """
    try:
        for element in arr:
            if isinstance(element, CellError):
                error = element
                return CellError(error.get_type(), error.get_detail())
    except Exception:    # TODO fix exception # pylint:disable=W0511,W0703
        return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
    return ""


class ExpressionHandler(lark.visitors.Interpreter):
    """
    The class to handle parsing

    Attributes:
       workbook       The parent workbook
       sheet          The current sheet

    Methods:
       add_expr       Parse addition and subtraction
       mul_expr       Parse multiplication and division
       concat_expr    Parse concatenation
       unary_op       Parse unary operator (-)
       cell           Parse cell
       number         Parse a number
       parens         Parse equation with parentheses
       string         Parse a string
       error          Parse an error
    """

    def __init__(self, workbook, sheet):
        self.workbook = workbook
        self.sheet = sheet

    def add_expr(self, tree):  # TODO standardize returns # pylint: disable=W0511,R1710,R0911,R0912
        """
        Parse formula with addition or subtraction
        """
        lhs, op, rhs = self.visit_children(tree)
        lhs, rhs = value_fixer(lhs, rhs, 0)
        if isinstance(lhs, CellError):
            return lhs

        if isinstance(lhs, str):
            if lhs[0] == "'":
                try:
                    lhs = decimal.Decimal(lhs[1:])
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
            else:
                try:
                    lhs = decimal.Decimal(lhs)
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")

        if isinstance(rhs, str):
            if rhs[0] == "'":
                try:
                    rhs = decimal.Decimal(rhs[1:]).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
            else:
                try:
                    rhs = decimal.Decimal(rhs).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")

        if op == '+':
            return lhs + rhs
        if op == '-':
            return lhs - rhs

    def mul_expr(self, tree):  # TODO standardize returns # pylint: disable=W0511,R1710,R0911,R0912
        """
        Parse formula with multiplication or division
        """
        lhs, op, rhs = self.visit_children(tree)

        lhs, rhs = value_fixer(lhs, rhs, 0)

        if isinstance(lhs, CellError):
            return lhs
        if isinstance(lhs, str):
            if lhs[0] == "'":
                try:
                    lhs = decimal.Decimal(lhs[1:]).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
            else:
                try:
                    lhs = decimal.Decimal(lhs).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")

        if isinstance(rhs, str):
            if rhs[0] == "'":
                try:
                    rhs = decimal.Decimal(rhs[1:]).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
            else:
                try:
                    rhs = decimal.Decimal(rhs).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")

        if op == '*':
            return lhs * rhs

        if op == '/':
            if rhs == 0:
                return CellError(
                    CellErrorType.DIVIDE_BY_ZERO,
                    "Dividing by 0")
            return lhs / rhs

    def comp_expr(self, tree):  # TODO standardize returns # pylint: disable=W0511,R1710,R0911,R0912
        """
        Parse formula with addition or subtraction
        """
        lhs, op, rhs = self.visit_children(tree)

        lhs, rhs = comp_value_fixer(lhs, rhs)

        if isinstance(lhs, CellError):
            return lhs

        comp = comp_helper(lhs, rhs)

        match op:  # pylint: disable=syntax-error  # noqa: E999
            case "=":
                return comp == 0
            case "==":
                return comp == 0
            case "!=":
                return comp != 0
            case "<>":
                return comp != 0
            case "<=":
                return comp <= 0
            case ">=":
                return comp >= 0
            case "<":
                return comp < 0
            case ">":
                return comp > 0

    def concat_expr(self, tree):
        """
        Parse formula with concat (&)
        """
        lhs, rhs = self.visit_children(tree)

        lhs, rhs = value_fixer(lhs, rhs, '')

        if isinstance(lhs, CellError):
            return lhs

        return str(lhs) + str(rhs)

    def unary_op(self, tree):
        """
        Parse formula with unary ops
        """
        lhs, rhs = self.visit_children(tree)
        lhs, rhs = value_fixer(lhs, rhs, 0)
        if isinstance(lhs, CellError):
            return lhs

        if isinstance(rhs, str):
            if rhs[0] == "'":
                try:
                    rhs = decimal.Decimal(rhs[1:]).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
            else:
                try:
                    rhs = decimal.Decimal(rhs).normalize()
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")

        return -decimal.Decimal(rhs).normalize()

    def cell(self, tree):
        """
        Parse cell
        """
        usedsheet = ""

        try:
            if tree.children[0].type == "SHEET_NAME":
                usedsheet = self.workbook.get_sheet(tree.children[0].value)
            elif tree.children[0].type == "QUOTED_SHEET_NAME":
                usedsheet = self.workbook.get_sheet(tree.children[0].value[1:-1])
            else:
                usedsheet = self.sheet
        except KeyError:
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "Encountered an Invalid Reference")
        rowcol = tree.children[-1].value.lower()
        rowcol = rowcol.replace("$", "")
        col, row, _ = re.split(r'(\d+)', rowcol)
        col = col_to_num(col, throw_errors=False)
        if (int(row) > 9999 or int(row) < 1) or (int(col) > col_to_num('zzzz') or int(col) < 1):
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "Encountered an Invalid Reference")
        location = usedsheet.name + '!' + rowcol
        location = location.lower()
        if location in self.workbook.cell_location_to_ref:
            try:
                cell = usedsheet.get_cell(rowcol)
            except KeyError:
                return decimal.Decimal(0)
            # value = cell.get_value
            match cell.get_type():  # noqa: E999
                case CellType.NUMBER:
                    return decimal.Decimal(cell.get_value()).normalize()
                case _:
                    return cell.get_value()
        return None

    # TODO Fix the error type for not strings, too many params or not enough params    # pylint: disable=W0511
    # TODO write docstring esmir   # pylint: disable=W0511
    def function(self, tree):  # pylint: disable=R0915,R0912,R0911,R0914
        """ TODO """
        children = self.visit_children(tree)
        if len(children) == 2:
            func, params = children
        else:
            return CellError(CellErrorType.TYPE_ERROR, "Incorrect number of parameters")
        func = func.lower()
        match func:
            case "and":
                params = condense_params(params)
                try:
                    params = arr_to_boolean(params)
                except (ValueError, TypeError):
                    return get_first_error(params)
                return bool(reduce(lambda x, y: x & y, params))
            case "or":
                params = condense_params(params)
                try:
                    params = arr_to_boolean(params)
                except (ValueError, TypeError):
                    return get_first_error(params)
                return bool(reduce(lambda x, y: x | y, params))
            case "not":
                try:
                    params = arr_to_boolean(params)
                except (ValueError, TypeError):
                    return get_first_error(params)
                return bool(not params[0])
            case "xor":
                params = condense_params(params)
                try:
                    params = arr_to_boolean(params)
                except (ValueError, TypeError):
                    return get_first_error(params)
                return bool(reduce(lambda x, y: x ^ y, params))
            case "exact":
                if len(params) != 2:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect number of parameters")
                try:
                    val1 = "" if params[0] is None else str(params[0])
                    val2 = "" if params[1] is None else str(params[1])
                except Exception:    # TODO fix exception # pylint:disable=W0511,W0703
                    return get_first_error(params)
                if isinstance(params[0], list) or isinstance(params[1], list):
                    return get_first_error(params)
                return val1 == val2
            case "if":
                if (len(params) != 2) and (len(params) != 3):
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect number of parameters")
                try:
                    cond = to_boolean(params[0])
                except Exception:    # TODO fix exception # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                if cond:
                    return params[1]
                if len(params) == 3:
                    return params[2]
                return False
            case "iferror":  # TODO Check specifics    # pylint: disable=W0511
                if len(params) != 1 and len(params) != 2:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                if isinstance(params[0], list):
                    return get_first_error(params)
                if not isinstance(params[0], CellError):
                    return params[0]
                if len(params) == 2:
                    return params[1]
                return ""
            case "choose":
                idx = params[0]
                try:
                    idx = decimal.Decimal(idx)
                except Exception:    # TODO fix exception  # TODO TYPE_ERROR CELL    # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                if idx <= 0 or idx > len(params):
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                try:
                    ret = params[int(idx)]
                except Exception:    # TODO fix exception # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return ret
            case "isblank":
                if len(params) != 1:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return params[0] is None
            case "iserror":
                if len(params) != 1:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                if isinstance(params, list):
                    return get_first_error(params)
                return isinstance(params[0], CellError)
            case "version":
                return version
            case "indirect":
                if not isinstance(params[0], str):
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                try:
                    cellref = params[0]
                    tree = self.workbook.parser.parse('=' + str(cellref))
                    value = self.visit(tree)
                    if isinstance(value, CellError):
                        return CellError(CellErrorType.BAD_REFERENCE, "Bad Reference!")
                    return value
                except Exception:    # TODO fix exception # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.BAD_REFERENCE, "Bad Reference!")
            case "min":
                params = condense_params(params)
                try:
                    params = arr_to_num(params)
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return min(params)
            case "max":
                params = condense_params(params)
                try:
                    params = arr_to_num(params)
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
                return max(params)
            case "sum":
                params = condense_params(params)
                try:
                    params = arr_to_num(params)
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return sum(params)
            case "average":
                params = condense_params(params)
                try:
                    params = arr_to_num(params)
                except decimal.InvalidOperation:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return sum(params) / len(params)
            case "hlookup":
                if len(params) != 3:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

                key, values, index = params
                # Make sure the index is an integer
                if index != to_number(int(index)):
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

                index = int(index)
                try:
                    for col in range(len(values[0])):
                        if key == values[0][col]:
                            return values[index - 1][col]
                except Exception:  # TODO fix exception # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")
                return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

            case "vlookup":
                if len(params) != 3:
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

                key, values, index = params
                # Make sure the index is an integer
                if index != to_number(int(index)):
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

                index = int(index)
                try:
                    for row in range(len(values)):  # pylint: disable=C0200
                        if key == values[row][0]:
                            return values[row][index - 1]
                except Exception:  # TODO fix exception # pylint:disable=W0511,W0703
                    return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

                return CellError(CellErrorType.TYPE_ERROR, "Invalid Type!")

            case _:
                return CellError(
                    CellErrorType.BAD_NAME,
                    "Invalid Function Name!")

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

        values = [[self.visit(tree) for tree in row] for row in cell_trees]

        return values

    def number(self, tree):
        """
        Parse number
        """
        return decimal.Decimal(tree.children[0]).normalize()

    def parens(self, tree):
        """
        Parse parentheses
        """
        return self.visit(tree.children[0])

    def string(self, tree):
        """
        Parse string. Truncates on either side due to quotes
        """
        return tree.children[0][1:-1]

    def error(self, tree):
        """
        Parse error
        """
        return match_errors(tree.children[0].upper())

    def boolean(self, tree):
        """
        Parse boolean
        """
        return tree.children[0].lower() == "true"
