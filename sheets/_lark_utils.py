"""
Caltech CS130 - Winter 2022

File containing the expression handler class
"""


from ._cell import CellError, CellErrorType


def match_errors(val):
    """
    Function to match cell value to proper error
    """
    match val:  # pylint: disable=syntax-error  # noqa: E999
        case "#ERROR!":
            return CellError(
                CellErrorType.PARSE_ERROR,
                "Could not Parse Cell")
        case "#CIRCREF!":
            return CellError(
                CellErrorType.CIRCULAR_REFERENCE,
                "Encountered a Circular Reference")
        case "#REF!":
            return CellError(
                CellErrorType.BAD_REFERENCE,
                "Encountered an Invalid Reference")
        case "#NAME?":
            return CellError(
                CellErrorType.BAD_NAME,
                "Incorrect Formula Name")
        case "#VALUE!":
            return CellError(
                CellErrorType.TYPE_ERROR,
                "Incorrect Type")
        case "#DIV/0!":
            return CellError(
                CellErrorType.DIVIDE_BY_ZERO,
                "Dividing by 0")
