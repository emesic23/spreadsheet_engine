# flake8: noqa # pylint: disable-all
"""Initialize workbook"""

__all__ = ['Workbook', 'CellError', 'CellErrorType']
from ._workbook import Workbook, CellError, CellErrorType

from ._version import version
