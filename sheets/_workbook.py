# pylint: disable=C0302
"""
Caltech CS130 - Winter 2022

File containing workbook class
"""

import json
import re
from typing import Callable, Iterable, Optional, TextIO, Tuple, List
import decimal
import uuid
import lark

from ._utils import is_number, is_error, is_string,  \
    add_quotes_sheet_name, get_cell_location, \
    get_target_area, num_to_col, \
    check_cell_inbounds
from ._cell import Cell, CellReference, CellType, CellError, CellErrorType, SortingCellRow
from ._sheet import Sheet
from ._expressionhandler import ExpressionHandler
from ._referencegetter import ReferenceGetter
from ._renamehandler import RenameHandler
from ._reconstructor import ContentReconstructor


class Workbook:  # TODO fix too many instance attributes # pylint: disable=R0902,W0511
    """
    A workbook containing zero or more named spreadsheets.

    Attributes:
       sheets                A dictionary mapping lowercase sheet names to
                             dictionaries mapping row/cell keys to the cell
       references_to         A dictionary mapping a key of lowercase sheet name
                             and the row/cell key to the list of cells that
                             cell references
       references_from       A dictionary mapping a key of lowercase sheet name
                             and the row/cell key to the list of cells that
                             cell is referenced by
       sheet_names           A list of sheet names in order, in original casing
       sheet_ids             A dictionary mapping lowercase sheet names to their ids
       cell_location_to_ref  Location of cell to cell object dictionary
       parser                A lark parser to parse formulas

    Public Methods:
       num_sheets            Returns a int of the number of sheets in the workbook.
       list_sheets           Returns a list of spreadsheet names with original casing.
       get_sheet_extent      Returns the extent of the sheet as a tuple of
                             (number rows, number columns).

       new_sheet             Adds a new sheet to the workbook. Returns the sheet
                             number and name in original casing.
       del_sheet             Deletes the sheet with the case-insensitive name.
       rename_sheet          Renames an existing sheet. Updates all cell references.
       get_sheet             Returns the specified sheet.

       get_cell_contents     Returns the contents of the specified cell as a string or
                             None if the cell is empty.
       get_cell_value        Returns the evaluated value of the specified cell as a string
                             or None if the cell is empty.
       set_cell_contents     Sets the contents of a cell.
       move_cells            Moves cells in a source area to a new location
       copy_cells            Copies cells in a source area to a new location

       load_workbook         Loads the workbook from a JSON file.
       save_workbook         Saves the workbook as a JSON file.

    Private Methods:
       __visit               Kosaraju Cycle Detection
       __handle_circ_ref     Sets cells within circular reference created to be circular
                             reference errors
       __update_cells        Update cells in topological order
       __topological_sort    Topological sort
       __split_sheet_cell    Splits a string of SheetName!CellLoc into (SheetName, CellLoc)
                             accounting for the possibility of quotes around the sheetname
       __check_sheet_name    Checks whether passed in sheet name is valid. If the name is not
                             specified, then a unique sheet name is auto-generated.
       __set_cell_contents   Sets the contents of a cell with optional immediate notifications.
                             Returns the changed cells.
       __move_copy_cells     Performs a move or a copy of a source area to a target area
    """

    def __init__(self):
        # Initializes a new empty workbook.

        # Dictionary of dictionaries mapping lowercase sheetname, row/col to cell value
        self.sheets = {}

        # Dictionaries of cell references
        self.references_to = {}
        self.referenced_from = {}

        # List of sheet names for ordering in original casing
        self.sheet_names = []

        # Mapping of sheet names to ids
        self.sheet_ids = {}

        # Mapping of cell locations to cell objects
        self.cell_location_to_ref = {}

        # Lark parse
        self.parser = lark.Lark.open('sheets/formulas.lark', start='formula', maybe_placeholders=False)

        self.valid_char = set('.?,:;!@#$%^&*()-_\
                               abcdefghijklmnopqrstuvwxyz\
                               ABCDEFGHIJKLMNOPQRSTUVWXYZ\
                               1234567890 ')

        self.notifiers = []

    # PUBLIC METHODS

    def num_sheets(self) -> int:
        """Returns the number of spreadsheets in the workbook."""
        return len(self.sheet_names)

    def list_sheets(self) -> list[str]:
        """
        Returns a list of the spreadsheet names in the workbook, with the
        capitalization specified at creation, and in the order that the
        sheets appear within the workbook.
        """
        return self.sheet_names.copy()

    def new_sheet(self, sheet_name: Optional[str] = None) -> Tuple[int, str]:
        """
        Adds a new sheet to the workbook with the specified name if given.
        Raises a ValueError if the name is not unique. If the name is not
        specified, then a unique sheet name is auto-generated.

        Returns a tuple of the sheet number and sheet name in original casing
        """

        # Check sheet name is valid or auto-generates a new one
        sheet_name = self.__check_sheet_name(sheet_name)

        # Add sheet to workbook
        cur_id = str(uuid.uuid4())
        self.sheet_names.append(sheet_name)
        self.sheet_ids[sheet_name.lower()] = cur_id
        self.sheets[cur_id] = Sheet(cur_id, sheet_name)

        changed_cells = []

        # Update cells that reference the new sheet (resurrection)
        for ref_cell, _ in self.referenced_from.items():
            if sheet_name.lower() == ref_cell.get_sheet():
                update_order = self.__topological_sort(ref_cell)
                changed_cells = self.__update_cells(update_order, changed_cells=changed_cells)

        # Notify cell updates
        if changed_cells != []:
            self.__notify(changed_cells)

        return (len(self.sheet_names) - 1, sheet_name)

    def del_sheet(self, sheet_name: str) -> None:
        """
        Delete the spreadsheet with the specified case-insensitive name. If
        the specified sheet name is not found, a KeyError is raised.

        Returns None
        """

        # Check that sheet exists
        if not self.get_sheet(sheet_name):
            raise KeyError("This Sheet does not exist")

        # Delete sheet references
        sheet = self.get_sheet(sheet_name)
        self.sheet_names.remove(sheet_name)
        self.sheet_ids.pop(sheet_name.lower())
        self.sheets.pop(sheet.sheet_id)

        ref_to_remove = []
        changed_cells = []

        # Get list of references to remove from references_to
        for ref_cell, _ in self.references_to.items():
            if sheet_name.lower() == ref_cell.get_sheet():
                ref_to_remove.append(ref_cell)

        # Delete references needed to be removed
        for ref_cell in ref_to_remove:
            if ref_cell in self.references_to:
                del self.references_to[ref_cell]
            if ref_cell.get_cell_loc_without_refs() in self.cell_location_to_ref:
                del self.cell_location_to_ref[ref_cell.get_cell_loc_without_refs()]

        # Update cells that reference deleted sheet and notify
        for ref_cell, _ in self.referenced_from.items():
            if sheet_name.lower() == ref_cell.get_sheet():
                update_order = self.__topological_sort(ref_cell)
                changed_cells = self.__update_cells(update_order, changed_cells)

        # Notify changed cells
        if changed_cells != []:
            self.__notify(changed_cells)

        # Clear references from that are coming from deleted sheet
        to_del = []
        for cell, references in self.referenced_from.items():
            self.referenced_from[cell] = list(filter(lambda val: sheet_name.lower() != val.get_sheet(), references))
            if not self.referenced_from[cell]:
                to_del.append(cell)

        for cell in to_del:
            del self.referenced_from[cell]

    def rename_sheet(self, sheet_name: str, new_sheet_name: str) -> None:  # TODO fix too many local vars and branches # pylint: disable=R0914,W0511,R0912,R0915
        """
        Rename the specified sheet to the new sheet name. All cell formulas
        that referenced the original sheet name are updated to reference the
        new sheet name (using the same case as the new sheet name, and
        single-quotes iff necessary). The case of the new_sheet_name is preserved
        by the workbook. If the sheet_name is not found, a KeyError is raised.
        If the new_sheet_name is an empty string or is otherwise invalid, a
        ValueError is raised.
        """

        # Check that sheet exists
        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet does not exist")

        # Check sheet name is valid or auto-generates a new one
        new_sheet_name = self.__check_sheet_name(new_sheet_name, auto_gen=False, check_unique=False)

        # Add sheet to workbook
        sheet_loc = 0
        for ii, temp_name in enumerate(self.sheet_names):
            if self.sheet_names[ii].lower() == sheet_name.lower():
                sheet_loc = ii
                break

        sheet_name = sheet_name.strip("'")
        # Copy new sheet with new sheet name
        if new_sheet_name.lower() != sheet_name.lower():
            self.new_sheet(new_sheet_name)

            og_id = self.sheet_ids[sheet_name.lower()]
            # new_id = self.sheet_ids[new_sheet_name.lower()]
            og_sheet = self.sheets[og_id]
            # new_sheet = self.sheets[new_id]

            for cur_cell in og_sheet.entries:
                self.set_cell_contents(new_sheet_name,
                                       cur_cell,
                                       og_sheet.entries[cur_cell].contents)

            self.del_sheet(sheet_name)
            self.sheet_names.remove(new_sheet_name)
        else:
            self.sheet_names.remove(sheet_name)
        self.sheet_names.insert(sheet_loc, new_sheet_name)

        # Get references needed to be changed
        to_change = []
        for ref_cell in self.referenced_from:
            if sheet_name.lower() == ref_cell.get_sheet():
                to_change.append(ref_cell)
        # Change references in referenced_from to match new name
        for ref_cell in to_change:
            new_ref = CellReference(new_sheet_name, ref_cell.get_col(), ref_cell.get_row())
            self.referenced_from[new_ref] = self.referenced_from.pop(ref_cell)
            for ii in range(len(self.referenced_from[new_ref])):
                ref = self.referenced_from[new_ref][ii]
                if ref.get_sheet() == sheet_name.lower():
                    self.referenced_from[new_ref][ii] = CellReference(new_sheet_name, ref.get_col(), ref.get_row())

        # Add quotes to old sheetname so that replacement works
        quoted_old_sheet_name = add_quotes_sheet_name(sheet_name)
        quoted_new_sheet_name = add_quotes_sheet_name(new_sheet_name)

        # Iterate through to get cells needed to update references_to

        to_set = []
        for ref_cell, references in self.referenced_from.items():
            if new_sheet_name.lower() == ref_cell.get_sheet():
                for temp_loc in references:
                    temp_name, temp_cell_loc = temp_loc.get_sheet(), temp_loc.get_rowcol()
                    cur_cell = self.get_sheet(temp_name).entries[temp_cell_loc.lower()]
                    RenameHandler(cur_cell.parse_tree, quoted_old_sheet_name, quoted_new_sheet_name).visit(cur_cell.parse_tree)
                    recon = ContentReconstructor()
                    recon.visit(cur_cell.parse_tree)
                    content = recon.content
                    to_set.append((temp_name, temp_cell_loc, content))
        to_set = set(to_set)

        # Set new cell contents (effectively updating references_to)
        # Also handles resurrecting sheets
        for (temp_name, temp_cell_loc, content) in to_set:
            self.set_cell_contents(temp_name, temp_cell_loc, content)

    def move_sheet(self, sheet_name: str, index: int) -> None:
        """
        Move the specified sheet to the specified index in the workbook's
        ordered sequence of sheets.  The index can range from 0 to
        workbook.num_sheets() - 1.  The index is interpreted as if the
        specified sheet were removed from the list of sheets, and then
        re-inserted at the specified index.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        If the specified sheet name is not found, a KeyError is raised.

        If the index is outside the valid range, an IndexError is raised.
        """

        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet does not exist")

        if index >= len(self.sheet_names) or index < 0:
            raise IndexError("The index is outside the valid range")

        for temp_name in self.sheet_names:
            if temp_name.lower() == sheet_name.lower():
                self.sheet_names.remove(temp_name)
                break

        self.sheet_names.insert(index, sheet_name)

    def copy_sheet(self, sheet_name: str) -> Tuple[int, str]:
        """
        Make a copy of the specified sheet, storing the copy at the end of the
        workbook's sequence of sheets.  The copy's name is generated by
        appending "_1", "_2", ... to the original sheet's name (preserving the
        original sheet name's case), incrementing the number until a unique
        name is found.  As usual, "uniqueness" is determined in a
        case-insensitive manner.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        The copy should be added to the end of the sequence of sheets in the
        workbook.  Like new_sheet(), this function returns a tuple with two
        elements:  (0-based index of copy in workbook, copy sheet name).  This
        allows the function to report the new sheet's name and index in the
        sequence of sheets.

        If the specified sheet name is not found, a KeyError is raised.
        """

        # Check if sheet exists
        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet does not exist")

        # Get next available sheet name (_x appended)
        sheet_names = set(map(str.lower, self.sheet_names))
        ii = 1
        while sheet_name.lower() + "_" + str(ii) in sheet_names:
            ii += 1
        new_sheet_name = sheet_name + "_" + str(ii)

        # Create new sheet
        ret = self.new_sheet(new_sheet_name)

        # Set cells of new sheet to match old sheet cells
        og_id = self.sheet_ids[sheet_name.lower()]
        # new_id = self.sheet_ids[new_sheet_name.lower()]
        og_sheet = self.sheets[og_id]
        # new_sheet = self.sheets[new_id]
        for cur_cell in og_sheet.entries:
            self.set_cell_contents(new_sheet_name, cur_cell, og_sheet.entries[cur_cell].contents)

        return ret

    def get_sheet(self, sheet_name: str) -> any:
        """Returns the the specified case-insensitive sheet."""
        return self.sheets[self.sheet_ids[sheet_name.lower()]]

    def get_sheet_extent(self, sheet_name: str) -> Tuple[int, int]:
        """
        Return a tuple (num-cols, num-rows) indicating the current extent of
        the specified spreadsheet. If the specified case-insensitive sheet
        name is not found, a KeyError is raised.
        """
        return self.get_sheet(sheet_name).get_extent()

    def set_cell_contents(self, sheet_name: str, location: str,
                          contents: Optional[str]) -> None:
        """
        Set the contents of the specified cell on the specified sheet. If the
        specified sheet name is not found, a KeyError is raised. If the cell
        location is invalid, a ValueError is raised. A cell may be set to "empty"
        by specifying a contents of None, "", or a string of whitespace. The
        leading and trailing whitespace around the contents are removed before
        storing the cells. If the cell contents is a formula and it is invalid,
        the cell will be set to the appropriate CellError.
        """
        self.__set_cell_contents(sheet_name, location, contents, notify_immediately=True)

    def get_cell_contents(self,
                          sheet_name: str,
                          location: str) -> Optional[str]:
        """
        Returns the contents of the specified cell on the specified
        case-insensitive sheet name. If the specified sheet name is not found,
        a KeyError is raised. If the cell location is invalid, a ValueError is
        raised. Empty cell values are indicated by a return of None
        """

        # Check exists
        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet Does Not Exist")

        get_cell_location(location)

        location = location.lower()
        temp_index = self.sheet_ids[sheet_name.lower()]
        try:
            return self.sheets[temp_index].get_cell(location).get_contents()
        except KeyError:
            return None

    def get_cell_value(self, sheet_name: str, location: str) -> any:
        """
        Returns the evaluated value of the specified cell on the specified
        case-insensitive sheet. If the specified sheet name is not found, a
        KeyError is raised. If the cell location is invalid, a ValueError is
        raised. The value of empty cells is None.
        """

        # Check exists
        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet Does Not Exist")

        get_cell_location(location)

        location = location.lower()
        temp_index = self.sheet_ids[sheet_name.lower()]

        try:
            return self.sheets[temp_index].get_cell(location).get_value()
        except KeyError:
            return None

    def move_cells(self, sheet_name: str, start_location: str,  # pylint: disable=R0913
                   end_location: str, to_location: str,
                   to_sheet: Optional[str] = None) -> None:
        """
        Move cells from one location to another, possibly moving them to
        another sheet.  All formulas in the area being moved will also have
        all relative and mixed cell-references updated by the relative
        distance each formula is being copied.

        Cells in the source area (that are not also in the target area) will
        become empty due to the move operation.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be moved.  The to_location specifies the
        top-left corner of the target area to move the cells to.

        Both corners are included in the area being moved; for example,
        copying cells A1-A3 to B1 would be done by passing
        start_location="A1", end_location="A3", and to_location="B1".

        The start_location value does not necessarily have to be the top left
        corner of the area to move, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to move.

        This function works correctly even when the destination area overlaps
        the source area.

        The sheet name matches are case-insensitive; the text must match but
        the case does not have to.

        If to_sheet is None then the cells are being moved to another
        location within the source sheet.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.

        If the target area would extend outside the valid area of the
        spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        no changes are made to the spreadsheet.

        If a formula being moved contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal in the formula.
        """
        self.__move_copy_cells(sheet_name, start_location, end_location, to_location, to_sheet, is_move=True)

    def copy_cells(self, sheet_name: str, start_location: str,  # pylint: disable=R0913
                   end_location: str, to_location: str,
                   to_sheet: Optional[str] = None) -> None:
        """
        Copy cells from one location to another, possibly copying them to
        another sheet.  All formulas in the area being copied will also have
        all relative and mixed cell-references updated by the relative
        distance each formula is being copied.

        Cells in the source area (that are not also in the target area) are
        left unchanged by the copy operation.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be copied.  The to_location specifies the
        top-left corner of the target area to copy the cells to.

        Both corners are included in the area being copied; for example,
        copying cells A1-A3 to B1 would be done by passing
        start_location="A1", end_location="A3", and to_location="B1".

        The start_location value does not necessarily have to be the top left
        corner of the area to copy, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to copy.

        This function works correctly even when the destination area overlaps
        the source area.

        The sheet name matches are case-insensitive; the text must match but
        the case does not have to.

        If to_sheet is None then the cells are being copied to another
        location within the source sheet.

        If any specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.

        If the target area would extend outside the valid area of the
        spreadsheet (i.e. beyond cell ZZZZ9999), a ValueError is raised, and
        no changes are made to the spreadsheet.

        If a formula being copied contains a relative or mixed cell-reference
        that will become invalid after updating the cell-reference, then the
        cell-reference is replaced with a #REF! error-literal in the formula.
        """
        self.__move_copy_cells(sheet_name, start_location, end_location, to_location, to_sheet)

    def sort_region(self,  # pylint: disable=R0914,R0912
                    sheet_name: str,
                    start_location: str,
                    end_location: str,
                    sort_cols: List[int]):
        """
        Sort the specified region of a spreadsheet with a stable sort, using
        the specified columns for the comparison.

        The sheet name match is case-insensitive; the text must match but the
        case does not have to.

        The start_location and end_location specify the corners of an area of
        cells in the sheet to be sorted.  Both corners are included in the
        area being sorted; for example, sorting the region including cells B3
        to J12 would be done by specifying start_location="B3" and
        end_location="J12".

        The start_location value does not necessarily have to be the top left
        corner of the area to sort, nor does the end_location value have to be
        the bottom right corner of the area; they are simply two corners of
        the area to sort.

        The sort_cols argument specifies one or more columns to sort on.  Each
        element in the list is the one-based index of a column in the region,
        with 1 being the leftmost column in the region.  A column's index in
        this list may be positive to sort in ascending order, or negative to
        sort in descending order.  For example, to sort the region B3..J12 on
        the first two columns, but with the second column in descending order,
        one would specify sort_cols=[1, -2].

        The sorting implementation is a stable sort:  if two rows compare as
        "equal" based on the sorting columns, then they will appear in the
        final result in the same order as they are at the start.

        If multiple columns are specified, the behavior is as one would
        expect:  the rows are ordered on the first column indicated in
        sort_cols; when multiple rows have the same value for the first
        column, they are then ordered on the second column indicated in
        sort_cols; and so forth.

        No column may be specified twice in sort_cols; e.g. [1, 2, 1] or
        [2, -2] are both invalid specifications.

        The sort_cols list may not be empty.  No index may be 0, or refer
        beyond the right side of the region to be sorted.

        If the specified sheet name is not found, a KeyError is raised.
        If any cell location is invalid, a ValueError is raised.
        If the sort_cols list is invalid in any way, a ValueError is raised.
        """

        # Check valid arguments
        # Check sheet name
        self.get_sheet(sheet_name)

        # Check sort_cols list
        if len(sort_cols) < 1:
            raise ValueError("Sort columns must contain at least one column")

        deduped_sort_cols = set()
        for sort_col in sort_cols:
            if abs(sort_col) in deduped_sort_cols:
                raise ValueError("Sort columns must be unique")
            deduped_sort_cols.add(abs(sort_col))

        if 0 in deduped_sort_cols:
            raise ValueError("Sort columns are 1-indexed, 0 is invalid")

        # Check and get source area
        source_area, num_cols, num_rows = get_target_area(start_location, end_location)

        # Check sort_col list bounds
        if max(deduped_sort_cols) > num_cols:
            raise ValueError("Sort columns out of bounds")

        if len(source_area) == 0:
            raise KeyError("No cells in the target area")

        # Create array of sorting rows
        delta_row_list = source_area[0][1]
        rows = []

        for row in range(0, num_rows):
            rows.append(SortingCellRow(sort_cols, row))

        for cell_loc in source_area:
            curr_cell = self.__get_cell_obj(sheet_name, cell_loc)
            rows[cell_loc[1] - delta_row_list].add_cell(curr_cell, cell_loc)

        # Sorting
        sorted_rows = sorted(rows)

        # Put together array of new cells
        new_cells = []
        for ii in range(0, len(sorted_rows)):  # pylint: disable=C0200
            row = sorted_rows[ii]
            original_row = row.original_row  # Relative to start of source area
            row_diff = ii - original_row

            new_row = original_row + delta_row_list + row_diff

            print(row)
            for s_cell in row.cells:
                new_col = s_cell.original_col

                new_loc_str = num_to_col(new_col) + str(new_row)
                old_loc_str = num_to_col(s_cell.original_col) + str(s_cell.original_row)

                new_contents = None
                if s_cell.cell:
                    new_contents = s_cell.cell.contents

                if new_contents and s_cell.cell.cell_type == CellType.EXPRESSION:
                    RenameHandler(s_cell.cell.parse_tree, cellchangefrom=old_loc_str, coldelta=0, rowdelta=row_diff).visit(s_cell.cell.parse_tree)

                    recon = ContentReconstructor()
                    recon.visit(s_cell.cell.parse_tree)

                    new_contents = recon.content

                new_cells.append([sheet_name, new_loc_str, new_contents])

        # Set new cells
        changed_cells = []
        for new_cell in new_cells:
            changed_cells = changed_cells + self.__set_cell_contents(new_cell[0], new_cell[1], new_cell[2], notify_immediately=False)

        self.__notify(changed_cells)

    def __get_cell_obj(self, sheet_name: str, original_loc: Tuple):
        """Helper function to get cell object for cell at specified location"""
        old_loc_str = num_to_col(original_loc[0]) + str(original_loc[1])
        old_sheet_loc_str = sheet_name.lower() + '!' + old_loc_str.lower()
        if old_sheet_loc_str in self.cell_location_to_ref:
            return self.cell_location_to_ref[old_sheet_loc_str]
        return None

    @staticmethod
    def load_workbook(file: TextIO):
        """Load workbook from a JSON file."""
        wbsheets = {}
        data = json.load(file)
        try:
            wbsheets = data['sheets'].copy()
        except KeyError as error:
            raise KeyError("Missing expected json value") from error
        # file.close()

        name = ""

        # Populate workbook with sheets
        workbook = Workbook()
        if not isinstance(data['sheets'], list):
            raise TypeError("Wrong Type! Sheets is not a list!")
        for sheet in wbsheets:
            try:
                name = sheet['name']
            except KeyError as error:
                raise KeyError("Missing expected json value") from error
            workbook.new_sheet(name)

        # Populate sheets with cells
        for sheet in wbsheets:
            name = sheet['name']
            try:
                cellcontents = sheet['cell-contents']
            except KeyError as error:
                raise KeyError("Missing expected json value") from error
            if not isinstance(cellcontents, dict):
                raise TypeError("Wrong Type! Cell contents is not a list!")
            for referenced_cell, val in cellcontents.items():
                if isinstance(name, str) and isinstance(referenced_cell, str) and isinstance(val, str):
                    workbook.set_cell_contents(name, referenced_cell, val)
                else:
                    raise TypeError("Wrong Type")
        return workbook

    def save_workbook(self, file: TextIO) -> None:
        """Save workbook in a JSON file."""

        jsondict = {"sheets": []}
        for sheet in self.sheets.values():
            sheetdict = {}
            sheetdict['name'] = sheet.name
            cellentries = {}
            for referenced_cell in sheet.entries:
                cellentries[referenced_cell] = self.get_cell_contents(sheet.name, referenced_cell)
            sheetdict['cell-contents'] = cellentries
            jsondict['sheets'].append(sheetdict)
        json.dump(jsondict, file, indent=4)
        # file.close()

    # String literals delay evaluation of the type, workbook called within itself needs to be delayed
    def notify_cells_changed(self,
                             notify_function: Callable[['Workbook',
                                                        Iterable[Tuple[str, str]]],
                                                       None]) -> None:
        """
        Request that all changes to cell values in the workbook are reported
        to the specified notify_function.  The values passed to the notify
        function are the workbook, and an iterable of 2-tuples of strings,
        of the form ([sheet name], [cell location]).  The notify_function is
        expected not to return any value; any return-value will be ignored.

        Multiple notification functions may be registered on the workbook;
        functions will be called in the order that they are registered.

        A given notification function may be registered more than once; it
        will receive each notification as many times as it was registered.

        If the notify_function raises an exception while handling a
        notification, this will not affect workbook calculation updates or
        calls to other notification functions.

        A notification function is expected to not mutate the workbook or
        iterable that it is passed to it.  If a notification function violates
        this requirement, the behavior is undefined.
        """
        self.notifiers.append(notify_function)

    def __notify(self, changed_cells):
        """Add notification functions to notifier list"""
        for notify_function in self.notifiers:
            try:
                notify_function(self, list(set(changed_cells)))
            # This needs to be a broad exception to handle all errors the notify
            # might throw
            except Exception:  # pylint: disable=W0703
                pass

    # PRIVATE METHODS

    # def __visit2(self, cell):
    #     ''' In Progress tarjan's algorithm for cycle detection'''
    #     st = []
    #     result = []
    #     comp = []
    #     visited = {}
    #     stack2 = [(cell, 0, len(visited))]
    #     while stack2:
    #         cur, pi, num = stack2.pop()
    #         if pi == 0:
    #             if cur in visited:
    #                 continue
    #             visited[cur] = num
    #             st.append(cur)
    #         if pi > 0:
    #             if cur in self.references_to:
    #                 visited[cur] = min(visited[cur], visited[self.references_to[cur][pi - 1]])
    #         if cur in self.references_to:
    #             if pi < len(self.references_to[cur]):
    #                 stack2.append((cur, pi + 1, num))
    #                 stack2.append((self.references_to[cur][pi], 0, len(visited)))
    #                 continue
    #         if num == visited[cur]:
    #             while True:
    #                 comp.append(st.pop())
    #                 if cur in self.references_to:
    #                     visited[comp[-1]] = len(self.references_to[cur])
    #                 if comp[-1] == cur:
    #                     break
    #             result.append(comp)
    #     return result[0]

    def __visit(self, cell):  # TODO fix branching  # pylint: disable=R0912,W0511
        """Kosaraju cycle detection"""

        # DFS to build inverse graph
        visited = set()
        processed = set()
        seen = []
        stack_to_visit = [cell]
        ii = 0
        while stack_to_visit:
            cur = stack_to_visit[-1]
            if cur not in visited:
                visited.add(cur)
                if cur in self.references_to:
                    stack_to_visit.extend(self.references_to[cur])
            elif cur not in processed:
                stack_to_visit.pop()
                processed.add(cur)
                seen.append(cur)
            else:
                stack_to_visit.pop()
            ii += 1
        components = [set()]
        component = 0
        interested_component = float('inf')

        # Cycle detection on inverse graph
        while seen:  # TODO fix nested blocks # pylint: disable=R1702,W0511
            cur = seen.pop()
            sub_stack = [cur]
            if cur in visited:
                visited.remove(cur)
                components[component].add(cur)
                while sub_stack:
                    sub_cur, done = sub_stack[-1], True
                    if sub_cur in self.referenced_from:
                        for vv in self.referenced_from[sub_cur]:
                            if vv in visited:
                                visited.remove(vv)
                                done = False
                                sub_stack.append(vv)
                                components[component].add(vv)
                                break
                    if done:
                        sub_stack.pop()
                    if sub_cur == cell:
                        interested_component = component
                component += 1
                components.append(set())

        return components[interested_component]

    def __handle_circ_ref(self, cur, circ_references, references):
        """Sets cells within circular reference created to be circular ref errors"""
        self_referral = False
        if len(references) > 0:
            self_referral = references[0] == cur
        if len(circ_references) > 1 or (len(circ_references) == 1 and self_referral):
            for referenced_cell in circ_references:
                if referenced_cell.get_cell_loc_without_refs() in self.cell_location_to_ref:
                    cell = self.cell_location_to_ref[referenced_cell.get_cell_loc_without_refs()]
                    cell.cell_type = CellErrorType.CIRCULAR_REFERENCE
                    cell.value = CellError(CellErrorType.CIRCULAR_REFERENCE,
                                           "Encountered a Circular Reference")

    def __update_cells(self, update_order, changed_cells=[], circ_references=[]):  # TODO fix danger # pylint: disable=W0102,W0511
        """Update cells in topological order"""
        for referenced_cell in update_order:
            if referenced_cell.get_cell_loc_without_refs() in self.cell_location_to_ref and referenced_cell not in circ_references:
                cell = self.cell_location_to_ref[referenced_cell.get_cell_loc_without_refs()]
                if cell.cell_type == CellType.EXPRESSION or isinstance(cell.cell_type, CellErrorType):
                    old_value = cell.value
                    cell.value = ExpressionHandler(self, cell.sheet).visit(cell.parse_tree)

                    # Check if cell value was updated
                    if old_value != cell.value:
                        changed_cells.append(self.__split_sheet_cell(referenced_cell.get_cell_loc_without_refs()))

                    # Check new value
                    if cell.value is None:
                        cell.value = decimal.Decimal(0)
                    elif isinstance(cell.value, CellError):
                        cell.cell_type = cell.value.get_type()
        return changed_cells

    def __topological_sort(self, referenced_cell):
        """Topological sort"""
        visited = set()
        post_order = []
        stack_to_visit = [referenced_cell]
        while stack_to_visit:
            cur = stack_to_visit[-1]
            if cur not in visited:
                visited.add(cur)

            if cur in self.referenced_from:
                for ref_cell in self.referenced_from[cur]:
                    if ref_cell not in visited:
                        stack_to_visit.append(ref_cell)

            if stack_to_visit[-1] == cur:
                stack_to_visit.pop()
                post_order.append(cur)
        return post_order[::-1]

    def __split_sheet_cell(self, location: str) -> tuple[str, str]:
        """
        Splits a string of SheetName!CellLoc into (SheetName, CellLoc) accounting
        for the possibility of quotes around the sheetname
        """

        last_exclam = location.rfind('!')
        return (self.get_sheet(location[:last_exclam]).name, location[last_exclam + 1:].upper())

    def __check_sheet_name(self, sheet_name: Optional[str], auto_gen=True, check_unique=True) -> str:
        """
        Checks whether passed in sheet name is valid. If the name is not
        specified, then a unique sheet name is auto-generated.

        Returns a sheet name.
        """

        # Check non-empty and unique
        if (sheet_name and sheet_name.strip() == "") or sheet_name == "":
            raise ValueError("Sheet name must be non-empty")

        # Check if string is None
        if not sheet_name and auto_gen:
            sheet_names = set(map(str.lower, self.sheet_names))
            ii = 1
            while "sheet" + str(ii) in sheet_names:
                ii += 1
            sheet_name = "Sheet" + str(ii)
        elif not sheet_name:
            raise ValueError("Sheet name must be not None when not auto-generating")

        # Check if string starts or ends with whitespace
        if sheet_name.strip() != sheet_name:
            raise ValueError("Sheet name must not start or end with whitespace")

        # Check if unique
        if check_unique and str.lower(sheet_name) in self.sheet_ids:
            raise ValueError("Sheet name must be unique")

        for cc in sheet_name:
            if cc not in self.valid_char:
                raise ValueError("Invalid Character in Sheetname")

        return sheet_name

    # TODO fix typing  # pylint:disable=W0511
    def __clean_contents(self, sheet: str, contents: Optional[str]):  # pylint: disable=R0912,R0915
        value = 0
        cell_type = 0
        if contents is None:
            return None, None, None, None

        if contents:
            contents = contents.strip()
        tree = None
        if contents.lower() == "true" or contents.lower() == "false":
            value = contents.lower() == "true"
            cell_type = CellType.BOOLEAN
        elif is_number(contents):  # Check if decimal
            value = decimal.Decimal(contents).normalize()
            cell_type = CellType.NUMBER
        elif is_error(contents):  # Check if error
            match contents.upper():  # pylint: disable=syntax-error  # noqa: E999
                case "#ERROR!":
                    value = CellError(CellErrorType.PARSE_ERROR, "Could not Parse Cell")
                    cell_type = CellErrorType.PARSE_ERROR
                case "#CIRCREF!":
                    value = CellError(CellErrorType.CIRCULAR_REFERENCE, "Encountered a Circular Reference")
                    cell_type = CellErrorType.CIRCULAR_REFERENCE
                case "#REF!":
                    value = CellError(CellErrorType.BAD_REFERENCE, "Encountered an Invalid Reference")
                    cell_type = CellErrorType.BAD_REFERENCE
                case "#NAME?":
                    value = CellError(CellErrorType.BAD_NAME, "Incorrect Formula Name")
                    cell_type = CellErrorType.BAD_NAME
                case "#VALUE!":
                    value = CellError(CellErrorType.TYPE_ERROR, "Incorrect Type")
                    cell_type = CellErrorType.TYPE_ERROR
                case "#DIV/0!":
                    value = CellError(CellErrorType.DIVIDE_BY_ZERO, "Dividing by 0")
                    cell_type = CellErrorType.DIVIDE_BY_ZERO
                case _:
                    value = contents
                    cell_type = CellType.STRING
        elif "#REF!" in contents.upper() and contents[0] != "'":
            value = CellError(CellErrorType.BAD_REFERENCE, "Encountered an Invalid Reference")
            cell_type = CellErrorType.BAD_REFERENCE
        elif is_string(contents):  # Check if string
            value = contents

            if len(value) > 0 and value[0] == "'":
                value = value[1:]
            # elif len(value) > 0 and value[0] == '"':
            #     value = value.strip('"')

            cell_type = CellType.STRING
        elif contents is None:
            value = None
            cell_type = CellType.NUMBER
        else:  # Else, either a formula or may be unparsable
            try:
                tree = self.parser.parse(contents)
                value = ExpressionHandler(self, sheet).visit(tree)
                if not isinstance(value, bool):
                    if value is None:
                        value = decimal.Decimal(0)
                cell_type = CellType.EXPRESSION
            except (lark.exceptions.UnexpectedEOF, lark.exceptions.UnexpectedCharacters) as error:
                if isinstance(error, ValueError):
                    raise error

                if contents == '':
                    value = ''
                    cell_type = CellType.STRING
                else:
                    value = CellError(CellErrorType.PARSE_ERROR, "Could not Parse Cell")
                    cell_type = CellErrorType.PARSE_ERROR
        return contents, value, cell_type, tree

    def __set_cell_contents(self, sheet_name: str, location: str,  # TODO fix too many branches,statements,locals # pylint: disable=R0914,W0511,R0912,R0915,R0913
                            contents: Optional[str], notify_immediately=True,
                            is_movecopy=False, value=None, cell_type=None,
                            tree=None) -> list[(str, str)]:
        """
        Set the contents of the specified cell on the specified sheet. If the
        specified sheet name is not found, a KeyError is raised. If the cell
        location is invalid, a ValueError is raised. A cell may be set to "empty"
        by specifying a contents of None, "", or a string of whitespace. The
        leading and trailing whitespace around the contents are removed before
        storing the cells. If the cell contents is a formula and it is invalid,
        the cell will be set to the appropriate CellError.

        Calls notification function immediately after setting contents if
        notify_immediately is True, otherwise returns the changed cells
        """
        # Find sheet and location of cell and check that it is inbounds
        sheet = self.get_sheet(sheet_name)
        get_cell_location(location)

        prev_cur_cell_value = self.get_cell_value(sheet_name, location)

        # Clean contents
        if not is_movecopy:
            contents, value, cell_type, tree = self.__clean_contents(sheet, contents)

        # Create cell, add to sheet and cell_location_to_ref dictionary
        location = location.lower()
        col, row, _ = re.split(r'(\d+)', location)
        cur_cell_ref = CellReference(sheet_name, col, row)
        cur_cell = Cell(sheet, contents, value, cell_type, tree)
        sheet.entries[location] = cur_cell
        self.cell_location_to_ref[cur_cell_ref.get_cell_loc_without_refs()] = cur_cell

        # Keep track of changed cells
        changed_cells = []
        if cur_cell.value != prev_cur_cell_value:
            changed_cells.append((sheet.name, location.upper()))

        # If it is a formula, populate cell references
        cell_references = []
        if cell_type == CellType.EXPRESSION:
            ref = ReferenceGetter(workbook=self, sheet=sheet)
            ref.visit(self.parser.parse(contents))
            cell_references = ref.references

            if cell_references:
                self.references_to[cur_cell_ref] = cell_references
            else:
                if cur_cell_ref in self.references_to:
                    del self.references_to[cur_cell_ref]

            for referenced_cell in cell_references:
                if referenced_cell in self.referenced_from:
                    if cur_cell_ref not in self.referenced_from[referenced_cell]:
                        self.referenced_from[referenced_cell].append(cur_cell_ref)
                else:
                    self.referenced_from[referenced_cell] = [cur_cell_ref]
        else:
            if cur_cell_ref in self.references_to:
                del self.references_to[cur_cell_ref]

        # Kosaraju for circular references, __topological_sort for update order
        circ_references = self.__visit(cur_cell_ref)
        update_order = self.__topological_sort(cur_cell_ref)

        # Update circular references
        self.__handle_circ_ref(cur_cell_ref, circ_references, cell_references)

        # Update cells in topological order
        self.__update_cells(update_order, changed_cells=changed_cells, circ_references=circ_references)

        # Notify changed cells
        if changed_cells and notify_immediately:
            self.__notify(changed_cells)

        # Set cell to empty if '' is given in the contents
        if contents == '' or contents is None:
            sheet.entries.pop(location)
            self.cell_location_to_ref.pop(cur_cell_ref.get_cell_loc_without_refs())

        return changed_cells

    def __move_copy_cells(self, sheet_name: str, start_location: str,  # pylint: disable=R0913,R0914,R0912
                          end_location: str, to_location: str,
                          to_sheet: Optional[str] = None, is_move: bool = False) -> None:

        # Get sheets (and check that both exist)
        self.get_sheet(sheet_name)

        if to_sheet:
            self.get_sheet(to_sheet)
        else:
            to_sheet = sheet_name

        # Figure out with cells are being grabbed (the source area)
        # and check that all cells are inbounds
        source_area, num_cols, num_rows = get_target_area(start_location, end_location)

        if len(source_area) == 0:
            raise KeyError("No cells in the target area")

        to_col, to_row, _, _ = get_cell_location(to_location)

        # If the top left corner and the sheet names match, no changes made
        top_left = source_area[0]
        if (top_left == (to_col, to_row) and sheet_name == to_sheet):
            return

        col_diff = to_col - top_left[0]
        row_diff = to_row - top_left[1]

        # Check that the target area is completely in bounds
        # Note to location already checked above
        if not (check_cell_inbounds(to_col + num_cols - 1, to_row)
                and check_cell_inbounds(to_col, to_row + num_rows - 1)
                and check_cell_inbounds(to_col + num_cols - 1, to_row + num_rows - 1)):
            raise ValueError("Target area extends out of bounds")

        # Iterate through to find the new cell values and store them
        new_cells = []
        to_delete = []
        for cell in source_area:
            # Get new location
            new_col = cell[0] + col_diff
            new_row = cell[1] + row_diff

            # New location string (just column and row)
            new_loc_str = num_to_col(new_col) + str(new_row)

            # Get old location string
            old_loc_str = num_to_col(cell[0]) + str(cell[1])
            old_sheet_loc_str = sheet_name.lower() + '!' + old_loc_str.lower()

            # Grab old cell contents if not blank
            new_contents = None
            if old_sheet_loc_str in self.cell_location_to_ref:
                cell_obj_original = self.cell_location_to_ref[old_sheet_loc_str]

                cell_obj = cell_obj_original
                if not is_move:  # If we are doing a copy we need to copy the cell
                    cell_obj = cell_obj_original.copy(sheet_name)

                original_contents = cell_obj.contents

                # If contents not blank, delete it
                if original_contents:
                    if not is_move:  # If we are doing a copy we need to copy the cell
                        cell_obj = cell_obj_original.copy(sheet_name)
                        original_contents = cell_obj.contents

                    to_delete.append(cell)

                    # If not a formula, just pass through the contents
                    new_contents = original_contents

                    # If cell is a formula, need to update references
                    if cell_obj.cell_type == CellType.EXPRESSION:
                        RenameHandler(cell_obj.parse_tree, cellchangefrom=old_loc_str, coldelta=col_diff, rowdelta=row_diff).visit(cell_obj.parse_tree)

                        recon = ContentReconstructor()
                        recon.visit(cell_obj.parse_tree)

                        new_contents = recon.content
            new_cells.append([to_sheet, new_loc_str, new_contents])

        changed_cells = []
        if is_move:
            for cell in to_delete:
                old_loc_str = num_to_col(cell[0]) + str(cell[1])
                changed_cells = changed_cells + self.__set_cell_contents(sheet_name, old_loc_str, "", notify_immediately=False, is_movecopy=True)

        # TODO pass in value, type, and tree since content doesn't need to be clean  # pylint: disable=W0511
        for new_cell in new_cells:
            changed_cells = changed_cells + self.__set_cell_contents(new_cell[0], new_cell[1], new_cell[2], notify_immediately=False)

        self.__notify(changed_cells)

    def get_cell_tree(self, sheet_name: str, location: str) -> any:
        """
        Returns the evaluated value of the specified cell on the specified
        case-insensitive sheet. If the specified sheet name is not found, a
        KeyError is raised. If the cell location is invalid, a ValueError is
        raised. The value of empty cells is None.
        """

        # Check exists
        if sheet_name.lower() not in self.sheet_ids:
            raise KeyError("Sheet Does Not Exist")

        get_cell_location(location)

        location = location.lower()
        temp_index = self.sheet_ids[sheet_name.lower()]
        try:
            return self.sheets[temp_index].get_cell(location).get_tree()
        except KeyError:
            return None

    # def __print_references(self):
    #     print("REFERENCED FROM")
    #     for ref_cell, references in self.referenced_from.items():
    #         print(ref_cell)
    #         print("The cell: " + ref_cell.get_cell_loc())
    #         for ref_cell_2 in references:
    #             print("Is referred to from: " + ref_cell_2.get_cell_loc())

    #     print("REFERENCED TO")
    #     for ref_cell, references in self.references_to.items():
    #         print(ref_cell)
    #         print("The cell: " + ref_cell.get_cell_loc())
    #         for ref_cell_2 in references:
    #             print("References: " + ref_cell_2.get_cell_loc())
