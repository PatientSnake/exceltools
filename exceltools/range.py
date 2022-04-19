"""
This module defines the Excel Range class
"""
import re

import exceltools.errors as err
from exceltools.row import Row
from exceltools.column import Column


class Range:
    """
    A class used to define an Excel range reference
    """
    RANGE_REGEX = re.compile(r"(^[a-zA-Z]{1,3})(\d+):([a-zA-Z]{1,3})(\d+$)|(^[a-zA-Z]{1,3}:[a-zA-Z]{1,3}$)")

    def __init__(self, range_reference: str = None, start_column: str | int = None, start_row: int = None,
                 end_column: str | int = None, end_row: int = None):

        self.start_column = None
        self.start_row = None
        self.end_column = None
        self.end_row = None
        self.range_reference = None

        self._validate_range(range_reference, start_column, start_row, end_column, end_row)

    def __repr__(self):
        return f"Range(start_column={self.start_column.col_reference!r}, start_row={self.start_row}, " \
               f"end_column={self.end_column.col_reference!r}, end_row={self.end_row}, "\
               f"range={self.range_reference!r})"

    def __str__(self):
        return self.range_reference

    def _validate_range(self, _range: str, start_col: str | int, start_row: int, end_col: str | int, end_row: int):
        """
        Ensures the range supplied is a valid Excel range -
        returns a string e.g. "A1:B2"
        """
        coords = (start_row, end_row, start_col, end_col)
        if _range is not None and all(coord is None for coord in coords):
            match = re.match(Range.RANGE_REGEX, _range)
            col_1 = match.group(1)
            row_1 = match.group(2)
            col_2 = match.group(3)
            row_2 = match.group(4)
            cols_only = match.group(5)

            if match is None:
                raise err.InvalidRangeError("range must be a valid Excel range string i.e. A1:B3 or A:A. "
                                            "Column references must be 3 chars max.")

            if cols_only is not None:
                self.range_reference = _range
                self.start_column, self.end_column = [Column(i) for i in _range.split(":")]
            else:
                col_1, col_2 = [Column(i) for i in (col_1, col_2)]
                row_1, row_2 = [Row(i) for i in (row_1, row_2)]

                if col_1 > col_2:
                    raise err.InvalidRangeError("Starting column cannot be greater than the ending column!")
                if row_1 > row_2:
                    raise err.InvalidRangeError("Starting row cannot be greater than the ending row!")

                self.start_column = col_1
                self.end_column = col_2
                self.start_row = row_1
                self.end_row = row_2
                self.range_reference = str(col_1) + str(row_1) + ":" + str(col_2) + str(row_2)

        else:
            if any(coord is not None for coord in coords) and any(coord is None for coord in coords):
                raise err.InvalidRangeError("All start and end col/row values must be supplied, "
                                            "only partial values detected.")
            if all(coord is not None for coord in coords) and _range is not None:
                raise err.InvalidRangeError("You cannot supply both an Excel range and start/end values. "
                                            "Please supply one or the other.")

            start_col, end_col = [Column(i) for i in (start_col, end_col)]
            start_row, end_row = [Row(i) for i in (start_row, end_row)]

            if start_col > end_col:
                raise err.InvalidRangeError("Starting column cannot be greater than the ending column!")
            if start_row > end_row:
                raise err.InvalidRangeError("Starting row cannot be greater than the ending row!")

            self.range_reference = str(start_col) + str(start_row) + ":" + str(end_col) + str(end_row)
            self.start_column = start_col
            self.end_column = end_col
            self.start_row = start_row
            self.end_row = end_row
