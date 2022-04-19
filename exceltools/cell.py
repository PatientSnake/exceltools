"""
This module defines the Excel Cell Reference class
"""
import re

import exceltools.errors as err
from exceltools.row import Row
from exceltools.column import Column


class CellReference:
    """
    A class used to define an Excel cell reference
    """
    CELL_REGEX = re.compile(r"^[a-zA-Z]\d+$")

    def __init__(self, cell_ref: str = None, column: str | int = None, row: int = None):
        if isinstance(cell_ref, int) and row is None:
            row = column
            column = cell_ref
            cell_ref = None

        self.column = None
        self.row = None
        self.reference = None

        self._validate_cell_ref(cell_ref, row, column)

    def __repr__(self):
        return f"CellReference(column={self.column.col_reference!r}, row={self.row.index!r}, "\
               f"reference={self.reference!r})"

    def __str__(self):
        return self.reference

    def _validate_cell_ref(self, cell_ref: str, row: int, col: str | int):
        """
        Ensures the cell_ref supplied is a valid Excel cell reference -
        returns a tuple of row and col values to be used.
        """
        if all(value is None for value in (row, col, cell_ref)):
            raise err.InvalidCellRefError("Please supply either a column and row value, "
                                          "or a cell reference e.g. \"A1\".")
        elif (all(value is not None for value in (row, col, cell_ref))
              or (any(value is not None for value in (row, col)) and cell_ref is not None)):
            raise err.InvalidCellRefError("Too many co-ordinates supplied."
                                          " Please supply either a cell reference or separate row and column values")
        elif (all(value is None for value in (row, col)) and cell_ref is not None
              and re.match(CellReference.CELL_REGEX, cell_ref) is None):
            raise err.InvalidCellRefError("Cell reference supplied is invalid.")
        else:
            if cell_ref is None:
                self.row = Row(row)
                self.column = Column(col)
                self.reference = str(self.column) + str(self.row)
            else:
                col, row = cell_ref[0], int(cell_ref[1])
                self.row = Row(row)
                self.column = Column(col)
                self.reference = cell_ref
