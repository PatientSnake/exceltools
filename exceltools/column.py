"""
A module to define the Column class
"""
import re

from exceltools import col2num, num2col


class Column:
    """
    An Excel Column
    """
    max_column_index = 18278

    def __init__(self, column: str | int):
        self.index = None
        self.col_reference = None

        self._validate_column(column)

    def __repr__(self):
        return f"Column(index={self.index!r}, col_reference={self.col_reference!r})"

    def __str__(self):
        return self.col_reference

    def __int__(self):
        return self.index

    def __eq__(self, other):
        return self.index == int(other)

    def __lt__(self, other):
        return self.index < int(other)

    def __gt__(self, other):
        return self.index > int(other)

    def __le__(self, other):
        return self.index >= int(other)

    def __ge__(self, other):
        return self.index <= int(other)

    def _validate_column(self, col: str | int):
        """
        Checks that a column reference supplied is valid, and returns it if true.
        String references such as "AB" are returned as integers.
        """
        if col is None:
            raise ValueError("Column reference cannot be \"None\"")

        if isinstance(col, str):
            if re.search(r"[^a-zA-Z0-9]", col):
                raise ValueError("Column reference must only contain alphanumeric characters"
                                 ", invalid column reference supplied")
            if len(col) > 3:
                raise ValueError("String must be no more than 3 characters")
            col = col2num(col)

        try:
            int(col)
        except ValueError as e:
            raise ValueError("Column reference could not be coerced to integer") from e

        if col > Column.max_column_index:
            idx = Column.max_column_index
            ref = num2col(Column.max_column_index)
            raise ValueError(f"Column reference is too large, {idx}/\"{ref}\" is the maximum width accepted")

        self.index = col
        self.col_reference = num2col(col)
