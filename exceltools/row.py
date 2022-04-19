"""
A module used to define the Row class
"""


class Row:
    """
    An Excel row
    """
    def __init__(self, row: int):
        self.index = None

        self._validate_row(row)

    def __repr__(self):
        return f"Row(index={self.index!r})"

    def __str__(self):
        return str(self.index)

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

    def _validate_row(self, row: int):
        """
        Ensures the value supplied is a valid Excel row number
        """
        if row is None:
            raise ValueError("Row index cannot be \"None\"")

        try:
            row = int(row)
        except ValueError as e:
            raise ValueError("Could not coerce row value to integer") from e

        if row < 1:
            raise ValueError("Row must be a positive integer greater than 0")

        self.index = row
