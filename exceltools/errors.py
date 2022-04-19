"""
Module used to handle custom error classes
"""
class Error(Exception):
    """Base Class for Exceptions"""
    def __init__(self, msg, *args):
        super().__init__()
        self.what = msg.format(*args)

    def __str__(self):
        return self.what


class NoWorkbookError(Exception):
    """Raised when there is no Excel Workbook Open/Active"""
    def __init__(self):
        super().__init__()
        self.what = "There are no active workbooks open"

    def __str__(self):
        return self.what


class InvalidSheetError(Error):
    """Raised when a worksheet does not exist/is invalid"""
    def __init__(self, msg, *args):
        super().__init__(msg, *args)


class ProtectedSheetError(InvalidSheetError):
    """Raised when a sheet is protected"""

    msg = "The sheet '{0}' is protected, please unprotect before attempting to write to it"

    def __init__(self, sheet_name):
        super().__init__(ProtectedSheetError.msg, sheet_name)


class InvalidCellRefError(Error):
    """Raised when a cell reference is invalid"""


class InvalidRangeError(Error):
    """Raised when a range is invalid"""
