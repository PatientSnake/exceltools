"""
exceltools - providing more user-friendly access to the pywin32 library
=============================================================================

exceltools is a Python module acting as a friendlier interface to the pywin32 library which
in itself is an API to the Windows COM client API.

exceltools does not provide the full functionality of pywin32, it only seeks to simplify some commonly used code.

exceltools is intended to work alongside pandas and numpy and aids in creating and
populating spreadsheets programmatically.

"""
import os
import re
import ast
import sys
import shutil
import warnings
import datetime as dt
from time import sleep
from pathlib import Path

# External Dependencies
import pythoncom
import numpy as np
import pandas as pd
import pandas.api.types as types
from win32com import client
from win32com.client import constants as c

# Local modules
import exceltools.errors as err
from range import Range
from column import Column
from cell import CellReference
from utils import col2num, num2col, excel_date, rgb2hex


class ExcelSpreadSheet:
    """
    A class built to simplify and streamline working with the win32client library.
    Example usage involves opening an existing workbook and saving a new copy without changing the original.
    New workbooks can also be created and originals can be overwritten.
    Example Usage:
        excel = ExcelSpreadSheet()
        excel.open("C:/Users/generic_user/Documents/master_file.xlsx")
        excel.write_dataframe(data, sheet="Sheet 1", start_col=1, start_row=2, headers=True)
        excel.write_cell("SomeString", sheet=1, row=1, col="A")
        excel.save_xlsx("C:/Users/generic_user/Documents/new_file.xlsx")
        excel.close(save_changes=False)
    """

    def __init__(self):
        global client
        try:
            self.excel = client.gencache.EnsureDispatch("Excel.Application")
        except Exception:
            # Remove cache and try again.
            module_list = [m.__name__ for m in sys.modules.values()]
            for module in module_list:
                if re.match(r"win32com\.gen_py\..+", module):
                    del sys.modules[module]
            shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
            from win32com import client
            self.excel = client.gencache.EnsureDispatch("Excel.Application")

        self.wb = None
        self._wb_open = 0
        self.active_sheet = None
        self.sheet_names = []
        self.null_arg = pythoncom.Empty
        self._wb_open = 0
        self.format_args = {
            "Condition": {
                "logic": "logic_dict[logic]",
                "value": "value",
                "value2": "value2"
            },
            "Format": {
                "interior_colour": "Interior.Color = self.rgb2hex(kwargs['interior_colour'])",
                "number_format": "NumberFormat = kwargs['number_format']",
                "bold": "Font.Bold = kwargs['bold']",
                "font_colour": "Font.Color = self.rgb2hex(kwargs['font_colour'])",
                "font_size": "Font.Size = kwargs['font_size']",
                "font_name": "Font.Name = kwargs['font_name']",
                "orientation": "Orientation = kwargs['orientation']",
                "underline": "Font.Underline = kwargs['underline']",
                "merge": "MergeCells = kwargs['merge']",
                "wrap_text": "WrapText = kwargs['wrap_text']",
                "h_align": "HorizontalAlignment = kwargs['h_align']",
                "v_align": "VerticalAlignment = kwargs['v_align']",
                "border_left": {
                    "line_style": "Borders(c.xlEdgeLeft).LineStyle = kwargs['border_left']['line_style']",
                    "weight": "Borders(c.xlEdgeLeft).Weight = kwargs['border_left']['weight']",
                    "colour": "Borders(c.xlEdgeLeft).Color = self.rgb2hex(kwargs['border_left']['colour'])",
                },
                "border_right": {
                    "line_style": "Borders(c.xlEdgeRight).LineStyle = kwargs['border_right']['line_style']",
                    "weight": "Borders(c.xlEdgeRight).Weight = kwargs['border_right']['weight']",
                    "colour": "Borders(c.xlEdgeRight).Color = self.rgb2hex(kwargs['border_right']['colour'])",
                },
                "border_top": {
                    "line_style": "Borders(c.xlEdgeTop).LineStyle = kwargs['border_top']['line_style']",
                    "weight": "Borders(c.xlEdgeTop).Weight = kwargs['border_top']['weight']",
                    "colour": "Borders(c.xlEdgeTop).Color = self.rgb2hex(kwargs['border_top']['colour'])",
                },
                "border_bot": {
                    "line_style": "Borders(c.xlEdgeBottom).LineStyle = kwargs['border_bot']['line_style']",
                    "weight": "Borders(c.xlEdgeBottom).Weight = kwargs['border_bot']['weight']",
                    "colour": "Borders(c.xlEdgeBottom).Color = self.rgb2hex(kwargs['border_bot']['colour'])",
                },
                "border_inside_h": {
                    "line_style": "Borders(c.xlInsideHorizontal).LineStyle = kwargs['border_inside_h']['line_style']",
                    "weight": "Borders(c.xlInsideHorizontal).Weight = kwargs['border_inside_h']['weight']",
                    "colour": "Borders(c.xlInsideHorizontal).Color = self.rgb2hex(kwargs['border_inside_h']['colour'])",
                },
                "border_inside_v": {
                    "line_style": "Borders(c.xlInsideVertical).LineStyle = kwargs['border_inside_v']['line_style']",
                    "weight": "Borders(c.xlInsideVertical).Weight = kwargs['border_inside_v']['weight']",
                    "colour": "Borders(c.xlInsideVertical).Color = self.rgb2hex(kwargs['border_inside_v']['colour'])",
                }
            }
        }

    @staticmethod
    def col2num(col_str: str) -> int:
        """
        Convert an Excel column string to an integer -> "A" == 1, "AA" == 27 e.t.c.
        """
        return col2num(col_str)

    @staticmethod
    def num2col(col_int: int) -> str:
        """
        Convert an Excel column index to a string -> 1 == "A", 27 == "AA" e.t.c.
        """
        return num2col(col_int)

    @staticmethod
    def rgb2hex(rgb: list | tuple) -> int:
        """
        Excel expects a hex value in order to fill cells
        This function allows you to supply standard RGB values to be converted to hex.
        """
        return rgb2hex(rgb)

    @staticmethod
    def excel_date(date: pd.Series | dt.datetime | dt.date) -> float:
        """
        Convert a datetime.datetime or pandas.Series object into an Excel date float
        """
        return excel_date(date)

    def _validate_workbook(self):
        """
        Ensure the current workbook is open and valid
        """
        if self._wb_open == 0:
            raise err.NoWorkbookError()

    def _validate_worksheet(self, sheet):
        """
        Make sure the sheet supplied is valid for the current open workbook
        """
        if isinstance(sheet, str):
            if sheet not in self.sheet_names:
                raise err.InvalidSheetError(f"A sheet with the name {sheet} does not exist")
        elif isinstance(sheet, int):
            if len(self.sheet_names) < sheet:
                raise err.InvalidSheetError(f"Invalid Sheet Index. Sheet index {sheet} is out of bounds.")

    def _cleanse_data(self, data):
        """
        Excel will print np.Nan as 65535.
        This function aims to cleanse any representations of NULL so that they print as expected to Excel.
        At this stage we also attempt to convert datetimes to a numeric value used by Excel.
        """
        if isinstance(data, pd.DataFrame):
            for column in data:

                _dtype = data[column].dtype

                if types.is_numeric_dtype(_dtype):
                    data.loc[:, column] = data[column].fillna(0)

                if types.is_string_dtype(_dtype):
                    data.loc[:, column] = data[column].fillna("")

                if types.is_datetime64_any_dtype(_dtype):
                    data.loc[:, column] = self.excel_date(data[column])

        elif isinstance(data, (pd.Series, list)):
            _dtype = pd.Series(data).dtype

            if types.is_numeric_dtype(_dtype):
                data = data.fillna(0)

            elif types.is_string_dtype(_dtype):
                data = data.fillna("")

            elif types.is_datetime64_any_dtype(_dtype):
                data = self.excel_date(data)

        return data

    def get_format_args(self):
        return list(self.format_args["Format"].keys())

    def open(self, file: str | Path):
        """
        Open an Excel workbook. If the workbook is already open, no action is taken.
        If the workbook does not exist, a new one is created
        args:
            file : The name of the workbook to be open.
        """
        if self._wb_open == 1:
            raise Exception("Only one active workbook can be open at a time, please close the current workbook")

        file = os.path.normpath(Path(file))
        if not os.path.isfile(file):
            self.wb = self.excel.Workbooks.Add()
            self.wb.SaveAs(file)
            self.wb = self.excel.Workbooks.Open(file)
        else:
            try:
                self.wb = self.excel.Workbooks(file)
            except Exception:
                try:
                    self.wb = self.excel.Workbooks.Open(file)
                except Exception:
                    # Remove cache and try again.
                    module_list = [m.__name__ for m in sys.modules.values()]
                    for module in module_list:
                        if re.match(r"win32com\.gen_py\..+", module):
                            del sys.modules[module]
                    shutil.rmtree(os.path.join(os.environ.get("LOCALAPPDATA"), "Temp", "gen_py"))
                    from win32com import client
                    try:
                        self.excel = client.gencache.EnsureDispatch("Excel.Application")
                        self.wb = self.excel.Workbooks.Open(file)
                    except Exception as e:
                        print(str(e))
                        self.wb = None
                        raise Exception from e

        # Wait until application is ready and has opened the file
        sleep(1)

        if self.wb is not None:
            self.excel.Visible = False
            self.excel.Interactive = False
            self.excel.DisplayAlerts = False
            self.excel.ScreenUpdating = False
            self.excel.EnableEvents = False
            self.excel.DisplayStatusBar = False
            self._wb_open = 1
            for sheet in self.wb.Sheets:
                self.sheet_names.append(sheet.Name)

    def write_dataframe(self, data: pd.DataFrame, sheet, cell_ref: str = None, start_row: int = None,
                        start_col: int | str = None, headers: bool = False):
        """
        Writes a pandas dataframe to an Excel worksheet. If the supplied data is not a dataframe the method will error.
        args:
            data : A pandas dataframe
            sheet : the name or index of the sheet to be populated, must be a valid sheet in the active Workbook.
            cell_ref: An Excel cell reference to start writing from - overrides any other arguments supplied
            start_row : the starting row to write data to (default=2)
            start_col : the starting column to write data to (default=1)
            headers  : Boolean flag, if true the column names of the dataframe will be printed (default=False)
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        cell_ref = CellReference(cell_ref, start_col, start_row)
        start_col = cell_ref.column.index
        start_row = cell_ref.row.index

        if not isinstance(data, pd.DataFrame):
            raise ValueError("Data supplied must be a pandas dataframe")

        data = self._cleanse_data(data)

        self.excel.Calculation = c.xlCalculationManual
        sheet = self.wb.Sheets(sheet)

        if sheet.ProtectContents:
            raise err.ProtectedSheetError(sheet.Name)

        # Write the data
        if headers:
            headers = np.array(data.columns)
            sheet.Range(sheet.Cells(start_row, start_col),
                        sheet.Cells(start_row, len(headers) + (start_col - 1))).Value = headers
            sheet.Range(sheet.Cells((start_row + 1), start_col),
                        sheet.Cells((len(data) + start_row), (len(headers) + (start_col - 1)))).Value = data.values
        else:
            sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells((len(data) + (start_row - 1)), (
                        len(data.columns) + (start_col - 1)))).Value = data.values

        self.excel.Calculate()
        self.excel.Calculation = c.xlCalculationAutomatic

    def write_cell(self, data: str | int | tuple | list | pd.Series, sheet: str | int,
                   cell_ref: str = None, col: str | int = None, row: int = None):
        """
        Write scalar data to a specific Cell in a workbook. Non-Scalar data will attempt to be coerced into a comma
        separated string.
        A Set cannot be passed to this method as it is un-ordered data.
        row is a row number, col can be supplied as a string Excel Reference i.e. "A" or column index.
        An Error should be returned if the object passed cannot be written.
            args:
                data : Variable to write to the cell
                cell_ref : An Excel cell reference to write to
                x : A row number to write to
                y : A column name or index to write to
                sheet : The sheet name or index to write to
        """

        if isinstance(data, (pd.DataFrame, set)):
            raise ValueError("Data supplied cannot be a pandas dataframe or set")
        self._validate_workbook()
        self._validate_worksheet(sheet)
        cell_ref = CellReference(cell_ref, col, row)
        col = cell_ref.column.index
        row = cell_ref.row.index

        if isinstance(data, tuple):
            data = str(data).lstrip("(").rstrip(")")
        elif isinstance(data, list):
            data = str(data).lstrip("[").rstrip("]")
        elif isinstance(data, pd.Series):
            data = data.to_string(index=False).strip().replace("\n", ",")

        data = self._cleanse_data(data)

        try:
            str(data)
        except ValueError:
            raise ValueError("Data could not be coerced to string - try supplying scalar, list, tuples or a series")

        sheet = self.wb.Sheets(sheet)

        if sheet.ProtectContents:
            raise err.ProtectedSheetError(sheet.Name)

        sheet.Cells(row, col).Value = data

    def write_row(self, data: pd.Series | list | tuple, sheet: str | int, cell_ref: str = None,
                  start_col: str | int = None, start_row: int = None, end_col: str | int = None):
        """
        Write list-like data to a specific Range in a workbook.
        Data structures will be coerced into a series.
        start_row is a row number, start_col can be supplied as a string Excel Reference i.e. "A" or column index.
        end_col is optional - a warning will be returned if the series length does not match the range supplied.
        An Error should be returned if the object passed cannot be written.
            args:
                data : Variable to write to the range
                sheet : The sheet name or index to write to
                cell_ref : An Excel cell reference to write from
                start_col : A column name or index to write to
                start_row : A row number to write to
                end_col : A column name or index to truncate the data to
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        cell_ref = CellReference(cell_ref, start_col, start_row)
        start_col = cell_ref.column.index
        start_row = cell_ref.row.index

        if not isinstance(data, (pd.Series, list, tuple)):
            raise ValueError("Data supplied must be a list-like structure")

        data = self._cleanse_data(data)

        try:
            data_series = pd.Series(data)
        except ValueError as e:
            print(str(e) + "\nData could not be coerced to Series")
            raise Exception from e

        if end_col is not None:
            end_col = Column(end_col)
            if end_col.index != len(data_series):
                warnings.warn("\nObject supplied differs in length to supplied range!\nExcess data will be truncated.",
                              UserWarning)
                if end_col.index < len(data_series):
                    data_series = data_series[:end_col.index]
        else:
            end_col = Column(len(data_series) + (start_col - 1))

        sheet = self.wb.Sheets(sheet)

        if sheet.ProtectContents:
            raise err.ProtectedSheetError(sheet.Name)

        sheet.Range(sheet.Cells(start_row, start_col), sheet.Cells(start_row, end_col.index)).Value = data_series.values

    def delete_sheet(self, sheet: str | int):
        """
        Delete a worksheet
            args:
            sheet : The sheet to delete
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        self.wb.Sheets(sheet).Delete()
        # Reset Sheets
        self.sheet_names = []
        for sheet in self.wb.Sheets:
            self.sheet_names.append(sheet.Name)

    def set_sheet_visibility(self, sheet: str | int, visibility: str | int):
        """
        Set a worksheets visibility.
            args:
            sheet : The sheet to change the visibility of
            visibility : The level of visibility to set. Provide either a string or int value 
                {"visible": -1, "hidden": 1, "very hidden": 2}
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        visibility_values = {"visible": -1, "hidden": 1, "very hidden": 2}

        if (isinstance(visibility, str) and visibility.lower() not in visibility_values or
                isinstance(visibility, int) and visibility not in [visibility_values.values()]):
            raise ValueError("Visibility value supplied is invalid, please provide a valid string or int value. "
                             + str(visibility_values))

        if visibility in visibility_values:
            visibility = visibility_values[visibility]

        self.wb.Sheets(sheet).Visible = visibility

    def protect_sheet(self, sheet: str | int, password: str = None, draw_objects: bool = True, contents: bool = True,
                      scenarios: bool = True, allow_sort: bool = False, allow_filter: bool = False,
                      enable_selection: bool = True):
        """
        Protect a worksheet
        args:
            sheet : The name/index of the worksheet to protect
            password : A password to lock the sheets with (Optional)
            draw_objects: Protect shapes (Optional: Default=True)
            contents: Protect contents (Optional: Default=True)
            scenarios: Protect scenarios (Optional: Default=True)
            allow_sort: Allow user to sort (Optional: Default=False)
            allow_filter: Allow user to filter (Optional: Default=False)
            enable_selection: Set as false to disable user selecting cells (Optional: Default=True)
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        if self.wb.Sheets(sheet).ProtectContents:
            warnings.warn("\nSheet {} is already protected.".format(sheet), UserWarning)
        else:
            if not enable_selection:
                self.wb.Sheets(sheet).EnableSelection = c.xlNoSelection

            self.wb.Sheets(sheet).Protect(
                password, draw_objects, contents, scenarios,
                False, False, False, False, False, False, False, False, False,  # Unimplemented positional args
                allow_filter, allow_sort, False
            )

    def unprotect_sheet(self, sheet: str | int, password: str = None):
        """
        Unprotect a worksheet
        args:
            sheet : The name/index of the worksheet to unprotect
            password : A password to unlock the sheets with (Optional)
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)

        if not self.wb.Sheets(sheet).ProtectContents:
            warnings.warn("\nSheet {} is not protected.".format(sheet), UserWarning)
        else:
            self.wb.Sheets(sheet).Unprotect(password)

    def protect_workbook(self, password: str = None):
        """
        Protect the current workbook and it's structure
        args:
            password : A password to lock the workbook with (Optional)
        """
        self._validate_workbook()
        self.wb.Protect(password, True)

    def unprotect_workbook(self, password: str = None):
        """
        Unprotect the current workbook and it's structure
        args:
            password : A password to unlock the workbook with (Optional)
        """
        self._validate_workbook()
        self.wb.Unprotect(password)

    def get_sheet_names(self):
        """
        Return a list of worksheet names
        """
        self._validate_workbook()

        return self.sheet_names

    def refresh_all(self):
        """
        Refresh all workbook connections and pivot tables
        """
        self._validate_workbook()

        # Refreshes the DB connections
        self.wb.RefreshAll()

        # Refresh all pivots - should be lighter than a second full refresh
        for sheet in self.wb.WorkSheets:
            for pivot in sheet.PivotTables:
                pivot.RefreshTable()
                pivot.Update()

    def read_dataframe(self, sheet: str | int, header: bool = True, excel_range: str = None,
                       start_col: str | int = None, start_row: int = None,
                       end_col: str | int = None, end_row: int = None):
        """
        Reads in a range of an Excel spreadsheet and attempts to return a pandas dataframe object.
        args:
            sheet: The sheet name/index to read from
            header: Does this range include column headers? Default == True
            range: An Excel range to read e.g. "A1:B5" - Supply this instead of start and end row values
            start_row: The starting row to read from
            end_row: The final row to read from
            start_col: The starting column to read from
            end_col: The final column to read from
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = Range(excel_range, start_row, end_row, start_col, end_col)

        data = self.wb.Sheets(sheet).Range(excel_range.range_reference).Value

        if hasattr(data, "__len__") and not isinstance(data, str):
            if header is True:
                data = pd.DataFrame(list(data[1:]), columns=data[0])
            else:
                data = pd.DataFrame(list(data))
        else:
            data = pd.DataFrame([data])

        return data

    def read_cell(self, sheet: str | int, cell_ref: str = None, col: str | int = None, row: int = None):
        """
        Reads in a range of an Excel spreadsheet and attempts to return a pandas dataframe object.
        args:
            sheet: The sheet name/index to read from
            cell_ref : An Excel cell reference to write from
            row : A row number to read from
            col : A column name or index to read from
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        cell_ref = CellReference(cell_ref, row, col)
        col = cell_ref.column.index
        row = cell_ref.row.index

        sheet = self.wb.Sheets(sheet)

        data = sheet.Range(sheet.Cells(row, col), sheet.Cells(row, col)).Value

        return data

    def conditional_formatting(self, sheet: str | int, condition: str = "cell_value", excel_range: str = None,
                               start_col: str | int = None, start_row: int = None, end_col: str | int = None,
                               end_row: int = None, **kwargs):
        """
        Add conditional formatting to a range. Current implementation means the formatting cannot be removed or modified once added.
        This is OK as it is not intended for use in interactive sessions at the moment.
        args:
            sheet: The sheet to add formatting to.
            excel_range: The range to format
            start_col, end_col, start_row, end_row: Integers representing the co-ordinates of a range
            condition: Specifies whether the conditional format is based on a cell value or an expression. See https://docs.microsoft.com/en-us/office/vba/api/excel.xlformatconditiontype
            logic: The conditional format operator
            value: The value or formula associated with the condition i.e. Cells in range less than "0.99"
            value2: Only used when a 'between' logical operator is passed, i.e. Cells in range between 0.80 and "0.99"
            **kwargs: snake_case formatting arguments i.e. font_colour, font_size, merge, bold, interior_colour ... (see self.get_format_args() for valid values)
        """
        arg_list = [i for keys in self.format_args.values() for i in keys]
        for k in kwargs.keys():
            if k not in arg_list:
                warnings.warn("Unknown parameter {!r} ignored.".format(k))
                del kwargs[k]
        for i in [keys for keys in self.format_args["Condition"].keys()]:  # Add required values
            if i not in kwargs.keys():
                try:
                    kwargs[i] = ast.literal_eval(i)
                except NameError:
                    kwargs[i] = self.null_arg
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = Range(excel_range, start_col, start_row, end_col, end_row)

        condition_dict = {
            "above_average": c.xlAboveAverageCondition,
            "is_blank": c.xlBlanksCondition,
            "cell_value": c.xlCellValue,
            "color_scale": c.xlColorScale,
            "data_bar": 4,
            "is_error": c.xlErrorsCondition,
            "expression": c.xlExpression,
            "icon_set": 6,
            "no_blanks": c.xlNoBlanksCondition,
            "no_errors": c.xlNoErrorsCondition,
            "text": c.xlTextString,
            "time_period": c.xlTimePeriod,
            "top_ten": c.xlTop10,
            "unique": c.xlUniqueValues
        }

        logic_dict = {
            "between": c.xlBetween,
            "equal_to": c.xlEqual,
            "greater_than": c.xlGreater,
            "greater_equal": c.xlGreaterEqual,
            "less_than": c.xlLess,
            "less_equal": c.xlLessEqual,
            "not_between": c.xlNotBetween,
            "not_equal": c.xlNotEqual,
            self.null_arg: self.null_arg
        }

        try:
            logic_dict[kwargs["logic"]]
        except KeyError:
            raise ValueError("Invalid 'logic' value supplied.")

        try:
            condition_dict[condition]
        except KeyError:
            raise ValueError("Invalid 'condition' value supplied.")

        wb_range = self.wb.Sheets(sheet).Range(excel_range.range_reference)

        _format = wb_range.FormatConditions.Add(Type=condition_dict[condition], Operator=logic_dict[kwargs["logic"]],
                                                Formula1=kwargs["value"], Formula2=kwargs["value2"])

        # Apply the actual formatting
        for key in kwargs.keys():
            if key in self.format_args["Format"]:
                if "border_" in key:
                    for k in kwargs[key].keys():
                        exec("_format.{}".format(self.format_args["Format"][key][k]))
                else:
                    exec("_format.{}".format(self.format_args["Format"][key]))

    def format_range(self, sheet: str | int, excel_range: str = None, start_col: str | int = None,
                     start_row: int = None, end_col: str | int = None, end_row: int = None, **kwargs):
        """
        Add formatting to a range.
        args:
            sheet: The sheet to add formatting to.
            excel_range: The range to format
            start_col, end_col, start_row, end_row: Integers representing the co-ordinates of a range
            **kwargs: snake_case'd formatting arguments i.e. font_colour, font_size, merge, bold, interior_colour ...
                      (see self.get_format_args for valid values)
        """
        arg_list = [keys for keys in self.format_args["Format"]]
        for k in kwargs.keys():
            if k not in arg_list:
                warnings.warn("Unknown parameter {!r} ignored.".format(k))
                del kwargs[k]
        self._validate_workbook()
        self._validate_worksheet(sheet)
        excel_range = Range(excel_range, start_col, start_row, end_col, end_row)
        wb_range = self.wb.Sheets(sheet).Range(excel_range.range_reference)

        for key in kwargs.keys():
            if key in self.format_args["Format"]:
                if "border_" in key:
                    for k in kwargs[key].keys():
                        exec("wb_range.{}".format(self.format_args["Format"][key][k]))
                else:
                    exec("wb_range.{}".format(self.format_args["Format"][key]))

    def reset_cursor(self, sheet: str | int = 1, cell_ref: str = "A1"):
        """
        Move the cursor to cell A1 on all sheets - useful to use before saving
        This is added to make reports open nicely
        args:
            sheet     : The sheet to open on after saving
            cell_ref  : The cell to select on each sheet on open
        """
        self._validate_workbook()
        self._validate_worksheet(sheet)
        cell_ref = CellReference(cell_ref)

        self.excel.EnableEvents = True

        for ws in self.wb.Sheets:
            if ws.Visible == -1 and ws.EnableSelection == c.xlNoRestrictions:
                ws.Activate()
                ws.Range(cell_ref.reference).Select()
                ws.Range(cell_ref.reference).Activate()
                self.excel.ActiveWindow.ScrollRow = 1
                self.excel.ActiveWindow.ScrollColumn = 1

        self.wb.Sheets(sheet).Activate()
        self.excel.EnableEvents = False

    def save_xlsx(self, outfile: str | Path):
        """
        Save the active workbook to a new .xlsx file.
        args:
            outfile : A valid Windows path to a file to export to.
        """
        self._validate_workbook()

        outfile = os.path.normpath(Path(outfile))
        self.wb.SaveAs(outfile, ConflictResolution=c.xlLocalSessionChanges)
        self.wb.Saved = True

    def save_pdf(self, outfile: str | Path, sheet: str | int = None):
        """
        Save the active workbook to a new .pdf file.
        You can specify a particular worksheet or list of worksheets to export.
        args:
            outfile : A valid Windows path to a file to export to, must have .pdf extension.
            sheet : specify a single worksheet to export by name or index (Optional)
        """
        self._validate_workbook()
        if sheet is not None:
            self._validate_worksheet(sheet)

        outfile = os.path.normpath(Path(outfile))
        if not outfile.endswith(".pdf"):
            raise ValueError("outfile argument must have a .pdf file extension")

        if sheet is None:
            self.wb.ExportAsFixedFormat(0, outfile)
        else:
            self.wb.Sheets(sheet).ExportAsFixedFormat(0, outfile)

    def close(self, save_changes: bool = False):
        """
        Close the current active workbook and the current Excel session after 1 second.
        args
            save_changes : Bool - If True save changes to the active workbook before closing (Defaults to False) (Optional)
        """
        sleep(1)
        if self._wb_open == 1:
            self.wb.Close(save_changes)
        self.excel.Quit()
        self.sheet_names = []
        self._wb_open = 0
        self.wb = None
