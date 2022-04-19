import datetime as dt
from pandas import Series, Timestamp


def col2num(col_str: str) -> int:
    """
    Convert an Excel column reference to an integer
    e.g. "A" = 1, "B" = 2 e.t.c.
    """
    if not isinstance(col_str, str):
        raise ValueError("Invalid data type supplied. Must supply a scalar string")
    expn = 0
    col_num = 0
    for char in reversed(col_str):
        col_num += (ord(char) - ord("A") + 1) * (26 ** expn)
        expn += 1
    return col_num


def num2col(col_int: int) -> str:
    """
    Convert an Excel column index to a string
    e.g. 1 == "A", 27 == "AA" e.t.c.
    """
    if not isinstance(col_int, int):
        raise ValueError("Invalid data type supplied. Must supply an integer")
    col_str = ""
    while col_int > 0:
        col_int, remainder = divmod(col_int - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str


def rgb2hex(rgb: list | tuple) -> int:
    """
    Excel expects a hex value in order to fill cells
    This function allows you to supply standard RGB values to be converted to hex.
    """

    if not isinstance(rgb, (tuple, list)):
        raise TypeError("Argument supplied must be a tuple or list of RGB values")
    bgr = (rgb[2], rgb[1], rgb[0])
    str_value = "%02x%02x%02x" % bgr
    hexcode = int(str_value, 16)
    return hexcode


def excel_date(date1: Series | dt.datetime | dt.date) -> float:
    """
    Convert a datetime.datetime or pandas.Series object into an Excel date float
    """
    if isinstance(date1, (dt.datetime, dt.date)):
        if isinstance(date1, dt.date):
            date1 = dt.datetime.combine(date1, dt.datetime.min.time())
        temp = dt.datetime(1899, 12, 30)  # Excels epoch. Note, not 31st Dec but 30th
        delta = date1 - temp
        return float(delta.days) + (float(delta.seconds) / 86400)
    elif isinstance(date1, Series):
        temp = Timestamp(dt.datetime(1899, 12, 30))
        delta = date1 - temp
        return delta.dt.days + (delta.dt.seconds / 86400)
    else:
        raise TypeError("Must supply datetime, date or pd.Series")
