# exceltools   
## providing more user-friendly access to the pywin32 library


exceltools is a Python module acting as a friendlier interface to the pywin32 library which in itself is an API 
to the Windows COM client API. exceltools does not provide the full functionality of pywin32, it only seeks to simplify 
some commonly used code.  
exceltools is intended to work alongside pandas and numpy and aids in creating and
populating spreadsheets programmatically.

## Sample Usage

```python

import exceltools
import pandas as pd

data = pd.read_csv("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/iris.csv")

# Open an existing workbook / Create a new one
excel = exceltools.ExcelSpreadSheet()
excel.open("C:/Users/generic_user/Documents/master_file.xlsx")

# Write data
excel.write_dataframe(data, sheet=1, start_col=1, start_row=2, headers=True)
excel.write_cell("SomeString", sheet=1, row=1, col="A")

# Apply formatting
excel.format_range(sheet=1, excel_range="A1:F1", interior_colour=(255, 0, 0))
excel.conditional_formatting(sheet=1, excel_range="A3:A10", logic="equal_to", value=5, interior_colour=(125, 125, 125), font_colour=(255, 255, 255))

# Protect worksheet
print(excel.get_sheet_names())
excel.protect_sheet(sheet=1, password="P@ssW0rd")

# Save and close
excel.save_xlsx("C:/Users/generic_user/Documents/new_file.xlsx")
excel.close(save_changes=False)
```

# Utility Functions
Excel tools comes with some basic utility functions to make working with Excel a little easier, these are outlined below.

 ----
## col2num
Converts an Excel column reference, to the relevant integer
```python
from exceltools import col2num

col2num("A")
# returns 1

col2num("XA")
# returns 625
```
## num2col
Converts an integer to an Excel column reference
```python
from exceltools import num2col

num2col(1)
# returns "A"

num2col(625)
# returns "XA"
```
## excel_date
Returns a python datetime or pandas Series object as an Excel date value as a float.
Excels epoch is notable 30-12-1899
```python
import datetime
import pandas as pd
from exceltools import excel_date

excel_date(datetime.date(2020, 1, 1))
# returns 43831.0

excel_date(pd.Series([pd.Timestamp(1999,6,26), pd.Timestamp(2020, 1, 1)]))
# returns
# 0    36337.0
# 1    43831.0
# dtype: float64
```
## rgb2hex
I found it easier to work with RGB values when picking colours, so this function converts an RGB colour tuple to an int
expected by pywin32 Excel calls.
```python
from exceltools import rgb2hex

rgb2hex((255, 0, 120))
# returns 7864575
```