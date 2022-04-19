# exceltools   
## providing more user-friendly access to the pywin32 library


exceltools is a Python module acting as a friendlier interface to the pywin32 library which in itself is an API 
to the Windows COM client API. exceltools does not provide the full functionality of pywin32, it only seeks to simplify 
some commonly used code.  
exceltools is intended to work alongside pandas and numpy and aids in creating and
populating spreadsheets programmatically.

## Usage

```python

from exceltools import exceltools
import pandas as pd

data = pd.read_csv("https://raw.githubusercontent.com/mwaskom/seaborn-data/master/iris.csv")

# Open an existing workbook / Create a new one
excel = exceltools.ExcelSpreadSheet()
excel.open("C:/Users/daflin/Documents/master_file.xlsx")

# Write data
excel.write_dataframe(data, sheet="Sheet 1", startcol=1, startrow=2, headers=True)
excel.write_cell("SomeString", sheet=1, row=1, col="A")

# Protect worksheet
print(excel.get_sheet_names())
excel.protect_sheet(sheet=1, password="P@ssW0rd")

# Apply formatting
excel.format_range(excel_range="A1:F1", interior_colour=(255, 255, 255))
excel.conditional_formatting(excel_range="A2:A5", logic="equal_to", value=5)

# Save and close
excel.save_xlsx("C:/Users/daflin/Documents/new_file.xlsx")
excel.close(save_changes=False)
```