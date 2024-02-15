# Excel_Automation_with_Python

* This is a compilation of projects, with randomized data, where Python is
being used to automate work flows within a preexisting Excel file.
* The scripts are for Windows systems with Excel installed
* If you don't have Python installed, I recommend installing the 64-bit
Anaconda distribution with Python 3.xx from: https://www.anaconda.com/download/
* These scripts will not support legacy Python

## Why Would Anyone Do This

Many of us work in environments where the exclusive consumption of data
is via Excel.  I image most, if not everything done in Excel, can be
more easily automated in Pandas.  When I write scripts for Excel, all of
the equations are in Excel (e.g. `=max(A1:A30)`) not just some value.

## Why Don't I Use an Existing Python Package for Excel

There are a number of Python packages for working in Excel:
* openpyxl
* xlrd
* xlwt
* http://www.python-excel.org/

Most of the packages work with constraints; all of the Excel functions aren't
available, or the package may not work with a preexisting file, so I use:

```python
import win32com.client as win32


excel = win32.gencache.EnsureDispatch('Excel.Application')
```

The com object exposes the full functionality of Excel.

