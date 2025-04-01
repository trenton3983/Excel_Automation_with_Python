![repo-image](excel_automation_repo_image.png)

# Excel Automation with Python

This repository contains a collection of scripts that automate workflows in existing Excel files using Python. The data in these examples is randomized, and the scripts are designed specifically for **Windows systems with Microsoft Excel installed**.

## Prerequisites

- Python 3 (Legacy Python is **not supported**)
- Windows OS
- Microsoft Excel
- [Anaconda (64-bit)](https://www.anaconda.com/download/) is recommended for managing Python environments and packages

## Why Automate Excel with Python?

Many workplaces still rely heavily on Excel for data analysis and reporting. While Python libraries like Pandas offer more scalable and flexible solutions, integrating automation directly into Excel workflows can:

- Improve productivity by reducing repetitive tasks
- Preserve familiar Excel interfaces while adding automation
- Allow use of native Excel formulas like `=MAX(A1:A30)` within automated processes

## Why Not Use Standard Excel Libraries?

Popular Excel libraries include:

- [`openpyxl`](https://openpyxl.readthedocs.io/en/stable/)
- [`xlrd`](https://xlrd.readthedocs.io/en/latest/)
- [`xlwt`](https://xlwt.readthedocs.io/)
- [python-excel.org](http://www.python-excel.org/)

While useful, these libraries have limitations:
- Limited support for formulas and complex formatting
- Inconsistent handling of existing Excel files
- Not all features of Excel are accessible

Instead, this project uses the `win32com` library to access Excel via the COM API:

```python
import win32com.client as win32

excel = win32.gencache.EnsureDispatch('Excel.Application')
```

This approach provides full control over the Excel application and allows direct manipulation of workbooks, formulas, and UI features.

