# coding=utf-8
"""
❯ time echo 'xlsx/SOX Controls Testing Template.xlsx' | python sheet_names_openpyxl.py
xlsx/SOX Controls Testing Template.xlsx >
/Users/patkujawa/.virtualenvs/xl/lib/python2.7/site-packages/openpyxl/workbook/names/named_range.py:125: UserWarning: Discarded range with reserved name
  warnings.warn("Discarded range with reserved name")
[u'Interim Testing', u'Year End Testing']
echo 'xlsx/SOX Controls Testing Template.xlsx'  0.00s user 0.00s system 41% cpu 0.002 total
python sheet_names_openpyxl.py  0.24s user 0.11s system 10% cpu 3.424 total

> \ls -1 **/*.xls? | python sheet_names_openpyxl.py --verbosity=info
> find . -print | python sheet_names_openpyxl.py --verbosity=info
> ls -d -1 **/*.*
"""
from sheet_names_api import main


def _openpyxl(filepath):
    import openpyxl
    # https://openpyxl.readthedocs.org/en/2.3.3/optimized.html
    try:
        wb = openpyxl.load_workbook(filepath, read_only=True)
    # InvalidFileException: openpyxl does not support .txt file format, please check you can open it with Excel first. Supported formats are: .xlsx,.xlsm,.xltx,.xltm
    # InvalidFileException: openpyxl does not support the old .xls file format, please use xlrd to read this file, or convert it to the more recent .xlsx file format.
    # BadZipfile: File is not a zip file
        return wb.get_sheet_names()
    except Exception as e:
        # Ignore unsupported files
        if 'openpyxl does not support' in repr(e):
            pass
        # if 'openpyxl does not support the old .xls file format' in repr(e):
        #     pass

if __name__ == '__main__':
    main(_openpyxl)
