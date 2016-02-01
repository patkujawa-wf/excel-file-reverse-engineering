# coding=utf-8
"""
> time \ls -1 **/*.xlsx | python sheet_names_xlrd.py

> for fname in **/*.xlsx; do time echo $fname | python sheet_names_xlrd.py; done

❯ time echo 'xlsx/SOX Controls Testing Template.xlsx' | python sheet_names_xlrd.py
[u'Interim Testing', u'Year End Testing']
echo 'xlsx/SOX Controls Testing Template.xlsx'  0.00s user 0.00s system 31% cpu 0.003 total
python sheet_names_xlrd.py  62.72s user 0.87s system 99% cpu 1:03.79 total

INFO:sheet_names_api:Slow (took longer than 0.5 seconds) files:
{
  "0.5802590847015381": "xlsm/Outline Check.xlsm",
  "0.7123560905456543": "xlsx/SOX Failure Listing Status.xlsx",
  "65.0460250377655": "xlsx/SOX Controls Testing Template.xlsx",
  "0.87471604347229": "xlsx/-hp8gt.xlsx",
  "1.2309041023254395": "xlsx/SOX Testing Status.xlsx",
  "1.5334298610687256": "xlsm/Compare XML.xlsm"
}
"""
from sheet_names_api import main


def _xlrd(filepath):
    import xlrd
    # https://secure.simplistix.co.uk/svn/xlrd/trunk/xlrd/doc/xlrd.html?p=4966
    # with xlrd.open_workbook(filepath, on_demand=True, ragged_rows=True) as wb:
    #     sheet_names = wb.sheet_names()
    #     return sheet_names

    # How about with memory mapping? Nope, blows up on both xls and xlsx
    import contextlib
    import mmap
    import os
    length = 2**10 * 4
    # length = 0  # whole file
    with open(filepath, 'rb') as f:
        # mmap throws if length is larger than file size
        length = min(os.path.getsize(filepath), length)
        with contextlib.closing(mmap.mmap(f.fileno(), length, access=mmap.ACCESS_READ)) as m,\
             xlrd.open_workbook(on_demand=True, file_contents=m) as wb:
            sheet_names = wb.sheet_names()
            return sheet_names

if __name__ == '__main__':
    main(_xlrd, ['xls/SMITH 2014 TRIP-Master List.xls'])
