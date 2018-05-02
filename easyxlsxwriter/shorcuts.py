# Author : Abhishek Munagekar
# Python 2 : Shortcut.py
# Shortcut will contain simple shortcuts for xl Handling
# Shourcut will contain plain wrappers

import xlsxwriter


# Newbook: Returns a xlsxwriter Workbook
# Args:
# Bookname : Python String
def newbook(bookname):
    return xlsxwriter.Workbook(bookname)
