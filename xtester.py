'''
Author: Abhishek Munagekar
xtester : Python2
This is a sample scipt demonstrating using of easyxlsxwriter
Which aims at keeping your python loops clean
by provinding python style access to an excel sheet
'''

import xlsxwriter
from easyxlsxwriter.misc import extension_chk
from easyxlsxwriter.sheetwriter import SheetWriter


headings = ['category', 'net', 'inp']
data0 = ['10', '20']  # Category  = X
data1 = [5, 35]  # Network Output = y1
data2 = [2, 4]  # Netowrk Input	= y2
data = [data0, data1, data2]

title = 'testgraph'
worksheetname = 'testsheet'
xaxis = '10mul'
yaxis = 'numeric'
'''
easyx.addworksheet('testsheet')
easyx.line_graph(headings, data, 'testsheet', title, xaxis, yaxis)
easyx.line_graph(headings, data, 'testsheet', title, xaxis, yaxis)
easyx.close()
'''

bookname = 'testbook.xlsx'
assert (extension_chk(bookname, 'xlsx') is True), 'Invalid extension'
book = xlsxwriter.Workbook(bookname)
writer = SheetWriter(book, 'NewSheet')
writer.writetable('head', data, orientation='col')
writer.line_graph(headings, data, title, xaxis, yaxis)
writer.writetable('head', data, orientation='col')
book.close()
