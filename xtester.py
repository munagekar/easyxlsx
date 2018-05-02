'''
Author: Abhishek Munagekar
xtester : Python2
This is a sample scipt demonstrating using of easyxlsxwriter
Which aims at keeping your python loops clean
by provinding python style access to an excel sheet
'''
from easyxlsxwriter.misc import extension_chk
from easyxlsxwriter.sheetwriter import SheetWriter
from easyxlsxwriter.shorcuts import newbook


headings = ['category', 'net', 'inp']
data0 = ['10', '20']  # Category  = X
data1 = [5, 35]  # Network Output = y1
data2 = [2, 4]  # Netowrk Input	= y2
data = [data0, data1, data2]

# Drawing a Graph

xaxisdata = ['Category', 'RedPSNR', 'BluePSNR', 'GreenPSNR']
Yaxisdata1 = ['Simple Method', 5, 6, 7]
Yaxisdata2 = ['Advanced Method', 50, 10, 15]
ydatas = [Yaxisdata1, Yaxisdata2]

title = 'testgraph'
worksheetname = 'testsheet'
xaxis = 'PSNRTypes'
yaxis = 'Gains'

bookname = 'testbook.xlsx'
assert (extension_chk(bookname, 'xlsx') is True), 'Invalid extension'
book = newbook(bookname)
writer = SheetWriter(book, 'NewSheet')
writer.writetable('head', data, orientation='col')
writer.line_graph(xaxisdata, ydatas, title, xaxis, yaxis)
writer.writetable('head', data, orientation='col')
book.close()
