# exw
Easy Xlsx : Wrapper around xlsxwriter & openpyxl module in python.

xlsxwriter spreadsheet class doesn't have its own pointer giving rise to this project.
Easy Xlsx aims to provide access to spreadsheets along with a pointer so that more can be done in less.
And Python Loops can be kept clean.
It provides a simpler interface to both xlsxwriter & openpxl importing both of them and providing a uniform interface

# Note

Most features of xlsxwriter & openpyxl will not be supported by this wrapper. If you want flexiblity stick with the original libraries which is quite great themselves. Anything that becomes complicated is straight away reimplemented. Example table writer is rewritten and doesn't match with that of xlsxwriter. 

# Limitations

- Xlsxwriter doesn't support does support appending to existing files , neither is is supported by this wrapper
- Openpyxl is currently used only reading.
- Don't read and write at the same time. It simply won't work. Different modules at the backend.
- Future Work would try to eliminate these limitations.

# Design Principles
- Limit exposure to xlsxwriter & openpyxl
- Provide File Pointer like control over Worksheets
- Simply Trivial Tasks
- Code : Python PEP 8
- Highly reusable for basic tasks.
- Performance : Doesn't matter. You shouldn't be using xl in the first place.


