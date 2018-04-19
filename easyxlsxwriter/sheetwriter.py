# Wrapper around Worksheet for File Pointer like access
from pointer import XPointer
from copy import deepcopy
from misc import sheetrange


# Worksheet + Pointer
class SheetWriter:
    _sheet = None
    _pointer = None
    # Some function need access to the workbook
    _workbook = None
    _name = None

    # Note name must be same as the
    def __init__(self, wbook, name, x='A', y=1):
        self._sheet = wbook.add_worksheet(name)
        self._pointer = XPointer(x, y)
        self._workbook = wbook
        self._name = name

    def newline(self):
        self._pointer.newline()

    def get_sheet(self):
        return self._sheet

    def get_pointer(self):
        return self._pointer

    def str_pointer(self):
        return str(self._pointer)

    def nextcell(self):
        self._pointer.nextcell()

    # 2D Row Major Data
    def write_rows(self, data):
        for datarow in data:
            self._sheet.write_row(str(self.get_pointer()), datarow)
            self.newline()

    # 2D Column Major Data
    def write_cols(self, data):
        height = 0
        for datacol in data:
            self._sheet.write_column(str(self.get_pointer()), datacol)
            height = max(height, len(datacol))
            self.nextcell()
        self.newline()
        self._pointer.v_jump(height - 1)

    def writetable(self, heading, data, orientation='row'):
        assert(orientation in ['row', 'col']), "Orientation should be row/col"
        print "Got Heading"
        print heading
        print "Got Data"
        print data
        if heading is not None:
            width = 0
            if orientation == 'row':
                width = len(data[0])
            else:
                width = len(data)
            hstart = str(self._pointer)
            hend = str(self._pointer.h_jump_cal(width - 1))
            merge_range = ':'.join([hstart, hend])
            self._sheet.merge_range(merge_range, heading)
            self.newline()

        if orientation == 'row':
            self.write_rows(data)
        else:
            self.write_cols(data)

    # Data for each heading is given in a row
    # Writes a Table to Excel given the Data and Sheetname
    '''
    Write to Excel
    H1  H2  H3
    D00 D10 D20
    D10 D11 D21

    Eg:
    headings = ['Number', 'Batch 1', 'Batch 2']
    data = [
    [2, 3, 4, 5, 6, 7], --under Number
    [10, 40, 50, 20, 10, 50],--under Batch1
    [30, 60, 70, 50, 40, 30],--under Batch2
    ]
    '''

    # Write the data into a table and return a chart
    # The Pointer Logic Inside this is cryptic, I couldn't write cleaner code
    # Hopefully you would never have to bother tweaking this code
    def line_graph(self, headings, data, title, xaxis, yaxis):
        chart = self._workbook.add_chart({'type': 'line'})
        # Chart Data Pointer: Points to the begining of the chart
        cd_pointer = deepcopy(self.get_pointer())
        cd_pointer.newline()
        # Where the graph will be drawn
        graph_cell = deepcopy(self.get_pointer()).h_jump_cal(len(headings))

        # Add the headings to the columns
        data = [[h] + d for h, d in zip(headings, data)]
        self.writetable(title, data, orientation='col')
        # Deal with Categories
        cd_pointer.newline()  # Go to line after chart title
        categorystr = sheetrange(
            self._name,
            cd_pointer,
            # -2 is due to Heading & Title
            cd_pointer.v_jump_cal(len(data[0]) - 2)
        )
        # Jump to data
        cd_pointer.nextcell()
        for datacol in data[1:]:
            jmp = len(datacol)
            chart.add_series({
                'name': sheetrange(self._name, cd_pointer.v_jump_cal(-1)),
                'categories': categorystr,
                'values': sheetrange(self._name,
                                     cd_pointer,
                                     cd_pointer.v_jump_cal(jmp))
            })
            cd_pointer.nextcell()

        # Set various chart options
        # Additional Configuration could be passed thorugh the method
        chart.set_title({'name': title})
        chart.set_x_axis({'name': xaxis})
        chart.set_y_axis({'name': yaxis})
        chart.set_style(10)
        offsets = {'x_offset': 25, 'y_offset': 15}
        self._sheet.insert_chart(str(graph_cell), chart, offsets)
