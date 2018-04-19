from excelmath import to_excel, from_excel


# Pointer Class for easy file pointer like navigation inside excel
class XPointer:
    x = 'A'
    y = 1

    def __str__(self):
        return(self.x + str(self.y))

    # X must be a string while y must be integer
    def __init__(self, x, y):
        # TODO : Add Safety Checks
        self.x = x
        self.y = int(y)

    # Calculate positon after units jump along x-axis
    def h_jump_cal(self, units):
        return XPointer(to_excel(from_excel(self.x) + units), self.y)

    def v_jump_cal(self, units):
        return XPointer(self.x, self.y + units)

    def h_jump(self, units):
        self.x = to_excel(from_excel(self.x) + units)

    def v_jump(self, units):
        self.y = self.y + units

    def newline(self):
        self.x = 'A'
        self.y += 1

    def nextcell(self):
        self.x = to_excel(from_excel(self.x) + 1)

    def move(self, x, y):
        self.x = x
        self.y = y
