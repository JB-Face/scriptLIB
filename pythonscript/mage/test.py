import py_cui
import py_cui.keys as KEYS
import tkinter as tk
from tkinter import filedialog
import xlrd
import xlwings as xw
 



class MouseApp:

    def __init__(self, root: py_cui.PyCUI):

        # Initialize our two widgets, a button and a mouse press log
        self.root = root
        self.button_presser = self.root.add_button('Press Me!', 0, 0)
        self.mouse_press_log = self.root.add_text_block('Mouse Presses', 0, 1, column_span=2)
        #self.button_presser.add_mouse_command(KEYS.LEFT_MOUSE_CLICK, self.print_left_press_with_coords)
        self.print_left_press_with_coords()
        
    def print_left_press_with_coords(self):
        root = tk.Tk()
        root.withdraw()
        
        file_path = filedialog. askopenfilename(title='选择你的英雄', initialdir='/', filetypes=[('xlsx','*.xlsx'),('xls','*.xls')])
        a = self.breakxlx(file_path)
        self.mouse_press_log.set_text(str(a))

    def breakxlx(self,path):
        data = xlrd.open_workbook(path)
        table = data.sheets()[0]
        maxrows = table.nrows
        maxcols = table.ncols
        pool = {}
        for i in range(maxrows):
            text = table.cell(i,1)
            #print(text)
            if '环节接口人评价' in text.value :
                pool['接口人'] = {'start':i}
            elif 'S-A级' in text.value :
                pool['sa'] = {'start':i}
                pool['接口人'] = {'end':i-3}
            elif 'B级' in text.value :
                pool['b'] = {'start':i}
                pool['sa'] = {'end':i-1}
            elif 'CD级' in text.value :
                pool['cd'] = {'start':i}
                pool['b'] = {'end':i-2}
        return pool



# root = py_cui.PyCUI(1, 3)
# MouseApp(root)
# root.start()

file_path = filedialog. askopenfilename(title='选择你的英雄', initialdir='/', filetypes=[('xlsx','*.xlsx'),('xls','*.xls')])
data = xlrd.open_workbook(file_path)
table = data.sheets()[0]
maxrows = table.nrows
maxcols = table.ncols
pool = {}
for i in range(maxrows):
    text = table.cell(i,1)
    #print(text)
    if '环节接口人评价' in text.value :
        pool['接口人'] = {'start':i}
    elif 'S-A级' in text.value :
        pool['sa'] = {'start':i}
        pool['接口人'] = {'end':i-3}
    elif 'B级' in text.value :
        pool['b'] = {'start':i}
        pool['sa'] = {'end':i-1}
    elif 'CD级' in text.value :
        pool['cd'] = {'start':i}
        pool['b'] = {'end':i-2}