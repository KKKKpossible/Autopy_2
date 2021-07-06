from openpyxl import Workbook
from openpyxl import load_workbook
from tkinter import *
import tkinter


class Excel:
    def __init__(self, path):
        self.path = path
        self.l_wb = ""
        self.l_ws_list = []
        self.w_wb = ""
        self.w_ws_list = []
        self.column_name = []

        self.load_xls_sheet_all()

    def load_xls_sheet_all(self):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        # Load all work sheet
        for i in range(0, len(ws_name)):
            self.l_ws_list.append(self.l_wb[ws_name[i]])

        # frequency, io offset, voltage get from work sheet 0
        self.get_offset()

    def load_xls_sheet_num(self, sheet_num):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        if (len(ws_name)-1 < sheet_num) or (sheet_num < 0):
            print("error in sheet num")
            pass
        else:
            self.l_ws_list[sheet_num] = self.l_wb[ws_name[sheet_num]]

    def save_xls_sheet_num(self, sheet_num):
        pass

    def get_offset(self):
        pass

class SjKim:
    def __init__(self):
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.g_root.mainloop()

    def root_close(self):
        self.g_root.quit()
