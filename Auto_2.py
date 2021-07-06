from openpyxl import Workbook  # openpy import
import tkinter  # import tkinter
from tkinter import filedialog as fd  # import file dialog
import os
import pyautogui as pg  # message box library
import Auto_Class as Ac


if __name__ == "__main__":
    sj_kim = Ac.SjKim()
    ex = ""

    dir_path = sj_kim.load_file_dialog()
    if dir_path != "-1":
        ex = Ac.Excel(dir_path)
        dir_path = sj_kim.save_file_dialog()

        if dir_path != "-1":
            ex.save_all_var(dir_path)




