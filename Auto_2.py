from openpyxl import Workbook  # openpy import
import tkinter  # import tkinter
from tkinter import filedialog as fd  # import file dialog
import os
import pyautogui as pg  # message box library
import Auto_Class as Ac


if __name__ == "__main__":

    dialog_root = tkinter.Tk()
    dialog_root.withdraw()

    dir_path = fd.askopenfilename(parent=dialog_root,
                                  initialdir=os.getcwd(),
                                  title='Select Config Excel File')  # 대화창 open
    dialog_root.quit()

    if len(dir_path) == 0:
        pass
    else:
        check_excel = dir_path[len(dir_path) - 5:len(dir_path) - 1]
        # print(check_excel)  # test code

        if check_excel == ".xls":
            ex_ = Ac.Excel(dir_path)
            sj_kim = Ac.SjKim()
        else:
            a = pg.alert(text=".xls 확장자 파일이 아닙니다.", title="Error", button="확인")
            print(a)
