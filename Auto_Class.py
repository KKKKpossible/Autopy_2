from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter
import pyautogui as pg  # message box library
from tkinter import filedialog as fd  # import file dialog
import os

class VariableDeclare:
    def __init__(self, path):
        self.load_clear = True
        self.path = path
        self.l_wb = ""
        self.l_ws_list = []
        self.w_wb = ""
        self.w_ws_list = []
        self.column_name = []

        self.freq_var = []
        self.input_offset_var = []
        self.output_offset_var = []
        self.voltage_var = []


class CommonVar:
    def __init__(self):
        self.excel_data_name_row = 1
        self.excel_data_unit_row = 2
        self.excel_data_start_row = 3
        self.excel_freq_unit = ["Hz", "KHz", "MHz", "GHz", "THz"]
        self.excel_input_unit = ["dB"]
        self.excel_output_unit = ["dB"]
        self.excel_power_unit = ["V"]
        self.rf_power = []

        self.excel_freq_column = "A"
        self.excel_input_column = "B"
        self.excel_output_column = "C"
        self.excel_power_column = "D"
        self.save_hor = True


class Excel(CommonVar, VariableDeclare):
    def __init__(self, path):
        CommonVar.__init__(self)
        VariableDeclare.__init__(self, path)

        self.load_xls_sheet_all()
        self.get_offset_procedure_version_0()

    def load_xls_sheet_all(self):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        # Load all work sheet
        for i in range(0, len(ws_name)):
            self.l_ws_list.append(self.l_wb[ws_name[i]])

    def load_xls_sheet_num(self, sheet_num):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        if (len(ws_name)-1 < sheet_num) or (sheet_num < 0):
            print("error in sheet num")
            self.load_clear = False
        else:
            self.l_ws_list[sheet_num] = self.l_wb[ws_name[sheet_num]]

    def save_all_var(self, path):
        # test var
        print(self.freq_var)
        print(self.input_offset_var)
        print(self.output_offset_var)
        print(self.voltage_var)

        if self.save_hor is True:
            self.w_wb = Workbook()
            self.w_ws_list.append(self.w_wb.create_sheet("ATR"))

            self.w_ws_list[0] = self.w_wb.active
            self.w_ws_list[0].cell(1, 1, "frequency")
            self.w_ws_list[0].cell(1, 2, "rf output_0")
            self.w_ws_list[0].cell(1, 3, "rf output_1")
            self.w_ws_list[0].cell(1, 4, "current")

            self.w_ws_list[0].cell(2, 1, "Hz")
            self.w_ws_list[0].cell(2, 2, "dBm")
            self.w_ws_list[0].cell(2, 3, "W")
            self.w_ws_list[0].cell(2, 4, "A")

            self.w_wb.save(path)

    # frequency, io offset, voltage get from work sheet 0
    def get_offset_procedure_version_0(self):
        # 현재 self.l_ws_list[0] frequency, input offset, output offset, voltage 정보가 들어 있다.
        # 이것을 추출하는 작업을 하려함.
        # 1. freq config
        # 2. check input offset unit
        # 3. check output offset unit
        # 4. check power voltage unit
        # 5. freq load
        # 6. input db load
        # 7. output db load
        # 8. Voltage load

        # 1. freq config
        multiple_hz = 0

        d_name_cell = self.excel_freq_column + str(self.excel_data_name_row)
        d_unit_cell = self.excel_freq_column + str(self.excel_data_unit_row)

        if self.l_ws_list[0][d_name_cell].value == 'frequency':
            print(self.l_ws_list[0][d_name_cell].value)
            if self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[0]:
                multiple_hz = 1
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[1]:
                multiple_hz = 1 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[2]:
                multiple_hz = 1 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[3]:
                multiple_hz = 1 * 1000 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[4]:
                multiple_hz = 1 * 1000 * 1000 * 1000 * 1000
            else:
                multiple_hz = 0
                print("Hz not defined")
                a = pg.alert(text="freq unit 설정되지 않음\n"
                                  + "ex) Hz, KHz, MHz, GHz, THz",
                             title="Error",
                             button="확인")
                self.load_clear = False
                return

        # 2. check input offset unit
        d_unit_cell = self.excel_input_column + str(self.excel_data_unit_row)
        if self.l_ws_list[0][d_unit_cell].value != self.excel_input_unit[0]:
            print("dB not defined")
            a = pg.alert(text="input offset unit 설정되지 않음\n"
                              + "ex) dB",
                         title="Error",
                         button="확인")
            self.load_clear = False
            return

        # 3. check output offset unit
        d_unit_cell = self.excel_output_column + str(self.excel_data_unit_row)
        if self.l_ws_list[0][d_unit_cell].value != self.excel_output_unit[0]:
            print("dB not defined")
            a = pg.alert(text="output offset unit 설정되지 않음\n"
                              + "ex) dB",
                         title="Error",
                         button="확인")
            self.load_clear = False
            return
        # 4. check power voltage unit
        d_unit_cell = self.excel_power_column + str(self.excel_data_unit_row)
        if self.l_ws_list[0][d_unit_cell].value != self.excel_power_unit[0]:
            print("V not defined")
            a = pg.alert(text="Power voltage unit 설정되지 않음\n"
                              + "ex) V",
                         title="Error",
                         button="확인")
            self.load_clear = False
            return

        # 5. freq load
        i = self.excel_data_start_row
        freq_data_cell = self.excel_freq_column + str(i)
        while self.l_ws_list[0][freq_data_cell].value is not None:
            print(self.l_ws_list[0][freq_data_cell].value)
            self.freq_var.append(self.l_ws_list[0][freq_data_cell].value * multiple_hz)
            i += 1
            freq_data_cell = self.excel_freq_column + str(i)

        # 6. input db load
        i = self.excel_data_start_row
        input_data_cell = self.excel_input_column + str(i)
        while self.l_ws_list[0][input_data_cell].value is not None:
            print(self.l_ws_list[0][input_data_cell].value)
            self.input_offset_var.append(self.l_ws_list[0][input_data_cell].value)

            i += 1
            input_data_cell = self.excel_input_column + str(i)

        # 7. output db load
        i = self.excel_data_start_row
        output_data_cell = self.excel_output_column + str(i)
        while self.l_ws_list[0][output_data_cell].value is not None:
            print(self.l_ws_list[0][output_data_cell].value)
            self.output_offset_var.append(self.l_ws_list[0][output_data_cell].value)

            i += 1
            output_data_cell = self.excel_output_column + str(i)

        # 8. Voltage load
        i = self.excel_data_start_row
        voltage_data_cell = self.excel_power_column + str(i)
        while self.l_ws_list[0][voltage_data_cell].value is not None:
            print(self.l_ws_list[0][voltage_data_cell].value)
            self.voltage_var.append(self.l_ws_list[0][voltage_data_cell].value)

            i += 1
            voltage_data_cell = self.excel_power_column + str(i)


class SjKim:
    def __init__(self):
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.g_root.withdraw()

    def root_close(self):
        self.g_root.quit()

    def load_file_dialog(self):
        dir_path = fd.askopenfilename(parent=self.g_root,
                                      initialdir=os.getcwd(),
                                      title='Select Config Excel File')  # 대화창 open
        if len(dir_path) == 0:
            return "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                return dir_path
            else:
                self.load_file_dialog()

    def save_file_dialog(self):

        dir_path = fd.asksaveasfilename(parent=self.g_root,
                                        initialdir=os.getcwd(),
                                        initialfile="atr.xlsx",
                                        title='Save Excel File',
                                        filetypes=[("excel files", "*.xlsx"), ("all", "*.*")]
                                        )  # 대화창 open
        if len(dir_path) == 0:
            return "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                return dir_path
            else:
                self.save_file_dialog()

