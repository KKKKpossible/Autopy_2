from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter
import pyautogui as pg  # message box library
from tkinter import filedialog as fd  # import file dialog
import os
from functools import partial


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


class CommonDialogVar:
    def __init__(self):
        self.remote_p_meter_id = 0
        self.remote_power_id = 1
        self.remote_network_id = 2
        self.remote_source_id = 3
        self.remote_spectrum_id = 4
        self.remote_all_id = 5

        self.atr_power_id = 6
        self.atr_network_id = 7
        self.atr_spectrum_id = 8

        self.remote_frame = ""
        self.atr_frame = ""

        self.remote_start_x = 30
        self.remote_start_y = 10
        self.remote_button_x = 10
        self.remote_y_gap = 10

        self.atr_start_x = 25
        self.atr_start_y = 10
        self.atr_button_x = 10
        self.atr_y_gap = 10

        self.main_butt_width = 10
        self.main_butt_height = 1


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
    def __init__(self, path=""):
        CommonVar.__init__(self)
        VariableDeclare.__init__(self, path)

    def load_excel_procedure(self, load_path):
        self.path = load_path
        self.__load_xls_sheet_all()
        self.__get_offset_procedure_version_0()

    def set_path(self, path):
        self.path = path

    def __load_xls_sheet_all(self):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        # Load all work sheet
        for i in range(0, len(ws_name)):
            self.l_ws_list.append(self.l_wb[ws_name[i]])

    # not finished
    # def load_xls_sheet_num(self, sheet_num):
    #     self.l_wb = load_workbook(self.path, data_only=True)
    #     ws_name = self.l_wb.get_sheet_names()
    #
    #     if (len(ws_name)-1 < sheet_num) or (sheet_num < 0):
    #         print("error in sheet num")
    #         self.load_clear = False
    #     else:
    #         self.l_ws_list[sheet_num] = self.l_wb[ws_name[sheet_num]]

    def save_all_var(self, save_path):
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
            self.w_ws_list[0].cell(2, 1, "rf output_0")
            self.w_ws_list[0].cell(3, 1, "rf output_1")
            self.w_ws_list[0].cell(4, 1, "current")

            self.w_ws_list[0].cell(1, 2, "Hz")
            self.w_ws_list[0].cell(2, 2, "dBm")
            self.w_ws_list[0].cell(3, 2, "W")
            self.w_ws_list[0].cell(4, 2, "A")

            self.w_wb.save(save_path)
        else: # save vertical atr
            pass

    # frequency, io offset, voltage get from work sheet 0
    def __get_offset_procedure_version_0(self):
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


class Dialog(CommonDialogVar):
    def __init__(self):
        CommonDialogVar.__init__(self)
        # excel instance call
        self.exel = Excel()
        # dialog set
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.g_root.withdraw()
        self.excel_path = ""

    def main_dialog_open(self):
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.__make_main_panel()
        self.g_root.mainloop()

    def __make_main_panel(self):
        self.g_root.title("Main Dialog")
        self.g_root.geometry("640x350")
        self.g_root.resizable(False, False)

        # self.remote_frame = tkinter.LabelFrame(self.g_root, text="Remote", padx=10, pady=10)
        # self.remote_frame.grid(row=0, column=1)
        # remote power meter button
        self.__add_remote("Open", self.remote_p_meter_id)
        # remote power button
        self.__add_remote("Open", self.remote_power_id)
        # remote network button
        self.__add_remote("Open", self.remote_network_id)
        # remote source button
        self.__add_remote("Open", self.remote_source_id)
        # remote spectrum button
        self.__add_remote("Open", self.remote_spectrum_id)
        # remote all button
        self.__add_remote("Open", self.remote_all_id)

        # atr label
        # self.atr_frame = tkinter.LabelFrame(self.g_root, text="ATR", padx=10, pady=10)
        # self.atr_frame.grid(row=0, column=0)
        # atr power button
        self.__add_atr("Open", self.atr_power_id)
        # atr network button
        self.__add_atr("Open", self.atr_network_id)
        # atr spectrum button
        self.__add_atr("Open", self.atr_spectrum_id)

    def __add_remote(self, text, _id):
        empty_row = 0
        empty_column = 0
        label_column = 1
        button_column = 2

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__button_clicked_remote, _id),
                                width=self.main_butt_width,
                                height=self.main_butt_height)
        if _id is None:
            pass
        elif _id == self.remote_p_meter_id:
            _row = 1
            # label
            label = tkinter.Label(self.g_root, text="Power Meter Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_start_y)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_start_y)
        elif _id == self.remote_power_id:
            _row = 2
            # label
            label = tkinter.Label(self.g_root, text="Power Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_y_gap)
        elif _id == self.remote_network_id:
            _row = 3
            # label
            label = tkinter.Label(self.g_root, text="Network Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_y_gap)
        elif _id == self.remote_source_id:
            _row = 4
            # label
            label = tkinter.Label(self.g_root, text="Source Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_y_gap)
        elif _id == self.remote_spectrum_id:
            _row = 5
            # label
            label = tkinter.Label(self.g_root, text="Spectrum Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_y_gap)
        elif _id == self.remote_all_id:
            _row = 6
            # label
            label = tkinter.Label(self.g_root, text="All Remote", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.remote_start_x, pady=self.remote_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.remote_button_x, pady=self.remote_y_gap)
        else:
            pass

    def __add_atr(self, text, _id):
        empty_row = 0
        label_column = 3
        button_column = 4

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__button_clicked_remote,
                                                _id),
                                width = self.main_butt_width,
                                height = self.main_butt_height)
        if _id == self.atr_power_id:
            _row = 1
            # label
            label = tkinter.Label(self.g_root, text="Power Meter ATR", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.atr_start_x, pady=self.atr_start_y)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.atr_button_x, pady=self.atr_start_y)
        elif _id == self.atr_network_id:
            _row = 2
            # label
            label = tkinter.Label(self.g_root, text="Network ATR", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.atr_start_x, pady=self.atr_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.atr_button_x, pady=self.atr_y_gap)
        elif _id == self.atr_spectrum_id:
            _row = 3
            # label
            label = tkinter.Label(self.g_root, text="Spectrum ATR", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.atr_start_x, pady=self.atr_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.atr_button_x, pady=self.atr_y_gap)
        else:
            pass

    def __button_clicked_remote(self, _id):
        if _id is None:
            pass
        elif _id == self.remote_p_meter_id:
            pass
        elif _id == self.remote_power_id:
            pass
        elif _id == self.remote_network_id:
            pass
        elif _id == self.remote_source_id:
            pass
        elif _id == self.remote_spectrum_id:
            pass
        elif _id == self.remote_all_id:
            pass
        elif _id == self.atr_power_id:
            # load excel dialog open
            self.load_file_dialog()
            if self.excel_path != "-1":
                self.exel.load_excel_procedure(load_path=self.excel_path)

                # save excel dialog open
                self.save_file_dialog()
                if self.excel_path != "-1":
                    self.exel.save_all_var(save_path=self.excel_path)

        elif _id == self.atr_network_id:
            pass
        elif _id == self.atr_spectrum_id:
            pass

    def root_close(self):
        self.g_root.withdraw()
        self.g_root.quit()

    def load_file_dialog(self):
        dir_path = fd.askopenfilename(parent=self.g_root,
                                      initialdir=os.getcwd(),
                                      title='Select Config Excel File')  # 대화창 open
        if len(dir_path) == 0:
            self.excel_path = "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                self.excel_path = dir_path
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
            self.excel_path = "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                self.excel_path = dir_path
            else:
                self.save_file_dialog()

