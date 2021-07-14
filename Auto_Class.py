from openpyxl import Workbook
from openpyxl import load_workbook
import tkinter
import pyautogui as pg  # message box library
from tkinter import filedialog as fd  # import file dialog
import os
from functools import partial
import Instrument
import time
import numpy as np


class ExcelVariableDeclare:
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
        self.power_voltage_var = ""
        self.power_current_var = ""
        self.power_current_limit_var = ""
        self.power_meter_probe_var = ""
        self.aging_var = ""
        self.aging_left_var = ""
        self.aging_done_var = False
        self.reference_done = False
        self.source_com_var = ""
        self.power_com_var = ""
        self.power_meter_com_var = ""

        self.multiple_freq = 0  # Hz, KHz, MHz, GHz, THz
        self.multiple_aging = 0  # sec, min, hour
        self.source_comm_option = ""  # GPIB, USB, Serial
        self.power_comm_option = ""  # GPIB, USB, Serial
        self.power_meter_comm_option = ""  # GPIB, USB, Serial

        self.multiple_select_freq = 0  # Hz, KHz, MHz, GHz, THz
        self.select_freq_var = []  # select freq var
        self.atr_start_input_var = ""
        self.p_sat_input_var = ""
        self.overdrive_input_var = ""
        self.pout_var = ""

        self.freq_cell_name = ""  # default Frequency
        self.input_cell_name = ""  # default Input offset
        self.output_cell_name = ""  # default Output offset
        self.power_voltage_cell_name = ""  # default Power Voltage
        self.power_current_cell_name = ""  # default Power Voltage
        self.power_meter_probe_cell_name = ""  # default Power Voltage
        self.aging_cell_name = ""  # default Aging time
        self.source_comm_cell_name = ""  # default Source generate
        self.power_comm_cell_name = ""  # default Power supply
        self.power_meter_comm_cell_name = ""  # default Power meter
        self.select_freq_cell_name = ""  # default Select Freq
        self.atr_start_input_dbm_name = ""  # default ATR start input dBm
        self.p_sat_input_name = ""  # default P_sat input
        self.overdrive_input_name = ""  # default Overdrive input
        self.pout_name = ""  # default Pout

        self.source_baud_rate = ""  # default Source generate
        self.power_baud_rate = ""  # default Power supply
        self.power_baud_rate = ""  # default Power meter


class CommonDialogVar:
    def __init__(self):
        # excel instance call
        self.excel = Excel()
        # dialog set
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.g_root.withdraw()
        self.excel_path = ""

        self.remote_p_meter_id = 0
        self.remote_power_id = 1
        self.remote_network_id = 2
        self.remote_source_id = 3
        self.remote_spectrum_id = 4
        self.remote_all_id = 5

        self.atr_power_id = 6
        self.atr_network_id = 7
        self.atr_spectrum_id = 8

        self.atr_start_button_id = 0
        self.atr_stop_button_id = 1
        self.atr_display_button_id = 2
        self.atr_save_button_id = 3

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

        # atr entry values
        self.atr_freq = tkinter.StringVar()
        self.atr_input_offset = tkinter.StringVar()
        self.atr_output_offset = tkinter.StringVar()
        self.atr_power_voltage = tkinter.StringVar()
        self.atr_power_current = tkinter.StringVar()
        self.atr_aging_time = tkinter.StringVar()
        self.atr_rf_output = tkinter.StringVar()
        self.atr_current = tkinter.StringVar()
        self.atr_aging_left = tkinter.StringVar()
        self.atr_system_option = False
        self.atr_sys_fwd = tkinter.StringVar()
        self.atr_sys_freq = tkinter.StringVar()
        self.atr_sys_temp = tkinter.StringVar()
        self.atr_sys_cmd = tkinter.StringVar()

        self.atr_source_com = tkinter.StringVar()
        self.atr_power_com = tkinter.StringVar()
        self.atr_power_meter_com = tkinter.StringVar()

        self.atr_rf_var = []
        self.atr_current_var = []

        self.atr_state = ""
        self.atr_p1_ref_input = ""
        self.atr_index = 0

        self.atr_p1_var = []

        self.inst_source = ""
        self.inst_power = ""
        self.inst_power_meter = ""

    def root_close(self):
        self.g_root.withdraw()
        self.g_root.quit()


class ExcelCommonVar:
    def __init__(self):
        self.excel_data_name_row = 1
        self.excel_data_unit_row = 2
        self.excel_data_start_row = 3
        self.excel_freq_unit = ["Hz", "KHz", "MHz", "GHz", "THz"]
        self.excel_input_unit = ["dB"]
        self.excel_output_unit = ["dB"]
        self.excel_power_unit = ["V"]
        self.excel_current_unit = ["A"]
        self.excel_power_meter_probe_unit = ["Channel"]
        self.excel_aging_unit = ["sec", "min", "hour"]
        self.excel_com_unit = ["GPIB", "USB", "Serial"]
        self.excel_dbm_unit = ["dBm"]
        self.rf_power = []

        self.excel_freq_column = "A"
        self.excel_input_column = "B"
        self.excel_output_column = "C"
        self.excel_power_voltage_column = "D"
        self.excel_power_current_column = "E"
        self.excel_power_meter_probe_column = "F"
        self.excel_aging_column = "G"
        self.excel_source_com_column = "H"
        self.excel_power_com_column = "I"
        self.excel_power_meter_com_column = "J"
        self.excel_select_freq_column = "K"
        self.excel_atr_start_input_column = "L"
        self.excel_p_sat_input_column = "M"
        self.excel_overdrive_input_column = "N"
        self.excel_pout_column = "O"
        self.save_hor = True


class Excel(ExcelCommonVar, ExcelVariableDeclare):
    def __init__(self, path=""):
        ExcelCommonVar.__init__(self)
        ExcelVariableDeclare.__init__(self, path)

    def load_power_atr_excel_procedure(self, load_path):
        self.path = load_path
        self.__load_xls_sheet_all()
        self.__get_offset_procedure_version_0()

    def set_path(self, path):
        self.path = path

    def __load_xls_sheet_all(self):
        self.l_wb = load_workbook(self.path, data_only=True)
        ws_name = self.l_wb.get_sheet_names()

        while len(self.l_ws_list) != 0:
            self.l_ws_list.pop()
            print("clean l_ws_list")
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

    def save_all_var(self, save_path, rf_var, curr_var):
        # test var
        print(self.freq_var)
        print(self.input_offset_var)
        print(self.output_offset_var)
        print(self.power_voltage_var)
        print(self.aging_var)
        print(self.source_com_var)
        print(self.power_com_var)
        print(self.power_meter_com_var)

        if self.save_hor is True:
            freq_row = 1
            rf_dbm_row = 2
            rf_w_row = 3
            current_row = 4
            atr_column = 1

            # test code
            # for var in range(0, len(self.freq_var)):
            #     rf_var.append(var)
            #     curr_var.append(len(self.freq_var) - var)

            if (len(rf_var) <= 0) and (len(curr_var) <= 0):
                pg.alert(text="No data exist",
                         title="Error",
                         button="확인")
                return

            if (len(rf_var) != len(self.freq_var)) or (len(rf_var) <= 0):
                pg.alert(text="RF data error\n" + "ex) ATR not started",
                         title="Error",
                         button="확인")
                return

            if (len(curr_var) != len(self.freq_var)) or (len(rf_var) <= 0):
                pg.alert(text="Current data error\n" + "ex) ATR not started",
                         title="Error",
                         button="확인")
                return

            self.w_wb = Workbook()
            self.w_ws_list.append(self.w_wb.create_sheet("ATR"))

            self.w_ws_list[0] = self.w_wb.active
            self.w_ws_list[0].cell(freq_row, atr_column, "frequency")
            self.w_ws_list[0].cell(rf_dbm_row, atr_column, "rf output_0")
            self.w_ws_list[0].cell(rf_w_row, atr_column, "rf output_1")
            self.w_ws_list[0].cell(current_row, atr_column, "current")
            atr_column += 1
            self.w_ws_list[0].cell(freq_row, atr_column, "Hz")
            self.w_ws_list[0].cell(rf_dbm_row, atr_column, "dBm")
            self.w_ws_list[0].cell(rf_w_row, atr_column, "W")
            self.w_ws_list[0].cell(current_row, atr_column, "A")
            atr_column += 1

            var_index = 0
            while len(self.freq_var) > var_index:
                self.w_ws_list[0].cell(freq_row, atr_column, self.freq_var[var_index])
                self.w_ws_list[0].cell(rf_dbm_row, atr_column, rf_var[var_index])
                self.w_ws_list[0].cell(rf_w_row, atr_column, (10 ** (rf_var[var_index] / 10) / 1000))
                self.w_ws_list[0].cell(current_row, atr_column, curr_var[var_index])
                var_index += 1
                atr_column += 1

            self.w_wb.save(save_path)
        else:  # save vertical atr
            pass

    # frequency, io offset, voltage get from work sheet 0
    def __get_offset_procedure_version_0(self):
        # 현재 self.l_ws_list[0] frequency, input offset, output offset, voltage 정보가 들어 있다.
        # 이것을 추출하는 작업을 하려함.
        # 1. freq config
        # 2. check input offset
        # 3. check output offset
        # 4. check power voltage set
        # 5. check power current limit
        # 6. check power meter probe channel
        # 7. check start aging time
        # 8. check source generate communication
        # 9. check power supply communication
        # 10. check power meter communication
        # 11. check select frequency
        # 12. check ATR start input dBm
        # 13. check P_sat input dBm
        # 14. check Overdrive input dBm
        # 15. check Pout output dBm
        # 12. load freq
        # 13. load input db
        # 14. load output db
        # 15. load Voltage set
        # 16. load Current limit
        # 17. load power meter probe channel
        # 18. load Start aging time
        # 19. load source generate communication
        # 20. load power supply communication
        # 21. load power meter communication
        # 22. load select frequency
        # 27. load ATR start input dBm
        # 28. load P_sat input dBm
        # 29. load Overdrive input dBm
        # 30. load Pout output dBm

        # 1. freq config
        self.__get_unit_from_column(self.excel_freq_column)
        # 2. check input offset
        self.__get_unit_from_column(self.excel_input_column)
        # 3. check output offset
        self.__get_unit_from_column(self.excel_output_column)
        # 4. check power voltage set
        self.__get_unit_from_column(self.excel_power_voltage_column)
        # 5. check power current limit
        self.__get_unit_from_column(self.excel_power_current_column)
        # 6. check power meter probe channel
        self.__get_unit_from_column(self.excel_power_meter_probe_column)
        # 7. check start aging time
        self.__get_unit_from_column(self.excel_aging_column)
        # 8. check source generate communication
        self.__get_unit_from_column(self.excel_source_com_column)
        # 9. check power supply communication
        self.__get_unit_from_column(self.excel_power_com_column)
        # 10. check power meter communication
        self.__get_unit_from_column(self.excel_power_meter_com_column)
        # 11. check select frequency
        self.__get_unit_from_column(self.excel_select_freq_column)
        # 12. check ATR start input dBm
        self.__get_unit_from_column(self.excel_atr_start_input_column)
        # 13. check P_sat input dBm
        self.__get_unit_from_column(self.excel_p_sat_input_column)
        # 14. check Overdrive input dBm
        self.__get_unit_from_column(self.excel_overdrive_input_column)
        # 15. check Pout output dBm
        self.__get_unit_from_column(self.excel_pout_column)
        # 16. load freq
        self.__load_var_from_column(self.excel_freq_column)
        # 17. load input db
        self.__load_var_from_column(self.excel_input_column)
        # 18. load output db
        self.__load_var_from_column(self.excel_output_column)
        # 19. load Voltage set
        self.__load_var_from_column(self.excel_power_voltage_column)
        # 20. load Current limit
        self.__load_var_from_column(self.excel_power_current_column)
        # 21. load power meter probe channel
        self.__load_var_from_column(self.excel_power_meter_probe_column)
        # 22. load Start aging time
        self.__load_var_from_column(self.excel_aging_column)
        # 23. load source generate communication
        self.__load_var_from_column(self.excel_source_com_column)
        # 24. load power supply communication
        self.__load_var_from_column(self.excel_power_com_column)
        # 25. load power meter communication
        self.__load_var_from_column(self.excel_power_meter_com_column)
        # 26. load select frequency
        self.__load_var_from_column(self.excel_select_freq_column)
        # 27. load ATR start input dBm
        self.__load_var_from_column(self.excel_atr_start_input_column)
        # 28. load P_sat input dBm
        self.__load_var_from_column(self.excel_p_sat_input_column)
        # 29. load Overdrive input dBm
        self.__load_var_from_column(self.excel_overdrive_input_column)
        # 30. load Pout output dBm
        self.__load_var_from_column(self.excel_pout_column)

    def __get_unit_from_column(self, column):
        d_name_cell = column + str(self.excel_data_name_row)
        d_unit_cell = column + str(self.excel_data_unit_row)
        if column == self.excel_freq_column:
            self.multiple_freq = 0
            self.freq_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[0]:
                self.multiple_freq = 1
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[1]:
                self.multiple_freq = 1 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[2]:
                self.multiple_freq = 1 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[3]:
                self.multiple_freq = 1 * 1000 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[4]:
                self.multiple_freq = 1 * 1000 * 1000 * 1000 * 1000
            else:
                print("Hz not defined")
                pg.alert(text="freq unit 설정되지 않음\n" + "ex) Hz, KHz, MHz, GHz, THz",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_input_column:
            self.input_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_input_unit[0]:
                print("dB not defined")
                pg.alert(text="input offset unit 설정되지 않음\n" + "ex) dB",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_output_column:
            self.output_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_output_unit[0]:
                print("dB not defined")
                pg.alert(text="output offset unit 설정되지 않음\n" + "ex) dB",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_power_voltage_column:
            self.power_voltage_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_power_unit[0]:
                print("V not defined")
                pg.alert(text="Power voltage unit 설정되지 않음\n" + "ex) V",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_power_current_column:
            self.power_current_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_current_unit[0]:
                print("A not defined")
                pg.alert(text="Power current unit 설정되지 않음\n" + "ex) A",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_power_meter_probe_column:
            self.power_meter_probe_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_power_meter_probe_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="Power meter probe channel 설정되지 않음\n" + "ex) Channel",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_aging_column:
            self.multiple_aging = 0
            self.aging_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_aging_unit[0]:
                self.multiple_aging = 1
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_aging_unit[1]:
                self.multiple_aging = 1 * 60
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_aging_unit[2]:
                self.multiple_aging = 1 * 60 * 60
            else:
                print("aging time not defined")
                pg.alert(text="Start aging time unit 설정되지 않음\n" + "ex) sec, min, hour",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_source_com_column:
            self.source_comm_option = ""
            self.source_comm_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[0]:
                self.source_comm_option = self.excel_com_unit[0]  # GPIB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[1]:
                self.source_comm_option = self.excel_com_unit[1]  # USB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[2]:
                self.source_comm_option = self.excel_com_unit[2]  # Serial
                self.source_baud_rate = self.l_ws_list[0][d_unit_cell + 1].value
            else:
                print("Source communication option error")
                pg.alert(text="Source communication unit 설정되지 않음\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate 필요",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_power_com_column:
            self.power_comm_option = ""
            self.power_comm_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[0]:
                self.power_comm_option = self.excel_com_unit[0]  # GPIB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[1]:
                self.power_comm_option = self.excel_com_unit[1]  # USB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[2]:
                self.power_comm_option = self.excel_com_unit[2]  # Serial
                self.source_baud_rate = self.l_ws_list[0][d_unit_cell + 1].value
            else:
                print("Power supply communication option error")
                pg.alert(text="Power supply communication unit 설정되지 않음\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate 필요",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_power_meter_com_column:
            self.power_meter_comm_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[0]:
                self.power_meter_comm_option = self.excel_com_unit[0]  # GPIB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[1]:
                self.power_meter_comm_option = self.excel_com_unit[1]  # USB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[2]:
                self.power_meter_comm_option = self.excel_com_unit[2]  # Serial
            else:
                print("Power meter communication option error")
                pg.alert(text="Power meter communication unit 설정되지 않음\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate 필요",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_select_freq_column:
            self.multiple_select_freq = 0
            self.select_freq_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[0]:
                self.multiple_select_freq = 1
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[1]:
                self.multiple_select_freq = 1 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[2]:
                self.multiple_select_freq = 1 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[3]:
                self.multiple_select_freq = 1 * 1000 * 1000 * 1000
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_freq_unit[4]:
                self.multiple_select_freq = 1 * 1000 * 1000 * 1000 * 1000
            else:
                print("Hz not defined")
                pg.alert(text="Select freq unit 설정되지 않음\n" + "ex) Hz, KHz, MHz, GHz, THz",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_atr_start_input_column:
            self.atr_start_input_dbm_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="atr start dBm unit 설정되지 않음\n" + "ex) dBm",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_p_sat_input_column:
            self.p_sat_input_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="p_sat input dBm unit 설정되지 않음\n" + "ex) dBm",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_overdrive_input_column:
            self.overdrive_input_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="overdrive dBm unit 설정되지 않음\n" + "ex) dBm",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        elif column == self.excel_pout_column:
            self.pout_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="overdrive dBm unit 설정되지 않음\n" + "ex) dBm",
                         title="Error",
                         button="확인")
                self.load_clear = False
                return
        else:
            print("알수없는 communication column")
            self.load_clear = False

    def __load_var_from_column(self, data_column):
        i = self.excel_data_start_row

        _data_cell = data_column + str(i)
        while self.l_ws_list[0][_data_cell].value is not None:
            print(self.l_ws_list[0][_data_cell].value)
            # check column
            if data_column == self.excel_freq_column:
                self.freq_var.append(self.l_ws_list[0][_data_cell].value * self.multiple_freq)
            elif data_column == self.excel_input_column:
                self.input_offset_var.append(self.l_ws_list[0][_data_cell].value)
            elif data_column == self.excel_output_column:
                self.output_offset_var.append(self.l_ws_list[0][_data_cell].value)
            elif data_column == self.excel_power_voltage_column:
                self.power_voltage_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_current_column:
                self.power_current_limit_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_meter_probe_column:
                self.power_meter_probe_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_aging_column:
                self.aging_var = self.l_ws_list[0][_data_cell].value * self.multiple_aging
            elif data_column == self.excel_source_com_column:
                self.source_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_com_column:
                self.power_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_meter_com_column:
                self.power_meter_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_select_freq_column:
                self.select_freq_var.append(self.l_ws_list[0][_data_cell].value * self.multiple_select_freq)
            elif data_column == self.excel_atr_start_input_column:
                self.atr_start_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_p_sat_input_column:
                self.p_sat_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_overdrive_input_column:
                self.overdrive_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_pout_column:
                self.pout_var = self.l_ws_list[0][_data_cell].value
            else:
                print("Error 등록되지 않은 id")
                self.load_clear = False
                return
            i += 1
            _data_cell = data_column + str(i)


class Dialog(CommonDialogVar):
    def __init__(self):
        CommonDialogVar.__init__(self)

    def main_dialog_open(self):
        self.g_root = tkinter.Tk()
        self.g_root.protocol("WM_DELETE_WINDOW", self.root_close)
        self.__make_main_panel()
        self.g_root.mainloop()

    def __make_main_panel(self):
        self.g_root.title("Main Dialog")
        self.g_root.geometry("640x350")
        self.g_root.resizable(False, False)
        self.g_root.iconbitmap("exodus.ico")  # exe 파일 폴더에 .ico 파일을 넣어주어야 한다.

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

        # atr power button
        self.__add_atr("Open", self.atr_power_id)
        # atr network button
        self.__add_atr("Open", self.atr_network_id)
        # atr spectrum button
        self.__add_atr("Open", self.atr_spectrum_id)

    def __add_remote(self, text, _id):
        label_column = 1
        button_column = 2

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__button_clicked, _id),
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
        label_column = 3
        button_column = 4

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__button_clicked,
                                                _id),
                                width=self.main_butt_width,
                                height=self.main_butt_height)
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

    def __button_clicked(self, _id):
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
                self.excel.load_power_atr_excel_procedure(load_path=self.excel_path)
                self.__new_window(_id=self.atr_power_id)
        elif _id == self.atr_network_id:
            pass
        elif _id == self.atr_spectrum_id:
            pass

    def __new_window(self, _id):
        new_window = tkinter.Toplevel(self.g_root)
        if _id == self.remote_power_id:
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
            new_window.geometry("800x500")
            new_window.title("Power ATR")
            new_window.iconbitmap("exodus.ico")
            new_window.resizable(False, False)

            entry_width = 15

            # line_0
            # panel label
            label_panel = tkinter.Label(new_window, text="", anchor="w")
            label_panel.grid(row=0, column=0, columnspan=5, sticky="NEWS", padx=10, pady=10)

            # line_1
            # frequency label
            label = tkinter.Label(new_window, text="Frequency")
            label.grid(row=1, column=0, sticky="NEWS")
            # input offset label
            label = tkinter.Label(new_window, text="Input offset")
            label.grid(row=1, column=1, sticky="NEWS")
            # output offset label
            label = tkinter.Label(new_window, text="Output offset")
            label.grid(row=1, column=2, sticky="NEWS")
            # Power voltage label
            label = tkinter.Label(new_window, text="Power voltage")
            label.grid(row=1, column=3, sticky="NEWS")
            # Power current label
            label = tkinter.Label(new_window, text="Power current")
            label.grid(row=1, column=4, sticky="NEWS")
            # Aging time label
            label = tkinter.Label(new_window, text="Aging time")
            label.grid(row=1, column=5, sticky="NEWS")

            # line_2
            # frequency entry
            self.atr_freq = tkinter.StringVar(new_window)
            if self.excel.freq_var[0] >= 1000000000000:
                self.atr_freq.set("{0} THz".format(self.excel.freq_var[0]/1000000000000))
            elif self.excel.freq_var[0] >= 1000000000:
                self.atr_freq.set("{0} GHz".format(self.excel.freq_var[0]/1000000000))
            elif self.excel.freq_var[0] >= 1000000:
                self.atr_freq.set("{0} MHz".format(self.excel.freq_var[0]/1000000))
            elif self.excel.freq_var[0] >= 1000:
                self.atr_freq.set("{0} KHz".format(self.excel.freq_var[0]/1000))
            else:
                self.atr_freq.set("{0} Hz".format(self.excel.freq_var[0]))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_freq,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=0, sticky="NEWS")
            # input offset entry
            self.atr_input_offset = tkinter.StringVar(new_window)
            self.atr_input_offset.set("{0} dB".format(self.excel.input_offset_var[0]))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_input_offset,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=1, sticky="NEWS")
            # output offset entry
            self.atr_output_offset = tkinter.StringVar(new_window)
            self.atr_output_offset.set("{0} dB".format(self.excel.output_offset_var[0]))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_output_offset,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=2, sticky="NEWS")
            # Power voltage entry
            self.atr_power_voltage = tkinter.StringVar(new_window)
            self.atr_power_voltage.set("{0} V".format(self.excel.power_voltage_var))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_power_voltage,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=3, sticky="NEWS")
            # Power current entry
            self.atr_power_current = tkinter.StringVar(new_window)
            self.atr_power_current.set("{0} A".format(self.excel.power_current_var))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_power_current,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=4, sticky="NEWS")
            # Aging time entry
            self.atr_aging_time = tkinter.StringVar(new_window)
            self.atr_aging_time.set("{0} sec".format(self.excel.aging_var))
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_aging_time,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=2, column=5, sticky="NEWS")

            # line_3
            # rf output label
            label = tkinter.Label(new_window, text="RF Output    ", anchor="e")
            label.grid(row=3, column=0, sticky="NEWS", columnspan=3)
            # current label
            label = tkinter.Label(new_window, text="Current")
            label.grid(row=3, column=3, sticky="NEWS")
            # aging time left label
            label = tkinter.Label(new_window, text="Aging left")
            label.grid(row=3, column=4, sticky="NEWS")

            # line_4
            # RF output entry
            self.atr_rf_output = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_rf_output,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=4, column=0, sticky="NES", columnspan=3)
            # Current entry
            self.atr_current = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_current,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=4, column=3, sticky="NEWS")
            # aging time left entry
            self.atr_aging_left = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_aging_left,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=4, column=4, sticky="NEWS")

            # line_5
            # rf output label
            label = tkinter.Label(new_window, text="  System option", anchor="w")
            label.grid(row=5, column=0, sticky="NEWS", columnspan=2)
            # System Forward Label
            label = tkinter.Label(new_window, text="Sys Fwd")
            label.grid(row=5, column=2, sticky="NEWS")
            # System Frequency Label
            label = tkinter.Label(new_window, text="Sys Freq")
            label.grid(row=5, column=3, sticky="NEWS")
            # System Temperature Label
            label = tkinter.Label(new_window, text="Sys Temp")
            label.grid(row=5, column=4, sticky="NEWS")

            # line_6
            # sys option checkbox
            checkbox = tkinter.Checkbutton(new_window,
                                           text="                      ",
                                           command=self.atr_sys_opt_clicked)
            checkbox.grid(row=6, column=0, columnspan=2, sticky="NEWS")
            # sys fwd entry
            self.atr_sys_fwd = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_sys_fwd,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=6, column=2, sticky="EW")
            # sys freq entry
            self.atr_sys_freq = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_sys_freq,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=6, column=3, sticky="EW")
            # sys freq temp
            self.atr_sys_temp = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_sys_temp,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=6, column=4, sticky="EW")

            # line 7
            # System cmd Label
            label = tkinter.Label(new_window, text="Sys cmd", anchor="w")
            label.grid(row=7, column=0, sticky="NEWS", columnspan=2)
            # Source communication label
            label = tkinter.Label(new_window, text="Source com")
            label.grid(row=7, column=2, sticky="NEWS")
            # Power supply communication label
            label = tkinter.Label(new_window, text="Power com")
            label.grid(row=7, column=3, sticky="NEWS")
            # Power meter communication label
            label = tkinter.Label(new_window, text="Power_m com")
            label.grid(row=7, column=4, sticky="NEWS")

            # line 8
            # System cmd text
            self.atr_sys_cmd = tkinter.StringVar(new_window)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_sys_cmd,
                                  width=entry_width*2,
                                  justify="center")
            entry.grid(row=8, column=0, sticky="EW", columnspan=2)
            # source communication entry
            self.atr_source_com = tkinter.StringVar(new_window)
            self.atr_source_com.set(self.excel.source_com_var)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_source_com,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=8, column=2, sticky="EW")
            # power supply communication entry
            self.atr_power_com = tkinter.StringVar(new_window)
            self.atr_power_com.set(self.excel.power_com_var)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_power_com,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=8, column=3, sticky="EW")
            # power meter communication entry
            self.atr_power_meter_com = tkinter.StringVar(new_window)
            self.atr_power_meter_com.set(self.excel.power_meter_com_var)
            entry = tkinter.Entry(new_window,
                                  textvariable=self.atr_power_meter_com,
                                  width=entry_width,
                                  state="readonly",
                                  justify="center")
            entry.grid(row=8, column=4, sticky="EW")

            # line 9
            # ATR start button
            button = tkinter.Button(new_window,
                                    text="START",
                                    width=self.main_butt_width,
                                    height=self.main_butt_height,
                                    command=partial(self.__atr_button_clicked, self.atr_start_button_id))
            button.grid(row=9, column=0, sticky="NEWS", columnspan=2)
            # ATR stop button
            button = tkinter.Button(new_window,
                                    text="STOP",
                                    width=self.main_butt_width,
                                    height=self.main_butt_height,
                                    command=partial(self.__atr_button_clicked, self.atr_stop_button_id))
            button.grid(row=9, column=2, sticky="NEWS", columnspan=2)

            # line 10
            # ATR start button
            button = tkinter.Button(new_window,
                                    text="DISPLAY",
                                    width=self.main_butt_width,
                                    height=self.main_butt_height,
                                    command=partial(self.__atr_button_clicked, self.atr_display_button_id))
            button.grid(row=10, column=0, sticky="NEWS", columnspan=2)
            # ATR stop button
            button = tkinter.Button(new_window,
                                    text="SAVE",
                                    width=self.main_butt_width,
                                    height=self.main_butt_height,
                                    command=partial(self.__atr_button_clicked, self.atr_save_button_id))
            button.grid(row=10, column=2, sticky="NEWS", columnspan=2)
        elif _id == self.atr_network_id:
            pass
        elif _id == self.atr_spectrum_id:
            pass

    def __atr_button_clicked(self, _id):
        if _id == self.atr_start_button_id:
            self.atr_start_clicked()
        elif _id == self.atr_stop_button_id:
            self.atr_stop_clicked()
        elif _id == self.atr_display_button_id:
            self.atr_display_clicked(rf_var=self.atr_rf_var,
                                     curr_var=self.atr_current_var)
        elif _id == self.atr_save_button_id:
            # save excel dialog open
            self.save_file_dialog()
            if self.excel_path != "-1":
                self.excel.save_all_var(save_path=self.excel_path,
                                        rf_var=self.atr_rf_var,
                                        curr_var=self.atr_current_var)
        else:
            print("atr button id error")

    def atr_start_clicked(self):
        self.excel.aging_done_var = False
        self.__atr_p1_data_procedure()

    def atr_stop_clicked(self):
        pass

    def atr_display_clicked(self, rf_var, curr_var):
        if (len(rf_var) <= 0) and (len(curr_var) <= 0):
            pg.alert(text="No data exist",
                     title="Error",
                     button="확인")
            return
        if (len(rf_var) != len(self.excel.freq_var)) or (len(rf_var) <= 0):
            pg.alert(text="RF data error\n" + "ex) ATR not started",
                     title="Error",
                     button="확인")
            return
        if (len(curr_var) != len(self.excel.freq_var)) or (len(rf_var) <= 0):
            pg.alert(text="Current data error\n" + "ex) ATR not started",
                     title="Error",
                     button="확인")
            return

    def atr_sys_opt_clicked(self):
        if self.atr_system_option:
            self.atr_system_option = False
        else:
            self.atr_system_option = True

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

    def __atr_p1_data_procedure(self):
        self.atr_state = "start"
        self.g_root.after(500, self.__atr_p1_clock_callback)

    def __atr_input_data_procedure(self):
        self.atr_state = "start"
        self.g_root.after(500, self.__atr_input_clock_callback)

    def __atr_p_sat_data_procedure(self):
        self.atr_state = "start"
        self.g_root.after(500, self.__atr_p_sat_clock_callback)

    def __atr_overdrive_data_procedure(self):
        self.atr_state = "start"
        self.g_root.after(500, self.__atr_overdrive_clock_callback)

    def __atr_check_p1_in_p_sat_ov(self):
        pass

    def __atr_p1_clock_callback(self):
        # 1. start
        # 2. power set
        # 3. power meter set
        # 4. source set
        # 5. aging set
        # 6. p1 check -> keep until p1 be done
        # 7. p1 done

        # 1. start
        if self.atr_state == "start":
            self.__atr_inst_call()
            # self.__atr_get_idn()  # test idn
            self.atr_state = "power_set"
            self.atr_index = 0
            self.atr_p1_ref_input = 0
            self.atr_p1_var.clear()

        # 2. power set
        elif self.atr_state == "power_set":
            self.inst_power.write_instrument(command="OUTP OFF")
            time.sleep(1)
            self.inst_power.write_instrument(
                command="VOLT:LEV {0}".format(self.excel.power_voltage_var))  # power voltage set
            time.sleep(1)
            self.inst_power.write_instrument(
                command="CURR:LEV {0}".format(self.excel.power_current_limit_var))  # power current limit set
            time.sleep(1)
            self.inst_power.write_instrument(command="OUTP ON")
            time.sleep(1)
            self.atr_state = "power_meter_set"
        # 3. power meter set
        elif self.atr_state == "power_meter_set":
            _send_freq = self.excel.select_freq_var[self.atr_index]
            _send_offset = float(self.__find_offset_in_table(send_freq=_send_freq, select="output_offset"))
            if _send_offset is not None:
                if _send_offset < 0:
                    _send_offset -= _send_offset
                self.inst_power_meter.write_instrument(
                    "SENS{0}:FREQ {1}HZ".format(self.excel.power_meter_probe_var, _send_freq))
                self.inst_power_meter.write_instrument(
                    "SENS{0}:CORR:LOSS{1} -{2}DB".format(self.excel.power_meter_probe_var,
                                                         self.excel.power_meter_probe_var,
                                                         _send_offset))  # set loss
                self.atr_state = "source_set"
            else:
                self.atr_state = "error"  # error 발생
                return
        # 4. source set
        elif self.atr_state == "source_set":
            _send_freq = self.excel.select_freq_var[self.atr_index]
            _send_offset = float(self.__find_offset_in_table(send_freq=_send_freq, select="input_offset"))
            _send_input = self.excel.atr_start_input_var
            if _send_offset is not None:
                if _send_offset < 0:
                    _send_offset -= _send_offset
                self.inst_source.write_instrument("OUTP:STAT OFF")  # turn off rf state
                time.sleep(1)
                self.inst_source.write_instrument("FREQ {0} Hz".format(_send_freq))  # set frequency
                time.sleep(1)
                self.inst_source.write_instrument("POW:OFFS -{0} DB".format(_send_offset))  # set offset
                time.sleep(1)
                self.inst_source.write_instrument("POW:AMPL {0} dBM".format(_send_input))  # set dBm
                time.sleep(1)
                self.inst_source.write_instrument("OUTP:STAT ON")  # turn off rf state
                time.sleep(1)
            self.atr_state = "aging_set"
        # 5. aging set
        elif self.atr_state == "aging_set":
            if self.excel.aging_done_var:
                self.atr_state = "reference_set"
        elif self.atr_state == "reference_set":
            self.inst_power_meter.write_instrument(command="CALC:REL:STAT OFF")
            time.sleep(1)
            self.inst_power_meter.write_instrument(command="CALC:REL:AUTO ONCE")
            time.sleep(1)
            self.inst_power_meter.write_instrument(command="CALC:REL:STAT OFF")
            print("reference return error")
            self.atr_state = "check_p1"
        # 6. p1 check -> keep until p1 be done
        elif self.atr_state == "check_p1":
            read_power = self.inst_power_meter.query_instrument(
                "FETC{0}?".format(self.excel.power_meter_probe_var))
            read_power = round(float(read_power), 2)
            time.sleep(1)
            adder = Dialog.get_adder(_out_ref=read_power, _input_ref=self.atr_p1_ref_input)
            if adder != 0:
                self.atr_p1_ref_input += adder
                self.inst_source.write_instrument("POW:AMPL {0} dBM".format(
                    self.excel.atr_start_input_var + self.atr_p1_ref_input))  # new input set
                time.sleep(1)
            else:
                self.atr_state = "done_p1"
        # 7. p1 done
        elif self.atr_state == "done_p1":
            self.inst_power_meter.write_instrument(command="OUTP:ROSC:STAT 0")
            time.sleep(1)
            read_power = self.inst_power_meter.query_instrument(
                "FETC{0}?".format(self.excel.power_meter_probe_var))
            time.sleep(1)
            self.atr_p1_var.append(read_power)
            self.atr_index += 1
            if self.atr_index == len(self.excel.freq_var):
                self.inst_source.write_instrument("OUTP:STAT OFF")  # turn off rf state
                time.sleep(1)
                pass
            else:
                self.atr_state = "reference_set"
        else:
            print("p1 id error")

        # time callback ms set
        if self.atr_state != "done_p1":
            if self.atr_state == "aging_set":
                if self.excel.aging_done_var:
                    pass
                else:
                    self.g_root.after(ms=int(self.excel.aging_var * 1000), func=self.__atr_p1_clock_callback)
                    self.excel.aging_done_var = True
            else:
                self.g_root.after(ms=500, func=self.__atr_p1_clock_callback)
        elif self.atr_state == "error":
            pg.alert(text="atr error 발생\n",
                     title="Error",
                     button="확인")
        elif self.atr_state == "done_p1":
            print("p1 done")
            self.__atr_input_data_procedure()
            pass

    def __atr_input_clock_callback(self):
        if self.atr_state == "done_p1":
            pass
        elif self.atr_state == "input_done":
            self.__atr_p_sat_data_procedure()

    def __atr_p_sat_clock_callback(self):
        if self.atr_state == "input_done":
            pass
        elif self.atr_state == "p_sat_done":
            self.__atr_overdrive_data_procedure()

    def __atr_overdrive_clock_callback(self):
        if self.atr_state == "p_sat_done":
            self.__atr_check_p1_in_p_sat_ov()
            pass

    @staticmethod
    def get_adder(_out_ref, _input_ref):
        _key_adder = 0
        compare = _out_ref + 1 - _input_ref
        if compare >= 0.5:      # over 0.5      # 5
            _key_adder = 1      # 1
        elif compare >= 0.25:   # 0.25 ~ 0.5    # 4
            _key_adder = 0.25   # 0.25
        elif compare >= 0.1:    # 0.1 ~ 0.25    # 3
            _key_adder = 0.1    # 0.1
        elif compare >= 0.01:   # 0.01 ~ 0.1    # 2
            _key_adder = 0.05   # 0.05
        elif compare > 0:       # 0 ~ 0.01
            _key_adder = 0.01   # 0.01          # 1
        elif compare == 0:      # 0             # 0
            _key_adder = 0      # 0
        elif compare >= -0.01:  # -0.01 ~ 0     # 1
            _key_adder = -0.01  # -0.01
        elif compare >= -0.1:   # -0.1 ~ -0.01  # 2
            _key_adder = -0.05  # -0.05
        elif compare >= -0.25:  # -0.25 ~ -0.1  #3
            _key_adder = -0.1
        elif compare >= -0.5:   # -0.5 ~ -0.25  # 4
            _key_adder = -0.25
        else:                   # ~ -0.5        # 5
            _key_adder = -1
        return _key_adder

    def __find_offset_in_table(self, send_freq, select):
        searched_index = np.abs(np.array(send_freq)-self.excel.freq_var).argmin()
        if select == "output_offset":
            if send_freq == self.excel.freq_var[searched_index]:
                return self.excel.output_offset_var[searched_index]
            elif send_freq > self.excel.freq_var[searched_index]:
                if searched_index == len(self.excel.freq_var) - 1:  # range 안에 들어오는지 체크
                    print("frequency range over error")
                    return
                else:
                    return self.excel.output_offset_var[searched_index]\
                        + (send_freq - self.excel.freq_var[searched_index])\
                        * (self.excel.output_offset_var[searched_index + 1]
                           - self.excel.output_offset_var[searched_index])\
                        / (self.excel.freq_var[searched_index + 1] - self.excel.freq_var[searched_index])
            elif send_freq < self.excel.freq_var[searched_index]:
                if searched_index == 0:  # range 안에 들어오는지 체크
                    print("frequency range over error")
                    return
                else:
                    return self.excel.output_offset_var[searched_index - 1]\
                        + (send_freq - self.excel.freq_var[searched_index - 1])\
                        * (self.excel.output_offset_var[searched_index]
                           - self.excel.output_offset_var[searched_index - 1])\
                        / (self.excel.freq_var[searched_index] - self.excel.freq_var[searched_index - 1])
        elif select == "input_offset":
            if send_freq == self.excel.freq_var[searched_index]:
                return self.excel.input_offset_var[searched_index]
            elif send_freq > self.excel.freq_var[searched_index]:
                if searched_index == len(self.excel.freq_var) - 1:  # range 안에 들어오는지 체크
                    print("frequency range over error")
                    return
                else:
                    return self.excel.input_offset_var[searched_index] \
                           + (send_freq - self.excel.freq_var[searched_index])\
                           * (self.excel.input_offset_var[searched_index + 1]
                              - self.excel.input_offset_var[searched_index])\
                           / (self.excel.freq_var[searched_index + 1]
                              - self.excel.freq_var[searched_index])
            elif send_freq < self.excel.freq_var[searched_index]:
                if searched_index == 0:  # range 안에 들어오는지 체크
                    print("frequency range over error")
                    return
                else:
                    return self.excel.input_offset_var[searched_index - 1]\
                           + (send_freq - self.excel.freq_var[searched_index - 1])\
                           * (self.excel.input_offset_var[searched_index]
                              - self.excel.input_offset_var[searched_index - 1])\
                           / (self.excel.freq_var[searched_index]
                              - self.excel.freq_var[searched_index - 1])
            else:
                print("error input, output offset call")

    def __atr_inst_call(self):
        self.inst_source = Instrument.Source()
        self.inst_power = Instrument.PowerSupply()
        self.inst_power_meter = Instrument.PowerMeter()

        # source open
        if self.excel.source_comm_option == "GPIB":
            self.inst_source.gpib_address = self.excel.source_com_var
            self.inst_source.open_instrument_gpib(gpib_address=self.inst_source.gpib_address)
        elif self.excel.source_comm_option == "USB":
            pass
        elif self.excel.source_comm_option == "SERIAL":
            pass

        # power supply open
        if self.excel.power_comm_option == "GPIB":
            self.inst_power.gpib_address = self.excel.power_com_var
            self.inst_power.open_instrument_gpib(gpib_address=self.inst_power.gpib_address)
        elif self.excel.power_comm_option == "USB":
            pass
        elif self.excel.power_comm_option == "SERIAL":
            pass

        # power meter open
        if self.excel.power_meter_comm_option == "GPIB":
            self.inst_power_meter.gpib_address = self.excel.power_meter_com_var
            self.inst_power_meter.open_instrument_gpib(gpib_address=self.inst_power_meter.gpib_address)
        elif self.excel.power_meter_comm_option == "USB":
            pass
        elif self.excel.power_meter_comm_option == "SERIAL":
            pass

    def __atr_get_idn(self):
        print(self.inst_source.query_instrument("*IDN?"))
        print(self.inst_power.query_instrument("*IDN?"))
        print(self.inst_power_meter.query_instrument("*IDN?"))

    def atr_error_state(self, text):
        self.atr_state = ""
        print("p1 procedure error")
        pg.alert(text=text+"\n",
                 title="Error",
                 button="확인")
