from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import tkinter
import tkinter.ttk
import pyautogui as pg  # message box library
from tkinter import filedialog as fd  # import file dialog
import os
from functools import partial
import Instrument
import time
import numpy as np
import MyCal


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
        self.power_cur_var = ""
        self.power_m_probe = ""
        self.aging_var = ""
        self.aging_left_var = ""
        self.aging_done_var = False
        self.reference_done = False
        self.source_com_var = ""
        self.power_com_var = ""
        self.power_meter_com_var = ""
        self.spectrum_com_var = ""
        self.network_com_var = ""

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
        self.output_offset_ref_power = ""

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
        self.spectrum_comm_cell_name = ""  # default Power meter
        self.network_comm_cell_name = ""  # default Power meter
        self.select_freq_cell_name = ""  # default Select Freq
        self.atr_start_input_dbm_name = ""  # default ATR start input dBm
        self.p_sat_input_name = ""  # default P_sat input
        self.output_off_ref_name = ""  # default P_sat input
        self.overdrive_input_name = ""  # default Overdrive input
        self.pout_name = ""  # default Pout

        self.source_baud_rate = ""  # default Source generate
        self.power_baud_rate = ""  # default Power supply
        self.power_meter_baud_rate = ""  # default Power meter

        self.source_ip_address = ""  # default Source generate
        self.power_ip_address = ""  # default Power supply
        self.power_ip_address = ""  # default Power meter


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

        self.atr_power_m_id = 6
        self.atr_power_id = 7
        self.atr_network_id = 8
        self.atr_spectrum_id = 9
        self.atr_offset_id = 10
        self.load_config_id = 11

        self.atr_start_button_id = 0
        self.atr_stop_button_id = 1
        self.atr_save_button_id = 2
        self.atr_measure_input_offset_id = 3
        self.atr_measure_output_offset_id = 4
        self.atr_save_all_offset_id = 5
        self.remote_set_button_id = 6
        self.remote_custom_freq_set_button_id = 7

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
        self.atr_rf_input = tkinter.StringVar()
        self.atr_rf_output = tkinter.StringVar()
        self.atr_current = tkinter.StringVar()
        self.atr_aging_left = tkinter.StringVar()
        self.atr_sys_input = tkinter.StringVar()
        self.atr_sys_output = tkinter.StringVar()
        self.atr_sys_curr = tkinter.StringVar()
        self.atr_sys_freq = tkinter.StringVar()
        self.atr_sys_temp = tkinter.StringVar()
        self.atr_sys_cmd = tkinter.StringVar()

        self.remote_p_meter_freq = tkinter.StringVar()
        self.remote_p_meter_power = tkinter.StringVar()
        self.remote_p_meter_offset = tkinter.StringVar()
        self.remote_p_meter_rel_state = tkinter.StringVar()

        self.remote_p_meter_custom_freq = tkinter.StringVar()
        self.remote_p_meter_custom_offset = tkinter.StringVar()

        self.atr_source_com = tkinter.StringVar()
        self.atr_power_com = tkinter.StringVar()
        self.atr_power_meter_com = tkinter.StringVar()

        self.remote_p_meter_freq_combobox = ""

        self.atr_rf_var = []
        self.atr_current_var = []

        self.atr_state = ""
        self.atr_p1_ref_input = 0
        self.atr_input_buff = ""
        self.atr_p_sat_input = ""
        self.atr_index = 0

        self.idq_var = []
        self.atr_p1_var = []
        self.atr_input_var = []
        self.atr_input_curr_var = []
        self.atr_p_sat_var = []
        self.atr_p_sat_curr_var = []
        self.atr_overdrive_var = []
        self.atr_overdrive_curr_var = []

        self.mes_input_offset = []
        self.mes_input_offset_under_2 = []
        self.mes_output_offset = []

        self.inst_source = Instrument.Source()
        self.inst_power = Instrument.PowerSupply()
        self.inst_power_meter = Instrument.PowerMeter()
        self.inst_network = Instrument.Network()
        self.inst_spectrum = Instrument.Spectrum()

        self.atr_stop = False

        self.atr_seq = 0
        self.rel_count = 0
        self.rel_count_limit = 3

        self.sort_count_out = 0
        self.compare_count_out = 0
        self.sort_count_out_limit = 10
        self.compare_count_limit = 10

        self.mes_output_offset_buff = ""
        self.atr_comp_sampling = 2  # 2 times per list

        self.adder = 0.0
        self.adder_input_direction = ""
        self.adder_input_dir_count = 0

        self.after_time_ms_250 = 250
        self.after_time_ms_500 = 500
        self.after_time_ms_1000 = 1000
        self.after_time_ms_2000 = 2000
        self.after_time_ms_3000 = 3000

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
        self.excel_com_unit = ["GPIB", "USB", "Serial", "Ethernet"]
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
        self.excel_spectrum_column = "K"
        self.excel_network_com_column = "L"
        self.excel_select_freq_column = "M"
        self.excel_output_off_ref_column = "N"
        self.excel_atr_start_input_column = "O"
        self.excel_p_sat_input_column = "P"
        self.excel_overdrive_input_column = "Q"
        self.excel_pout_column = "R"
        self.save_hor = True
        self.excel_load_done = False


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

    # frequency, io offset, voltage get from work sheet 0

    def __check_config(self, option):
        self.__get_unit_from_column(option)
        if self.load_clear is not True:
            return -1

    def __get_offset_procedure_version_0(self):
        # ?????? self.l_ws_list[0] frequency, input offset, output offset, voltage ????????? ?????? ??????.
        # ????????? ???????????? ????????? ?????????.

        # 0. var reset
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
        # 11. check spectrum communication
        # 12. check network communication
        # 13. check select frequency
        # 14. check output offset ref power
        # 15. check ATR start input dBm
        # 16. check P_sat input dBm
        # 17. check Overdrive input dBm
        # 18. check Pout output dBm
        # 19. load freq
        # 20. load input db
        # 21. load output db
        # 22. load Voltage set
        # 23. load Current limit
        # 24. load power meter probe channel
        # 25. load Start aging time
        # 26. load source generate communication
        # 27. load power supply communication
        # 28. load power meter communication
        # 29. load spectrum communication
        # 30. load network communication
        # 31. load select frequency
        # 32. load output offset ref power
        # 33. load ATR start input dBm
        # 34. load P_sat input dBm
        # 35. load Overdrive input dBm
        # 36. load Pout output dBm

        # 0. var reset
        self.__var_reset()
        # 1. freq config
        if self.__check_config(self.excel_freq_column) == -1:
            return
        # 2. check input offset
        if self.__check_config(self.excel_input_column) == -1:
            return
        # 3. check output offset
        if self.__check_config(self.excel_output_column) == -1:
            return
        # 4. check power voltage set
        if self.__check_config(self.excel_power_voltage_column) == -1:
            return
        # 5. check power current limit
        if self.__check_config(self.excel_power_current_column) == -1:
            return
        # 6. check power meter probe channel
        if self.__check_config(self.excel_power_meter_probe_column) == -1:
            return
        # 7. check start aging time
        if self.__check_config(self.excel_aging_column) == -1:
            return
        # 8. check source generate communication
        if self.__check_config(self.excel_source_com_column) == -1:
            return
        # 9. check power supply communication
        if self.__check_config(self.excel_power_com_column) == -1:
            return
        # 10. check power meter communication
        if self.__check_config(self.excel_power_meter_com_column) == -1:
            return
        # 11. check spectrum communication
        if self.__check_config(self.excel_spectrum_column) == -1:
            return
        # 12. check network communication
        if self.__check_config(self.excel_network_com_column) == -1:
            return
        # 13. check select frequency
        if self.__check_config(self.excel_select_freq_column) == -1:
            return
        # 14. check output offset ref power
        if self.__check_config(self.excel_output_off_ref_column) == -1:
            return
        # 15. check ATR start input dBm
        if self.__check_config(self.excel_atr_start_input_column) == -1:
            return
        # 16. check P_sat input dBm
        if self.__check_config(self.excel_p_sat_input_column) == -1:
            return
        # 17. check Overdrive input dBm
        if self.__check_config(self.excel_overdrive_input_column) == -1:
            return
        # 18. check Pout output dBm
        if self.__check_config(self.excel_pout_column) == -1:
            return
        # 19. load freq
        self.__load_var_from_column(self.excel_freq_column)
        # 20. load input db
        self.__load_var_from_column(self.excel_input_column)
        # 21. load output db
        self.__load_var_from_column(self.excel_output_column)
        # 22. load Voltage set
        self.__load_var_from_column(self.excel_power_voltage_column)
        # 23. load Current limit
        self.__load_var_from_column(self.excel_power_current_column)
        # 24. load power meter probe channel
        self.__load_var_from_column(self.excel_power_meter_probe_column)
        # 25. load Start aging time
        self.__load_var_from_column(self.excel_aging_column)
        # 26. load source generate communication
        self.__load_var_from_column(self.excel_source_com_column)
        # 27. load power supply communication
        self.__load_var_from_column(self.excel_power_com_column)
        # 28. load power meter communication
        self.__load_var_from_column(self.excel_power_meter_com_column)
        # 29. load spectrum communication
        self.__load_var_from_column(self.excel_spectrum_column)
        # 30. load network communication
        self.__load_var_from_column(self.excel_network_com_column)
        # 31. load select frequency
        self.__load_var_from_column(self.excel_select_freq_column)
        # 32. load output offset ref power
        self.__load_var_from_column(self.excel_output_off_ref_column)
        # 33. load ATR start input dBm
        self.__load_var_from_column(self.excel_atr_start_input_column)
        # 34. load P_sat input dBm
        self.__load_var_from_column(self.excel_p_sat_input_column)
        # 35. load Overdrive input dBm
        self.__load_var_from_column(self.excel_overdrive_input_column)
        # 36. load Pout output dBm
        self.__load_var_from_column(self.excel_pout_column)

    def __var_reset(self):
        self.freq_var = []
        self.input_offset_var = []
        self.output_offset_var = []
        self.power_voltage_var = ""
        self.power_current_var = ""
        self.power_cur_var = ""
        self.power_m_probe = ""
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
        self.output_offset_ref_power = ""  # select freq var
        self.atr_start_input_var = 0
        self.p_sat_input_var = 0
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
        self.load_clear = True

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
                pg.alert(text="freq unit ???????????? ??????\n" + "ex) Hz, KHz, MHz, GHz, THz",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_input_column:
            self.input_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_input_unit[0]:
                pg.alert(text="input offset unit ???????????? ??????\n" + "ex) dB",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_output_column:
            self.output_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_output_unit[0]:
                pg.alert(text="output offset unit ???????????? ??????\n" + "ex) dB",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_power_voltage_column:
            self.power_voltage_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_power_unit[0]:
                pg.alert(text="Power voltage unit ???????????? ??????\n" + "ex) V",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_power_current_column:
            self.power_current_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_current_unit[0]:
                pg.alert(text="Power current unit ???????????? ??????\n" + "ex) A",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_power_meter_probe_column:
            self.power_meter_probe_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_power_meter_probe_unit[0]:
                pg.alert(text="Power meter probe channel ???????????? ??????\n" + "ex) Channel",
                         title="Error",
                         button="??????")
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
                pg.alert(text="Start aging time unit ???????????? ??????\n" + "ex) sec, min, hour",
                         title="Error",
                         button="??????")
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
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[3]:
                self.source_comm_option = self.excel_com_unit[3]  # Ethernet
                self.source_baud_rate = self.l_ws_list[0][d_unit_cell + 1].value
            else:
                pg.alert(text="Source communication unit ???????????? ??????\n"
                              + "ex) GPIB, USB, Serial, Ethernet\nSerial: baud rate ??????\nEthernet: ip address ??????",
                         title="Error",
                         button="??????")
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
                pg.alert(text="Power supply communication unit ???????????? ??????\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate ??????",
                         title="Error",
                         button="??????")
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
                pg.alert(text="Power meter communication unit ???????????? ??????\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate ??????",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_spectrum_column:
            self.spectrum_comm_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[0]:
                self.power_meter_comm_option = self.excel_com_unit[0]  # GPIB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[1]:
                self.power_meter_comm_option = self.excel_com_unit[1]  # USB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[2]:
                self.power_meter_comm_option = self.excel_com_unit[2]  # Serial
            else:
                pg.alert(text="Spectrum communication unit ???????????? ??????\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate ??????",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_network_com_column:
            self.network_comm_cell_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[0]:
                self.power_meter_comm_option = self.excel_com_unit[0]  # GPIB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[1]:
                self.power_meter_comm_option = self.excel_com_unit[1]  # USB
            elif self.l_ws_list[0][d_unit_cell].value == self.excel_com_unit[2]:
                self.power_meter_comm_option = self.excel_com_unit[2]  # Serial
            else:
                pg.alert(text="Network communication unit ???????????? ??????\n"
                              + "ex) GPIB, USB, Serial\nSerial baud rate ??????",
                         title="Error",
                         button="??????")
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
                pg.alert(text="Select freq unit ???????????? ??????\n" + "ex) Hz, KHz, MHz, GHz, THz",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_output_off_ref_column:
            self.output_off_ref_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                pg.alert(text="p_sat input dBm unit ???????????? ??????\n" + "ex) dBm",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_atr_start_input_column:
            self.atr_start_input_dbm_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                pg.alert(text="atr start dBm unit ???????????? ??????\n" + "ex) dBm",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_p_sat_input_column:
            self.output_off_ref_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                pg.alert(text="p_sat input dBm unit ???????????? ??????\n" + "ex) dBm",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_overdrive_input_column:
            self.overdrive_input_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="overdrive dBm unit ???????????? ??????\n" + "ex) dBm",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        elif column == self.excel_pout_column:
            self.pout_name = d_name_cell
            if self.l_ws_list[0][d_unit_cell].value != self.excel_dbm_unit[0]:
                print("Probe channel not defined")
                pg.alert(text="overdrive dBm unit ???????????? ??????\n" + "ex) dBm",
                         title="Error",
                         button="??????")
                self.load_clear = False
                return
        else:
            print("???????????? communication column")
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
                self.power_current_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_meter_probe_column:
                self.power_m_probe = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_aging_column:
                self.aging_var = self.l_ws_list[0][_data_cell].value * self.multiple_aging
            elif data_column == self.excel_source_com_column:
                self.source_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_com_column:
                self.power_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_power_meter_com_column:
                self.power_meter_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_spectrum_column:
                self.spectrum_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_network_com_column:
                self.network_com_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_select_freq_column:
                self.select_freq_var.append(self.l_ws_list[0][_data_cell].value * self.multiple_select_freq)
            elif data_column == self.excel_output_off_ref_column:
                self.output_offset_ref_power = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_atr_start_input_column:
                self.atr_start_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_p_sat_input_column:
                self.p_sat_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_overdrive_input_column:
                self.overdrive_input_var = self.l_ws_list[0][_data_cell].value
            elif data_column == self.excel_pout_column:
                self.pout_var = self.l_ws_list[0][_data_cell].value
            else:
                print("Error ???????????? ?????? id")
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
        self.g_root.iconbitmap("exodus.ico")  # exe ?????? ????????? .ico ????????? ??????????????? ??????.

        # remote power meter button
        self.__add_main_cor_1("Open", self.remote_p_meter_id)
        # remote power button
        self.__add_main_cor_1("Open", self.remote_power_id)
        # remote network button
        self.__add_main_cor_1("Open", self.remote_network_id)
        # remote source button
        self.__add_main_cor_1("Open", self.remote_source_id)
        # remote spectrum button
        self.__add_main_cor_1("Open", self.remote_spectrum_id)
        # remote all button
        self.__add_main_cor_1("Open", self.remote_all_id)

        # atr power button
        self.__add_main_cor_2("Open", self.atr_power_id)
        # atr network button
        self.__add_main_cor_2("Open", self.atr_network_id)
        # atr spectrum button
        self.__add_main_cor_2("Open", self.atr_spectrum_id)
        # atr offset button
        self.__add_main_cor_2("Open", self.atr_offset_id)
        # load config button
        self.__add_main_cor_2("Open", self.load_config_id)

    def __add_main_cor_1(self, text, _id):
        label_column = 1
        button_column = 2

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__main_window_button_clicked, _id),
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

    def __add_main_cor_2(self, text, _id):
        label_column = 3
        button_column = 4

        button = tkinter.Button(self.g_root,
                                text=text,
                                command=partial(self.__main_window_button_clicked,
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
        elif _id == self.atr_offset_id:
            _row = 4
            # label
            label = tkinter.Label(self.g_root, text="Offset", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.atr_start_x, pady=self.atr_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.atr_button_x, pady=self.atr_y_gap)
        elif _id == self.load_config_id:
            _row = 5
            # label
            label = tkinter.Label(self.g_root, text="Load config", anchor="w")
            label.grid(row=_row, column=label_column, sticky="NEWS",
                       padx=self.atr_start_x, pady=self.atr_y_gap)
            # button
            button.grid(row=_row, column=button_column, sticky="NEWS",
                        padx=self.atr_button_x, pady=self.atr_y_gap)

        else:
            pass

    def __main_window_button_clicked(self, _id):
        if _id == self.load_config_id:
            self.excel.excel_load_done = False
            self.excel.excel_load_done = False

        if self.excel.excel_load_done is False:
            self.load_file_dialog()
            self.excel.load_clear = True
            if self.excel_path != "-1":
                self.excel.load_power_atr_excel_procedure(load_path=self.excel_path)
                if self.excel.load_clear is True:
                    self.excel.excel_load_done = True
            else:
                self.excel.load_clear = False
        if self.excel.load_clear is True:
            # load excel dialog open
            if _id is None:
                print("unknown window id")
            elif _id == self.remote_p_meter_id:
                self.__new_window_create(_id=self.remote_p_meter_id)
            elif _id == self.remote_power_id:
                self.__new_window_create(_id=self.remote_power_id)
            elif _id == self.remote_network_id:
                self.__new_window_create(_id=self.remote_network_id)
            elif _id == self.remote_source_id:
                self.__new_window_create(_id=self.remote_source_id)
            elif _id == self.remote_spectrum_id:
                self.__new_window_create(_id=self.remote_spectrum_id)
            elif _id == self.remote_all_id:
                self.__new_window_create(_id=self.remote_all_id)
            elif _id == self.atr_power_id:
                self.__new_window_create(_id=self.atr_power_id)
            elif _id == self.atr_network_id:
                pass
            elif _id == self.atr_spectrum_id:
                pass
            elif _id == self.atr_offset_id:
                self.__new_window_create(_id=self.atr_offset_id)

    def __atr_power_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("800x350")
        new_window.title("Power ATR")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

        label_pad_x = (20, 0)
        entry_pad_x = (20, 0)
        sys_pad_y = (20, 0)
        button_pad_y = (30, 0)

        entry_width = 15

        # line_0
        # panel label
        label_panel = tkinter.Label(new_window, text="", anchor="w")
        label_panel.grid(row=0, column=0, columnspan=6, sticky="NEWS", pady=10)
        # line_1
        # frequency label
        label = tkinter.Label(new_window, text="Frequency")
        label.grid(row=1, column=0, sticky="NEWS", padx=label_pad_x)
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
            self.atr_freq.set("{0} THz".format(self.excel.freq_var[0] / 1000000000000))
        elif self.excel.freq_var[0] >= 1000000000:
            self.atr_freq.set("{0} GHz".format(self.excel.freq_var[0] / 1000000000))
        elif self.excel.freq_var[0] >= 1000000:
            self.atr_freq.set("{0} MHz".format(self.excel.freq_var[0] / 1000000))
        elif self.excel.freq_var[0] >= 1000:
            self.atr_freq.set("{0} KHz".format(self.excel.freq_var[0] / 1000))
        else:
            self.atr_freq.set("{0} Hz".format(self.excel.freq_var[0]))
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_freq,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=2, column=0, sticky="NEWS", padx=entry_pad_x)
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
        # rf input label
        label = tkinter.Label(new_window, text="RF Input", anchor="center")
        label.grid(row=3, column=0, sticky="NEWS", padx=label_pad_x)
        # rf output label
        label = tkinter.Label(new_window, text="RF Output", anchor="center")
        label.grid(row=3, column=1, sticky="NEWS")
        # current label
        label = tkinter.Label(new_window, text="Current")
        label.grid(row=3, column=2, sticky="NEWS")
        # aging time left label
        label = tkinter.Label(new_window, text="Aging left")
        label.grid(row=3, column=3, sticky="NEWS")

        # line_4
        # RF input entry
        self.atr_rf_input = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_rf_input,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=4, column=0, sticky="NES")
        # RF output entry
        self.atr_rf_output = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_rf_output,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=4, column=1, sticky="NES")
        # Current entry
        self.atr_current = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_current,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=4, column=2, sticky="NEWS")
        # aging time left entry
        self.atr_aging_left = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_aging_left,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=4, column=3, sticky="NEWS")

        # line_5
        # Source communication label
        label = tkinter.Label(new_window, text="Source com")
        label.grid(row=5, column=0, sticky="NEWS", padx=label_pad_x)
        # Power supply communication label
        label = tkinter.Label(new_window, text="Power com")
        label.grid(row=5, column=1, sticky="NEWS")
        # Power meter communication label
        label = tkinter.Label(new_window, text="Power_m com")
        label.grid(row=5, column=2, sticky="NEWS")

        # line_6
        # source communication entry
        self.atr_source_com = tkinter.StringVar(new_window)
        self.atr_source_com.set(self.excel.source_com_var)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_source_com,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=6, column=0, sticky="EW", padx=entry_pad_x)
        # power supply communication entry
        self.atr_power_com = tkinter.StringVar(new_window)
        self.atr_power_com.set(self.excel.power_com_var)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_power_com,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=6, column=1, sticky="EW")
        # power meter communication entry
        self.atr_power_meter_com = tkinter.StringVar(new_window)
        self.atr_power_meter_com.set(self.excel.power_meter_com_var)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_power_meter_com,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=6, column=2, sticky="EW")

        # line_7
        # System Frequency Label
        label = tkinter.Label(new_window, text="Sys Freq")
        label.grid(row=7, column=0, sticky="NEWS", padx=label_pad_x, pady=sys_pad_y)
        # System Input Label
        label = tkinter.Label(new_window, text="Sys Input")
        label.grid(row=7, column=1, sticky="NEWS", pady=sys_pad_y)
        # System Forward Label
        label = tkinter.Label(new_window, text="Sys Output")
        label.grid(row=7, column=2, sticky="NEWS", pady=sys_pad_y)
        # System Current Label
        label = tkinter.Label(new_window, text="Sys Current")
        label.grid(row=7, column=3, sticky="NEWS", pady=sys_pad_y)
        # System Temperature Label
        label = tkinter.Label(new_window, text="Sys Temp")
        label.grid(row=7, column=4, sticky="NEWS", pady=sys_pad_y)
        # System cmd Label
        label = tkinter.Label(new_window, text="Sys cmd")
        label.grid(row=7, column=5, sticky="NEWS", pady=sys_pad_y)

        # line_8
        # sys freq entry
        self.atr_sys_freq = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_freq,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=8, column=0, sticky="EW", padx=entry_pad_x)
        # sys input entry
        self.atr_sys_input = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_input,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=8, column=1, sticky="EW")
        # sys output entry
        self.atr_sys_output = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_output,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=8, column=2, sticky="EW")
        # sys current
        self.atr_sys_curr = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_curr,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=8, column=3, sticky="EW")
        # sys temperature
        self.atr_sys_temp = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_temp,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=8, column=4, sticky="EW")
        # System cmd entry
        self.atr_sys_cmd = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.atr_sys_cmd,
                              width=entry_width,
                              justify="center")
        entry.grid(row=8, column=5, sticky="EW")

        # line 9
        # ATR start button
        button = tkinter.Button(new_window,
                                text="START",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_start_button_id))
        button.grid(row=9, column=0, sticky="NEWS", columnspan=2, padx=entry_pad_x, pady=button_pad_y)
        # ATR stop button
        button = tkinter.Button(new_window,
                                text="STOP",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_stop_button_id))
        button.grid(row=9, column=2, sticky="NEWS", columnspan=2, pady=button_pad_y)
        # ATR save button
        button = tkinter.Button(new_window,
                                text="SAVE",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_save_button_id))
        button.grid(row=9, column=4, sticky="NEWS", columnspan=2, pady=button_pad_y)

    def __remote_p_meter_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("650x600")
        new_window.title("Remote power meter")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

        label_pad_x = (20, 0)
        entry_pad_x = (20, 0)

        entry_width = 15

        # line_0
        # panel label
        label_panel = tkinter.Label(new_window, text="Power meter Status", anchor="w")
        label_panel.grid(row=0, column=0, columnspan=5, sticky="NEWS", padx=10, pady=20)
        # line_1
        # frequency label
        label = tkinter.Label(new_window, text="Frequency")
        label.grid(row=1, column=1, sticky="NEWS", padx=label_pad_x)
        # RF Power label
        label = tkinter.Label(new_window, text="RF Power")
        label.grid(row=1, column=2, sticky="NEWS", padx=label_pad_x)
        # Offset label
        label = tkinter.Label(new_window, text="Offset")
        label.grid(row=1, column=3, sticky="NEWS", padx=label_pad_x)
        # Rel offset state
        label = tkinter.Label(new_window, text="Rel offset state")
        label.grid(row=1, column=4, sticky="NEWS", padx=label_pad_x)

        # line_2
        # remote power meter freq
        self.remote_p_meter_freq = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_freq,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=2, column=1, padx=entry_pad_x)
        # remote power meter power
        self.remote_p_meter_power = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_power,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=2, column=2, padx=entry_pad_x)
        # remote power meter power
        self.remote_p_meter_offset = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_offset,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=2, column=3, padx=entry_pad_x)
        # remote power meter power
        self.remote_p_meter_rel_state = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_rel_state,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=2, column=4, padx=entry_pad_x)
        # line_3
        # panel label
        label_panel = tkinter.Label(new_window, text="Control", anchor="w")
        label_panel.grid(row=3, column=0, columnspan=5, sticky="NEWS", padx=10, pady=(40, 10))
        # line_4
        # control frequency label
        label = tkinter.Label(new_window, text="Frequency")
        label.grid(row=4, column=1, sticky="NEWS", padx=label_pad_x)
        # set combobox label
        label = tkinter.Label(new_window, text="Set")
        label.grid(row=4, column=2, padx=label_pad_x)
        # line_5
        # combo box
        values = []
        unit = ""
        for data in self.excel.select_freq_var:
            if data >= 1000000000000:
                data /= 1000000000000
                unit = "THz"
            elif data >= 1000000000:
                data /= 1000000000
                unit = "GHz"
            elif data >= 1000000:
                data /= 1000000
                unit = "MHz"
            elif data >= 1000:
                data /= 1000
                unit = "KHz"
            data = "{0} {1}".format(data, unit)
            values.append(data)
        self.remote_p_meter_freq_combobox = tkinter.ttk.Combobox(new_window, values=values, width=entry_width-3)
        self.remote_p_meter_freq_combobox.grid(row=5, column=1, padx=entry_pad_x)
        # set button
        button = tkinter.Button(new_window,
                                text="Set",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.remote_set_button_id))
        button.grid(row=5, column=2, sticky="EW", columnspan=1, padx=entry_pad_x)
        # line_6
        # panel label
        label_panel = tkinter.Label(new_window, text="Custom", anchor="w")
        label_panel.grid(row=6, column=0, columnspan=5, sticky="NEWS", padx=10, pady=(40, 10))
        # line_7
        # custom label line_0
        label = tkinter.Label(new_window, text="Frequency")
        label.grid(row=7, column=1, sticky="NEWS", padx=label_pad_x)
        label = tkinter.Label(new_window, text="Unit")
        label.grid(row=7, column=2, sticky="NEWS", padx=label_pad_x)
        label = tkinter.Label(new_window, text="Set")
        label.grid(row=7, column=3, sticky="NEWS", padx=label_pad_x)
        # line_8
        # custom wizet line_0
        # custom freq entry
        self.remote_p_meter_custom_freq = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_custom_freq,
                              width=entry_width,
                              justify="center")
        entry.grid(row=8, column=1, padx=entry_pad_x)
        # unit combo box
        unit = ["GHz", "MHz", "KHz", "Hz"]
        self.remote_p_meter_freq_combobox = tkinter.ttk.Combobox(new_window, values=unit, width=entry_width-3)
        self.remote_p_meter_freq_combobox.grid(row=8, column=2, padx=entry_pad_x)
        # set button
        button = tkinter.Button(new_window,
                                text="Set",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked,
                                                self.remote_custom_freq_set_button_id))
        button.grid(row=8, column=3, sticky="EW", columnspan=1, padx=entry_pad_x)
        # line_9
        # custom label line_1
        label = tkinter.Label(new_window, text="Offset")
        label.grid(row=9, column=1, sticky="NEWS", padx=label_pad_x)
        label = tkinter.Label(new_window, text="Unit")
        label.grid(row=9, column=2, sticky="NEWS", padx=label_pad_x)
        label = tkinter.Label(new_window, text="Set")
        label.grid(row=9, column=3, sticky="NEWS", padx=label_pad_x)
        # line_10
        # custom wizet line_1
        # custom offset entry
        self.remote_p_meter_custom_offset = tkinter.StringVar(new_window)
        entry = tkinter.Entry(new_window,
                              textvariable=self.remote_p_meter_custom_offset,
                              width=entry_width,
                              justify="center")
        entry.grid(row=10, column=1, padx=entry_pad_x)
        # unit entry
        _var = tkinter.StringVar(new_window)
        _var.set("dB")
        entry = tkinter.Entry(new_window,
                              textvariable=_var,
                              width=entry_width,
                              state="readonly",
                              justify="center")
        entry.grid(row=10, column=2, padx=entry_pad_x)
        button = tkinter.Button(new_window,
                                text="Set",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked,
                                                self.remote_custom_freq_set_button_id))
        button.grid(row=10, column=3, sticky="EW", columnspan=1, padx=entry_pad_x)
        # line 11
        # custom label line_2
        label = tkinter.Label(new_window, text="Rel Set")
        label.grid(row=11, column=1, sticky="NEWS", padx=label_pad_x)
        label = tkinter.Label(new_window, text="Rel Off")
        label.grid(row=11, column=2, sticky="NEWS", padx=label_pad_x)
        # line_12
        # custom wizet line_2
        # custom offset button
        button = tkinter.Button(new_window,
                                text="Set",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked,
                                                self.remote_custom_freq_set_button_id))
        button.grid(row=12, column=1, sticky="EW", columnspan=1, padx=entry_pad_x)
        button = tkinter.Button(new_window,
                                text="Set",
                                width=self.main_butt_width,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked,
                                                self.remote_custom_freq_set_button_id))
        button.grid(row=12, column=2, sticky="EW", columnspan=1, padx=entry_pad_x)

    def __remote_power_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("350x200")
        new_window.title("Remote power supply")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

    def __remote_network_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("350x200")
        new_window.title("Remote network")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

    def __remote_source_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("350x200")
        new_window.title("Remote source")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

    def __remote_spectrum_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("350x200")
        new_window.title("Remote spectrum")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

    def __remote_all_new_window(self):
        self.__remote_p_meter_new_window()
        self.__remote_power_new_window()
        self.__remote_network_new_window()
        self.__remote_source_new_window()
        self.__remote_spectrum_new_window()

    def __offset_new_window(self):
        new_window = tkinter.Toplevel(self.g_root)
        new_window.geometry("350x200")
        new_window.title("Offset")
        new_window.iconbitmap("exodus.ico")
        new_window.resizable(False, False)

        # input offset start button
        button = tkinter.Button(new_window,
                                text="Measure Input",
                                width=self.main_butt_width * 3,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_measure_input_offset_id))
        button.grid(row=0, column=0, sticky="NEWS", columnspan=1, padx=30, pady=15)

        # output offset start button
        button = tkinter.Button(new_window,
                                text="Measure Output",
                                width=self.main_butt_width * 3,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_measure_output_offset_id))
        button.grid(row=1, column=0, sticky="NEWS", columnspan=1, padx=30, pady=15)

        # save all offset var button
        button = tkinter.Button(new_window,
                                text="Save All",
                                width=self.main_butt_width * 3,
                                height=self.main_butt_height,
                                command=partial(self.__new_window_button_clicked, self.atr_save_all_offset_id))
        button.grid(row=2, column=0, sticky="NEWS", columnspan=1, padx=30, pady=15)

    def __new_window_create(self, _id):
        if _id == self.remote_p_meter_id:
            self.__remote_p_meter_new_window()
        elif _id == self.remote_power_id:
            self.__remote_power_new_window()
        elif _id == self.remote_network_id:
            self.__remote_network_new_window()
        elif _id == self.remote_source_id:
            self.__remote_source_new_window()
        elif _id == self.remote_spectrum_id:
            self.__remote_spectrum_new_window()
        elif _id == self.remote_all_id:
            self.__remote_all_new_window()
        elif _id == self.atr_power_id:
            self.__atr_power_new_window()
        elif _id == self.atr_network_id:
            pass
        elif _id == self.atr_spectrum_id:
            pass
        elif _id == self.atr_offset_id:
            self.__offset_new_window()
        elif _id == self.load_config_id:  # nothing to do
            pass
        else:
            print("unknown new window id")

    def __new_window_button_clicked(self, _id):
        if _id == self.atr_start_button_id:
            self.atr_start_clicked()
        elif _id == self.atr_stop_button_id:
            self.atr_stop_clicked()
        elif _id == self.atr_save_button_id:
            # save excel dialog open
            self.save_file_dialog(_id=self.atr_save_button_id)
            if self.excel_path != "-1":
                self.save_all_var(save_path=self.excel_path)
        elif _id == self.atr_measure_input_offset_id:
            self.measure_offset_clicked(_id=_id)
        elif _id == self.atr_measure_output_offset_id:
            self.measure_offset_clicked(_id=_id)
        elif _id == self.atr_save_all_offset_id:
            self.save_all_offset_clicked()
        elif _id == self.remote_set_button_id:
            self.remote_p_meter_clicked()
        else:
            print("atr button id error")

    def offset_procedure_input(self):
        source_power = 0
        # 0. set var
        # 1. set instrument instance
        # 2. reset var
        # 3. reset instrument
        # 4. set instrument

        # 0. set var
        _send_freq = int(self.excel.select_freq_var[self.atr_index])
        _send_offset_i = 0
        _send_offset_o = 0
        if _send_offset_i is not None:
            if _send_offset_i < 0:
                _send_offset_i = -_send_offset_i
        if _send_offset_o is not None:
            if _send_offset_o < 0:
                _send_offset_o = -_send_offset_o
        # 1. set instrument
        self.atr_seq += 1
        if self.atr_seq == 1:
            # 1. set instrument instance
            self.__atr_inst_call(source=True, p_meter=True)
            # 2. reset var
            self._reset_input_offset_var()
        # 3. reset instrument_0
        elif self.atr_seq == 2:
            self.inst_source.set_output(on_off=False)
            self.inst_power_meter.set_rel(on_off=False)  # set power meter rel off
        # 4. reset instrument_1
        elif self.atr_seq == 3:
            self.inst_source.set_offset_state(state=True)
            self.inst_power_meter.set_offset_state(state=True, ch=self.excel.power_m_probe)
        # 5. set instrument
        elif self.atr_seq == 4:
            self.inst_source.set_offset(offset=0)
            self.inst_power_meter.set_offset(offset=0, ch=self.excel.power_m_probe)  # set power meter offset
        elif self.atr_seq == 5:
            self.inst_source.set_dbm(dbm=source_power)
        elif self.atr_seq == 6:
            self.inst_source.set_output(on_off=True)
        elif self.atr_seq == 7:
            self.inst_source.set_freq(freq=_send_freq)
            self.inst_power_meter.set_freq(freq=_send_freq, ch=self.excel.power_m_probe)  # set power_m freq
        elif self.atr_seq == 8:
            ret = self.inst_power_meter.get_output(display_ch=1, round_num=2)
            self.mes_input_offset.append(round(ret, 2))
            self.mes_input_offset_under_2.append(ret)
            if len(self.mes_input_offset) == len(self.excel.select_freq_var):
                # save excel dialog open
                self.inst_source.set_output(on_off=False)  # set source off
                pg.alert(text="Done",
                         title="Done",
                         button="??????")
                return
            else:  # next freq
                self.atr_index += 1
                self.atr_seq -= 2
        else:
            print("input offset seq error")
        after_time = self.set_after_time_input_offset()
        self.after_call(ms=after_time, func=self.offset_procedure_input)

    def set_after_time_input_offset(self):
        ret = 0
        if self.atr_seq == 0:
            ret = self.after_time_ms_250
        elif self.atr_seq == 1:
            ret = self.after_time_ms_250
        elif self.atr_seq == 2:
            ret = self.after_time_ms_250
        elif self.atr_seq == 3:
            ret = self.after_time_ms_250
        elif self.atr_seq == 4:
            ret = self.after_time_ms_250
        elif self.atr_seq == 5:
            ret = self.after_time_ms_250
        elif self.atr_seq == 6:
            ret = self.after_time_ms_250
        elif self.atr_seq == 7:
            ret = self.after_time_ms_2000  # ready for fetch
        elif self.atr_seq == 8:
            ret = self.after_time_ms_250
        return ret

    def offset_procedure_output(self):
        source_power = self.excel.output_offset_ref_power
        # 0. set var
        # 1. set instrument instance
        # 2. reset var
        # 3. reset instrument
        # 4. set instrument

        # 0. set var
        _send_freq = int(self.excel.select_freq_var[self.atr_index])
        _send_offset_i = 0
        _send_offset_o = 0
        if _send_offset_i is not None:
            if _send_offset_i < 0:
                _send_offset_i = -_send_offset_i
        if _send_offset_o is not None:
            if _send_offset_o < 0:
                _send_offset_o = -_send_offset_o
        # 1. set instrument
        self.atr_seq += 1
        if self.atr_seq == 1:
            # 1. set instrument instance
            self.__atr_inst_call(source=True, p_meter=True)
            # 2. reset var
            self._reset_output_offset_var()
        # 3. reset instrument_0
        elif self.atr_seq == 2:
            self.inst_source.set_output(on_off=False)  # set source off
            self.inst_power_meter.set_rel(on_off=False)  # set power meter rel off
        # 4. reset instrument_1
        elif self.atr_seq == 3:
            self.inst_source.set_offset_state(state=True)
            self.inst_power_meter.set_offset_state(state=True, ch=self.excel.power_m_probe)
        # 4. set instrument
        elif self.atr_seq == 4:
            if len(self.mes_input_offset) != 0:
                self.inst_source.set_offset(offset=round(self.mes_input_offset[self.atr_index], 2))  # set source offset
            else:
                self.inst_source.set_offset(offset=0)  # set source offset
            self.inst_power_meter.set_offset(offset=0, ch=self.excel.power_m_probe)  # set power meter offset
        elif self.atr_seq == 5:
            self.inst_source.set_dbm(dbm=source_power)  # set dBm
        elif self.atr_seq == 6:
            self.inst_source.set_output(on_off=True)  # set source on
        elif self.atr_seq == 7:
            self.inst_source.set_freq(freq=_send_freq)  # set source frequency
            self.inst_power_meter.set_freq(freq=_send_freq, ch=self.excel.power_m_probe)  # set power_m freq
        elif self.atr_seq == 8:
            self.mes_output_offset.append(round(self.inst_power_meter.get_output(display_ch=1, round_num=2)
                                                - self.excel.output_offset_ref_power, 2))
            if len(self.mes_output_offset) == len(self.excel.select_freq_var):
                # save excel dialog open
                self.inst_source.set_output(on_off=False)  # set source off
                pg.alert(text="Done",
                         title="Done",
                         button="??????")
                return
            else:  # next freq
                self.atr_index += 1
                self.atr_seq -= 5
        else:
            print("output offset atr seq error")
        after_time = self.set_after_time_output_offset()
        self.after_call(ms=after_time, func=self.offset_procedure_output)

    def set_after_time_output_offset(self):
        ret = 0
        if self.atr_seq == 0:
            ret = self.after_time_ms_250
        elif self.atr_seq == 1:
            ret = self.after_time_ms_250
        elif self.atr_seq == 2:
            ret = self.after_time_ms_250
        elif self.atr_seq == 3:
            ret = self.after_time_ms_250
        elif self.atr_seq == 4:
            ret = self.after_time_ms_250
        elif self.atr_seq == 5:
            ret = self.after_time_ms_250
        elif self.atr_seq == 6:
            ret = self.after_time_ms_250
        elif self.atr_seq == 7:
            ret = self.after_time_ms_3000  # ready for fetch
        elif self.atr_seq == 8:
            ret = self.after_time_ms_250
        return ret

    def offset_start(self, offset):
        if offset == "input":
            self.offset_procedure_input()
        elif offset == "output":
            self.offset_procedure_output()
        else:
            print("offset argument error")

    def save_all_offset_clicked(self):
        # save excel dialog open
        self.save_file_dialog(_id=self.atr_save_all_offset_id)
        if self.excel_path != "-1":
            self.save_option(save_path=self.excel_path, _id=self.atr_save_all_offset_id)

    def remote_p_meter_clicked(self):
        print("{0}".format(self.remote_p_meter_freq_combobox.get()))

    def measure_offset_clicked(self, _id):
        self.atr_seq = 0
        if _id == self.atr_measure_input_offset_id:
            self.offset_start(offset="input")
        elif _id == self.atr_measure_output_offset_id:
            self.offset_start(offset="output")

    def atr_start_clicked(self):
        self.atr_seq = 0
        self.atr_ready()  # set config, power on, rf on

    def atr_ready(self):
        # 0. set var
        # 1. set instrument instance
        # 2. reset atr var
        # 3. reset instrument
        # 4. set instrument

        # 0. set var
        _send_freq = int(self.excel.select_freq_var[self.atr_index])
        _send_offset_i = float(self.__find_offset_in_table(send_freq=_send_freq, select="input_offset"))
        _send_offset_o = float(self.__find_offset_in_table(send_freq=_send_freq, select="output_offset"))
        if (_send_offset_i is None) or (_send_offset_o is None):
            print("offset error 0")
            return
        # set instrument
        self.atr_seq += 1
        if self.atr_seq == 1:
            # 1. set instrument instance
            self.__atr_inst_call(source=True, p_meter=True, power=True)
            # 2. reset atr var
            self._reset_atr_var()
        elif self.atr_seq == 2:  # set config_0
            self.inst_source.set_output(on_off=False)  # turn off source
            self.inst_power_meter.set_rel(on_off=False)  # reference off
            self.inst_power.set_output(on_off=False)  # turn off power
        elif self.atr_seq == 3:  # set config_1
            self.inst_source.set_offset_state(state=True)
            self.inst_power_meter.set_offset_state(state=True, ch=self.excel.power_m_probe)
        elif self.atr_seq == 4:  # set config_2
            self.inst_power.set_voltage(voltage=self.excel.power_voltage_var)  # power voltage set
            self.dialog_var_set(var_name="atr_power_voltage", value=self.excel.power_voltage_var)  # set dialog var
            self.inst_source.set_freq(freq=_send_freq)  # set frequency
            self.inst_power_meter.set_freq(freq=_send_freq, ch=self.excel.power_m_probe)  # set frequency
            self.dialog_var_set(var_name="atr_freq", value=_send_freq)  # set dialog var
        elif self.atr_seq == 5:  # set config_3
            self.inst_power.set_current(current=self.excel.power_current_var)  # power current set
            self.dialog_var_set(var_name="atr_power_current", value=self.excel.power_current_var)  # set dialog var
            self.inst_source.set_offset(offset=_send_offset_i)  # set offset
            self.inst_power_meter.set_offset(offset=_send_offset_o, ch=self.excel.power_m_probe)  # set loss
            self.dialog_var_set(var_name="atr_input_offset", value=_send_offset_i)  # set dialog var
            self.dialog_var_set(var_name="atr_output_offset", value=_send_offset_o)  # set dialog var
        elif self.atr_seq == 6:  # set config_3
            self.inst_source.set_dbm(self.excel.atr_start_input_var)  # set dBm
            self.dialog_var_set(var_name="atr_rf_input", value=self.excel.atr_start_input_var)  # set dialog var
        elif self.atr_seq == 7:  # release_0
            self.inst_power.set_output(on_off=True)  # turn on power
        elif self.atr_seq == 8:
            self.idq_var.append(self.inst_power.get_current(round_num=2))
        elif self.atr_seq == 9:  # release_1
            self.inst_source.set_output(on_off=True)  # turn on rf state
        else:
            self.atr_seq = 0
            self.after_call(ms=self.after_time_ms_250, func=self.aging())
            return
        after_time = self.after_call_atr_ready()
        self.after_call(ms=after_time, func=self.atr_ready)

    def after_call_atr_ready(self):
        ret = 0
        if self.atr_seq == 0:
            ret = self.after_time_ms_250
        elif self.atr_seq == 1:
            ret = self.after_time_ms_250
        elif self.atr_seq == 2:
            ret = self.after_time_ms_250
        elif self.atr_seq == 3:
            ret = self.after_time_ms_250
        elif self.atr_seq == 4:
            ret = self.after_time_ms_250
        elif self.atr_seq == 5:
            ret = self.after_time_ms_250
        elif self.atr_seq == 6:
            ret = self.after_time_ms_1000
        elif self.atr_seq == 7:
            ret = self.after_time_ms_1000
        elif self.atr_seq == 8:
            ret = self.after_time_ms_250
        elif self.atr_seq == 9:
            ret = self.after_time_ms_250
        else:
            print("after time atr ready error")
        return ret

    def aging(self):
        self.atr_seq = 0
        self.after_call(ms=self.after_time_ms_250, func=self.atr_start())

    def set_after_time_aging(self):
        ret = 0
        if self.atr_seq == 0:
            ret = self.after_time_ms_250
        elif self.atr_seq == 1:
            ret = self.after_time_ms_250
        elif self.atr_seq == 2:
            ret = self.after_time_ms_250
        elif self.atr_seq == 3:
            ret = self.after_time_ms_250
        elif self.atr_seq == 4:
            ret = self.after_time_ms_250
        elif self.atr_seq == 5:
            ret = self.after_time_ms_250
        elif self.atr_seq == 6:
            ret = self.after_time_ms_250
        elif self.atr_seq == 7:
            ret = self.after_time_ms_2000  # ready for fetch
        elif self.atr_seq == 8:
            ret = self.after_time_ms_250
        return ret

    def atr_start(self):
        # 0 atr index check
        # 1. set var
        # 2. set instrument default
        # 3. p1 ready
        # 4. p1 get
        # 5. input ready
        # 6. input get
        # 7. p_sat ready
        # 8. p_sat get
        # 9. ready overdrive
        # 10. get overdrive
        # 11. atr index increase
        # 12. atr_seq increase
        # 13. groot after set

        # 1. set var
        _send_freq = int(self.excel.select_freq_var[self.atr_index])
        _send_offset_in = round(float(self.__find_offset_in_table(send_freq=_send_freq, select="input_offset")), 2)
        _send_offset_out = round(float(self.__find_offset_in_table(send_freq=_send_freq, select="output_offset")), 2)
        _send_input = self.excel.atr_start_input_var
        if (_send_offset_in is None) or (_send_offset_out is None):
            print("offset error 1")
            return
        else:
            pass
        # 2. set instrument default
        self.atr_seq += 1
        if self.atr_seq == 1:
            self.inst_source.set_output(on_off=False)  # turn off rf state
        # 3. p1 ready
        elif self.atr_seq == 2:
            self.inst_source.set_freq(freq=_send_freq)  # set frequency
            self.inst_power_meter.set_freq(ch=self.excel.power_m_probe, freq=_send_freq)
            self.dialog_var_set(var_name="atr_freq", value=_send_freq)  # set dialog var
        elif self.atr_seq == 3:
            self.inst_source.set_offset(offset=_send_offset_in)  # set offset
            self.inst_power_meter.set_offset(offset=_send_offset_out, ch=self.excel.power_m_probe)  # set loss
            self.dialog_var_set(var_name="atr_input_offset", value=_send_offset_in)  # set dialog var
            self.dialog_var_set(var_name="atr_output_offset", value=_send_offset_out)  # set dialog var
        elif self.atr_seq == 4:
            self.inst_source.set_dbm(dbm=self.excel.atr_start_input_var)  # set dBm
            self.dialog_var_set(var_name="atr_rf_input", value=self.excel.atr_start_input_var)
            self.inst_power_meter.set_rel(on_off=False)  # reference off
            self.atr_input_buff = self.excel.atr_start_input_var
        elif self.atr_seq == 5:
            self.inst_source.set_output(on_off=True)  # turn on rf state
            self.rel_count = 0
        elif self.atr_seq == 6:
            self.inst_power_meter.set_rel(on_off=True)  # reference on
            self.atr_p1_ref_input = 0
            self.adder_input_direction = ""
            self.adder_input_dir_count = 0
            if self.rel_count < self.rel_count_limit:
                self.atr_seq -= 1
                self.rel_count += 1
        # 4. p1 get
        elif self.atr_seq == 7:
            fetch = self.inst_power_meter.get_rel(display_ch=1, round_num=2)
            self.dialog_var_set(var_name="atr_rf_output", value=fetch)
            self.adder = MyCal.get_p1_adder(
                _out_ref=fetch, _input_ref=self.atr_p1_ref_input)
            if self.adder is not None:
                if self.adder > 0:
                    if self.adder_input_direction == "":
                        self.adder_input_direction = "forward"
                    if self.adder_input_direction == "backward":
                        self.adder_input_dir_count += 1
                        self.adder_input_direction = "forward"
                else:
                    if self.adder_input_direction == "":
                        self.adder_input_direction = "backward"
                    if self.adder_input_direction == "forward":
                        self.adder_input_dir_count += 1
                        self.adder_input_direction = "backward"
                if self.adder != 0:
                    if self.adder_input_dir_count > 1:
                        if self.adder_input_direction == "backward":
                            self.inst_power_meter.set_rel(on_off=False)  # reference off
                            self.adder_input_direction = ""
                            self.adder_input_dir_count = 0
                        else:
                            self.p1_procedure()
                    else:
                        self.p1_procedure()
                else:
                    self.inst_power_meter.set_rel(on_off=False)  # reference off
                    self.adder_input_direction = ""
                    self.adder_input_dir_count = 0
        elif self.atr_seq == 8:
            fetch = self.inst_power_meter.get_output(display_ch=1, round_num=2)
            self.dialog_var_set(var_name="atr_rf_output", value=fetch)
            self.atr_p1_var.append(fetch)
        elif self.atr_seq == 9:
            fetch = self.inst_power_meter.get_output(display_ch=1, round_num=2)
            self.dialog_var_set(var_name="atr_rf_output", value=fetch)
            self.adder = MyCal.get_input_adder(_output_now=fetch,
                                               _output_goal=self.excel.pout_var)
            if self.adder > 0:
                if self.adder_input_direction == "":
                    self.adder_input_direction = "forward"
                if self.adder_input_direction == "backward":
                    self.adder_input_dir_count += 1
                    self.adder_input_direction = "forward"
            else:
                if self.adder_input_direction == "":
                    self.adder_input_direction = "backward"
                if self.adder_input_direction == "forward":
                    self.adder_input_dir_count += 1
                    self.adder_input_direction = "backward"
            if self.adder is not None:
                if self.adder_input_dir_count > 1:
                    if self.adder_input_direction == "backward":
                        self.atr_input_var.append(self.atr_input_buff)
                        _current = self.inst_power.get_current(round_num=2)
                        self.atr_input_curr_var.append(_current)
                        self.dialog_var_set(var_name="atr_power_current", value=_current)
                        self.adder_input_direction = ""
                        self.adder_input_dir_count = 0
                    else:
                        self.input_procedure_0()
                else:
                    self.input_procedure_0()
            else:
                print("input adder error")
                self.atr_stop = True
        # 7. p_sat ready
        elif self.atr_seq == 10:
            self.inst_source.set_dbm(dbm=self.excel.p_sat_input_var)  # set p_sat input dBm
            self.dialog_var_set(var_name="atr_rf_input", value=self.excel.p_sat_input_var)
        # 8. p_sat get
        elif self.atr_seq == 11:
            fetch = self.inst_power_meter.get_output(display_ch=1, round_num=2)
            self.dialog_var_set(var_name="atr_rf_output", value=fetch)
            self.atr_p_sat_var.append(fetch)
            # load curr from instrument
            _current = self.inst_power.get_current(round_num=2)
            self.atr_p_sat_curr_var.append(_current)
            self.dialog_var_set(var_name="atr_power_current", value=_current)
        # 9. ready overdrive
        elif self.atr_seq == 12:
            self.inst_source.set_dbm(dbm=self.excel.overdrive_input_var)
            self.dialog_var_set(var_name="atr_rf_input", value=self.excel.overdrive_input_var)
        # 10. get overdrive
        elif self.atr_seq == 13:
            fetch = self.inst_power_meter.get_output(display_ch=1, round_num=2)
            self.dialog_var_set(var_name="atr_rf_output", value=fetch)
            self.atr_overdrive_var.append(fetch)
            _current = self.inst_power.get_current(round_num=2)
            self.atr_overdrive_curr_var.append(_current)
            self.dialog_var_set(var_name="atr_power_current", value=_current)
        elif self.atr_seq == 14:
            if self.atr_index >= len(self.excel.select_freq_var) - 1:
                print("atr start end")
                self.atr_seq = 0
                self.after_call(ms=self.after_time_ms_250, func=self.atr_end)
                return
            else:
                self.atr_seq = 0
                # 11. atr index increase
                self.atr_index += 1
        else:
            print("unknown sequence number")
            self.atr_stop = True
        # 13. groot after set
        after_time = self.set_after_time_atr_start()
        self.after_call(ms=after_time, func=self.atr_start)

    def input_procedure_0(self):
        if self.adder != 0:
            before_input_buff = self.atr_input_buff
            self.atr_input_buff += self.adder
            if self.atr_input_buff <= self.excel.p_sat_input_var:
                self.inst_source.set_dbm(self.atr_input_buff)  # new input set
                self.dialog_var_set(var_name="atr_rf_input", value=self.atr_input_buff)
                self.atr_seq -= 1
            else:
                print("p input set over range\n before = {0}, sender = {1}".format(
                    before_input_buff, self.atr_input_buff))
                # self.atr_stop = True
                self.atr_input_buff = round(self.atr_input_buff, 1)
                self.atr_input_var.append(-999)
                _current = self.inst_power.get_current(round_num=2)
                self.atr_input_curr_var.append(_current)
                self.dialog_var_set(var_name="atr_power_current", value=_current)
                self.adder_input_direction = ""
                self.adder_input_dir_count = 0
        else:
            self.atr_input_buff = round(self.atr_input_buff, 1)
            self.atr_input_var.append(self.atr_input_buff)
            _current = self.inst_power.get_current(round_num=2)
            self.atr_input_curr_var.append(_current)
            self.dialog_var_set(var_name="atr_power_current", value=_current)
            self.adder_input_direction = ""
            self.adder_input_dir_count = 0

    def p1_procedure(self):
        self.atr_p1_ref_input += self.adder
        send_var = round(float(self.excel.atr_start_input_var + self.atr_p1_ref_input), 1)
        if send_var <= self.excel.p_sat_input_var:
            self.inst_source.set_dbm(dbm=send_var)  # new input set
            self.atr_input_buff = send_var
            self.dialog_var_set(var_name="atr_rf_input", value=send_var)
            self.atr_seq -= 1
        else:
            print("p1 input set over range")
            self.inst_power_meter.set_rel(on_off=False)  # reference off
            self.atr_p1_var.append(-999)
            self.inst_source.set_dbm(dbm=self.excel.p_sat_input_var)
            self.atr_input_buff = self.excel.p_sat_input_var

            self.adder_input_direction = ""
            self.adder_input_dir_count = 0
            self.atr_seq += 1

    def set_after_time_atr_start(self):
        ret = 0
        if self.atr_seq == 0:
            ret = self.after_time_ms_250
        elif self.atr_seq == 1:
            ret = self.after_time_ms_250
        elif self.atr_seq == 2:
            ret = self.after_time_ms_250
        elif self.atr_seq == 3:
            ret = self.after_time_ms_250
        elif self.atr_seq == 4:
            ret = self.after_time_ms_250
        elif self.atr_seq == 5:
            ret = self.after_time_ms_2000  # ready for rel on
        elif self.atr_seq == 6:
            ret = self.after_time_ms_2000  # ready for rel fetch sampling
        elif self.atr_seq == 7:
            ret = self.after_time_ms_2000  # ready for rel off fetch append
        elif self.atr_seq == 8:
            ret = self.after_time_ms_2000  # input loop
        elif self.atr_seq == 9:
            ret = self.after_time_ms_2000  # set for p sat
        elif self.atr_seq == 10:
            ret = self.after_time_ms_2000  # ready for p sat fetch
        elif self.atr_seq == 11:
            ret = self.after_time_ms_2000  # set for overdrive
        elif self.atr_seq == 12:
            ret = self.after_time_ms_2000  # ready for overdrive
        elif self.atr_seq == 13:
            ret = self.after_time_ms_250
        elif self.atr_seq == 14:
            ret = self.after_time_ms_250
        else:
            print("after time start atr error")
        return ret

    def atr_end(self):
        # 1. reset instrument
        if self.atr_seq == 0:
            self.inst_power.set_output(on_off=False)  # turn off power
        elif self.atr_seq == 1:
            self.inst_source.set_output(on_off=False)  # turn off rf state
        elif self.atr_seq == 2:
            self.inst_power_meter.set_rel(on_off=False)  # reference off
        else:
            print("atr end end")
            return
        self.atr_seq += 1
        self.after_call(ms=self.after_time_ms_250, func=self.atr_end)

    def after_call(self, func=None, ms=250):
        if self.atr_stop:
            self.atr_stop = False
            self.inst_power.set_output(on_off=False)  # turn off power
            self.inst_source.set_output(on_off=False)  # turn off rf state
            self.inst_power_meter.set_rel(on_off=False)  # reference off
            pg.alert(text="ATR Stop called\n",
                     title="Stop",
                     button="??????")
        else:
            if func is not None:
                self.g_root.after(ms=ms, func=func)

    def _reset_input_offset_var(self):
        self.atr_index = 0
        self.mes_input_offset.clear()
        self.mes_input_offset_under_2.clear()
        self.mes_output_offset.clear()

    def _reset_output_offset_var(self):
        self.atr_index = 0
        self.mes_output_offset.clear()

    def _reset_atr_var(self):
        self.excel.aging_done_var = False
        self.atr_state = "start"
        self.atr_index = 0
        self.atr_p1_var.clear()
        self.atr_input_var.clear()
        self.atr_input_curr_var.clear()
        self.atr_p_sat_var.clear()
        self.atr_p_sat_curr_var.clear()
        self.atr_overdrive_var.clear()
        self.atr_overdrive_curr_var.clear()
        self.atr_input_buff = 0.0
        self.excel.aging_left_var = self.excel.aging_var
        self.sort_count_out = 0
        self.compare_count_out = 0
        self.adder_input_dir_count = 0
        self.adder_input_direction = ""
        self.rel_count = 0
        self.idq_var.clear()

    def dialog_var_set(self, var_name=None, value=None):
        if (var_name is not None) and (value is not None):
            if var_name == "atr_freq":
                if value >= 1000000000000:
                    self.atr_freq.set("{0} THz".format(round(value / 1000000000000, 12)))
                elif self.excel.freq_var[0] >= 1000000000:
                    self.atr_freq.set("{0} GHz".format(round(value / 1000000000, 9)))
                elif self.excel.freq_var[0] >= 1000000:
                    self.atr_freq.set("{0} MHz".format(round(value / 1000000, 6)))
                elif self.excel.freq_var[0] >= 1000:
                    self.atr_freq.set("{0} KHz".format(round(value / 1000, 3)))
                else:
                    self.atr_freq.set("{0} Hz".format(int(value)))
            elif var_name == "atr_input_offset":
                self.atr_input_offset.set("{0} dB".format(round(value, 1)))
            elif var_name == "atr_output_offset":
                self.atr_output_offset.set("{0} dB".format(round(value, 2)))
            elif var_name == "atr_rf_input":
                self.atr_rf_input.set("{0} dBm".format(round(value, 1)))
            elif var_name == "atr_rf_output":
                self.atr_rf_output.set("{0} dBm".format(round(value, 2)))
            elif var_name == "atr_power_voltage":
                self.atr_power_voltage.set("{0} V".format(round(value, 1)))
            elif var_name == "atr_power_current":
                self.atr_current.set("{0} A".format(round(value, 2)))
            elif var_name == "aging_time_left":
                self.atr_aging_left.set("{0} sec".format(int(value)))
            else:
                print("{0} is unknown dialog var_name".format(var_name))

    @staticmethod
    def thread_sleep(sec=0.0):
        time.sleep(sec)

    def atr_stop_clicked(self):
        self.atr_stop = True

    @staticmethod
    def _atr_display_for_loop(var):
        text = ""
        for var in var:
            text += str(var) + " "
        text += "\n"
        return text

    def load_file_dialog(self):
        dir_path = fd.askopenfilename(parent=self.g_root,
                                      initialdir=os.getcwd(),
                                      title='Select Config Excel File')  # ????????? open
        if len(dir_path) == 0:
            self.excel_path = "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                self.excel_path = dir_path
            else:
                self.load_file_dialog()

    def save_file_dialog(self, _id):
        if _id == self.atr_save_button_id:
            init_name = "atr.xlsx"
        elif _id == self.atr_measure_input_offset_id:
            init_name = "input_offset.xlsx"
        elif _id == self.atr_measure_output_offset_id:
            init_name = "output_offset.xlsx"
        elif _id == self.atr_save_all_offset_id:
            init_name = "offset.xlsx"
        else:
            init_name = ".xlsx"
        dir_path = fd.asksaveasfilename(parent=self.g_root,
                                        initialdir=os.getcwd(),
                                        initialfile=init_name,
                                        title='Save Excel File',
                                        filetypes=[("excel files", "*.xlsx"), ("all", "*.*")]
                                        )  # ????????? open
        if len(dir_path) == 0:
            self.excel_path = "-1"
        else:
            check_excel = dir_path[len(dir_path) - 5:len(dir_path)]
            if check_excel == ".xlsx":
                self.excel_path = dir_path
            else:
                self.save_file_dialog(_id=_id)

    def save_excel_loop_col_watt(self, var_s, row, column):
        for var in var_s:
            self.excel.w_ws_list[0].cell(row, column, round((10 ** (var / 10)) / 1000, 1))
            column += 1

    def save_excel_loop_col(self, var_s, row, column):
        for var in var_s:

            if var == -999:  # out of range red color
                color = PatternFill(start_color="ff0000", end_color="ff0000", fill_type="solid")
                self.excel.w_ws_list[0].cell(row, column, var).fill = color
            else:
                self.excel.w_ws_list[0].cell(row, column, var)
            column += 1

    def save_excel_loop(self, var_s, column):
        index_start = 1
        for var in var_s:
            self.excel.w_ws_list[0].cell(index_start, ord(column) + 1 - ord("A"), var)
            index_start += 1

    def save_option(self, save_path, _id):
        names = []
        units = ["Hz", "dB", "dB"]
        if _id == self.atr_measure_input_offset_id:
            names = ["frequency", "input offset"]
        elif _id == self.atr_measure_output_offset_id:
            names = ["frequency", "output offset"]
        elif _id == self.atr_save_all_offset_id:
            names = ["frequency", "input offset", "output offset"]
        else:
            print("unknown save option error")
        self.excel.w_wb = Workbook()
        self.excel.w_ws_list.append(self.excel.w_wb.create_sheet("OFFSET"))
        self.excel.w_ws_list[0] = self.excel.w_wb.active

        if self.excel.save_hor is True:
            freq_row = 1
            offset_row_0 = 2
            offset_row_1 = 3

            name_column = "A"
            unit_column = "B"
            data_start_column = "C"

            self.save_excel_loop(var_s=names, column=name_column)
            self.excel.w_ws_list[0].column_dimensions[name_column].width = 15  # column set

            self.save_excel_loop(var_s=units, column=unit_column)

            __column = ord(data_start_column) + 1 - ord("A")
            self.save_excel_loop_col(var_s=self.excel.select_freq_var, column=__column, row=freq_row)

            if _id == self.atr_measure_input_offset_id:
                self.save_excel_loop_col(var_s=self.mes_input_offset, column=__column, row=offset_row_0)
            elif _id == self.atr_measure_output_offset_id:
                self.save_excel_loop_col(var_s=self.mes_output_offset, column=__column, row=offset_row_0)
            elif _id == self.atr_save_all_offset_id:
                self.save_excel_loop_col(var_s=self.mes_input_offset, column=__column, row=offset_row_0)
                self.save_excel_loop_col(var_s=self.mes_output_offset, column=__column, row=offset_row_1)
        self.excel.w_wb.save(save_path)

    def save_all_var(self, save_path):
        # test var
        print(self.excel.freq_var)
        print(self.excel.input_offset_var)
        print(self.excel.output_offset_var)
        print(self.excel.power_voltage_var)
        print(self.excel.aging_var)
        print(self.excel.source_com_var)
        print(self.excel.power_com_var)
        print(self.excel.power_meter_com_var)
        print(self.atr_p1_var)
        print(self.atr_input_var)
        print(self.atr_input_curr_var)
        print(self.atr_p_sat_var)
        print(self.atr_p_sat_curr_var)
        print(self.atr_overdrive_var)
        print(self.atr_overdrive_curr_var)
        print(self.idq_var)

        # test
        # self.atr_p1_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 41]
        # self.atr_input_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 44]
        # self.atr_input_curr_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 47]
        # self.atr_p_sat_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 50]
        # self.atr_p_sat_curr_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 53]
        # self.atr_overdrive_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 56]
        # self.atr_overdrive_curr_var = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 59]

        if self.excel.save_hor is True:
            freq_row = 1
            p1_row = 2
            input_row = 3
            input_watt_row = 4
            input_current_row = 5
            p_sat_row = 6
            p_sat_watt_row = 7
            p_sat_current_row = 8
            overdrive_row = 9
            overdrive_watt_row = 10
            overdrive_current_row = 11
            idq_row = 12

            name_column = "A"
            unit_column = "B"
            data_start_column = "C"

            names = ["frequency", "p1", "input dBm", "input watt", "input current",
                     "p_sat dBm", "p_sat watt", "p_sat current", "overdrive dBm", "overdrive watt",
                     "overdrive current", "idq"]
            units = ["Hz", "dBm", "dBm", "W", "A",
                     "dBm", "W", "A", "dBm", "W",
                     "A", "A"]

            self.excel.w_wb = Workbook()
            self.excel.w_ws_list.append(self.excel.w_wb.create_sheet("ATR"))

            self.excel.w_ws_list[0] = self.excel.w_wb.active

            self.save_excel_loop(var_s=names, column=name_column)
            self.excel.w_ws_list[0].column_dimensions[name_column].width = 15  # column set

            self.save_excel_loop(var_s=units, column=unit_column)

            __column = ord(data_start_column) + 1 - ord("A")
            self.save_excel_loop_col(var_s=self.excel.select_freq_var, column=__column, row=freq_row)
            self.save_excel_loop_col(var_s=self.atr_p1_var, column=__column, row=p1_row)
            self.save_excel_loop_col(var_s=self.atr_input_var, column=__column, row=input_row)
            self.save_excel_loop_col_watt(var_s=self.atr_input_var, column=__column, row=input_watt_row)
            self.save_excel_loop_col(var_s=self.atr_input_curr_var, column=__column, row=input_current_row)
            self.save_excel_loop_col(var_s=self.atr_p_sat_var, column=__column, row=p_sat_row)
            self.save_excel_loop_col_watt(var_s=self.atr_p_sat_var, column=__column, row=p_sat_watt_row)
            self.save_excel_loop_col(var_s=self.atr_p_sat_curr_var, column=__column, row=p_sat_current_row)
            self.save_excel_loop_col(var_s=self.atr_overdrive_var, column=__column, row=overdrive_row)
            self.save_excel_loop_col_watt(var_s=self.atr_overdrive_var, column=__column, row=overdrive_watt_row)
            self.save_excel_loop_col(var_s=self.atr_overdrive_curr_var, column=__column, row=overdrive_current_row)
            self.save_excel_loop_col(var_s=self.idq_var, column=__column, row=idq_row)
            self.excel.w_wb.save(save_path)
        else:  # save vertical atr
            pass

    def __find_offset_in_table(self, send_freq, select):
        searched_index = np.abs(np.array(send_freq) - self.excel.freq_var).argmin()
        if select == "output_offset":
            if send_freq == self.excel.freq_var[searched_index]:
                return self.excel.output_offset_var[searched_index]
            elif send_freq > self.excel.freq_var[searched_index]:
                if searched_index == len(self.excel.freq_var) - 1:  # range ?????? ??????????????? ??????
                    print("frequency range over error")
                    return
                else:
                    return self.excel.output_offset_var[searched_index] \
                           + (send_freq - self.excel.freq_var[searched_index]) \
                           * (self.excel.output_offset_var[searched_index + 1]
                              - self.excel.output_offset_var[searched_index]) \
                           / (self.excel.freq_var[searched_index + 1] - self.excel.freq_var[searched_index])
            elif send_freq < self.excel.freq_var[searched_index]:
                if searched_index == 0:  # range ?????? ??????????????? ??????
                    print("frequency range over error")
                    return
                else:
                    return self.excel.output_offset_var[searched_index - 1] \
                           + (send_freq - self.excel.freq_var[searched_index - 1]) \
                           * (self.excel.output_offset_var[searched_index]
                              - self.excel.output_offset_var[searched_index - 1]) \
                           / (self.excel.freq_var[searched_index] - self.excel.freq_var[searched_index - 1])
        elif select == "input_offset":
            if send_freq == self.excel.freq_var[searched_index]:
                return self.excel.input_offset_var[searched_index]
            elif send_freq > self.excel.freq_var[searched_index]:
                if searched_index == len(self.excel.freq_var) - 1:  # range ?????? ??????????????? ??????
                    print("frequency range over error")
                    return
                else:
                    return self.excel.input_offset_var[searched_index] \
                           + (send_freq - self.excel.freq_var[searched_index]) \
                           * (self.excel.input_offset_var[searched_index + 1]
                              - self.excel.input_offset_var[searched_index]) \
                           / (self.excel.freq_var[searched_index + 1]
                              - self.excel.freq_var[searched_index])
            elif send_freq < self.excel.freq_var[searched_index]:
                if searched_index == 0:  # range ?????? ??????????????? ??????
                    print("frequency range over error")
                    return
                else:
                    return self.excel.input_offset_var[searched_index - 1] \
                           + (send_freq - self.excel.freq_var[searched_index - 1]) \
                           * (self.excel.input_offset_var[searched_index]
                              - self.excel.input_offset_var[searched_index - 1]) \
                           / (self.excel.freq_var[searched_index]
                              - self.excel.freq_var[searched_index - 1])
            else:
                print("error input, output offset call")

    def __atr_inst_call(self, source=False, power=False, p_meter=False, spectrum=False, network=False):
        if source is True:
            self.inst_source = Instrument.Source()
            # source open
            if self.excel.source_comm_option == "GPIB":
                self.inst_source.gpib_address = self.excel.source_com_var
                self.inst_source.open_instrument_gpib(gpib_address=self.inst_source.gpib_address)
            elif self.excel.source_comm_option == "USB":
                pass
            elif self.excel.source_comm_option == "SERIAL":
                pass
        if power is True:
            self.inst_power = Instrument.PowerSupply()
            # power supply open
            if self.excel.power_comm_option == "GPIB":
                self.inst_power.gpib_address = self.excel.power_com_var
                self.inst_power.open_instrument_gpib(gpib_address=self.inst_power.gpib_address)
            elif self.excel.power_comm_option == "USB":
                pass
            elif self.excel.power_comm_option == "SERIAL":
                pass
        if p_meter is True:
            self.inst_power_meter = Instrument.PowerMeter()
            # power meter open
            if self.excel.power_meter_comm_option == "GPIB":
                self.inst_power_meter.gpib_address = self.excel.power_meter_com_var
                self.inst_power_meter.open_instrument_gpib(gpib_address=self.inst_power_meter.gpib_address)
            elif self.excel.power_meter_comm_option == "USB":
                pass
            elif self.excel.power_meter_comm_option == "SERIAL":
                pass
        if spectrum is True:
            print("not done spectrum")
            pass
        if network is True:
            print("not done network")
            pass

    def __atr_get_idn(self):
        print(self.inst_source.query_instrument("*IDN?"))
        print(self.inst_power.query_instrument("*IDN?"))
        print(self.inst_power_meter.query_instrument("*IDN?"))

    def atr_error_state(self, text):
        self.atr_state = ""
        print("p1 procedure error")
        pg.alert(text=text + "\n",
                 title="Error",
                 button="??????")
