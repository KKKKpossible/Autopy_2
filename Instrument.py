import pyvisa
import pyautogui as pg  # message box library


class Instrument:
    def __init__(self):
        self.instrument_name = ""
        self.gpib_address = ""
        self.usb_address = ""
        self.serial_port_name = ""
        self.com_opt = ""
        self.com_rm = ""
        self.com_inst = ""
        self.opened = False
        self.test_except = False
        self.anritsu = False
        self.agilent = False
        self.hp = False
        self.thandar = False

    def open_instrument_gpib(self, gpib_address):
        if self.test_except:
            pass
        else:
            self.gpib_address = gpib_address
            self.__open_instrument(com_opt="GPIB")
            checker = str(self.query_instrument(command="*IDN?"))
            if checker.find("ANRITSU") != -1:
                self.anritsu = True
                self.agilent = False
                self.hp = False
                self.thandar = False
            elif checker.find("Agilent") != -1:
                self.anritsu = False
                self.agilent = True
                self.hp = False
                self.thandar = False
            elif checker.find("HP") != -1:
                self.anritsu = False
                self.agilent = False
                self.hp = True
                self.thandar = False
            elif checker.find("THANDAR") != -1:
                self.anritsu = False
                self.agilent = False
                self.hp = False
                self.thandar = True

    def __open_instrument(self, com_opt):
        if self.test_except:
            pass
        else:
            if com_opt == "GPIB":
                address = "GPIB0::" + str(self.gpib_address) + "::INSTR"
                self.com_opt = com_opt
                self.com_rm = pyvisa.ResourceManager()
                self.com_inst = self.com_rm.open_resource(address)
                self.opened = True
            elif com_opt == "USB":
                pass
            elif com_opt == "SERIAL":
                pass
            else:
                return
            print("instrument opened")

    def query_instrument(self, command):
        if self.test_except:
            pass
        else:
            if self.opened:
                pass
            else:
                pg.alert(text="Send Error_0",
                         title="Error",
                         button="확인")
                return
            return self.com_inst.query(command)

    def write_instrument(self, command):
        if self.test_except:
            pass
        else:
            if self.opened:
                pass
            else:
                pg.alert(text="Send Error_1",
                         title="Error",
                         button="확인")
                return
            self.com_inst.write(command)

    def read_instrument(self):
        if self.test_except:
            pass
        else:
            if self.opened:
                pass
            else:
                pg.alert(text="Read Error",
                         title="Error",
                         button="확인")
                return
            return self.com_inst.read()


class PowerMeter(Instrument):
    def __init__(self):
        Instrument.__init__(self)
        self.instrument_name = "POWER METER"
        self.offset = 0.0
        self.rel_state = False
        self.frequency = 0
        self.probe_channel = 1

    def set_rel(self, on_off):
        if self.agilent is True:
            self.set_rel_agilent(on_off=on_off)
        elif self.anritsu is True:
            self.set_rel_anritsu(on_off=on_off)
        else:
            print("error 0")

    def set_freq(self, freq, ch=1):
        if self.agilent is True:
            self.set_freq_agilent(freq=freq, ch=ch)
            self.frequency = freq
        elif self.anritsu is True:
            self.set_freq_anritsu(freq=freq, ch=ch)
            self.frequency = freq
        else:
            print("error 1")

    def set_offset(self, offset, ch=1):
        if self.agilent is True:
            self.set_offset_agilent(offset=offset, ch=ch)
        elif self.anritsu is True:
            self.set_offset_anritsu(offset=offset, ch=ch)

    def set_offset_state(self, state, ch):
        if self.agilent is True:
            self.set_offset_state_agilent(state=state, ch=ch)
        elif self.anritsu is True:
            self.set_offset_state_anritsu(state=state, ch=ch)

    def set_rel_agilent(self, on_off):
        if on_off:
            self.write_instrument(command="CALC:REL:AUTO ONCE")
            self.rel_state = True
        else:
            self.write_instrument(command="CALC:REL:STAT OFF")
            self.rel_state = False

    def set_rel_anritsu(self, on_off):
        if on_off:
            self.write_instrument(command="REL 1 1")
            self.rel_state = True
        else:
            self.write_instrument(command="REL 1 0")
            self.rel_state = False

    def set_freq_agilent(self, freq, ch=1):
        self.write_instrument("SENS{0}:FREQ {1}HZ".format(ch, freq))
        self.frequency = freq

    def set_freq_anritsu(self, freq, ch=1):
        self.write_instrument("CFFRQ A, {0}HZ".format(freq))  # A = config A, 이모델에서 config B는 사용하지 않는듯
        self.frequency = freq

    def set_offset_agilent(self, offset, ch=1):
        self.write_instrument(command="SENS{0}:CORR:LOSS2 {1}DB".format(ch, offset))  # LOSS1은 cal 에 사용된다고함

    def set_offset_anritsu(self, offset, ch=1):
        self.write_instrument(command="OFFFIX A, {0}DB".format(offset))  # anritsu 24버전은 config A 만 존재

    def set_offset_state_agilent(self, state, ch):
        if state is True:
            self.write_instrument(command="SENS{0}:CORR:LOSS2:STAT ON".format(ch))
        elif state is False:
            self.write_instrument(command="SENS{0}:CORR:LOSS2:STAT OFF".format(ch))
        else:
            print("unknown state in set offset state agilent")

    def set_offset_state_anritsu(self, state, ch):
        if state is True:
            self.write_instrument(command="OFFTYP A, FIXED")
        elif state is False:
            self.write_instrument(command="OFFTYP A, OFF")
        else:
            print("unknown state in set offset state agilent")

    def get_output(self, display_ch=1, round_num=2):
        if self.agilent is True:
            return self.get_output_agilent(display_ch=display_ch, round_num=round_num)
        elif self.anritsu is True:
            return self.get_output_anritsu(display_ch=display_ch, round_num=round_num)
        else:
            print("error 2")

    def get_rel(self, display_ch=1, round_num=2):
        if self.agilent is True:
            return self.get_rel_agilent(display_ch=display_ch, round_num=round_num)
        elif self.anritsu is True:  # anritsu 는 rel read 가 따로 없다.
            return self.get_output_anritsu(display_ch=display_ch, round_num=round_num)
        else:
            print("error 3")

    def get_output_agilent(self, display_ch=1, round_num=2):
        return round(float(self.query_instrument("FETC{0}?".format(display_ch))), round_num)

    def get_output_anritsu(self, display_ch=1, round_num=2):
        return round(float(self.query_instrument("O {0}?".format(display_ch))), round_num)

    def get_rel_agilent(self, display_ch=1, round_num=2):
        return round(float(self.query_instrument("FETC{0}:REL?".format(display_ch))), round_num)


class PowerSupply(Instrument):
    def __init__(self):
        Instrument.__init__(self)
        self.instrument_name = "POWER SUPPLY"
        self._voltage_set = 0.0
        self._current_set = 0.0
        self._output_state = False
        self.test_except = False

    def set_output(self, on_off, ch=1):
        if self.hp is True:
            self.set_output_hp_6x74a(on_off=on_off)
        elif self.thandar is True:
            self.set_output_thandar_cpx400dp(on_off=on_off, ch=ch)

    def set_output_thandar_cpx400dp(self, on_off, ch=1):
        if self.test_except:
            pass
        else:
            if on_off:
                self.write_instrument(command="OP{0} {1}".format(ch, 1))  # turn off power
                self._output_state = True
            else:
                self.write_instrument(command="OP{0} {1}".format(ch, 0))  # turn off power
                self._output_state = False

    def set_output_hp_6x74a(self, on_off):
        if self.test_except:
            pass
        else:
            if on_off:
                self.write_instrument(command="OUTP ON")  # turn off power
                self._output_state = True
            else:
                self.write_instrument(command="OUTP OFF")  # turn off power
                self._output_state = False

    def set_voltage(self, voltage, ch=1):
        if self.hp is True:
            self.set_voltage_hp_6x74a(voltage=voltage)
        elif self.thandar is True:
            self.set_voltage_thandar_cpx400dp(voltage=voltage, ch=ch)

    def set_voltage_hp_6x74a(self, voltage):
        if self.test_except:
            pass
        else:
            self.write_instrument(command="VOLT:LEV {0}".format(voltage))
            self._voltage_set = voltage

    def set_voltage_thandar_cpx400dp(self, voltage, ch=1):
        if self.test_except:
            pass
        else:
            self.write_instrument(command="V{0} {1}".format(ch, voltage))
            self._voltage_set = voltage

    def set_current(self, current, ch=1):
        if self.hp is True:
            self.set_current_hp_6x74a(current=current)
        elif self.thandar is True:
            self.set_current_thandar_cpx400dp(current=current, ch=ch)

    def set_current_hp_6x74a(self, current):
        if self.test_except:
            pass
        else:
            self.write_instrument(command="CURR:LEV {0}".format(current))
            self._current_set = current

    def set_current_thandar_cpx400dp(self, current, ch=1):
        if self.test_except:
            pass
        else:
            self.write_instrument(command="I{0} {1}".format(ch, current))
            self._current_set = current

    def get_current(self, round_num=2, ch=1):
        if self.hp is True:
            return self.get_current_hp_6x74a(round_num=round_num)
        elif self.thandar is True:
            return self.get_current_thandar_cpx400dp(round_num=round_num, ch=ch)

    def get_current_hp_6x74a(self, round_num=2):
        if self.test_except:
            return -1667
        else:
            return round(float(self.query_instrument(command="MEAS:CURR?")), round_num)

    def get_current_thandar_cpx400dp(self, round_num=2, ch=1):
        if self.test_except:
            return -1667
        else:
            ret = self.query_instrument(command="I{0}O?".format(ch))
            return round(float(ret.split("A")[0]), round_num)


class Source(Instrument):
    def __init__(self):
        Instrument.__init__(self)
        self._instrument_name = "SOURCE"
        self._offset = 0.0
        self._set_dbm = 0.0
        self._output_state = False
        self._frequency = 0
        self.anritsu = False
        self.agilent = False

    def set_output(self, on_off):
        if self.agilent is True:
            self.set_output_agilent(on_off=on_off)
        elif self.anritsu is True:
            self.set_output_anritsu(on_off=on_off)

    def set_freq(self, freq):
        if self.agilent is True:
            self.set_freq_agilent(freq=freq)
        elif self.anritsu is True:
            self.set_freq_anritsu(freq=freq)

    def set_offset(self, offset):
        if self.agilent is True:
            self.set_offset_agilent(offset=offset)
        elif self.anritsu is True:
            self.set_offset_anritsu(offset=offset)

    def set_dbm(self, dbm):
        if self.agilent is True:
            self.set_dbm_agilent(dbm=dbm)
        elif self.anritsu is True:
            self.set_dbm_anritsu(dbm=dbm)

    def set_output_agilent(self, on_off):
        if on_off:
            self.write_instrument(command="OUTP:STAT ON")
            self._output_state = True
        else:
            self.write_instrument(command="OUTP:STAT OFF")
            self._output_state = False

    def set_freq_agilent(self, freq):
        self.write_instrument(command="FREQ {0} Hz".format(freq))
        self._frequency = freq

    def set_offset_agilent(self, offset):
        self.write_instrument(command="POW:OFFS {0} DB".format(offset))
        self._offset = offset

    def set_dbm_agilent(self, dbm):
        self.write_instrument(command="POW:AMPL {0} dBM".format(dbm))
        self._set_dbm = dbm

    def set_output_anritsu(self, on_off):
        if on_off:
            self.write_instrument(command="RF1")
            self._output_state = True
        else:
            self.write_instrument(command="RF0")
            self._output_state = False

    def set_freq_anritsu(self, freq):
        self.write_instrument(command="CF0 {0} HZ".format(freq))  # Hz
        self._frequency = freq

    def set_offset_anritsu(self, offset):
        self.write_instrument(command="LOS {0} DB".format(offset))
        self._offset = offset

    def set_offset_state(self, state):
        if self.agilent is True:
            self.set_offset_state_agilent(state=state)
        else:
            self.set_offset_state_anritsu(state=state)

    @staticmethod
    def set_offset_state_agilent(state):
        # if state is True:
        #     self.write_instrument(command="POW:OFFS:STAT ON")
        # elif state is False:
        #     self.write_instrument(command="POW:OFFS:STAT OFF")
        # else:
        #     print("unknown state_ set offset state agilent")
        print("agilent source doesn't have offset state command")

    def set_offset_state_anritsu(self, state):
        if state is True:
            self.write_instrument(command="LO1")
        elif state is False:
            self.write_instrument(command="LO0")
        else:
            print("unknown state_ set offset state anritsu")

    def set_dbm_anritsu(self, dbm):
        self.write_instrument(command="L0 {0} DM".format(dbm))
        self._set_dbm = dbm


class Spectrum(Instrument):
    def __init__(self):
        Instrument.__init__(self)


class Network(Instrument):
    def __init__(self):
        Instrument.__init__(self)


if __name__ == "__main__":
    # rm = pyvisa.ResourceManager()
    # for var in rm.list_resources():
    #     print(var)
    test = PowerSupply()
    test.open_instrument_gpib(gpib_address="11")
    test.set_output(on_off=False, ch=1)
    test.set_output(on_off=False, ch=2)
    print(test.query_instrument(command="OCP1?"))
    # test.set_voltage(voltage=1.23, ch=1)
    # test.set_voltage(voltage=1.23, ch=2)
    # test.set_current(current=1.23, ch=1)
    # test.set_current(current=1.23, ch=2)
    # test.set_output(on_off=False, ch=1)
    # test.set_output(on_off=False, ch=2)
    # print(test.get_current(ch=1))
    # print(test.get_current(ch=2))
    # test.set_output(on_off=False, ch=1)
    # test.set_output(on_off=False, ch=2)
    # print(test.get_current(ch=1))
    # print(test.get_current(ch=2))

    # test = PowerMeter()
    # test.open_instrument_gpib(gpib_address="13")
    # print(test.query_instrument(command="*idn?"))
    # print(test.query_instrument(command="CHCFG? 1"))  # return sent ch's config(1번채널은 config A 또는 B이다)
    # print(test.query_instrument(command="CHCFG? 2"))  # return sent ch's config(1번채널은 config A 또는 B이다)
    # print(test.query_instrument(command="CHUNIT? 1"))  # return ch unit
    # print(test.query_instrument(command="CHUNIT? 2"))  # return ch unit ch가 위아래 화면이네
    # print(test.write_instrument(command="DISP OFF"))  # return ch config, channel
    # print(test.write_instrument(command="DISP ON"))  # return ch config, channel
    # print(test.write_instrument(command="FROFF ON"))  # 화면에 frequency, offset 이 출력된다.
    #
    # print(test.query_instrument(command="O 1"))  # 출력을 읽는것
    # print(test.query_instrument(command="OFFFIX? A"))  # OFFSET config A의 offset 을 읽는것 24모델은 config A 만 존재하는듯
    # test.write_instrument(command="OFFFIX A, -10")
    # print(test.query_instrument(command="OFFFIX? A"))  # OFFSET config A의 offset 을 읽는것 24모델은 config A 만 존재하는듯
    #
    # print(test.query_instrument(command="OFFVAL A"))  # offfix의 뒤에 나오는 값을 그대로 출력해줌
    # print(test.query_instrument(command="OFFTYP? A"))  # Offset settting
    # print(test.write_instrument(command="OFFTYP A,OFF"))  # Offset settting
    # print(test.write_instrument(command="OFFTYP A,FIXED"))  # Offset settting
    # print(test.query_instrument(command="OFFTYP? A"))  # Offset settting
    #
    # print("before rel on = {0}".format(test.query_instrument(command="O 1")))  # 출력을 읽는것
    # print(test.write_instrument(command="REL 1, 1"))  # Offset settting
    # print("after rel on = {0}".format(test.query_instrument(command="O 1")))  # 출력을 읽는것
    # print(test.write_instrument(command="REL 1, 0"))  # Offset settting
    # print("after rel off = {0}".format(test.query_instrument(command="O 1")))  # 출력을 읽는것
    # print(test.query_instrument(command="CFFRQ? A"))  # Offset settting
    # print(test.write_instrument(command="CFFRQ A, 25000000000HZ"))  # Offset settting
    # print(test.query_instrument(command="CFFRQ? A"))  # Offset settting


    # test = PowerMeter()
    # test.open_instrument_gpib("13")
    # print(test.get_rel())

    # power supply test
    # test.write_instrument(command="*IDN?")
    # test.read_instrument()
    # time.sleep(2)
    # test.write_instrument(command="SYST:REM")
    # time.sleep(2)
    # test = PowerSupply()
    # test.open_instrument_gpib(gpib_address="5")
    # test.query_instrument(command="*IDN?")
    # test.write_instrument(command="OUTP OFF")
    # time.sleep(2)
    # test.write_instrument(command="OUTP ON")
    # time.sleep(2)
    # test.write_instrument(command="VOLT:LEV 4.5") # voltage set
    # time.sleep(2)
    # test.write_instrument(command="CURR:LEV 4.5") # current set
    # time.sleep(2)
    # test.write_instrument(command="VOLT?") # voltage set
    # time.sleep(2)
    # test.write_instrument(command="CURR?") # current set
    # time.sleep(2)
    # test.write_instrument(command="MEAS:VOLT?") # voltage set
    # time.sleep(2)
    # test.write_instrument(command="MEAS:CURR?") # current set
    # time.sleep(2)

    # source test
    # test = Source()
    # test.open_instrument_gpib("19")
    # test.write_instrument("FREQ 500 kHz")  # set frequency
    # time.sleep(2)
    # test.query_instrument("FREQ:CW?")  # get frequency
    # time.sleep(2)
    # test.write_instrument("POW:AMPL -2.3 dBM")  # set dBm
    # time.sleep(2)
    # test.query_instrument("POW:AMPL?")  # get dBm
    # time.sleep(2)

    # test.write_instrument("OUTP:STAT OFF")  # turn off rf state
    # time.sleep(2)
    # test.query_instrument("OUTP?")  # read rf state
    # time.sleep(2)
    # test.write_instrument("OUTP:STAT ON")
    # time.sleep(2)
    # test.query_instrument("OUTP?")  # read rf state
    # time.sleep(2)
    # test.write_instrument("OUTP:STAT OFF")  # turn off rf state
    # time.sleep(2)
    # test.query_instrument("OUTP?")  # read rf state
    # time.sleep(2)
    # test.write_instrument("POW:OFFS -10 DB")  # set offset
    # test.query_instrument("POW:OFFS?")  # read offset

    # power meter test
    # pyvisa test
    # rm = pyvisa.ResourceManager()
    # for var in rm.list_resources():
    #     print(var)
    # inst = rm.open_resource("GPIB0::13::INSTR")
    # print(inst.query("*IDN?"))  # get idn
    # print(inst.query("FETC1?"))  # fetch channel 1
    # print(inst.query("*RST"))  # reset
    # inst.write("*CLS") # clear
    # print(inst.write(":SYST:ERR?"))  # get error
    # print(inst.read())  # read returns
    # print(inst.query("SERV:SENS2:TYPE?"))  # get sensor adapter part number
    # print(inst.write("SENS2:FREQ 500MHZ"))  # set channel freq no return
    # print(inst.query("SENS2:CORR:LOSS2?"))  # get loss
    # print(inst.write("SENS2:CORR:LOSS2 -30DB"))  # set loss
    # print(inst.query("SENS2:CORR:LOSS2?"))  # get loss
    # print(inst.write("SENS2:CORR:LOSS2:STAT OFF"))  # offset off
    # print(inst.write("SENS2:CORR:LOSS2:STAT ON"))  # offset on
    # self.inst_power_meter.write_instrument(command="CALC:REL:AUTO ONCE")
    # time.sleep(1)
    # self.inst_power_meter.write_instrument(command="CALC:REL:STAT OFF")
