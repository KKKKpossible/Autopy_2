import time
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
        print("instrument opend")

    def open_instrument_gpib(self, gpib_address):
        self.gpib_address = gpib_address
        self.__open_instrument(com_opt="GPIB")

    def __open_instrument(self, com_opt):
        if com_opt == "GPIB":
            address = "GPIB0::" + str(self.gpib_address) + "::INSTR"
            self.com_opt = com_opt
            self.com_rm = pyvisa.ResourceManager()
            self.com_inst = self.com_rm.open_resource(address)
            self.opened = True
            a = 2
        elif com_opt == "USB":
            pass
        elif com_opt == "SERIAL":
            pass
        else:
            return

    def query_instrument(self, command):
        if self.opened:
            pass
        else:
            pg.alert(text="Send Error_0",
                     title="Error",
                     button="확인")
            return
        return self.com_inst.query(command)

    def write_instrument(self, command):
        if self.opened:
            pass
        else:
            pg.alert(text="Send Error_1",
                     title="Error",
                     button="확인")
            return
        self.com_inst.write(command)

    def read_instrument(self):
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


class PowerSupply(Instrument):
    def __init__(self):
        Instrument.__init__(self)
        self.instrument_name = "POWER SUPPLY"


class Source(Instrument):
    def __init__(self):
        Instrument.__init__(self)
        self.instrument_name = "SOURCE"


if __name__ == "__main__":
    test = PowerMeter()
    test.open_instrument_gpib(gpib_address="13")
    test.query_instrument(command="*IDN?")

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

    # sourece test
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