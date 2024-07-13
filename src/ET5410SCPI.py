import serial
import serial.tools.list_ports
from time import sleep
import sys

sys_ver = 0
serial_holder = serial.Serial()


def port2list():
    l_ports = serial.tools.list_ports.comports()
    result = [element.device for element in l_ports]
    return result


def sh_wr(tx_str):
    print(sys_ver)
    serial_holder.write(tx_str.encode() if (sys_ver >= 3) else tx_str)


def sh_rd():
    print(sys_ver)
    s = serial_holder.read_until()
    return s.decode() if (sys_ver >= 3) else s


def main():
    global sys_ver
    global serial_holder

    sys_ver = sys.version_info[0]

    print(port2list())

    serial_holder = serial.Serial('COM11', 115200, timeout=1)
    port_status = serial_holder.isOpen()
    if port_status:
        serial_holder.close()
    serial_holder.open()

    print('\n\nStart HERE\n')

    sh_wr('*IDN?\x0A')
    print(sh_rd())

    sh_wr('SYSTem:VERSion?\x0A')
    print(sh_rd())

    # sh_wr('SYST:BEEP\x0A')
    # print(sh_rd())

    # sh_wr('SYST:LOCA\x0A')
    # print(sh_rd())

    # Constant resistance
    # sh_wr('RESI:CRCV?\x0A')
    # print(sh_rd())

    # Constant CC-CV
    sh_wr('VOLT:VMAX?\x0A')
    print(sh_rd())
    sh_wr('VOLT:VMAX 14.00\x0A')
    print(sh_rd())
    sh_wr('VOLT:VMAX?\x0A')
    print(sh_rd())

    sh_wr('CURR:IMAX?\x0A')
    print(sh_rd())
    sh_wr('CURR:IMAX 7.00\x0A')
    print(sh_rd())
    sh_wr('CURR:IMAX?\x0A')
    print(sh_rd())

    sh_wr('CURR:CC?\x0A')
    print(sh_rd())
    print('Set CC = 5')
    sh_wr('CURR:CC 5.00\x0A')
    print(sh_rd())
    sh_wr('CURR:CC?\x0A')
    print(sh_rd())

    sh_wr('VOLT:CV?\x0A')
    print(sh_rd())

    # sh_wr('CURR:CCCV?\x0A')
    # print(sh_rd())
    # sh_wr('CURR:CCCV 12\x0A')
    # print(sh_rd())
    # sh_wr('CURR:CCCV?\x0A')
    # print(sh_rd())

    sh_wr('CH:SW?\x0A')
    print(sh_rd())
    sh_wr('CH:SW ON\x0A')
    print(sh_rd())
    sh_wr('CH:SW?\x0A')
    print(sh_rd())

    sleep(1)

    sh_wr('MEAS:VOLT?\x0A')
    print(sh_rd())
    sh_wr('MEAS:CURR?\x0A')
    print(sh_rd())
    sh_wr('MEAS:POWE?\x0A')
    print(sh_rd())

    sh_wr('CH:SW OFF\x0A')
    print(sh_rd())
    sh_wr('CH:SW?\x0A')
    print(sh_rd())

    # MODE
    # sh_wr('LIST:MODE?\x0A')
    # print(sh_rd())
    #
    # sh_wr('LIST:NUM?\x0A')
    # print(sh_rd())
    #
    # sh_wr('LOAD:TRIG?\x0A')
    # print(sh_rd())
    #
    # sh_wr('LOAD:VRAN?\x0A')
    # print(sh_rd())
    #
    # sh_wr('LOAD:CRAN?\x0A')
    # print(sh_rd())
    #
    # sh_wr('LOAD:ABNO?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:TEST?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:OUT?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:VHIG?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:VLOW?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:CHIG?\x0A')
    # print(sh_rd())
    #
    # sh_wr('QUAL:CLOW?\x0A')
    # print(sh_rd())


if __name__ == '__main__':
    main()
