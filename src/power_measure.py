import spd1000_series as _spd
import csv
import math
from datetime import datetime
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
    # print(sys_ver)
    serial_holder.write(tx_str.encode() if (sys_ver >= 3) else tx_str)


def sh_rd():
    # print(sys_ver)
    s = serial_holder.read_until()
    s = s.decode().strip() if (sys_ver >= 3) else s.strip()
    s = s.split(' ')
    return s


def psu_on(spd_hdr):
    while True:
        spd_hdr.spdSetup(b'OUTP CH1,ON')
        res = spd_hdr.spdQuery(b'SYST:STAT?')
        if int(res, 16) & 0x10 == 0x10:
            break


def psu_off(spd_hdr):
    while True:
        spd_hdr.spdSetup(b'OUTP CH1,OFF')
        res = spd_hdr.spdQuery(b'SYST:STAT?')
        if int(res, 16) & 0x10 == 0x00:
            break


def main():
    spd_hdr = _spd.SPD1000()
    spd_hdr.spd_ip = "192.168.1.214"

    global sys_ver
    global serial_holder

    sys_ver = sys.version_info[0]
    print(sys_ver)
    print(port2list())

    serial_holder = serial.Serial('COM12', 115200, timeout=1)
    port_status = serial_holder.isOpen()
    if port_status:
        serial_holder.close()

    serial_holder.open()
    spd_hdr.spdConnect()

    dt = datetime.now()
    ts = math.floor(datetime.timestamp(dt))

    sh_wr('CH:SW OFF\x0A')
    print(sh_rd())
    sh_wr('CH:SW?\x0A')
    print(sh_rd())

    spd_hdr.spdSetup(b'*UNLOCK')
    spd_hdr.spdSetup(b'MODE:SET 2W')

    filename = f'SEPIC_24VO_20VI_ms_{ts}.csv'
    print(filename)

    with open(filename, 'w', newline='')as f:
        writer = csv.writer(f)
        writer.writerow([str(sys.version_info)])
        sh_wr('*IDN?\x0A')
        writer.writerow([str(sh_rd())])
        sh_wr('SYSTem:VERSion?\x0A')
        writer.writerow([str(sh_rd())])
        writer.writerow([str(spd_hdr.spdQuery(b'*IDN?'))])

        for j in range(28, 32, 2):

            writer.writerow([])

            psu_off(spd_hdr)
            sh_wr('CH:SW OFF\x0A')
            print(sh_rd())

            spd_hdr.spdSetup('CH1:VOLT {0:.2f}'.format(j).encode())
            spd_hdr.spdSetup(b'CH1:CURR 2.50')

            psu_on(spd_hdr)
            writer.writerow(['PSU SET', str(spd_hdr.spdQuery(b'CH1:VOLT?')), str(spd_hdr.spdQuery(b'CH1:CURR?'))])

            psu_load = ['PS Mes',
                        spd_hdr.spdQuery(b'MEAS:VOLT?'), spd_hdr.spdQuery(b'MEAS:CURR?'),
                        spd_hdr.spdQuery(b'MEAS:POWE?')]
            writer.writerow(psu_load)

            prg_load = ['PL Mes']
            sh_wr('MEAS:VOLT?\x0A')
            prg_load.append(sh_rd()[-1])
            sh_wr('MEAS:CURR?\x0A')
            prg_load.append(sh_rd()[-1])
            sh_wr('MEAS:POWE?\x0A')
            prg_load.append(sh_rd()[-1])
            writer.writerow(prg_load)

            sh_wr('CURR:IMAX 4.50\x0A')
            print(sh_rd())
            sh_wr('VOLT:VMAX 35.00\x0A')
            print(sh_rd())

            for i in range(0, 105, 5):
                sh_wr('CH:SW ON\x0A')
                print(sh_rd())
                sh_wr('CURR:CC {0:.3f}\x0A'.format(i/100))
                print(sh_rd())
                sh_wr('CURR:CC?\x0A')
                data = ['CC SET', sh_rd()[-1][1:]]

                sleep(0.01)

                prg_load = ['PL Mes']
                sh_wr('MEAS:VOLT?\x0A')
                prg_load.append(sh_rd()[-1])
                sh_wr('MEAS:CURR?\x0A')
                prg_load.append(sh_rd()[-1])
                sh_wr('MEAS:POWE?\x0A')
                prg_load.append(sh_rd()[-1])
                data += prg_load

                psu_load = ['PS Mes',
                            spd_hdr.spdQuery(b'MEAS:VOLT?'), spd_hdr.spdQuery(b'MEAS:CURR?'),
                            spd_hdr.spdQuery(b'MEAS:POWE?')]
                data += psu_load
                # writer.writerow(data)
                try:
                    pwr_cal = float(prg_load[-1])/float(psu_load[-1])*100
                except ZeroDivisionError:
                    pwr_cal = 0
                data += ['Eff', pwr_cal]
                writer.writerow(data)

            psu_off(spd_hdr)
            sh_wr('CH:SW OFF\x0A')
            print(sh_rd())

    serial_holder.close()
    spd_hdr.spdClose()


if __name__ == '__main__':
    main()
