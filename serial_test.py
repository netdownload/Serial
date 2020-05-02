import serial
import time
import libscrc


INIT_PORT = b'\x00\x00\x01\xB0'
INIT_PORT2 = b'\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81'
REQUEST = b'\x00\x64\xE4\xC4\x97\x07\x00\xED\xDE'
REQUEST2 = b'\x00\x08'
REQUEST3 = b'\x00\x76\x00'
DELAY = 0.1

crc16 = libscrc.modbus(b'\x00\x64\xE4\xC4\x97\x07\x00')
print(hex(crc16)[1::2]+hex(crc16)[2::2])


with serial.Serial('COM3', 9600, timeout=1) as ser:
    ser.write(INIT_PORT)
    time.sleep(DELAY)
    print(ser.readall().hex())
    time.sleep(DELAY)
    ser.write(INIT_PORT2)
    time.sleep(DELAY)
    print(ser.readall().hex())
    time.sleep(DELAY)
    ser.write(REQUEST)
    time.sleep(DELAY)
    print(ser.readall().hex())
    time.sleep(DELAY)
    ser.write(REQUEST2+REQUEST3)
    time.sleep(DELAY)
    print(ser.readall().hex())
