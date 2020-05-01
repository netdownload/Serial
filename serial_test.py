import serial
import time

INIT_PORT = b'\x00\x00\x01\xB0'
INIT_PORT2 = b'\x00\x01\x01\x01\x01\x01\x01\x01\x01\x77\x81'
REQUEST = b'\x00\x64\xE4\xC4\x97\x07\x00\xED\xDE'
REQUEST2 = b'\x00\x08\x00\x76\x00'

with serial.Serial('COM3', timeout=1) as ser:
    ser.write(INIT_PORT)
    time.sleep(0.1)
    print(ser.readall())
    time.sleep(0.1)
    ser.write(INIT_PORT2)
    time.sleep(0.1)
    ser.read_all()
    time.sleep(0.1)
    ser.write(REQUEST)
    time.sleep(0.1)
    ser.read_all()
    time.sleep(0.1)
    ser.write(REQUEST2)
    time.sleep(0.1)
    print(ser.readall().hex())
    time.sleep(0.1)
