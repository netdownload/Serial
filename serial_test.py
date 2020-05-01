import serial

with serial.Serial('COM3', 19200, timeout=1) as ser:
    x = ser.read(160)
    print(x)
