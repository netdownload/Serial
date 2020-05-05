# TODO Возникают ситуации, что при считывании данных из прибора не находится почасовой срез данных
#      Необходимо определить и записать в базу ближайщие  по времени значения +1 секунда
# TODO Запись результата работы в лог
# TODO Вывод сообщений в дискорд/телеграм/whatsup/viber

# Библитотека для работы с COM портом
import serial
import time
import datetime
# Библиотека для расчета контрольной суммы
import libscrc
# Библиотека для работы с базой данных MySQL
import pymysql
import re

INIT_PORT = b'\x2F\x3F\x21\x0D\x0A'
OPEN_DEVICE = b'\x06\x30\x36\x31\x0D\x0A'
CLOSE_DEVICE = b'\x01\x42\x30\x03\x71\x01\x42\x30\x03\x71'
DEVICE_NAME = b'\x2f\x45\x6c\x73\x36\x45\x4b\x32\x36\x30\x0d\x0a'
BEGIN_REQ = b'\x01\x52\x33\x02\x33\x3a\x56\x2e\x30\x28\x33\x3b'
END_REQ = b'1)' + b'\x03'
DELAY = 5
NUMBER_OF_ATTEMPTS = 10
CRC_OK = 'CRC Ok'


def delta_hours(date_end, date_begin):
    if date_begin <= date_end:
        delta_in_hours = (date_end - date_begin).total_seconds() // 3600
        return int(delta_in_hours)
    else:
        return 0


def create_request_string(begin_request, date_request, end_request):
    date_request_str = str(date_request)[0:10] + ',' + str(date_request)[11:19] + ';'
    request_string_for_calc_crc = begin_request + date_request_str.encode() + date_request_str.encode() + end_request
    return request_string_for_calc_crc


def calculate_crc(request_string):
    crc_int = request_string[1]
    for item in range(len(request_string) - 2):
        # print(hex(request_string[item + 2]))
        crc_int = crc_int ^ request_string[item + 2]
        # print('crc = ' + hex(crc))
    crc = crc_int.to_bytes(1, byteorder='big')
    return crc


def get_last_date_from_database():
    connection = pymysql.connect(host='10.1.1.50',
                                 user='user',
                                 password='qwerty123',
                                 db='resources')
    try:
        with connection.cursor() as cursor:
            sql = 'SELECT gas_datetime FROM gas ORDER BY gas_datetime DESC LIMIT 0, 1'
            cursor.execute(sql)
            result = cursor.fetchone()
            datetime_last = result[0]
    finally:
        connection.close()
    return datetime_last


def insert_values_into_database():  # TODO Insert function
    pass


def check_crc_and_date(answer_for_check_crc):               # TODO Дабавить проверку по дате
    check = split_answer_into_values(answer_for_check_crc)
    if check[-2] == CRC_OK:
        print(CRC_OK)
    else:
        print('Error')
    return 0


def split_answer_into_values(answer):
    values = re.split(r'\)\(|\(|\)', answer.decode())
    # values.remove('')
    # values.pop()
    return values


# crc16 = libscrc.modbus(INIT_PORT)
# print(hex(crc16)[1::2] + hex(crc16)[2::2])  # TODO разобраться в срезах

date_now = datetime.datetime.now()
date_last = get_last_date_from_database()
print(delta_hours(date_now, date_last))
print(date_last)
print(date_last + datetime.timedelta(minutes=60))

is_read_ok = False

while not is_read_ok and NUMBER_OF_ATTEMPTS > 0:
    with serial.Serial('COM4', 19200, parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_ONE,
                       bytesize=serial.SEVENBITS,
                       timeout=1) as ser:
        ser.write(INIT_PORT)
        time.sleep(DELAY)
        device_name_from_com = ser.readall().hex()

        if device_name_from_com == DEVICE_NAME.hex():
            is_read_ok = True
        else:
            is_read_ok = False
            NUMBER_OF_ATTEMPTS -= 1
            ser.write(CLOSE_DEVICE)

        ser.write(OPEN_DEVICE)
        time.sleep(DELAY)
        print(ser.readall())
        time.sleep(DELAY)

        for hours in range(delta_hours(date_now, date_last)):
            time.sleep(DELAY)
            request_string = create_request_string(BEGIN_REQ, date_last + datetime.timedelta(minutes=60 * (hours + 1)),
                                                   END_REQ)
            crc = calculate_crc(request_string)
            ser.write(request_string + crc)
            time.sleep(DELAY)
            check_crc_and_date((ser.readall()))
            insert_values_into_database()
        ser.write(CLOSE_DEVICE)
        time.sleep(DELAY)
        print(ser.readall())
