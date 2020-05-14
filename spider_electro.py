# Библитотека для работы с COM портом
import datetime
import time
# Библиотека для работы с логаи
import logging
# Библиотека для расчета контрольной суммы
import libscrc
# Библиотека для работы с базой данных MySQL
import pymysql
import serial
import sys

DEVICE_NUMBER = b'\x5E'
TEST_PORT = DEVICE_NUMBER + b'\x00'  # запрос на тестирование порта, в ответе должно прийти то же значение
INIT_PORT = DEVICE_NUMBER + b'\x01\x01\x01\x01\x01\x01\x01\x01'
DATE_MEMORY_REQUEST = DEVICE_NUMBER + b'\x08\x13'
# В ответе проверить байт стостояния, возможно он отвечает за то какой банк памяти брать
DELAY = 0.1
COM = 'COM6'
COM_SPEED = 9600
DATABASE_HOST = '10.1.1.99'
DATABASE_USER = 'user'
DATABASE_PASSWORD = 'qwerty123'
DATABASE = 'resources'
MEMORY_BANK1 = b'\x06\x03'
MEMORY_BANK2 = b'\x06\x83'
PERIOD = b'\x1E'
PERIOD_HEX = 16

crc16 = libscrc.modbus(INIT_PORT)
init_port_with_crc = INIT_PORT + crc16.to_bytes(2, byteorder='little')
crc16 = libscrc.modbus(TEST_PORT)
test_port_with_crc = TEST_PORT + crc16.to_bytes(2, byteorder='little')
crc16 = libscrc.modbus(DATE_MEMORY_REQUEST)
date_request_with_crc = DATE_MEMORY_REQUEST + crc16.to_bytes(2, byteorder='little')


# TODO настроить время ротации и поправить отображение лога (время, дата)
# Функция для настройки логирования
def log_setup():
    log_handler = logging.handlers.WatchedFileHandler('spider_electro.txt')
    formatter = logging.Formatter(
        '%(asctime)s spider_electro.py [%(process)d]: %(message)s',
        '%b %d %H:%M:%S')
    formatter.converter = time.gmtime  # if you want UTC time
    log_handler.setFormatter(formatter)
    logger = logging.getLogger()
    logger.addHandler(log_handler)
    logger.setLevel(logging.DEBUG)


# Функция для проверки доступности COM порта
def check_com_port():
    try:
        serial.Serial(COM, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                      bytesize=serial.EIGHTBITS, timeout=1)
        logging.debug('Порт доступен')
        return 0
    except serial.serialutil.SerialException:
        logging.debug('Порт недоступен')
        sys.exit("Порт недоступен")


def convert_date(date_memory_answer_hex):
    date_answer = hex(date_memory_answer_hex[6])[2:4] + '.' + hex(date_memory_answer_hex[7])[2:4] + '.' + hex(
        date_memory_answer_hex[8])[2:4] + ' ' + hex(date_memory_answer_hex[4])[2:4] + ':' + hex(
        date_memory_answer_hex[5])[2:4]
    datetime_object = datetime.datetime.strptime(date_answer, '%d.%m.%y %H:%M')
    return datetime_object


def convert_memory(date_memory_answer_hex):
    memory_answer = date_memory_answer_hex[1:3]
    return memory_answer


# Функция берет последнее значение времени из базы
def get_last_date_from_database():
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            sql = 'SELECT electro42_datetime FROM electro42 ORDER BY electro42_datetime DESC LIMIT 0, 1'
            cursor.execute(sql)
            result = cursor.fetchone()
            datetime_last = result[0]
            logging.debug('Последняя дата из базы данных: ' + str(datetime_last))
    finally:
        connection.close()
    return datetime_last


def get_date_memory_from_device():
    with serial.Serial(COM, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                       bytesize=serial.EIGHTBITS, timeout=1) as ser:
        logging.debug('COM порт открыт')
        ser.write(init_port_with_crc)
        time.sleep(DELAY)
        ser.readall()
        time.sleep(DELAY)
        ser.write(date_request_with_crc)
        time.sleep(DELAY)
        date_memory_answer_hex = ser.readall()
        time.sleep(DELAY)
        logging.debug('COM порт закрыт')
        return date_memory_answer_hex


# Функция расчитывает разницу в часах между последним временем из базы данных и текущим временем
def delta_period(date_end, date_begin):
    if date_begin <= date_end:
        delta_in_period = (date_end - date_begin).total_seconds() // 3600 * 2
        logging.debug('Разница в периодах(30 минут): ' + str(delta_in_period))
        return int(delta_in_period)
    else:
        logging.debug('Разница в часах равна нулю или последнее время из базы <= текущему времени')
        return 0


def split_result_datetime(power_profile_answer):
    result_datetime = hex(power_profile_answer[6])[2:4] + '.' + hex(power_profile_answer[5])[2:4] + '.' + hex(
        power_profile_answer[4])[2:4] + ' ' + hex(power_profile_answer[2])[2:4] + ':' + hex(
        power_profile_answer[3])[2:4]
    result_datetime_object = datetime.datetime.strptime(result_datetime, '%d.%m.%y %H:%M')
    return result_datetime_object


def split_active_power(power_profile_answer):
    result_active_power = power_profile_answer[8] + power_profile_answer[9]
    active_power = float(result_active_power / 1000)
    return active_power


def split_reactive_power(power_profile_answer):
    result_reactive_power = power_profile_answer[12] + power_profile_answer[13]
    reactive_power = float(result_reactive_power / 1000)
    return reactive_power


def get_start_memory(memory_from_device, delta_in_period):
    int_memory = int.from_bytes(memory_from_device, byteorder='big')
    start_memory = int_memory - delta_in_period * PERIOD_HEX
    start_memory_hex = start_memory.to_bytes(2, byteorder='little')
    return start_memory_hex


check_com_port()
date_memory_answer_hex = get_date_memory_from_device()
convert_date(date_memory_answer_hex)
memory_hex = convert_memory(date_memory_answer_hex)
delta_period_int = delta_period(convert_date(date_memory_answer_hex), get_last_date_from_database())
print(get_start_memory(memory_hex, delta_period_int))

# byte_memory = int_memory.to_bytes(2, byteorder='big')


power_profile_request = DEVICE_NUMBER + MEMORY_BANK1 + convert_memory(date_memory_answer_hex) + PERIOD
crc16 = libscrc.modbus(power_profile_request)
power_profile_request_with_crc = power_profile_request + crc16.to_bytes(2, byteorder='little')

# with serial.Serial(COM, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
#                    bytesize=serial.EIGHTBITS, timeout=1) as ser:
#     logging.debug('COM порт открыт')
#     ser.write(init_port_with_crc)
#     time.sleep(DELAY)
#     ser.readall()
#     time.sleep(DELAY)
#     ser.write(power_profile_request_with_crc)
#     time.sleep(DELAY)
#     power_profile_answer_from_device = ser.readall()
#     time.sleep(DELAY)
#     ser.write(byte_memory)
#     time.sleep(DELAY)
#     ser.readall()
#     time.sleep(DELAY)
#     # for period in range(delta_period_int + 1):
#     #     print(period)
#     logging.debug('COM порт закрыт')
#
# print(split_result_datetime(power_profile_answer_from_device))
# print(split_active_power(power_profile_answer_from_device))
# print(split_reactive_power(power_profile_answer_from_device))
