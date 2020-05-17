# bd 1d f3 09 22 30 14 05 20 1e 6f 54
# Не понятно как рассчитывать последнюю запись в ячейке
# Возможно надо ориентироваться по 3 байту "09" и брать ячейку d f3 0
# в этом случае будет использоваться банк памяти 83
# Если 3 байт равен "19" то брать ячейку 1d f3 (в этом случае тут будет значение кратное 16), банк памяти брать 03
# В документации данного момента не нашел

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

DEVICE_NUMBER = b'\xBD'  # ТП-2
TEST_PORT = DEVICE_NUMBER + b'\x00'  # запрос на тестирование порта, в ответе должно прийти то же значение
INIT_PORT = DEVICE_NUMBER + b'\x01\x01\x01\x01\x01\x01\x01\x01'
DATE_MEMORY_REQUEST = DEVICE_NUMBER + b'\x08\x13'
# В ответе проверить байт стостояния, возможно он отвечает за то какой банк памяти брать
DELAY = 0.2
COM = 'COM7'
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
    memory_block_for_check = date_memory_answer_hex[3]
    if memory_block_for_check == 9:
        memory_answer_text1 = "{0:#0{1}x}".format(date_memory_answer_hex[1], 4)[3]
        memory_answer_text2 = "{0:#0{1}x}".format(date_memory_answer_hex[2], 4)[2:4]
        memory_answer_text3 = "{0:#0{1}x}".format(date_memory_answer_hex[3], 4)[2]
        memory_text = memory_answer_text1 + memory_answer_text2 + memory_answer_text3
        memory_answer = bytes.fromhex(memory_text)
    elif memory_block_for_check == 25:
        memory_answer_text1 = "{0:#0{1}x}".format(date_memory_answer_hex[1], 4)[2:4]
        memory_answer_text2 = "{0:#0{1}x}".format(date_memory_answer_hex[2], 4)[2:4]
        memory_text = memory_answer_text1 + memory_answer_text2
        memory_answer = bytes.fromhex(memory_text)
    else:
        memory_answer = ''
    return memory_answer


# Функция берет последнее значение времени из базы
def get_last_date_from_database():
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            sql = 'SELECT electro56_datetime FROM electro56 ORDER BY electro56_datetime DESC LIMIT 0, 1'
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


def delta_period(date_end, date_begin):
    if date_begin <= date_end:
        delta_in_period = (date_end - date_begin).total_seconds() // 1800
        logging.debug('Разница в периодах(30 минут): ' + str(delta_in_period))
        return int(delta_in_period)
    else:
        logging.debug('Разница в часах равна нулю или последнее время из базы <= текущему времени')
        return 0


def split_result_datetime(power_profile_answer):
    # TODO Обработать исключение, если в ответе будет '', т.е. из порта не поступят данные
    result_datetime = hex(power_profile_answer[4])[2:4] + '.' + hex(power_profile_answer[5])[2:4] + '.' + hex(
        power_profile_answer[6])[2:4] + ' ' + hex(power_profile_answer[2])[2:4] + ':' + hex(
        power_profile_answer[3])[2:4]
    print(result_datetime)
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
    start_memory_hex = start_memory.to_bytes(2, byteorder='big')
    return start_memory_hex


def get_next_memory(memory):
    int_memory = int.from_bytes(memory, byteorder='big')
    next_memory = int_memory + PERIOD_HEX
    next_memory_hex = next_memory.to_bytes(2, byteorder='big')
    return next_memory_hex


def create_profile_request(memory_request):
    power_profile_request = DEVICE_NUMBER + MEMORY_BANK2 + memory_request + PERIOD
    crc16 = libscrc.modbus(power_profile_request)
    power_profile_request_with_crc = power_profile_request + crc16.to_bytes(2, byteorder='little')
    return power_profile_request_with_crc


# TODO Функция проверки перехода с одного банка памяти на другой
def check_memory_bank():
    pass


# Функция записывает считаные данные в базу
def insert_values_into_database(gas_datetime, active_power, reactive_power):
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            sql = "INSERT INTO `electro56` (`electro56_datetime`, `electro56_active`, `electro56_reactive`) VALUES (" \
                  "%s, %s, %s)"
            cursor.execute(sql, (gas_datetime, active_power, reactive_power))
            connection.commit()
            logging.debug('Запись значений в базу')
    finally:
        connection.close()


check_com_port()

date_memory_answer_hex = get_date_memory_from_device()
memory_hex = convert_memory(date_memory_answer_hex)
delta_period_int = delta_period(convert_date(date_memory_answer_hex), get_last_date_from_database())
memory_start = get_start_memory(memory_hex, delta_period_int)

with serial.Serial(COM, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                   bytesize=serial.EIGHTBITS, timeout=1) as ser:
    logging.debug('COM порт открыт')
    ser.write(init_port_with_crc)
    time.sleep(DELAY)
    ser.readall()
    for period in range(delta_period_int):
        memory_start = get_next_memory(memory_start)
        power_profile_request_with_crc = create_profile_request(memory_start)
        time.sleep(DELAY)
        ser.write(power_profile_request_with_crc)
        time.sleep(DELAY)
        power_profile_answer_from_device = ser.readall()
        time.sleep(DELAY)
        print(memory_start)
        gas_datetime = split_result_datetime(power_profile_answer_from_device)
        print(gas_datetime)
        active_power = split_active_power(power_profile_answer_from_device)
        print(active_power)
        reactive_power = split_reactive_power(power_profile_answer_from_device)
        print(reactive_power)
        insert_values_into_database(gas_datetime, active_power, reactive_power)
    logging.debug('COM порт закрыт')
