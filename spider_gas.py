# TODO Ротация лога
# TODO Вывод сообщений в дискорд/телеграм/whatsup/viber
# TODO Проверка по количеству байтов в ответе
# TODO Вынести настройки порта в константы


# Библиотека для работы с датой (для расчета разницы во времени и прибавления времени)
import datetime
# Библиотека для работы с логаи
import logging
import logging.handlers
# Библиотека для работы с регулярными выражениями (умная разбивка строки)
import re

import sys
# Библиотека для работы со временем (паузы между запросами данных из порта)
import time
# Библиотека для работы с базой данных MySQL
import pymysql
# Библитотека для работы с COM портом
import serial

INIT_PORT = b'\x2F\x3F\x21\x0D\x0A'
OPEN_DEVICE = b'\x06\x30\x36\x31\x0D\x0A'
CLOSE_DEVICE = b'\x01\x42\x30\x03\x71\x01\x42\x30\x03\x71'
DEVICE_NAME = b'\x2f\x45\x6c\x73\x36\x45\x4b\x32\x36\x30\x0d\x0a'
BEGIN_REQ = b'\x01\x52\x33\x02\x33\x3a\x56\x2e\x30\x28\x33\x3b'
END_REQ = b'1)' + b'\x03'
DELAY = 2
NUMBER_OF_ATTEMPTS = 10
CRC_OK = 'CRC Ok'
NOT_FOUND = '#0103'
COM = 'COM5'
COM_SPEED = 19200
DATABASE_HOST = '10.1.1.50'
DATABASE_USER = 'user'
DATABASE_PASSWORD = 'qwerty123'
DATABASE = 'resources'


# TODO настроить время ротации и поправить отображение лога (время, дата)
# Функция для настройки логирования
def log_setup():
    log_handler = logging.handlers.WatchedFileHandler('spider_gas.txt')
    formatter = logging.Formatter(
        '%(asctime)s spider_gas.py [%(process)d]: %(message)s',
        '%b %d %H:%M:%S')
    formatter.converter = time.gmtime  # if you want UTC time
    log_handler.setFormatter(formatter)
    logger = logging.getLogger()
    logger.addHandler(log_handler)
    logger.setLevel(logging.DEBUG)


# Функция расчитывает разницу в часах между последним временем из базы данных и текущим временем
def delta_hours(date_end, date_begin):
    if date_begin <= date_end:
        delta_in_hours = (date_end - date_begin).total_seconds() // 3600
        logging.debug('Разница в часах: ' + str(delta_in_hours))
        return int(delta_in_hours)
    else:
        logging.debug('Разница в часах равна нулю или последнее время из базы <= текущему времени')
        return 0


# Функция формирует строку запроса перед расчитыванием контрольного значения CRC
def create_request_string(begin_request, date_request, end_request):
    date_request_str = str(date_request)[0:10] + ',' + str(date_request)[11:19] + ';'
    request_string_for_calc_crc = begin_request + date_request_str.encode() + date_request_str.encode() + end_request
    logging.debug('Сформированная строка для расчета CRC: ' + str(request_string_for_calc_crc))
    return request_string_for_calc_crc


# Функция расчитывает контрольное значение CRC
def calculate_crc(request_string_for_crc):
    crc_int = request_string_for_crc[1]
    for item in range(len(request_string_for_crc) - 2):
        crc_int = crc_int ^ request_string_for_crc[item + 2]
    crc_calc = crc_int.to_bytes(1, byteorder='big')
    logging.debug('Расчитанная котрольная сумма CRC: ' + str(crc_calc))
    return crc_calc


# Функция берет последнее значение времени из базы
def get_last_date_from_database():
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            sql = 'SELECT gas_datetime FROM gas ORDER BY gas_datetime DESC LIMIT 0, 1'
            cursor.execute(sql)
            result = cursor.fetchone()
            datetime_last = result[0]
            logging.debug('Последняя дата из базы данных: ' + str(datetime_last))
    finally:
        connection.close()
    return datetime_last


# Функция записывает считаные данные в базу
def insert_values_into_database(answer_for_database):
    gas_mark_gray = 1  # TODO прочитать в документации почему маркируется серым
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            sql = 'SELECT `gas_v_r_s`, `gas_v_st_s` FROM `gas` ORDER BY gas_datetime DESC LIMIT 0, 1'
            cursor.execute(sql)
            result = cursor.fetchone()
            gas_v_r_s = result[0]
            gas_v_st_s = result[1]
            gas_v_r_p = format(float(answer_for_database[7]) - float(gas_v_r_s), '.4f')
            gas_v_st_p = format(float(answer_for_database[5]) - float(gas_v_st_s), '.4f')
            sql = "INSERT INTO `gas` (`gas_n`, `gas_n2`, `gas_datetime`," \
                  "`gas_v_r_s`, `gas_v_st_s`, `gas_pressure`, " \
                  "`gas_temperature`, `gas_kkor`, `gas_sys_status`, " \
                  "`gas_status_vr`, `gas_status_vst`, `gas_status_p`, " \
                  "`gas_status_t`, `gas_crc_ok`, `gas_v_r_p`, " \
                  "`gas_v_st_p`, `gas_mark_gray`)" \
                  " VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s, %s)"
            cursor.execute(sql, (answer_for_database[1], answer_for_database[2], answer_for_database[3],
                                 answer_for_database[7], answer_for_database[5], answer_for_database[8][:6],
                                 answer_for_database[9], answer_for_database[11][:6], answer_for_database[12],
                                 answer_for_database[13], answer_for_database[14], answer_for_database[15],
                                 answer_for_database[16], answer_for_database[18], gas_v_r_p,
                                 gas_v_st_p, gas_mark_gray))
            connection.commit()
            logging.debug('Запись значений в базу')
    finally:
        connection.close()


# Функция проверки значений по CRC и времени
def check_crc_and_date(answer_for_check_crc):  # TODO Добавить проверку по дате
    check = split_answer_into_values(answer_for_check_crc)
    if check[-2] == CRC_OK:
        return CRC_OK
    if check[-2] == NOT_FOUND:
        return NOT_FOUND


# Функция разбивки строки на значения
def split_answer_into_values(answer):
    values = re.split(r'\)\(|\(|\)', answer.decode())
    # values.remove('')
    # values.pop()
    return values


# Функция для проверки доступности COM порта
def check_com_port():
    try:
        serial.Serial(COM, COM_SPEED, parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_ONE,
                      bytesize=serial.SEVENBITS, timeout=None)
        logging.debug('Порт доступен')
        return 0
    except serial.serialutil.SerialException:
        logging.debug('Порт недоступен')
        sys.exit("Порт недоступен")


# Функция для проверки доступности базы данных
def check_database_connection():
    try:
        pymysql.connect(host=DATABASE_HOST,
                        user=DATABASE_USER,
                        password=DATABASE_PASSWORD,
                        db=DATABASE)
        logging.debug('База данных доступна')
    except pymysql.err.OperationalError:
        logging.debug('База данных недоступна')
        sys.exit("База данных недоступна")


log_setup()

logging.debug('---------------------------------------')

# Проверяем доступность COM порта и базы данных, в случае неудачи выходим из программы
check_com_port()
check_database_connection()

date_now = datetime.datetime.now()
date_last = get_last_date_from_database()

with serial.Serial(COM, COM_SPEED, parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_ONE,
                   bytesize=serial.SEVENBITS, timeout=1) as ser:
    logging.debug('COM порт открыт')
    ser.write(INIT_PORT)
    time.sleep(DELAY)
    ser.readall()
    time.sleep(DELAY)
    ser.write(OPEN_DEVICE)
    time.sleep(DELAY)
    ser.readall()
    time.sleep(DELAY)

    for hours in range(delta_hours(date_now, date_last)):
        time.sleep(DELAY)
        date_req = date_last + datetime.timedelta(minutes=60 * (hours + 1))
        request_string = create_request_string(BEGIN_REQ, date_req, END_REQ)
        crc = calculate_crc(request_string)
        ser.write(request_string + crc)
        time.sleep(DELAY)
        answer_from_device = ser.readall()

        if check_crc_and_date(answer_from_device) == NOT_FOUND:
            # TODO Что делать в датой больше на 1 секунду, какое значение писать в базу и как проверять
            date_req = date_req + datetime.timedelta(seconds=1)
            request_string = create_request_string(BEGIN_REQ, date_req, END_REQ)
            crc = calculate_crc(request_string)
            ser.write(request_string + crc)
            time.sleep(DELAY)
            answer_from_device = ser.readall()
        logging.debug('Полученные данные со счетчика: ' + str(answer_from_device))
        insert_values_into_database(split_answer_into_values(answer_from_device))
    ser.write(CLOSE_DEVICE)
    logging.debug('COM порт закрыт')
