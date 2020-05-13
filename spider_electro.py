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
TEST_PORT = DEVICE_NUMBER + b'\x00' # запрос на тестирование порта, в ответе должно прийти то же значение
INIT_PORT = DEVICE_NUMBER  + b'\x01\x01\x01\x01\x01\x01\x01\x01'
DELAY = 0.1
COM = 'COM9'
COM_SPEED = 9600
DATABASE_HOST = '10.1.1.99'
DATABASE_USER = 'user'
DATABASE_PASSWORD = 'qwerty123'
DATABASE = 'resources'


crc16 = libscrc.modbus(INIT_PORT)
init_port_with_crc = INIT_PORT + crc16.to_bytes(2, byteorder='little')
crc16 = libscrc.modbus(TEST_PORT)
test_port_with_crc = TEST_PORT + crc16.to_bytes(2, byteorder='little')

# TODO настроить время ротации и поправить отображение лога (время, дата)
# Функция для настройки логирования
def log_setup():
    log_handler = logging.handlers.WatchedFileHandler('spider_electro.txt')
    formatter = logging.Formatter(
        '%(asctime)s spider_gas.py [%(process)d]: %(message)s',
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


check_com_port()

with serial.Serial(COM, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                   bytesize=serial.EIGHTBITS, timeout=1) as ser:
    logging.debug('COM порт открыт')
    ser.write(init_port_with_crc)
    time.sleep(DELAY)
    ser.readall()
    time.sleep(DELAY)
    logging.debug('COM порт закрыт')
