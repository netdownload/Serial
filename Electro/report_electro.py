# TODO Если понадобится конвертация в xls https://python-forum.io/Thread-how-to-convert-xlsx-to-xls
# TODO Добавить проверки по показаниям счетчиков
# pip install XlsxWriter
import xlsxwriter
# Библиотека для работы с базой данных MySQL
import pymysql
import sys
import datetime
import time
# Библиотека для расчета контрольной суммы
import libscrc
import serial
import sys

COMPANY = 'ОАО "Аньковское"'
ENGINEER = 'Главный энергетик_________________'
ENGINEER2 = 'Исполнитель_________________'
NAME = 'Кабанов С.В.'
PHONE_NUMBER = '(49353) 33102'
CONTRACT_NUMBER = '29'
CONTRACT_NUMBER2 = 'ЭСК-29'
CONTRACT_DATE = '01.07.2014'
# TODO Месяц и год надо брать из запроса
MONTH = 'Апрель'
YEAR = 2020
# ------------------------
DEVICE55 = 'Меркурий 230 ART-01'
DEVICE56 = 'Меркурий 230 ART-01'
DEVICE42 = 'Меркурий 230 ART-01'
DEVICE_NAME55_1 = '"Аньково" ВЛ-102 ЗТП-400\nТ-1 активный'
DEVICE_NAME55_2 = '"Аньково" ВЛ-102 ЗТП-400\nТ-1 реактивный'
DEVICE_NAME56_1 = '"Аньково" ВЛ-102 ЗТП-400\nТ-2 активный'
DEVICE_NAME56_2 = '"Аньково" ВЛ-102 ЗТП-400\nТ-2 реактивный'
DEVICE_NAME42_1 = '"Аньково" ВЛ-102 КТП-400\nОчистные сооружения'
# --------------------------
# TODO Коэффициент трансформации определять по запросу из счетчика
RATIO55 = 120
RATIO56 = 120
RATIO42 = 20
# --------------------------
DELAY = 1
COM42 = 'COM6'
COM55 = 'COM8'
COM56 = 'COM7'
COM_SPEED = 9600
DEVICE_ADDRESS42 = b'\x5E'  # Очистные
DEVICE_ADDRESS55 = b'\x5D'  # ТП-1
DEVICE_ADDRESS56 = b'\xBD'  # ТП-2
TEST_PORT = b'\x00'  # запрос на тестирование порта, в ответе должно прийти то же значение
INIT_PORT = b'\x01\x01\x01\x01\x01\x01\x01\x01'
DEVICE_NUMBER_REQUEST = b'\x08\x00'
# --------------------------
ACTIVE_LOSS = 3777
REACTIVE_LOSS = 20464
ACTIVE_LOSS2 = 3474
ACTIVE_LOSS3 = 303
# --------------------------
POWER_MONTH_JANUARY = b'\x05\x31\x00'
POWER_MONTH_FEBRUARY = b'\x05\x32\x00'
POWER_MONTH_MARCH = b'\x05\x33\x00'
POWER_MONTH_APRIL = b'\x05\x34\x00'
POWER_MONTH_MAY = b'\x05\x35\x00'
POWER_MONTH_JUNE = b'\x05\x36\x00'
POWER_MONTH_JULY = b'\x05\x37\x00'
POWER_MONTH_AUGUST = b'\x05\x38\x00'
POWER_MONTH_SEPTEMBER = b'\x05\x39\x00'
POWER_MONTH_OCTOBER = b'\x05\x3A\x00'
POWER_MONTH_NOVEMBER = b'\x05\x3B\x00'
POWER_MONTH_DECEMBER = b'\x05\x3C\x00'
# --------------------------
POWER_MONTH_JANUARY_BEGIN = b'\x05\xB1\x00'
POWER_MONTH_FEBRUARY_BEGIN = b'\x05\xB2\x00'
POWER_MONTH_MARCH_BEGIN = b'\x05\xB3\x00'
POWER_MONTH_APRIL_BEGIN = b'\x05\xB4\x00'
POWER_MONTH_MAY_BEGIN = b'\x05\xB5\x00'
POWER_MONTH_JUNE_BEGIN = b'\x05\xB6\x00'
POWER_MONTH_JULY_BEGIN = b'\x05\xB7\x00'
POWER_MONTH_AUGUST_BEGIN = b'\x05\xB8\x00'
POWER_MONTH_SEPTEMBER_BEGIN = b'\x05\xB9\x00'
POWER_MONTH_OCTOBER_BEGIN = b'\x05\xBA\x00'
POWER_MONTH_NOVEMBER_BEGIN = b'\x05\xBB\x00'
POWER_MONTH_DECEMBER_BEGIN = b'\x05\xBC\x00'
# --------------------------
DATABASE_HOST = '10.1.1.99'
DATABASE_USER = 'user'
DATABASE_PASSWORD = 'qwerty123'
DATABASE = 'resources'
# --------------------------
date_time_begin = '2020-03-01 00:30:00'
date_time_begin_obj = datetime.datetime.strptime(date_time_begin, '%Y-%m-%d %H:%M:%S')
date_time_end = '2020-04-01 00:00:00'
date_time_end_obj = datetime.datetime.strptime(date_time_end, '%Y-%m-%d %H:%M:%S')


# Функция для проверки доступности базы данных
def check_database_connection():
    try:
        pymysql.connect(host=DATABASE_HOST,
                        user=DATABASE_USER,
                        password=DATABASE_PASSWORD,
                        db=DATABASE)
        print('База данных доступна')
    except pymysql.err.OperationalError:
        print('База данных недоступна')
        sys.exit("База данных недоступна")


# Функция для проверки доступности COM порта
def check_com_port(com_port):
    try:
        serial.Serial(com_port, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                      bytesize=serial.EIGHTBITS, timeout=1)
        print('Порт доступен')
        return 0
    except serial.serialutil.SerialException:
        print('Порт недоступен')
        sys.exit("Порт недоступен")


def get_device_number(com_port, device_address):
    crc16 = libscrc.modbus(device_address + INIT_PORT)
    init_port_with_crc = device_address + INIT_PORT + crc16.to_bytes(2, byteorder='little')
    device_number_request = device_address + DEVICE_NUMBER_REQUEST
    crc16 = libscrc.modbus(device_number_request)
    device_number_request_with_crc = device_number_request + crc16.to_bytes(2, byteorder='little')
    with serial.Serial(com_port, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                       bytesize=serial.EIGHTBITS, timeout=1) as ser:
        ser.write(init_port_with_crc)
        time.sleep(DELAY)
        ser.readall()
        ser.write(device_number_request_with_crc)
        time.sleep(DELAY)
        device_number_hex = ser.readall()
        # TODO попробовать избавиться от проверок при помощи {}.format
        device_number1 = str(device_number_hex[1])
        if len(device_number1) < 2:
            device_number1 = '0' + device_number1
        device_number2 = str(device_number_hex[2])
        if len(device_number2) < 2:
            device_number2 = '0' + device_number2
        device_number3 = str(device_number_hex[3])
        if len(device_number3) < 2:
            device_number3 = '0' + device_number3
        device_number4 = str(device_number_hex[4])
        if len(device_number4) < 2:
            device_number4 = '0' + device_number4
        device_number_text = device_number1 + device_number2 + device_number3 + device_number4
        return device_number_text


def get_power_month_hex(com_port, device_address, power_month):
    crc16 = libscrc.modbus(device_address + INIT_PORT)
    init_port_with_crc = device_address + INIT_PORT + crc16.to_bytes(2, byteorder='little')
    power_month_request = device_address + power_month
    crc16 = libscrc.modbus(power_month_request)
    power_month_request_with_crc = power_month_request + crc16.to_bytes(2, byteorder='little')
    with serial.Serial(com_port, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                       bytesize=serial.EIGHTBITS, timeout=1) as ser:
        ser.write(init_port_with_crc)
        time.sleep(DELAY)
        ser.readall()
        ser.write(power_month_request_with_crc)
        time.sleep(DELAY)
        power_month_hex = ser.readall()
    return power_month_hex


def get_active_power_from_hex(power_answer, ratio):
    active_power_hex1 = str(hex(power_answer[2])[2:])
    active_power_hex2 = str(hex(power_answer[1])[2:])
    active_power_hex3 = str(hex(power_answer[4])[2:])
    active_power_hex4 = str(hex(power_answer[3])[2:])
    if len(active_power_hex1) < 2:
        active_power_hex1 = '0' + active_power_hex1
    if len(active_power_hex2) < 2:
        active_power_hex2 = '0' + active_power_hex2
    if len(active_power_hex3) < 2:
        active_power_hex3 = '0' + active_power_hex3
    if len(active_power_hex4) < 2:
        active_power_hex4 = '0' + active_power_hex4
    active_power_hex = active_power_hex1 + active_power_hex2 + active_power_hex3 + active_power_hex4
    active_power = int(active_power_hex, 16) / 1000 * ratio
    return active_power


def get_reactive_power_from_hex(power_answer, ratio):
    reactive_power_hex1 = str(hex(power_answer[10])[2:])
    reactive_power_hex2 = str(hex(power_answer[9])[2:])
    reactive_power_hex3 = str(hex(power_answer[12])[2:])
    reactive_power_hex4 = str(hex(power_answer[11])[2:])
    if len(reactive_power_hex1) < 2:
        reactive_power_hex1 = '0' + reactive_power_hex1
    if len(reactive_power_hex2) < 2:
        reactive_power_hex2 = '0' + reactive_power_hex2
    if len(reactive_power_hex3) < 2:
        reactive_power_hex3 = '0' + reactive_power_hex3
    if len(reactive_power_hex4) < 2:
        reactive_power_hex4 = '0' + reactive_power_hex4
    reactive_power_hex = reactive_power_hex1 + reactive_power_hex2 + reactive_power_hex3 + reactive_power_hex4
    reactive_power = int(reactive_power_hex, 16) / 1000 * ratio
    return reactive_power


def get_power_month_begin(com_port, device_address, month_begin):
    crc16 = libscrc.modbus(device_address + INIT_PORT)
    init_port_with_crc = device_address + INIT_PORT + crc16.to_bytes(2, byteorder='little')
    power_month_begin_request = device_address + month_begin
    crc16 = libscrc.modbus(power_month_begin_request)
    power_month_begin_request_with_crc = power_month_begin_request + crc16.to_bytes(2, byteorder='little')
    with serial.Serial(com_port, COM_SPEED, parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                       bytesize=serial.EIGHTBITS, timeout=1) as ser:
        ser.write(init_port_with_crc)
        time.sleep(DELAY)
        ser.readall()
        ser.write(power_month_begin_request_with_crc)
        time.sleep(DELAY)
        power_month_begin_hex = ser.readall()
    return power_month_begin_hex


# Функция возвращает список значений активной мощности 42 счетчика
def get_power_values_from_database42(date_begin, date_end):
    answer_from_database42 = []
    delta_time = delta_period(date_end, date_begin)
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            for values in range(delta_time + 1):
                sql = 'SELECT electro42_active FROM electro42 WHERE electro42_datetime=%s LIMIT 0, 1'
                cursor.execute(sql, date_begin)
                result = cursor.fetchone()
                power_active42 = result
                try:
                    answer_from_database42.append(power_active42[0])
                except TypeError:
                    answer_from_database42.append(0)
                date_begin = date_begin + datetime.timedelta(minutes=30)
    finally:
        connection.close()
    return answer_from_database42


# Функция возвращает список значений активной мощности 55 счетчика
def get_power_values_from_database55(date_begin, date_end):
    answer_from_database55 = []
    delta_time = delta_period(date_end, date_begin)
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            for values in range(delta_time + 1):
                sql = 'SELECT electro55_active FROM electro55 WHERE electro55_datetime=%s LIMIT 0, 1'
                cursor.execute(sql, date_begin)
                result = cursor.fetchone()
                power_active55 = result
                try:
                    answer_from_database55.append(power_active55[0])
                except TypeError:
                    answer_from_database55.append(0)
                date_begin = date_begin + datetime.timedelta(minutes=30)
    finally:
        connection.close()
    return answer_from_database55


# Функция возвращает список значений активной мощности 56 счетчика
def get_power_values_from_database56(date_begin, date_end):
    answer_from_database56 = []
    delta_time = delta_period(date_end, date_begin)
    connection = pymysql.connect(host=DATABASE_HOST,
                                 user=DATABASE_USER,
                                 password=DATABASE_PASSWORD,
                                 db=DATABASE)
    try:
        with connection.cursor() as cursor:
            for values in range(delta_time + 1):
                sql = 'SELECT electro56_active FROM electro56 WHERE electro56_datetime=%s LIMIT 0, 1'
                cursor.execute(sql, date_begin)
                # TODO Дабавить считывание данных из базы при неполном срезе, данные будут между датами
                result = cursor.fetchone()

                power_active56 = result
                try:
                    answer_from_database56.append(power_active56[0])
                except TypeError:
                    answer_from_database56.append(0)
                date_begin = date_begin + datetime.timedelta(minutes=30)
    finally:
        connection.close()
    return answer_from_database56


def delta_period(date_end, date_begin):
    if date_begin <= date_end:
        delta_in_period = (date_end - date_begin).total_seconds() // 1800
        print('Разница в периоде 30 минут: ' + str(delta_in_period))
        return int(delta_in_period)
    else:
        print('Разница в равна нулю или последнее время из базы <= текущему времени')
        return 0


check_database_connection()

check_com_port(COM55)
device_number55 = get_device_number(COM55, DEVICE_ADDRESS55)
power_answer_month = get_power_month_hex(COM55, DEVICE_ADDRESS55, POWER_MONTH_MARCH)
active_power_month55 = get_active_power_from_hex(power_answer_month, RATIO55)
reactive_power_month55 = get_reactive_power_from_hex(power_answer_month, RATIO55)
power_answer_month_begin = get_power_month_begin(COM55, DEVICE_ADDRESS55, POWER_MONTH_MARCH_BEGIN)
active_power_begin_month55 = get_active_power_from_hex(power_answer_month_begin, 1)
reactive_power_begin_month55 = get_reactive_power_from_hex(power_answer_month_begin, 1)
power_answer_month_end = get_power_month_begin(COM55, DEVICE_ADDRESS55, POWER_MONTH_APRIL_BEGIN)
active_power_end_month55 = get_active_power_from_hex(power_answer_month_end, 1)
reactive_power_end_month55 = get_reactive_power_from_hex(power_answer_month_end, 1)

check_com_port(COM56)
device_number56 = get_device_number(COM56, DEVICE_ADDRESS56)
power_answer_month = get_power_month_hex(COM56, DEVICE_ADDRESS56, POWER_MONTH_MARCH)
active_power_month56 = get_active_power_from_hex(power_answer_month, RATIO56)
reactive_power_month56 = get_reactive_power_from_hex(power_answer_month, RATIO56)
power_answer_month_begin = get_power_month_begin(COM56, DEVICE_ADDRESS56, POWER_MONTH_MARCH_BEGIN)
active_power_begin_month56 = get_active_power_from_hex(power_answer_month_begin, 1)
reactive_power_begin_month56 = get_reactive_power_from_hex(power_answer_month_begin, 1)
power_answer_month_end = get_power_month_begin(COM56, DEVICE_ADDRESS56, POWER_MONTH_APRIL_BEGIN)
active_power_end_month56 = get_active_power_from_hex(power_answer_month_end, 1)
reactive_power_end_month56 = get_reactive_power_from_hex(power_answer_month_end, 1)

check_com_port(COM42)
device_number42 = get_device_number(COM42, DEVICE_ADDRESS42)
power_answer_month = get_power_month_hex(COM42, DEVICE_ADDRESS42, POWER_MONTH_MARCH)
active_power_month42 = get_active_power_from_hex(power_answer_month, RATIO42)
reactive_power_month42 = get_reactive_power_from_hex(power_answer_month, RATIO42)
power_answer_month_begin = get_power_month_begin(COM42, DEVICE_ADDRESS42, POWER_MONTH_MARCH_BEGIN)
active_power_begin_month42 = get_active_power_from_hex(power_answer_month_begin, 1)
reactive_power_begin_month42 = get_reactive_power_from_hex(power_answer_month_begin, 1)
power_answer_month_end = get_power_month_begin(COM42, DEVICE_ADDRESS42, POWER_MONTH_APRIL_BEGIN)
active_power_end_month42 = get_active_power_from_hex(power_answer_month_end, 1)
reactive_power_end_month42 = get_reactive_power_from_hex(power_answer_month_end, 1)

answer42 = get_power_values_from_database42(date_time_begin_obj, date_time_end_obj)
answer55 = get_power_values_from_database55(date_time_begin_obj, date_time_end_obj)
answer56 = get_power_values_from_database56(date_time_begin_obj, date_time_end_obj)

delta = delta_period(date_time_end_obj, date_time_begin_obj)

workbook = xlsxwriter.Workbook('demo.xlsx')
worksheet = workbook.add_worksheet()
worksheet.name = 'Worksheet'

bold = workbook.add_format({'bold': True})

format_center = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_center_bold = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_center_without_borders = workbook.add_format({
    'bold': 0,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_center_without_borders_bold = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_center_with_borders = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_right = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'right',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_right_bold = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'right',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_right_with_borders_bold_times_12_number3 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'right',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
    'num_format': 3,
    # 'num_format': x -> https://xlsxwriter.readthedocs.io/format.html#num-format-categories
})

format_right_without_borders_bold = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'right',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_left_bold = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
})

format_left_without_borders_bold = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 0,
})

format_left_with_borders_bold_underline_times_12 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
    'underline': 1,
})

format_left_with_borders_bold_times_14_bg_color = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 14,
    'bg_color': '#FFFF99',
})

format_left_with_borders_bold_times_12 = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
})

format_left_with_borders_times_12 = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
})

format_center_without_borders_bold_times_16 = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 16,
})

format_center_without_borders_bold_times_14 = workbook.add_format({
    'bold': 1,
    'border': 0,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 14,
})

format_center_without_borders_bold_times_16_number4_bg_color = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 16,
    'bg_color': '#FFFF99',
    'num_format': 3,
})

format_center_with_borders_times_12 = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
})

format_center_with_borders_times_12_number = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 12,
    'num_format': 4,
    # 'num_format': x -> https://xlsxwriter.readthedocs.io/format.html#num-format-categories
})

format_center_with_borders_times_8 = workbook.add_format({
    'bold': 0,
    'border': 1,
    'align': 'center',
    'valign': 'vcenter',
    'text_wrap': 1,
    'font_name': 'Times New Roman',
    'font_size': 8,
})

worksheet.set_column('A:A', 23.3)
worksheet.set_column('B:Z', 11.3)

worksheet.write('A1', COMPANY, bold)
worksheet.write('A2', 'Отчет за потребленную электроэнергию и мощность, ' + MONTH + ' ' + str(YEAR) + ' г.', bold)

worksheet.write('A3', 'Счетчик №')
worksheet.write('B3', device_number55, format_right_without_borders_bold)
worksheet.write('A4', 'Тр.тока (коэф)')
worksheet.write('B4', RATIO55, bold)

worksheet.merge_range('B5:B6', 'Число расчетного месяца', format_center)
worksheet.set_row(5, 30)
worksheet.merge_range('C5:Z5', 'Время суток', format_center)

for cells in range(0, 24, 1):
    worksheet.write(5, cells + 2, str(cells) + '.00-' + str(cells + 1) + '.00', format_center)

for rows in range(0, 31, 1):
    worksheet.write(rows + 6, 1, rows + 1, format_center)

for rows in range(0, 31, 1):
    for cells in range(0, 24, 1):
        worksheet.write(rows + 6, cells + 2, None, format_right)

cells_in = 0

for rows in range(0, int((delta + 1) / 24 / 2), 1):
    for cells in range(0, 48, 2):
        worksheet.write(rows + 6, (cells + 2) / 2 + 1, (answer55[cells_in] + answer55[cells_in + 1]) / 2, format_right)
        print('cell = ', cells_in)
        print('A1 = ', answer55[cells_in], ' A2 = ', answer55[cells_in + 1])
        print((answer55[cells_in] + answer55[cells_in + 1]) / 2)
        cells_in += 2

worksheet.write('Y38', 'ИТОГО', format_left_bold)
worksheet.write('Z38', '=SUM(C7:Z37)*120', format_right_bold)

worksheet.write('A39', 'Счетчик №')
worksheet.write('B39', device_number56, format_right_without_borders_bold)
worksheet.write('A40', 'Тр.тока (коэф)')
worksheet.write('B40', RATIO56, bold)

for rows in range(0, 31, 1):
    worksheet.write(rows + 40, 1, rows + 1, format_center)

for rows in range(0, 31, 1):
    for cells in range(0, 24, 1):
        worksheet.write(rows + 40, cells + 2, None, format_right)
cells_in = 0

for rows in range(0, int((delta + 1) / 24 / 2), 1):
    for cells in range(0, 48, 2):
        worksheet.write(rows + 40, (cells + 2) / 2 + 1, (answer56[cells_in] + answer56[cells_in + 1]) / 2, format_right)
        cells_in += 2

worksheet.write('Y72', 'ИТОГО', format_left_bold)
worksheet.write('Z72', '=SUM(C41:Z71)*120', format_right_bold)

worksheet.write('A73', 'Счетчик №')
worksheet.write('B73', device_number42, format_right_without_borders_bold)
worksheet.write('A74', 'Тр.тока (коэф)')
worksheet.write('B74', RATIO42, bold)

for rows in range(0, 31, 1):
    worksheet.write(rows + 74, 1, rows + 1, format_center)

for rows in range(0, 31, 1):
    for cells in range(0, 24, 1):
        worksheet.write(rows + 74, cells + 2, None, format_right)

cells_in = 0

for rows in range(0, int((delta + 1) / 24 / 2), 1):
    for cells in range(0, 48, 2):
        worksheet.write(rows + 74, (cells + 2) / 2 + 1, (answer42[cells_in] + answer42[cells_in + 1]) / 2, format_right)
        cells_in += 2

worksheet.write('Y106', 'ИТОГО', format_left_bold)
# TODO в MS Excel 2003 не происходит расчет формул на первом листе
worksheet.write('Z106', "=SUM(C75:Z105)*20", format_right_bold)

worksheet.write('B109', ENGINEER + NAME)

# -----------------NEXT SHEET 'Summa'-----------------

worksheet2 = workbook.add_worksheet()
worksheet2.name = 'Summa'

worksheet2.set_column('A:A', 2.3)
worksheet2.set_column('B:Z', 11.3)

worksheet2.merge_range('A1:Z1', 'Сведения', format_center_without_borders)
worksheet2.merge_range('G2:T2', 'о фактическом почасовом расходе электрической энергии за '
                       + MONTH + ' ' + str(YEAR) + ' года ' + COMPANY, format_center_without_borders)
worksheet2.merge_range('B3:D3', 'Договор №' + CONTRACT_NUMBER)
worksheet2.merge_range('B5:B6', 'Число расчетного месяца', format_center)
worksheet2.set_row(5, 30)
worksheet2.merge_range('C5:Z5', 'Время суток', format_center)
worksheet2.merge_range('X3:Z3', 'Уровень СН2')

for cells in range(0, 24, 1):
    worksheet2.write(5, cells + 2, str(cells) + '.00-' + str(cells + 1) + '.00', format_center)

for rows in range(0, 31, 1):
    worksheet2.write(rows + 6, 1, rows + 1, format_center)

for rows in range(0, 31, 1):
    # TODO Попробовать цикл по алфавиту
    #  https://stackoverflow.com/questions/17182656/how-do-i-iterate-through-the-alphabet
    worksheet2.write(rows + 6, 2,
                     '=Worksheet!C' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!C' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!C' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 3,
                     '=Worksheet!D' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!D' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!D' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 4,
                     '=Worksheet!E' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!E' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!E' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 5,
                     '=Worksheet!F' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!F' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!F' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 6,
                     '=Worksheet!G' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!G' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!G' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 7,
                     '=Worksheet!H' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!H' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!H' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 8,
                     '=Worksheet!I' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!I' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!I' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 9,
                     '=Worksheet!J' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!J' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!J' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 10,
                     '=Worksheet!K' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!K' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!K' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 11,
                     '=Worksheet!L' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!L' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!L' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 12,
                     '=Worksheet!M' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!M' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!M' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 13,
                     '=Worksheet!N' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!N' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!N' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 14,
                     '=Worksheet!O' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!O' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!O' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 15,
                     '=Worksheet!P' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!P' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!P' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 16,
                     '=Worksheet!Q' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!Q' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!Q' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 17,
                     '=Worksheet!R' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!R' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!R' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 18,
                     '=Worksheet!S' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!S' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!S' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 19,
                     '=Worksheet!T' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!T' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!T' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 20,
                     '=Worksheet!U' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!U' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!U' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 21,
                     '=Worksheet!V' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!V' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!V' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 22,
                     '=Worksheet!W' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!W' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!W' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 23,
                     '=Worksheet!X' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!X' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!X' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 24,
                     '=Worksheet!Y' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!Y' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!Y' + str(rows + 75) + '*' + str(RATIO42), format_right)
    worksheet2.write(rows + 6, 25,
                     '=Worksheet!Z' + str(rows + 7) + '*' + str(RATIO55) + ' + Worksheet!Z' + str(rows + 41)
                     + '*' + str(RATIO56) + ' + Worksheet!Z' + str(rows + 75) + '*' + str(RATIO42), format_right)

worksheet2.write('Y38', 'Сумма', format_left_bold)
worksheet2.write('Z38', '=SUM(C7:Z37)', format_right_bold)

worksheet2.write('Q41', 'М.П.')
worksheet2.write('B41', ENGINEER + NAME)

# ---------------NEXT SHEET 'Rashod'-------------
worksheet3 = workbook.add_worksheet()
worksheet3.name = 'Rashod'

worksheet3.set_column('A:A', 9.3)
worksheet3.set_column('B:B', 24.3)
worksheet3.set_column('C:C', 14.3)
worksheet3.set_column('D:D', 14.3)
worksheet3.set_column('E:E', 14.3)
worksheet3.set_column('F:F', 14.3)
worksheet3.set_column('G:G', 14.3)

worksheet3.merge_range('A1:G1', 'СВЕДЕНИЯ', format_center_without_borders_bold)
worksheet3.merge_range('A2:G2', 'о расходе электроэнергии по ' + COMPANY + ', за ' + MONTH + ' ' + str(YEAR)
                       + ' года.', format_center_without_borders_bold)
worksheet3.set_row(2, 30)
worksheet3.merge_range('A3:A4', '№ п/п', format_center_bold)
worksheet3.merge_range('B3:B4', 'Наименование точки учета', format_center_bold)
worksheet3.merge_range('C3:C4', 'Номер прибора учета', format_center_bold)
worksheet3.merge_range('D3:E3', 'Показания прибора учета', format_center_bold)
worksheet3.write('D4', 'Конечное', format_center_bold)
worksheet3.write('E4', 'Начальное', format_center_bold)
worksheet3.merge_range('F3:F4', 'Расчетный к-т', format_center_bold)
worksheet3.merge_range('G3:G4', 'Расход э/энергии (кВтч)', format_center_bold)
for cells in range(0, 7):
    worksheet3.write(4, cells, str(cells + 1), format_center_bold)
worksheet3.write('A6', '1.', format_center_bold)
worksheet3.merge_range('B6:G6', 'Потребление электроэнергии Потребителем', format_left_bold)
worksheet3.write('A7', '1.1', format_center_bold)
worksheet3.write('B7', DEVICE_NAME55_1, format_center_bold)
worksheet3.write('C7', device_number55, format_center_bold)
worksheet3.write('D7', active_power_end_month55, format_center_bold)
worksheet3.write('E7', active_power_begin_month55, format_center_bold)
worksheet3.write('F7', RATIO55, format_center_bold)
worksheet3.write('G7', active_power_month55, format_center_bold)
worksheet3.write('A8', '1.2', format_center_bold)
worksheet3.write('B8', DEVICE_NAME55_2, format_center_bold)
worksheet3.write('C8', device_number55, format_center_bold)
worksheet3.write('D8', reactive_power_end_month55, format_center_bold)
worksheet3.write('E8', reactive_power_begin_month55, format_center_bold)
worksheet3.write('F8', RATIO55, format_center_bold)
worksheet3.write('G8', reactive_power_month55, format_center_bold)
worksheet3.write('A9', '1.3', format_center_bold)
worksheet3.write('B9', DEVICE_NAME56_1, format_center_bold)
worksheet3.write('C9', device_number56, format_center_bold)
worksheet3.write('D9', active_power_end_month56, format_center_bold)
worksheet3.write('E9', active_power_begin_month56, format_center_bold)
worksheet3.write('F9', RATIO56, format_center_bold)
worksheet3.write('G9', active_power_month56, format_center_bold)
worksheet3.write('A10', '1.4', format_center_bold)
worksheet3.write('B10', DEVICE_NAME56_2, format_center_bold)
worksheet3.write('C10', device_number56, format_center_bold)
worksheet3.write('D10', reactive_power_end_month56, format_center_bold)
worksheet3.write('E10', reactive_power_begin_month56, format_center_bold)
worksheet3.write('F10', RATIO56, format_center_bold)
worksheet3.write('G10', reactive_power_month56, format_center_bold)
worksheet3.write('A11', '1.5', format_center_bold)
worksheet3.write('B11', DEVICE_NAME42_1, format_center_bold)
worksheet3.write('C11', device_number42, format_center_bold)
worksheet3.write('D11', active_power_end_month42, format_center_bold)
worksheet3.write('E11', active_power_begin_month42, format_center_bold)
worksheet3.write('F11', RATIO42, format_center_bold)
worksheet3.write('G11', active_power_month42, format_center_bold)
worksheet3.write('A12', '2.', format_center_bold)
worksheet3.merge_range('B12:F12', 'Потери э/э оплачиваемые активные', format_left_bold)
worksheet3.write('G12', ACTIVE_LOSS, format_center_bold)
worksheet3.write('A13', '2.1', format_center_bold)
for cells in range(1, 7):
    worksheet3.write(12, cells, None, format_center_bold)
worksheet3.write('A14', '3.', format_center_bold)
worksheet3.merge_range('B14:F14', 'Потери э/э оплачиваемые реактивные', format_left_bold)
worksheet3.write('G14', REACTIVE_LOSS, format_center_bold)
worksheet3.write('A15', '3.1', format_center_bold)
worksheet3.write('B15', 'Реактивное потребление\nпо "Правилу"', format_center_bold)
for cells in range(2, 7):
    worksheet3.write(14, cells, None, format_center_bold)
worksheet3.write('A16', '3.2', format_center_bold)
for cells in range(1, 7):
    worksheet3.write(15, cells, None, format_center_bold)
worksheet3.write('A17', '4.', format_center_bold)
worksheet3.merge_range('B17:G17', 'Замена прибора учета э/э', format_left_bold)
worksheet3.write('A18', '4.1', format_center_bold)
for cells in range(1, 7):
    worksheet3.write(17, cells, None, format_center_bold)
worksheet3.write('A19', '4.2', format_center_bold)
worksheet3.merge_range('B19:F19', 'ИТОГО ОБЩИЙ РАСХОД (С ПОТЕРЯМИ), АКТИВНЫЙ', format_left_bold)
worksheet3.write('G19', active_power_month55 + active_power_month56 + active_power_month42 + ACTIVE_LOSS,
                 format_center_bold)
worksheet3.write('B20', None, format_center_bold)
worksheet3.merge_range('B20:F20', 'ИТОГО ОБЩИЙ РАСХОД (С ПОТЕРЯМИ), РЕАКТИВНЫЙ', format_left_bold)
worksheet3.write('G20', reactive_power_month55 + reactive_power_month56 + REACTIVE_LOSS, format_center_bold)
for cells in range(0, 7):
    worksheet3.write(20, cells, None, format_center_bold)
worksheet3.write('F23', 'М. П.', format_left_without_borders_bold)
worksheet3.write('B25', ENGINEER2 + NAME, format_left_without_borders_bold)
worksheet3.write('B27', 'телефон ' + PHONE_NUMBER, format_left_without_borders_bold)

# ---------------NEXT SHEET 'Svedeniya'-------------

worksheet4 = workbook.add_worksheet()
worksheet4.name = 'Svedeniya'

worksheet4.set_column('A:A', 61.6)
worksheet4.set_column('B:B', 28)
worksheet4.set_column('C:C', 9)
worksheet4.set_column('D:D', 22.7)
worksheet4.set_column('E:E', 15.6)
worksheet4.set_column('F:F', 16.6)
worksheet4.set_column('G:G', 12.7)
worksheet4.set_column('H:H', 12.7)
worksheet4.set_column('I:I', 3)
worksheet4.set_column('J:J', 6.6)
worksheet4.set_column('K:K', 7.7)
worksheet4.set_column('L:L', 18.4)

worksheet4.set_row(2, 7.5)
worksheet4.set_row(4, 5.25)
worksheet4.set_row(5, 6)
worksheet4.set_row(6, 27)
worksheet4.set_row(7, 21)
worksheet4.set_row(10, 13.5)
worksheet4.set_row(12, 30)
worksheet4.set_row(13, 30)
worksheet4.set_row(14, 30)
worksheet4.set_row(15, 30)
worksheet4.set_row(16, 30)
worksheet4.set_row(17, 19)
worksheet4.set_row(18, 30)
worksheet4.set_row(19, 30)
worksheet4.set_row(20, 30)
worksheet4.set_row(21, 30)
worksheet4.set_row(22, 30)
worksheet4.set_row(23, 0)
worksheet4.set_row(24, 30)
worksheet4.set_row(25, 30)
worksheet4.set_row(26, 15)

worksheet4.merge_range('A1:L1', 'СВЕДЕНИЯ', format_center_without_borders_bold_times_16)
worksheet4.merge_range('A2:L2', 'о расходе электроэнергии за ' + MONTH + ' ' + str(YEAR) + ' г.',
                       format_center_without_borders_bold_times_14)
worksheet4.merge_range('A4:L4', COMPANY + ' № ' + CONTRACT_NUMBER2 + ' от ' + CONTRACT_DATE + ' г.',
                       format_center_without_borders_bold_times_14)
worksheet4.merge_range('A7:A8', 'Наименование объекта', format_center_with_borders_times_12)
worksheet4.merge_range('B7:B8', '№\nфидера', format_center_with_borders_times_12)
worksheet4.merge_range('C7:C8', '№\nКТП', format_center_with_borders_times_12)
worksheet4.merge_range('D7:D8', 'Тип счетчика', format_center_with_borders_times_12)
worksheet4.merge_range('E7:E8', '№ счетчика', format_center_with_borders_times_12)
worksheet4.merge_range('F7:F8', 'Дата снятия показаний счетчика', format_center_with_borders_times_12)
worksheet4.merge_range('G7:H7', 'Показания счетчика', format_center_with_borders_times_12)
worksheet4.write('G8', 'начальное', format_center_with_borders_times_12)
worksheet4.write('H8', 'конечное', format_center_with_borders_times_12)
worksheet4.merge_range('I7:J8', 'Разность показаний', format_center_with_borders_times_12)
worksheet4.merge_range('K7:K8', 'Коэффи-\nциент счетчика', format_center_with_borders_times_12)
worksheet4.merge_range('L7:L8', 'Расход электроэнергии\n(кВтч)', format_center_with_borders_times_12)
for cells in range(0, 7):
    worksheet4.write(8, cells, cells + 1, format_center_with_borders_times_8)
worksheet4.merge_range('I9:J9', '9', format_center_with_borders_times_8)
worksheet4.write('K9', '10', format_center_with_borders_times_8)
worksheet4.write('L9', '11', format_center_with_borders_times_8)
worksheet4.merge_range('A10:L10', 'Потребление электроэнергии Потребителем',
                       format_left_with_borders_bold_times_14_bg_color)
worksheet4.merge_range('A11:L11', None, format_center_with_borders_times_12)
worksheet4.write('A12', 'Потребители с макс. мощностью  от 150кВт до 670кВт, СН-2',
                 format_left_with_borders_bold_underline_times_12)
for cells in range(1, 6):
    worksheet4.write(11, cells, None, format_center_with_borders_times_12)
worksheet4.merge_range('G12:K12', None, format_center_with_borders_times_12)
worksheet4.write('L12', None, format_center_with_borders_times_12)
worksheet4.merge_range('A13:A14', '1.1. ПС "Аньково", ВЛ-10кВ №102, 107', format_left_with_borders_bold_times_12)
worksheet4.merge_range('B13:B14', 'РУ-0,4кВ ЗТП 2*400кВА', format_center_with_borders_times_12)
worksheet4.write('C13', 'Т-1', format_center_with_borders_times_12)
worksheet4.write('C14', 'Т-2', format_center_with_borders_times_12)
worksheet4.write('D13', DEVICE55, format_center_with_borders_times_12)
worksheet4.write('D14', DEVICE56, format_center_with_borders_times_12)
worksheet4.write('E13', device_number55, format_center_with_borders_times_12)
worksheet4.write('E14', device_number56, format_center_with_borders_times_12)
worksheet4.write('F13', date_time_begin_obj.strftime('%d.%m.%Y'), format_center_with_borders_times_12)
worksheet4.write('F14', date_time_begin_obj.strftime('%d.%m.%Y'), format_center_with_borders_times_12)
worksheet4.merge_range('G13:K13', 'по профилю потребителя', format_center_with_borders_times_12)
worksheet4.merge_range('G14:K14', 'по профилю потребителя', format_center_with_borders_times_12)
worksheet4.write('L13', active_power_month55, format_center_with_borders_times_12_number)
worksheet4.write('L14', active_power_month56, format_center_with_borders_times_12_number)
worksheet4.write('A15', 'ИТОГО:', format_left_with_borders_bold_times_12)
worksheet4.merge_range('B15:L15', (active_power_month55 + active_power_month56),
                       format_right_with_borders_bold_times_12_number3)
worksheet4.write('A16', 'Потери э/э оплачиваемые:', format_left_with_borders_bold_times_12)
worksheet4.merge_range('B16:L16', ACTIVE_LOSS2, format_right_with_borders_bold_times_12_number3)
worksheet4.write('A17', 'ВСЕГО: от 150кВт до 670кВт', format_left_with_borders_bold_times_12)
worksheet4.merge_range('B17:L17', (active_power_month55 + active_power_month56 + ACTIVE_LOSS2),
                       format_right_with_borders_bold_times_12_number3)
worksheet4.merge_range('A18:L18', None, format_center_with_borders_times_12)
worksheet4.merge_range('A19:A21', 'Потребители с макс. Мощностью менее 150кВт, СН-2',
                       format_left_with_borders_bold_times_12)
worksheet4.merge_range('B19:B21', None, format_center_with_borders_times_12)
worksheet4.merge_range('C19:C21', None, format_center_with_borders_times_12)
worksheet4.merge_range('D19:D21', None, format_center_with_borders_times_12)
worksheet4.merge_range('E19:E21', None, format_center_with_borders_times_12)
worksheet4.merge_range('F19:F21', None, format_center_with_borders_times_12)
worksheet4.merge_range('G19:K21', None, format_center_with_borders_times_12)
worksheet4.merge_range('L19:L21', None, format_center_with_borders_times_12)
worksheet4.merge_range('A22:A23', '1.1. "Очистные сооружения"', format_left_with_borders_times_12)
worksheet4.merge_range('B22:B23', 'РУ-0,4кВ КТП-100кВа', format_center_with_borders_times_12)
worksheet4.merge_range('C22:C23', None, format_center_with_borders_times_12)
worksheet4.write('D22', DEVICE42, format_center_with_borders_times_12)
worksheet4.write('D23', None, format_center_with_borders_times_12)
worksheet4.write('E22', device_number42, format_center_with_borders_times_12)
worksheet4.write('E23', None, format_center_with_borders_times_12)
worksheet4.write('F22', date_time_begin_obj.strftime('%d.%m.%Y'), format_center_with_borders_times_12)
worksheet4.write('F23', None, format_center_with_borders_times_12)
worksheet4.merge_range('G22:K22', 'по профилю потребителя', format_center_with_borders_times_12)
worksheet4.merge_range('G23:K23', None, format_center_with_borders_times_12)
worksheet4.write('L22', active_power_month42, format_center_with_borders_times_12)
worksheet4.write('L23', None, format_center_with_borders_times_12)
worksheet4.merge_range('A25:A27', 'Потери э/э оплачиваемые:', format_left_with_borders_times_12)
worksheet4.merge_range('B25:L27', ACTIVE_LOSS3, format_right_with_borders_bold_times_12_number3)
worksheet4.merge_range('A28:A30', 'ИТОГО: менее 150кВт', format_left_with_borders_bold_times_12)
worksheet4.merge_range('B28:L30', (active_power_month42 + ACTIVE_LOSS3),
                       format_right_with_borders_bold_times_12_number3)
worksheet4.write('A31', None, format_center_with_borders_times_12)
worksheet4.write('B31', None, format_center_with_borders_times_12)
worksheet4.write('C31', None, format_center_with_borders_times_12)
worksheet4.write('D31', None, format_center_with_borders_times_12)
worksheet4.write('E31', None, format_center_with_borders_times_12)
worksheet4.write('F31', None, format_center_with_borders_times_12)
worksheet4.merge_range('G31:K31', None, format_center_with_borders_times_12)
worksheet4.write('L31', None, format_center_with_borders_times_12)
worksheet4.write('A32', None, format_center_with_borders_times_12)
worksheet4.write('B32', None, format_center_with_borders_times_12)
worksheet4.write('C32', None, format_center_with_borders_times_12)
worksheet4.write('D32', None, format_center_with_borders_times_12)
worksheet4.write('E32', None, format_center_with_borders_times_12)
worksheet4.write('F32', None, format_center_with_borders_times_12)
worksheet4.merge_range('G32:K32', None, format_center_with_borders_times_12)
worksheet4.write('L32', None, format_center_with_borders_times_12)
worksheet4.write('A33', None, format_center_with_borders_times_12)
worksheet4.write('B33', None, format_center_with_borders_times_12)
worksheet4.write('C33', None, format_center_with_borders_times_12)
worksheet4.write('D33', None, format_center_with_borders_times_12)
worksheet4.write('E33', None, format_center_with_borders_times_12)
worksheet4.write('F33', None, format_center_with_borders_times_12)
worksheet4.merge_range('G33:K33', None, format_center_with_borders_times_12)
worksheet4.write('L33', None, format_center_with_borders_times_12)
worksheet4.write('A34', None, format_center_with_borders_times_12)
worksheet4.write('B34', None, format_center_with_borders_times_12)
worksheet4.write('C34', None, format_center_with_borders_times_12)
worksheet4.write('D34', None, format_center_with_borders_times_12)
worksheet4.write('E34', None, format_center_with_borders_times_12)
worksheet4.write('F34', None, format_center_with_borders_times_12)
worksheet4.merge_range('G34:K34', None, format_center_with_borders_times_12)
worksheet4.write('L34', None, format_center_with_borders_times_12)
worksheet4.merge_range('A35:K35', 'ИТОГО по договору №ЭСК-29', format_left_with_borders_bold_times_14_bg_color)
worksheet4.write('L35', (active_power_month55 + active_power_month56 + active_power_month42 + ACTIVE_LOSS),
                 format_center_without_borders_bold_times_16_number4_bg_color)
worksheet4.merge_range('A37:L37', 'Генеральный директор____________________________________________________',
                       format_center_without_borders)
worksheet4.merge_range('B40:D40', 'Исполнитель:______________________________', format_center_without_borders)
worksheet4.merge_range('I40:L40', '______________________________ /дата/', format_right)
workbook.close()
