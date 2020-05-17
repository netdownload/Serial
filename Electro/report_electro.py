# pip install XlsxWriter
import xlsxwriter
# Библиотека для работы с базой данных MySQL
import pymysql
import sys
import datetime
import time

COMPANY = 'ОАО "Аньковское"'
ENGINEER = 'Главный энергетик_______________'
NAME = 'Кабанов С.В.'
CONTRACT_NUMBER = '29'
# TODO Месяц и год надо брать из запроса
MONTH = 'Апрель'
YEAR = 2020
# ------------------------
# TODO Номера приборов определять по запросу из счетчика
DEVICE_NUMBER1 = 25506893
DEVICE_NUMBER2 = 1753886
DEVICE_NUMBER3 = 4479066
# --------------------------
# TODO Коэффициент трансформации определять по запросу из счетчика
RATIO1 = 120
RATIO2 = 120
RATIO3 = 20
# --------------------------
DATABASE_HOST = '10.1.1.99'
DATABASE_USER = 'user'
DATABASE_PASSWORD = 'qwerty123'
DATABASE = 'resources'
# --------------------------
date_time_begin = '2020-04-01 00:30:00'
date_time_begin_obj = datetime.datetime.strptime(date_time_begin, '%Y-%m-%d %H:%M:%S')
date_time_end = '2020-05-01 00:00:00'
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

format_center_without_borders = workbook.add_format({
    'bold': 0,
    'border': 0,
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

format_left = workbook.add_format({
    'bold': 1,
    'border': 1,
    'align': 'left',
    'valign': 'vcenter',
    'text_wrap': 1,
})


worksheet.set_column('A:A', 23.3)
worksheet.set_column('B:Z', 11.3)

worksheet.write('A1', COMPANY, bold)
worksheet.write('A2', 'Отчет за потребленную электроэнергию и мощность, ' + MONTH + ' ' + str(YEAR) + ' г.', bold)

worksheet.write('A3', 'Счетчик №')
worksheet.write('B3', DEVICE_NUMBER1, bold)
worksheet.write('A4', 'Тр.тока (коэф)')
worksheet.write('B4', RATIO1, bold)

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

worksheet.write('Y38', 'ИТОГО', format_left)
worksheet.write('Z38', '=SUM(C7:Z37)*' + str(RATIO1), format_right_bold)

worksheet.write('A39', 'Счетчик №')
worksheet.write('B39', DEVICE_NUMBER2, bold)
worksheet.write('A40', 'Тр.тока (коэф)')
worksheet.write('B40', RATIO2, bold)

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

worksheet.write('Y72', 'ИТОГО', format_left)
worksheet.write('Z72', '=SUM(C41:Z71)*' + str(RATIO2), format_right_bold)

worksheet.write('A73', 'Счетчик №')
worksheet.write('B73', DEVICE_NUMBER3, bold)
worksheet.write('A74', 'Тр.тока (коэф)')
worksheet.write('B74', RATIO3, bold)

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

worksheet.write('Y106', 'ИТОГО', format_left)
worksheet.write('Z106', '=SUM(C75:Z105)*' + str(RATIO3), format_right_bold)

worksheet.write('B109', ENGINEER + NAME)

# -----------------NEXT SHEET 'Summa'-----------------

worksheet2 = workbook.add_worksheet()
worksheet2.name = 'Summa'

worksheet2.set_column('A:A', 2.3)
worksheet2.set_column('B:Z', 11.3)

worksheet2.merge_range('A1:Z1', 'Сведения', format_center_without_borders)
worksheet2.merge_range('G2:T2', 'о фактическом почасовом расходе электрической энергии за Апрель 2020 года ОАО '
                                '"Аньковское"', format_center_without_borders)
worksheet2.merge_range('B3:D3', 'Договор №' + CONTRACT_NUMBER)
worksheet2.merge_range('B5:B6', 'Число расчетного месяца', format_center)
worksheet2.merge_range('C5:Z5', 'Время суток', format_center)
worksheet2.merge_range('X3:Z3', 'Уровень СН2')


for cells in range(0, 24, 1):
    worksheet2.write(5, cells + 2, str(cells) + '.00-' + str(cells + 1) + '.00', format_center)

for rows in range(0, 31, 1):
    worksheet2.write(rows + 6, 1, rows + 1, format_center)

for rows in range(0, 31, 1):
    # TODO Попробовать цикл по алфавиту
    #  https://stackoverflow.com/questions/17182656/how-do-i-iterate-through-the-alphabet
    worksheet2.write(rows + 6, 2, '=Worksheet!C' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!C' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!C' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 3, '=Worksheet!D' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!D' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!D' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 4, '=Worksheet!E' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!E' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!E' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 5, '=Worksheet!F' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!F' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!F' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 6, '=Worksheet!G' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!G' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!G' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 7, '=Worksheet!H' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!H' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!H' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 8, '=Worksheet!I' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!I' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!I' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 9, '=Worksheet!J' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!I' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!I' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 10, '=Worksheet!K' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!K' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!K' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6,11, '=Worksheet!L' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!L' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!L' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 12, '=Worksheet!M' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!M' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!M' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 13, '=Worksheet!N' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!N' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!N' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 14, '=Worksheet!O' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!P' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!O' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 15, '=Worksheet!P' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!P' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!P' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 16, '=Worksheet!Q' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!Q' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!Q' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 17, '=Worksheet!R' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!R' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!R' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 18, '=Worksheet!S' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!S' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!S' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 19, '=Worksheet!T' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!T' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!T' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 20, '=Worksheet!U' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!U' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!U' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 21, '=Worksheet!V' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!V' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!V' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 22, '=Worksheet!W' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!W' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!W' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 23, '=Worksheet!X' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!X' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!X' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 24, '=Worksheet!Y' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!Y' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!Y' + str(rows + 75) + '*' + str(RATIO3), format_right)
    worksheet2.write(rows + 6, 25, '=Worksheet!Z' + str(rows + 7) + '*' + str(RATIO1) + ' + Worksheet!Z' + str(rows + 41)
                     + '*' + str(RATIO2) + ' + Worksheet!Z' + str(rows + 75) + '*' + str(RATIO3), format_right)

worksheet2.write('Y38', 'Сумма', format_left)
worksheet2.write('Z38', '=SUM(C7:Z37)', format_right_bold)

worksheet2.write('Q41', 'М.П.')
worksheet2.write('B41', ENGINEER + NAME)
workbook.close()
