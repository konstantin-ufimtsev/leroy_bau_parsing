import openpyxl
from openpyxl import Workbook
from config import *
import psycopg2
from datetime import datetime
from openpyxl.styles import Font
from openpyxl.styles import Alignment


# функция выгружает данные из базы result в excel
def from_database_to_excel():
    book = openpyxl.Workbook()
    sheet = book.active

    sheet.column_dimensions['A'].width = 5  # количество символов
    sheet.column_dimensions['B'].width = 15  # количество символов
    sheet.column_dimensions['C'].width = 15  # количество символов
    sheet.column_dimensions['D'].width = 10  # количество символов
    sheet.column_dimensions['E'].width = 9  # количество символов
    sheet.column_dimensions['F'].width = 50  # количество символов
    sheet.column_dimensions['G'].width = 7  # количество символов
    sheet.column_dimensions['H'].width = 6  # количество символов
    sheet.column_dimensions['I'].width = 50  # количество символов
    sheet.column_dimensions['J'].width = 7  # количество символов
    sheet.column_dimensions['K'].width = 6  # количество символов
    sheet.column_dimensions['M'].width = 8  # количество символов
    sheet.column_dimensions['N'].width = 8  # количество символов
    sheet.column_dimensions['O'].width = 6  # количество символов

    sheet.row_dimensions[1].height = 60  # устанавливаем высоту певрой строки строки  - 30

    font = Font(bold=True, size=11, color='000000')
    # sheet['A1'].alignment = Alignment(wrapText=True)

    sheet.row_dimensions[1].font = font  # устанавливаем жирный шрифт для первой строки
    sheet.row_dimensions[1].alignment = Alignment(vertical='center')  # перенос по иширне в первой строке
    sheet.row_dimensions[1].alignment = Alignment(wrapText=True)  # перенос по иширне в первой строке

    sheet['A1'].value = 'Номер'
    sheet['B1'].value = 'Регион'
    sheet['C1'].value = 'Дата получения'
    sheet['D1'].value = 'Время получения'
    sheet['E1'].value = 'Артикул Бау'
    sheet['F1'].value = 'Наименование Бау'
    sheet['G1'].value = 'Цена Бау'
    sheet['H1'].value = 'Артикул ЛМ'
    sheet['I1'].value = 'Наименование ЛМ'
    sheet['J1'].value = 'Цена ЛМ'
    sheet['K1'].value = 'Допустимый процент'
    sheet['L1'].value = 'Фактический процент'
    sheet['M1'].value = 'Новая цена'
    sheet['N1'].value = 'Новая цена с округление'
    sheet['O1'].value = 'Новый процент'

    try:
        connection = psycopg2.connect(
            host=host,
            user=user,
            password=password,
            database=db_name
        )
        connection.autocommit = True
        with connection.cursor() as cursor:
            cursor.execute('SELECT * FROM result')
            record = cursor.fetchall()
            i = 2
            for row in record:
                sheet[f'A{i}'].value = row[0]  # N
                sheet[f'B{i}'].value = row[1]  # region
                sheet[f'C{i}'].value = row[2]  # date_of_parsing
                sheet[f'D{i}'].value = row[3]  # time_of_parsing
                sheet[f'E{i}'].value = row[4]  # bau_article
                sheet[f'F{i}'].value = row[5]  # bau_name
                sheet[f'G{i}'].value = row[6]  # bau_price
                sheet[f'H{i}'].value = row[7]  # lm_article
                sheet[f'I{i}'].value = row[8]  # lm_name
                sheet[f'J{i}'].value = row[9]  # lm_price
                sheet[f'K{i}'].value = row[10]  # percent
                sheet[f'L{i}'].value = f"=G{i}/J{i} - 1"  # percent_deviation

                sheet[f'L{i}'].number_format = '0.0%'  # устанавливаем формат процентов

                sheet[f'M{i}'].value = f"=J{i} * (1 + K{i}/100)"  # new_price
                sheet[f'N{i}'].value = f"=IF(AND(M{i}<1000,M{i}>10),ROUND(M{i},-1)-1,IF(M{i}>=1000,ROUND(M{i},-2)-10,IF(M{i}<=10,ROUND(M{i},0)+1)))"  # new_price_rounded
                sheet[f'O{i}'].value = f"=N{i}/J{i} - 1"  # new_percent

                sheet[f'O{i}'].number_format = '0.0%'  # устанавливаем формат процентов

                i += 1


    except Exception as _ex:
        print('[INFO] Error while working with PostgreSQL,', _ex)

    book.save('result.xlsx')


def main():
    from_database_to_excel()


if __name__ == '__main__':
    main()
