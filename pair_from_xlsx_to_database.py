import openpyxl
from config import *
import psycopg2
from datetime import datetime


# создание таблица для хранения пар для парсинга
def create_pairs_table():
    try:
        connection = psycopg2.connect(
            host=host,
            user=user,
            password=password,
            database=db_name

        )
        connection.autocommit = True
        with connection.cursor() as cursor:
            cursor.execute(
                """CREATE TABLE pairs(
                    id serial PRIMARY KEY,
                    bau_article VARCHAR(50),
                    bau_name VARCHAR,
                    lm_article VARCHAR(50),
                    lm_name VARCHAR,
                    region VARCHAR(50),
                    percent NUMERIC);
                """
            )
    except Exception as _ex:
        print('[INFO] Error while working with PostgreSQL,', _ex)




# функция записывает данные из таблицы пар экселя в базу данных "pairs"
def from_excel_to_database(filename):
    book = openpyxl.load_workbook(filename=filename)
    sheet = book.active

    try:
        connection = psycopg2.connect(
            host=host,
            user=user,
            password=password,
            database=db_name
        )
        connection.autocommit = True
        with connection.cursor() as cursor:
            for i in range(2, sheet.max_row + 1):
                cursor.execute(
                    'INSERT INTO pairs (bau_article, bau_name, lm_article, lm_name, region, percent) VALUES (%s, %s, %s, %s, %s, %s)',
                    (sheet[f'A{i}'].value, sheet[f'B{i}'].value, sheet[f'C{i}'].value, sheet[f'D{i}'].value,
                     sheet[f'E{i}'].value, sheet[f'F{i}'].value))

    except Exception as _ex:
        print('[INFO] Error while working with PostgreSQL,', _ex)


# pair.xlsx - файл с парами для загрузки
def main():
    pass

#   create_result_table()
#   create_pairs_table()
#   from_excel_to_database("pair.xlsx")


if __name__ == '__main__':
    main()
