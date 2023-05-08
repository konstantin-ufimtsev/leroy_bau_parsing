from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
import time
from datetime import datetime
from config import *
import psycopg2

cur_date = datetime.now().strftime('%d.%m.%Y')
cur_time = datetime.now().strftime('%H:%M:%S')


# создание таблицы для хранения результатов парсинга
def create_result_table():
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
                """CREATE TABLE result(
                    id serial PRIMARY KEY,
                    region VARCHAR(50),
                    date_of_parsing DATE,
                    time_of_parsing TIME, 
                    bau_article VARCHAR(50),
                    bau_name VARCHAR,
                    bau_price NUMERIC,
                    lm_article VARCHAR(50),
                    lm_name VARCHAR,
                    lm_price NUMERIC,
                    percent NUMERIC);
                """
            )
    except Exception as _ex:
        print('[INFO] Error while working with PostgreSQL,', _ex)


# загрузка пар в базу данных - сделано
# экспорт в эксель с формулами из базы данных - сделано

# парсинг
# 1. создание таблицы для записи результата - сделано
# 2. парсинг
# 3. запись в базу данных



options = webdriver.ChromeOptions()
options.add_argument("--disable-blink-features=AutomationControlled")
# options.add_experimental_option("excludeSwitches", ["enable-automation"])
# options.add_experimental_option('useAutomationExtension', False)

#s = Service(executable_path='path_to_chromedriver')
driver = webdriver.Chrome(options=options)

driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    'source': '''
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Array;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Promise;
        delete window.cdc_adoQpoasnfa76pfcZLmcfl_Symbol;
  '''
})

try:
    art_list = ["18669562", "18669546", "82576840", "82108181", "18669554", "18669571", "82138510", "82033830", "89837349", "82138042", "82138511", "82033832", "82138036"]
    for art in art_list:
        driver.maximize_window()
        driver.get(f'https://kaliningrad.leroymerlin.ru/search/?q={art}')
        time.sleep(0.5)
        price = driver.find_element(By.XPATH, '//*[@id="root"]/div/main/div[2]/div[2]/div/section/div[4]/section/div/div/div[2]/div/div/p[1]')
        name = driver.find_element(By.XPATH, '//*[@id="root"]/div/main/div[2]/div[2]/div/section/div[4]/section/div/div/div[1]/a/span/span')
        print(name.text)
        print(price.text)

except Exception as ex:
    print(ex)
finally:
    driver.close()
    driver.quit()



