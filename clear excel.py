import re
import shutil

import openpyxl
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

# Загружаем входной файл
input_table = '/Users/egoryurihin/Downloads/Боевая таблица за 2015 год.xlsx'

# Определяем место расположения выходного файла
output_table = '/Users/egoryurihin/Downloads/Новая Боевая таблица за 2015 год.xlsx'

# Создаем выходной файл с метаданными путем копирования входного
shutil.copyfile(input_table, output_table)

# Создаем список страниц входного файла
sheet_names = openpyxl.load_workbook(input_table).sheetnames

# Запускаем браузер без графики
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('--disable-notifications')
options.add_argument("--disable-gpu")
driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)

# Задаем адрес страницы
driver.get("https://eduscan.net/help/phone")

# Находим форму ввода
elem = driver.find_element_by_name("num")

# Создаем счетчик для отсчета итераций при ожидании
time_1 = len(sheet_names)-1

# Создаем список с наименованием столбцов
c_names = ['ID', 'Дата', 'Запрос/Клиент', 'Менеджер', 'Источник информации о нас', 'Чем интересуется', 'Итог',
               'Прим.', 'Повторные контакты', 'Данные', 'опроса', 'клиента', 'Комментарии']

# Стартуем постраничную обработку файла
for s_name in sheet_names:
    # Считываем страницу
    df = pd.read_excel(input_table, sheet_name=s_name)
    columns_count = len(df.columns)
    df = df.drop([0])
    # Присваиваем столбцам названия
    c_names_1 = []
    for i in range(0, columns_count):
        c_names_1.append(c_names[i])
    df.columns = c_names_1

    # Заполняем пропуски нулями (для определения только рабочих строк таблицы)
    df['ID'] = df['ID'].fillna(0)
    df = df.loc[df['ID'] != 0]

    # Заполняем отсутствующие данные
    df = df.fillna('!Данные не указаны!')

    # Забираем данные на разбор
    client = df['Запрос/Клиент'], str.split(',')

    # Вытаскиваем почту
    e_mails = []
    for e in client[0]:
        if '@' in e:
            e_mails.append(re.findall('[a-zA-z0-9\.]+[@][a-zA-z0-9\S]+[.][a-zA-z]*\w', e))
        else:
            e_mails.append('0')
    e_mail = []
    for i in e_mails:
        e_mail.append(i)
    for i in range(0, len(e_mail)):
        if e_mail[i] == '0':
            e_mail[i] = '!Почта не указана!'
    df['Электронная почта'] = e_mail # Добавляем новый столбец и вносим в него почтовые адреса

    # Обрабатываем номера телефонов
    phones = []
    for i in client[0]:
        phones.append(re.findall('[0-9+]+\d', i))
    phone = []
    for i in phones:
        phone.append(''.join(i))
    for i in range(0, len(phone)):
        if len(phone[i]) < 11:
            phone[i] = '0'
    for i in range(0, len(phone)):
        if '+' in phone[i]:
            phone[i] = phone[i].split('+')
    for i in phone:
        if len(i) > 1 and i[0] == '':
            i.remove(i[0])
    for i in range(0, len(phone)):
        if type(phone[i]) == list and len(phone[i]) == 1:
            phone[i] = phone[i][0]
    for i in range(0, len(phone)):
        if phone[i][0] == '8':
            phone[i] = phone[i].replace('8', '7', 1)
    for i in range(0, len(phone)):
        if phone[i] == '0':
            phone[i] = '!Телефон не указан!'
    for i in range(0, len(phone)):
        if type(phone[i]) == list:
            phone[i] = phone[i][0]

    # Определяем по телефону регион путем заполнения формы на сайте https://eduscan.net/help/phone и считывания региона
    region_name = []
    for i in phone:
        if i != '!Телефон не указан!':
            elem.send_keys(i[:10])
            elem.send_keys(Keys.ENTER)
            region = driver.find_element_by_id("sstring3")
            region_name.append(region.text)
            elem.clear()
        else:
            region_name.append('!Регион не определен!')
    for i in range(0, len(region_name)):
        if region_name[i] == '':
            region_name[i] = '!Регион не определен!'
    df['Регион'] = region_name # Создаем новый столбец и заполняем его наименованием регионов

    # Сохраняем данные в файл в новую страницу
    book = load_workbook(output_table)
    writer = pd.ExcelWriter(output_table, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, s_name)
    writer.save()

    # Выводим сообщение для информирования пользователя
    print('Ожидайте...',time_1)
    time_1 = time_1 - 1

print('!!!Форматирование завершено!!!')
driver.close()
