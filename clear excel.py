import re
import shutil

import openpyxl
import pandas as pd
from openpyxl import load_workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager

input_table: str = '/Users/egoryurihin/Downloads/Боевая таблица за 2015 год.xlsx'
output_table = '/Users/egoryurihin/Downloads/Новая Боевая таблица за 2015 год.xlsx'
shutil.copyfile(input_table, output_table)
sheet_names = openpyxl.load_workbook(input_table).sheetnames
options = webdriver.ChromeOptions()
options.add_argument('headless')
options.add_argument('--disable-notifications')
options.add_argument("--disable-gpu")
driver = webdriver.Chrome(ChromeDriverManager().install(), chrome_options=options)
driver.get("https://eduscan.net/help/phone")
elem = driver.find_element_by_name("num")
time_1: int = len(sheet_names)-1
for s_name in sheet_names:
    df = pd.read_excel(input_table, sheet_name=s_name)
    columns_count = len(df.columns)
    df = df.drop([0])
    c_names = ['ID', 'Дата', 'Запрос/Клиент', 'Менеджер', 'Источник информации о нас', 'Чем интересуется', 'Итог',
               'Прим.', 'Повторные контакты', 'Данные', 'опроса', 'клиента', 'Комментарии']
    c_names_1 = []
    for i in range(0, columns_count):
        c_names_1.append(c_names[i])
    df.columns = c_names_1
    df['ID'] = df['ID'].fillna(0)
    df = df.loc[df['ID'] != 0]
    df = df.fillna('!Данные не указаны!')
    client = df['Запрос/Клиент'], str.split(',')
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
    df['Электронная почта'] = e_mail
    phones = []
    for i in client[0]:
        phones.append(re.findall('[0-9+]+\d', i))
    phone = []
    for i in phones:
        phone.append(''.join(i))
    for i in range(0, len(phone)):
        if len(phone[i]) < 11:
            phone[i] = '0'
    # for i in range(0,len(phone)):
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
    df['Регион'] = region_name
    book = load_workbook(output_table)
    writer = pd.ExcelWriter(output_table, engine='openpyxl')
    writer.book = book
    writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
    df.to_excel(writer, s_name)
    writer.save()
    print('Ожидайте...',time_1)
    time_1 = time_1 - 1
print('Форматирование завершено!')
driver.close()
