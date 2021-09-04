import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import openpyxl

""" Программа для поиска лотов на сайте ГосЗакупок
    Лоты считываются с файла lots.xlsx
    Результат работы программы выгружается в файл result.xlsx
"""


def str_finder(string, param):
    if param == 'Окончание подачи заявок':
        if string[(string.find(param) + param.__len__() + 1):].__len__() > 15:
            return 'Не указано'
        else:
            return string[(string.find(param) + param.__len__() + 1):]
    else:
        return string[
               (string.find(param) + param.__len__() + 1):string.find("\n", string.find(param) + param.__len__() + 1)]


# Открываем файл с лотами
book = openpyxl.open("lots.xlsx", read_only=True)
# Выбираем первый лист ( По умолчанию первый лист - active)
sheet = book.active
# Объявляем список лотов
lots = []
# Заполняем данными из файлика
for i in range(1, sheet.max_row + 1):
    lots.append(sheet[i][0].value)
# Запускаем Chrome
browser = webdriver.Chrome()
# Будем искать следующие запросы
lots_params = ['№', 'Заказчик', 'Размещено',
               'Окончание подачи заявок',
               'Объект закупки', 'Начальная цена']
# Закрываем файл с лотами
book.close()
# Создаем новый файл
book = openpyxl.Workbook()
# Выбираем активным первый лист
sheet = book.active
# Формируем первую строку
first_string = ['Лот']
# Добавляем параметры из списка
first_string = first_string.__add__(lots_params)
# Добавляем изменения на лист Excel файла
sheet.append(first_string)
for lot in lots:
    # Для каждого лота заходим на стартовую страничку сайта
    browser.get('https://zakupki.gov.ru')
    # Находим элемент для ввода запроса
    search = browser.find_element_by_xpath("//input[@role='search']")
    # Нажимаем на него
    search.click()
    # Отправляем строку из списка lots
    search.send_keys(f'{lot}')
    # Симулируем нажатие клавиши ENTER
    search.send_keys(Keys.ENTER)
    # На странице всего 10 элементов, поэтому проходим по ним в цикле
    for i in range(1, 11):
        # Начинаем формировать строку
        row = [f'{lot}']
        # Создаем переменную для того, чтобы иметь возможность скроллить
        html = browser.find_element_by_tag_name('html')
        # Скроллим вниз
        for down in range(8):
            html.send_keys(Keys.DOWN)
        # Находим элемент по xPath
        # По данному xPath находится 10 элементов на странице, то есть 10 лотов
        # Соответственно мы будем проходить по каждому в цикле, присваивая значение i
        # И двигаясь дальше
        lot_block = browser.find_element_by_xpath(
            f"//div[@class='search-registry-entry-block box-shadow-search-input'][{i}]")
        # Заполняем строку далее по параметрам
        for lot_param in lots_params:
            # вызываем функцию за каждый из внесенных параметров поиска
            item = str_finder(lot_block.text, lot_param)
            # Добавляем в список, формирующий строку
            row.append(item)
        # И под конец каждой итерации цикла добавляем строку на страницу
        sheet.append(row)
# Сохраняем файл c именем времени заверщения
time = datetime.datetime.now()
name_of_file = str(time.day) + '_' + str(time.month) + '_' + str(time.year) + '__' + str(time.hour) + '_' + str(
    time.minute) + '.xlsx'
book.save(name_of_file)
# Закрываем файл
book.close()
print("Выполнено успешно!")
