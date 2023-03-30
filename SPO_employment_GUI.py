import tkinter as tk
"""
Скрипт для подсчета показателей трудоустройства выпускников
"""
import pandas as pd
import os
import openpyxl
import math

from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time

# Отображать все колонки в пандас
pd.set_option('display.max_columns', None)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
     Для того чтобы упаковать картинку в exe"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def select_files_data():
    """
    Функция для выбора файлов с данными параметры из которых нужно подсчитать
    :return: Путь к файлам с данными
    """
    global name_file_data
    # Получаем путь к файлу
    name_file_data = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder():
    """
    Функция для выбора папки куда будут генерироваться файл  с результатом подсчета и файл с проверочной инфомрацией
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()


def calculate_data():
    """
    Функция для подсчета данных из файлов
    :return:
    """

    current_time = time.strftime('%H_%M_%S')
    # Создаем список всех специальностей которые могут встретиться
    all_code_spec = []
    # Колонки для базового датафрейма
    columns = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12',
               '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24',
               '25', '26', '27', '28', '29', '30', '31', '32', '33', '34', '35', '36',
               '37', '38', '39', '40', '41', '42', '43', '44', '45', '46', '47', '48',
               '49', '50']
    base_correct_df = pd.DataFrame(columns=columns)




    # Итерирумся по списку файлов для обработки

    try:

        # Открываем лист со списком СПО, с опцией для чтения чтобы ускорить процесс
        temp = openpyxl.load_workbook(f'{name_file_data}', read_only=True)
        # Получаем список листов
        temp_sheets = temp.sheetnames
        # Очищаем список от лишних листов
        # Создаем кортеж с названиями листов которые не будем обрабатывать
        del_sheets = ('Форма 1', 'Форма 2', 'Коды и наименования программ', 'СВОД общий','Сводная таблица 1')

        sheets = [sheet for sheet in temp_sheets if sheet not in del_sheets]
        # Закрываем лист после чтения
        temp.close()
       # Создаем словарь верхнего уровня

        high_dct = dict.fromkeys(sheets, dict())

        # Создаем  общий список для специальностей


        # Открываем файл с указанным листом и сохраняем его в датафрейм пропуская  первые 8 строк
        for sheet in sheets:
            # Счетчик чтобы определять на какой строке возникла ошибка(чаще всего просто пустая строка)
            count_rows = 10
            print(sheet)

            df = pd.read_excel(f'{name_file_data}', sheet_name=sheet, skiprows=8,
                               dtype={'Код профессии, специальности*': str})

            # Добавляем полученный датафрейм в базовый
            base_correct_df = base_correct_df.append(df, ignore_index=True)

            # Создаем список специальность которые есть на данном листе
            raw_codes_spec = df['04'].unique()

            # Очищаем получившиеся коды от nan
            codes_spec = [code for code in raw_codes_spec if str(code) != 'nan']
            # Очищаем коды от пустых строк
            codes_spec = [code for code in codes_spec if code != ' ']

            # Добавляем полученный список в общий список специальностей
            all_code_spec.extend(codes_spec)

            # Создаем словарь через comprehnsions
            # Отбираем только строковые, отбрасывая ключ nan. В качестве значения присваиваем копию словаря низкого уровня
            # Забыл про ссылочную модель и не мог долго понять почему значения дублируются
            spec_code_dct = {code: create_low_dct() for code in codes_spec if type(code) == str}

            # Присваиваем полученный словарь со специальностями словарю high_dct
            """
            В итоге получается такая струтура {'КТИНЗ': {'29.01.07': {}, '35.01.23': {}, '43.01.09': {}, '54.01.06': {}, '54.01.03': {}, '35.02.07': {}, '54.02.02': {}}, 'ТСГХ': {}, 'БФКК': {}}
            """
            high_dct[sheet] = spec_code_dct
            # Итерируемся по полученному датафрейму через itertuples
            for row in df.itertuples():
                # Получаем код специальности для итерируемой строки. Отбрасываем пустые строки, просто проверяя есть ли такой ключ в главном словаре
                spec_code = row[4]
                if spec_code in high_dct[sheet]:

                    name_param = row[6].strip()
                    high_dct[sheet][spec_code][name_param]['Выпуск в 2021'] = check_data(row[7])
                    high_dct[sheet][spec_code][name_param]['Трудоустроено человек'] = check_data(row[8])
                    high_dct[sheet][spec_code][name_param]['Индивидуальные предприниматели'] = check_data(row[10])
                    high_dct[sheet][spec_code][name_param]['Самозанятые'] = check_data(row[12])
                    high_dct[sheet][spec_code][name_param]['Призваны в Вооруженные силы'] = check_data(row[14])
                    high_dct[sheet][spec_code][name_param]['Продолжили обучение'] = check_data(row[16])
                    high_dct[sheet][spec_code][name_param]['Находятся в отпуске по уходу за ребенком'] = check_data(
                        row[18])
                    high_dct[sheet][spec_code][name_param]['Находящиеся под риском нетрудоустройства '] = check_data(
                        row[20])
                    high_dct[sheet][spec_code][name_param][
                        'в том числе (из гр. 22): состоят на учете в центрах занятости в качестве ищущих работу или безработных'] = check_data(
                        row[22])
                    high_dct[sheet][spec_code][name_param][
                        'Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.***'] = check_data(
                        row[23])
                    high_dct[sheet][spec_code][name_param][
                        'перечислить причины, указав число человек'] = check_data_text(row[25])
                    high_dct[sheet][spec_code][name_param][
                        'Не определились (ожидают результатов приемной кампании, ожидают призыва, находятся в активном поиске работы, собирают документы для открытия ИП. Выпускники временно не заняты, но их занятости ничего не угрожает)'] = check_data(
                        row[26])
                    high_dct[sheet][spec_code][name_param]['Прогноз Трудоустройство'] = check_data(row[28])
                    high_dct[sheet][spec_code][name_param]['Прогноз Индивидуальные предприниматели'] = check_data(
                        row[30])
                    high_dct[sheet][spec_code][name_param]['Прогноз Самозанятые'] = check_data(row[32])
                    high_dct[sheet][spec_code][name_param]['Прогноз Продолжили обучение'] = check_data(row[34])
                    high_dct[sheet][spec_code][name_param]['Прогноз Призваны в Вооруженные силы'] = check_data(row[36])
                    high_dct[sheet][spec_code][name_param][
                        'Прогноз Находятся в отпуске по уходу за ребенком'] = check_data(
                        row[38])
                    high_dct[sheet][spec_code][name_param][
                        'Прогноз Находящиеся под риском нетрудоустройства выпускники'] = check_data(row[40])
                    high_dct[sheet][spec_code][name_param][
                        'Прогноз в том числе (из гр. 42): состоят на учете в центрах занятости в качестве ищущих работу или безработных'] = check_data(
                        row[42])
                    high_dct[sheet][spec_code][name_param][
                        'Прогноз Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.*** '] = check_data(
                        row[43])
                    high_dct[sheet][spec_code][name_param]['Прогноз перечислить причины'] = check_data_text(row[45])
                    high_dct[sheet][spec_code][name_param][
                        'Причины, по которым выпускники находятся под риском нетрудоустройства, и принимаемые меры (тезисно)'] = check_data_text(
                        row[46])
                    count_rows += 1

                else:
                    with open(f'{path_to_end_folder}/ERRORS {current_time}.txt', 'a', encoding='utf-8') as f:
                        f.write(f'Строка {count_rows} в листе  {sheet} не обработана!!!\n')
                    count_rows += 1
                    continue


        # Удаляем дубликаты
        unique_code_spec = list(set(all_code_spec))

        # Подсчитываем количество по каждой специальности
        itog_dct = calculation_data(high_dct, unique_code_spec)
        # Обрабатываем и сохраняем результат подсчета
        prepare_final_table(itog_dct)
    # Проверяем на корректность(совпадают ли суммы)

    except NameError:
        messagebox.showerror('Трудоустройство СПО 2021', 'Выберите файл с параметрами,обрабатываемые данные, конечную папку')

    check_correct_data(base_correct_df)
    messagebox.showinfo('Трудоустройство СПО 2021',f'Обработка файлов  завершена! ')


def check_data(cell):
    """
    Метод для проверки значения ячейки
    :param cell: значение ячейки
    :return: число в формате int
    """
    # Проверям на строку
    if type(cell) == str:
        return 0
    # Проверяем на пустую ячейку
    if math.isnan(cell):
        return 0
    if type(cell) == int:
        return cell
    if type(cell) == float:
        return int(cell)
    else:
        return 0


def create_low_dct():
    """
    Функция для создания словаря низкого уровня вида 8 ключей:{23 ключа:0}
    :return: Словарь нужной структуры
    """
    # Создаем базовые структы данных
    # Список категорий по которым будет идти подсчет
    base_cat = ['Всего', 'Лица с ограниченными возможностями здоровья',
                'из них (из строки 02): инвалиды и дети-инвалиды',
                'Инвалиды и дети-инвалиды (кроме учтенных в строке 03)',
                'Имеют договор о целевом обучении',
                'из них (из строки 05): Лица с ограниченными возможностями здоровья (имеющие договор о целевом обучении)',
                'из строки 06: инвалиды и дети-инвалиды (имеющие договор о целевом обучении)',
                'из строки 05 инвалиды и дети-инвалиды (кроме учтенных в строке 07) (имеющие договор о целевом обучении)']

    # Вот это словарь жутковато выглядит.Словарь показателей
    base_dct = {'Выпуск в 2021': 0, 'Трудоустроено человек': 0, 'Индивидуальные предприниматели': 0, 'Самозанятые': 0,
                'Призваны в Вооруженные силы': 0, 'Продолжили обучение': 0,
                'Находятся в отпуске по уходу за ребенком': 0,
                'Находящиеся под риском нетрудоустройства ': 0,
                'в том числе (из гр. 22): состоят на учете в центрах занятости в качестве ищущих работу или безработных': 0,
                'Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.***': 0,
                'перечислить причины, указав число человек': '',
                'Не определились (ожидают результатов приемной кампании, ожидают призыва, находятся в активном поиске работы, собирают документы для открытия ИП. Выпускники временно не заняты, но их занятости ничего не угрожает)': 0,
                'Прогноз Трудоустройство': 0, 'Прогноз Индивидуальные предприниматели': 0, 'Прогноз Самозанятые': 0,
                'Прогноз Продолжили обучение': 0, 'Прогноз Призваны в Вооруженные силы': 0,
                'Прогноз Находятся в отпуске по уходу за ребенком': 0,
                'Прогноз Находящиеся под риском нетрудоустройства выпускники': 0,
                'Прогноз в том числе (из гр. 42): состоят на учете в центрах занятости в качестве ищущих работу или безработных': 0,
                'Прогноз Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.*** ': 0,
                'Прогноз перечислить причины': '',
                'Причины, по которым выпускники находятся под риском нетрудоустройства, и принимаемые меры (тезисно)': ''}

    # Создаем словарь нижнего уровня
    # ССЫЛОЧНАЯ МОДЕЛЬ!!! Я ССЫЛАЛСЯ НА ОДИН И ТОТ ЖЕ СЛОВАРЬ УУУУ
    temp_dct = {key: base_dct.copy() for key in base_cat}
    return temp_dct


def calculation_data(dct, unique_code_lst):
    """
    Функция для подсчета получившихся значений для каждой специальности по всем ПОО.
    :param dct: словарь в котором находятся все данные по трудоустройству
            uniqie_code_lst: список всех уникальных специальностей
    :return: новый словарь где главным ключом будет являться код специальности.
    """
    # Создаем словарь где ключами верхнего уровня будут выступать коды специальности
    calcul_dct = {code: create_low_dct() for code in unique_code_lst}

    # Итерируемся по верхнему уровню словаря
    for poo, spec in dct.items():
        # Итерируемся по вложенному словарю вида {код специальности:{данные}}
        for cod_spec, data in spec.items():

            # Итерируемся по базовому словарю
            for cat, papam_dct in data.items():

                # Итерируемся по словарю параметров. Ха четверная итерация
                for param, value in papam_dct.items():
                    # Заполняем итоговый словарь
                    calcul_dct[cod_spec][cat][param] += value

    return calcul_dct


def check_data_text(cell):
    """
    Функция для обработки ячеек в которых могут встретиться как текстовые так и числовые данные. Колонки 25 и 45 как пример
    :param cell: значение ячейки
    :return:
    """

    if type(cell) == str:
        if cell !='nan':
            return f'{cell};'
        else:
            return ' ;'
    if type(cell) == float:
        if cell == 0.0:
            return f' ;'
    if type(cell) == int:
        return f'{str(cell)};'
    else:
        return ' ;'


def prepare_final_table(dct):
    """
    Функция для обработки словаря с посчитанными данными по студентам в удобочитаемую таблицу
    :param dct: словарь с суммированными данными по каждой специальности
    :return: таблица Excel
    """
    # Превращаем словарь в датафрейм
    df = pd.DataFrame.from_dict(dct, orient='index')
    stack_df = df.stack()
    # название такое выбрал потому что было лень заменять значения из блокнота юпитера
    frame = stack_df.to_frame()
    # Извлекаем данные из словаря в колонке 0
    frame['Выпуск в 2021'] = frame[0].apply(lambda x: x.get('Выпуск в 2021'))
    frame['Трудоустроено человек'] = frame[0].apply(lambda x: x.get('Трудоустроено человек'))
    frame['Индивидуальные предприниматели'] = frame[0].apply(lambda x: x.get('Индивидуальные предприниматели'))
    frame['Самозанятые'] = frame[0].apply(lambda x: x.get('Самозанятые'))
    frame['Призваны в Вооруженные силы'] = frame[0].apply(lambda x: x.get('Продолжили обучение'))
    frame['Находятся в отпуске по уходу за ребенком'] = frame[0].apply(
        lambda x: x.get('Находятся в отпуске по уходу за ребенком'))
    frame['Находящиеся под риском нетрудоустройства '] = frame[0].apply(
        lambda x: x.get('Находящиеся под риском нетрудоустройства '))
    frame['в том числе (из гр. 22): состоят на учете в центрах занятости в качестве ищущих работу или безработных'] = \
    frame[0].apply(lambda x: x.get(
        'в том числе (из гр. 22): состоят на учете в центрах занятости в качестве ищущих работу или безработных'))
    frame[
        'Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.***'] = \
    frame[0].apply(lambda x: x.get(
        'Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.***'))
    frame['перечислить причины, указав число человек'] = frame[0].apply(
        lambda x: x.get('перечислить причины, указав число человек'))
    frame[
        'Не определились (ожидают результатов приемной кампании, ожидают призыва, находятся в активном поиске работы, собирают документы для открытия ИП. Выпускники временно не заняты, но их занятости ничего не угрожает)'] = \
    frame[0].apply(lambda x: x.get(
        'Не определились (ожидают результатов приемной кампании, ожидают призыва, находятся в активном поиске работы, собирают документы для открытия ИП. Выпускники временно не заняты, но их занятости ничего не угрожает)'))
    frame['Прогноз Трудоустройство'] = frame[0].apply(lambda x: x.get('Прогноз Трудоустройство'))
    frame['Прогноз Индивидуальные предприниматели'] = frame[0].apply(
        lambda x: x.get('Прогноз Индивидуальные предприниматели'))
    frame['Прогноз Самозанятые'] = frame[0].apply(lambda x: x.get('Прогноз Самозанятые'))
    frame['Прогноз Продолжили обучение'] = frame[0].apply(lambda x: x.get('Прогноз Продолжили обучение'))
    frame['Прогноз Призваны в Вооруженные силы'] = frame[0].apply(
        lambda x: x.get('Прогноз Призваны в Вооруженные силы'))
    frame['Прогноз Находятся в отпуске по уходу за ребенком'] = frame[0].apply(
        lambda x: x.get('Прогноз Находятся в отпуске по уходу за ребенком'))
    frame['Прогноз Находящиеся под риском нетрудоустройства выпускники'] = frame[0].apply(
        lambda x: x.get('Прогноз Находящиеся под риском нетрудоустройства выпускники'))
    frame[
        'Прогноз в том числе (из гр. 42): состоят на учете в центрах занятости в качестве ищущих работу или безработных'] = \
    frame[0].apply(lambda x: x.get(
        'Прогноз в том числе (из гр. 42): состоят на учете в центрах занятости в качестве ищущих работу или безработных'))
    frame[
        'Прогноз Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.*** '] = \
    frame[0].apply(lambda x: x.get(
        'Прогноз Прочее: смерть, переезд за пределы Российской Федерации, семейные обстоятельства, по состоянию здоровья и др.*** '))
    frame['Прогноз перечислить причины'] = frame[0].apply(lambda x: x.get('Прогноз перечислить причины'))
    frame['Причины, по которым выпускники находятся под риском нетрудоустройства, и принимаемые меры (тезисно)'] = \
    frame[0].apply(lambda x: x.get(
        'Причины, по которым выпускники находятся под риском нетрудоустройства, и принимаемые меры (тезисно)'))

    # Удаляем колонку со словарем
    frame = frame.drop([0], axis=1)
    #Сохраняем в файл Excel
    frame.to_excel(f'{path_to_end_folder}/Итоговый результат.xlsx')

def count_checked_cells(lst):
    """
    Функция для подсчета значений из ячеек
    :param lst: список данных из проверяемых ячеек
    :return: результат в формате int
    """
    quantity = 0
    for value in lst:
        if type(value) == str:
            quantity +=0
        elif math.isnan(value):
            quantity += 0
        else:
            quantity += int(value)
    return quantity

def check_correct_data(df):
    """
    Функция для проверки совпадения сумм фактического и планируемого распределения
    :param df:датафрейм в котором находятся все данные по трудоустройству
    :return:
    """
    # Итерируемся по датафрейму
    # Добавляем 2 колонки куда будем записывать статус проверки
    df.insert(5,'Проверка фактического распределения','')
    df.insert(6,'Проверка планового распределения','')
    for row in df.itertuples():
        # Проверяем на заполненность
        # print(row[8],row[10],row[12],row[14],row[16],row[18],row[20],row[23],row[26])
        # print(type(row[8]),type(row[10]),type(row[12]),type(row[14]),type(row[16]),type(row[18]),type(row[20]),type(row[23]),type(row[26]))
        if type(row[4]) == float:
            df.loc[row[0],'Проверка фактического распределения'] = 'НЕ ЗАПОЛНЕНО'
            df.loc[row[0],'Проверка планового распределения'] = 'НЕ ЗАПОЛНЕНО'
        else:
            # не очень выглядит да?

            fact_lst = [row[10],row[12],row[14],row[16],row[18],row[20],row[22],row[25],row[28]]
            plan_lst = [row[30],row[32],row[34],row[36],row[38],row[40],row[42],row[45]]
            check_summa_fact = count_checked_cells(fact_lst)
            check_summa_plan = count_checked_cells(plan_lst)

            all_summa = 0 if math.isnan(row[9]) else int(row[9])
            if all_summa == check_summa_fact:
                df.loc[row[0], 'Проверка фактического распределения'] = 'принято'
            else:
                df.loc[row[0], 'Проверка фактического распределения'] = 'НЕПРАВИЛЬНАЯ СУММА'
            if all_summa == check_summa_plan:
                df.loc[row[0], 'Проверка планового распределения'] = 'принято'
            else:
                df.loc[row[0], 'Проверка планового распределения'] = 'НЕПРАВИЛЬНАЯ СУММА'
    # Сохраняем проверочный датафрейм
    df.to_excel(f'{path_to_end_folder}/Результат проверки.xlsx',index=False)

if __name__ == '__main__':
    window = Tk()
    window.title('СПО Трудоустройство выпускников 2021 Ver 1.01')
    window.geometry('600x800')
    window.resizable(False, False)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных СПО
    tab_calculate_data = ttk.Frame(tab_control)
    tab_control.add(tab_calculate_data, text='Обработка данных')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_calculate_data,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия\nОбработка данных по трудоустройству студентов СПО за 2021 год',
                      font=25)
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')
    img = PhotoImage(file=path_to_img)
    Label(tab_calculate_data,
          image=img
          ).grid(column=0, row=1, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_select_files_data = Button(tab_calculate_data, text='1) Выбрать файлы с данными', font=('Arial Bold', 20),
                                   command=select_files_data
                                   )
    btn_select_files_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_calculate_data, text='2) Выбрать конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)


    # Создаем кнопку для запуска обработки файлов

    btn_calculate = Button(tab_calculate_data, text='3) Обработать', font=('Arial Bold', 20),
                           command=calculate_data
                           )
    btn_calculate.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()
