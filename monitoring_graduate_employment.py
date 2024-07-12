# -*- coding: utf-8 -*-
"""
Модуль для обработки таблиц мониторинга занятости выпускников используемого для загрузки на сайт СССР
"""
from check_functions import base_check_file
from support_functions import convert_to_int

import pandas as pd
import numpy as np
import os
import warnings
from tkinter import messagebox
import time

pd.options.mode.chained_assignment = None  # default='warn'
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', category=FutureWarning)
import copy
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def prepare_graduate_employment(path_folder_data: str, path_result_folder: str):
    """
    Функция для обработки мониторинга занятости
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    dct_code_and_name = dict()
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    requred_columns_first_sheet = ['1', '1.1', '1.2', '2', '3', '3.1', '3.2', '4', '5', '6', '7', '8', '9', '10',
                                   '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                                   '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
                                   '31', '32', '33', '34', '35', '36', '37', '38', '39', '40',
                                   '41', '42', '43', '44', '45', '46', '47', '48', '49', '50',
                                   '51', '52', '53', '54', '55', '56', '57', '58', '59', '60',
                                   '61', '62', '63', '64', '65', '66', '67', '68', '69', '70',
                                   '71', '72', '73']

    requred_columns_second_sheet = ['1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11', '12', '13']

    # Создаем словарь для базовой проверки файла (расширение, наличие листов, наличие колонок)
    """
    {Название листа:{Количество строк заголовка:int,'Обязательные колонки':список колонок,'Текст ошибки':'Описание ошибки'}}
    """
    check_required_dct = {'Выпуск-СПО': {'Количество строк заголовка': 2,
                                         'Обязательные колонки':requred_columns_first_sheet,
                                         'Не найден лист': 'В файле не найден лист с названием Выпуск-СПО',
                                         'Нет колонок': 'На листе Выпуск-СПО не найдены колонки:'},
                          'Выпуск-Целевое': {'Количество строк заголовка': 3,
                                             'Обязательные колонки': requred_columns_second_sheet,
                                          'Не найден лист': 'В файле не найден лист с названием Выпуск-Целевое',
                                             'Нет колонок': 'На листе Выпуск-Целевое не найдены колонки:'}}

    try:
        for file in os.listdir(path_folder_data):
            if not file.startswith('~$'):
                # Проверяем файл на расширение, наличие нужных листов и колонок
                file_error_df = base_check_file(file, path_folder_data, check_required_dct)
                error_df = pd.concat([error_df, file_error_df], axis=0, ignore_index=True)
                if len(file_error_df) != 0:
                    continue
                print(file)
                # Обрабатываем данные с листа Выпуск-СПО
                df_first_sheet = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='Выпуск-СПО',
                                               skiprows=check_required_dct['Выпуск-СПО']['Количество строк заголовка']) # Считываем прошедший базовую проверку файл
                df_first_sheet.columns = list(map(str, df_first_sheet.columns))  # делаем названия колонок строковыми

                # Приводим все колонки кроме первой к инту
                df_first_sheet[requred_columns_first_sheet[1:]] = df_first_sheet[requred_columns_first_sheet[1:]].applymap(convert_to_int)
                # очищаем первую колонку от пробельных символов вначаче и конце
                df_first_sheet['1'] = df_first_sheet['1'].apply(lambda x:x.strip() if isinstance(x,str) else x)
                print(df_first_sheet['1'])





        print(error_df)
        t = time.localtime()  # получаем текущее время
        current_time = time.strftime('%H_%M_%S', t)
        error_df.to_excel(f'{path_result_folder}/Ошибки {current_time}.xlsx',index=False)
    except ZeroDivisionError:
        print('dssd')


if __name__ == '__main__':
    main_data_folder = 'c:/Users/1/PycharmProjects/Dodger_2023/data/Мониторинг занятости выпускников/Файлы/'
    main_result_folder = 'c:/Users/1/PycharmProjects/Dodger_2023/data/Мониторинг занятости выпускников/Результат'
    prepare_graduate_employment(main_data_folder, main_result_folder)

    print('Lindy Booth')
