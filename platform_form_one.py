# -*- coding: utf-8 -*-
"""
Скрипт для подсчета данных для заполненения мониторинга трудоустройства выпускников для цифровой платформы ИРПО
Форма с разделением выпускников по отраслям
"""
from cass_check_functions import * # импортируем функции проверки
from support_functions import * # импортируем вспомогательные функции и исключения
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


def prepare_platform_form_one_employment(path_folder_data:str,result_folder:str):
    """
    Функция для обработки формы 1 по отраслям
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    dct_code_and_name = dict()
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    for file in os.listdir(path_folder_data):
        if not file.startswith('~$') and not file.endswith('.xlsx'):
            name_file = file.split('.xls')[0]
            temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                'Расширение файла НЕ XLSX! Программа обрабатывает только XLSX ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                         columns=['Название файла', 'Строка или колонка с ошибкой',
                                                  'Описание ошибки'])
            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
            continue
        if not file.startswith('~$') and file.endswith('.xlsx'):
            name_file = file.split('.xlsx')[0]
            print(name_file)
            # получаем название первого листа
            temp_wb = openpyxl.load_workbook(f'{path_folder_data}/{file}', read_only=True)
            lst_temp_sheets = temp_wb.sheetnames  # получаем листы в файле
            temp_wb.close()
            if 'Выпуск-СПО (все)' not in lst_temp_sheets:  # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием Выпуск-СПО (все) !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='Выпуск-СПО (все)', skiprows=4, dtype=str)
            df.columns = list(map(str, df.columns))  # делаем названия колонок строковыми
            check_cols = ['1', '1.1', '1.2', '2', '3', '4', '5', '6', '7', '8', '9', '10',
                          '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                          '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
                          '31', '32', '33', '34', '35', '36', '37', '38', '39', '40',
                          '41', '42', '43', '44', '45', '46', '47', '48', '49', '50',
                          '51', '52', '53', '54', '55', '56', '57', '58', '59', '60',
                          '61', '62', '63', '64', '65', '66', '67', '68', '69', '70',
                          '71', '72', '73', '74', '75', '76', '77']
            diff_cols = set(check_cols).difference(set(df.columns))
            if len(diff_cols) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols}',
                                                    'Структура формы сбора данных отличается от эталонной.Строка с номерами колонок (1,1.1,1.2,2,3,4 ... 76,77 как в исходной форме)\n должна находиться на 5 строке!\n'
                                                    ' не хватает указанных колонок.Колонки с названимем Unnamed означаеют что на листе есть данные без заголовка в виде цифр на 5 строке  ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue





    print(error_df)
    error_df.to_excel(f'{result_folder}/Ошибки.xlsx',index=False)

if __name__ == '__main__':
    main_data_folder = 'data/example/Форма 1 Выпуск-СПО все'
    main_result_folder = 'data/result/Платформа Форма 1'
    prepare_platform_form_one_employment(main_data_folder,main_result_folder)

    print('Lindy Booth')