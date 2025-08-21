"""
Скрипт для обработки мониторинга занятости выпускников на сентябрь 2025
"""

# -*- coding: utf-8 -*-
"""
Скрипт для обработки данных Формы 2 нозология (15 строк) мониторинга занятости выпускников
"""
import numpy as np
# проверка основного листа
import pandas as pd
import copy
import os
import warnings
from tkinter import messagebox
import time
pd.options.mode.chained_assignment = None  # default='warn'
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', category=FutureWarning)
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


class NotCorrectFile(Exception):
    """
    Исключения для обработки случая когда нет ни одного корректного файла
    """
    pass




def prepare_september_2025(path_folder_data:str,path_to_end_folder):
    """
    Фунция для обработки данных мониторинга сентябрь 2025
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}

    # создаем базовый датафрейм для нозологии
    main_nose_df = pd.DataFrame(
        columns=['Наименование файла', '1', '2', '3', '4', '5', '6', '7'])


    # создаем базовый датафрейм для целевиков
    main_target_df = pd.DataFrame(
        columns=['Наименование файла', '1', '2', '3', '4', '5', '6', '7', '8', '9', '10'])

    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    dct_code_and_name = dict()

    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    target_dct_code_and_name = dict()

    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    # Создаем датафрейм для контроля дублирующихся кодов специальностей
    main_dupl_df = pd.DataFrame(columns=['Название файла', 'Полное наименование'])

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

            if '1. Форма сбора' not in lst_temp_sheets:  # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием 1. Форма сбора !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            if '2. Нозологии' not in lst_temp_sheets:  # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием 2. Нозологии !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            if '3. Целевики' not in lst_temp_sheets:  # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием 3. Целевики !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            # Считываем данные листа с общими данными
            try:
                df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='1. Форма сбора', skiprows=3, dtype=str)
            except:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не удалось прочитать файл! Отключите фильтры в файле, проверьте файл на целостность и соответствие шаблону']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            df.columns = list(map(str, df.columns))  # делаем названия колонок строковыми

            # создаем множество колонок наличие которых мы проверяем
            check_cols = {'1', '2', '3', '3.1', '3.2', '3.3',
                          '4', '4.1', '4.2', '4.3', '5', '6', '7', '8', '9', '10', '11', '12',
                          '13', '14', '15', '16', '17'}
            diff_cols = check_cols.difference(list(df.columns))
            if len(diff_cols) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols}',
                                                    'На листе 1. Форма сбора не найдены указанные колонки. Проверьте соответствие файла шаблону с сайта, строка с номерами колонок должна быть на четвертой строке ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            # Считываем данные листа Нозологии
            try:
                nose_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='2. Нозологии', skiprows=1, dtype=str)
            except:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не удалось прочитать файл! Отключите фильтры в файле, проверьте файл на целостность и соответствие шаблону']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            nose_df.columns = list(map(str, nose_df.columns))  # делаем названия колонок строковыми
            # создаем множество колонок наличие которых мы проверяем
            check_cols_nose = {'1', '2', '3', '4', '5', '6', '7'}
            diff_cols_nose = check_cols_nose.difference(list(nose_df.columns))
            if len(diff_cols_nose) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols_nose}',
                                                    'На листе 2. Нозологии не найдены указанные колонки. Проверьте соответствие файла шаблону с сайта, строка с номерами колонок должна быть на второй строке ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            # Считываем данные листа с целевиками
            try:
                target_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='3. Целевики', skiprows=1, dtype=str)
            except:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не удалось прочитать файл! Отключите фильтры в файле, проверьте файл на целостность и соответствие шаблону']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            target_df.columns = list(map(str, target_df.columns))  # делаем названия колонок строковыми
            # создаем множество колонок наличие которых мы проверяем
            check_cols_target = {'1', '2', '3', '4', '5', '6', '7', '8', '9', '10'}
            diff_cols_target = check_cols_target.difference(list(target_df.columns))
            if len(diff_cols_target) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols_target}',
                                                    'На листе 3. Целевики не найдены указанные колонки. Проверьте соответствие файла шаблону с сайта, строка с номерами колонок должна быть на второй строке ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue










    error_df.to_excel(f'{path_to_end_folder}/dfg.xlsx')








if __name__ == '__main__':
    main_path_folder_data = 'data/2025'
    main_path_end_folder = 'data/РЕЗУЛЬТАТ'
    prepare_september_2025(main_path_folder_data,main_path_end_folder)
    print('Lindy Booth')




