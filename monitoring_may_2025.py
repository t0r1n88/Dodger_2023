# -*- coding: utf-8 -*-
"""
Скрипт для обработки данных Формы 2 нозология (15 строк) мониторинга занятости выпускников
"""
from cass_support_functions import * # импортируем вспомогательные функции и исключения
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

class NotCorrectFile(Exception):
    """
    Исключения для обработки случая когда нет ни одного корректного файла
    """
    pass


def prepare_may_2025(path_folder_data:str,path_to_end_folder):
    """
    Фугкция для обработки данных мониторинга 2025
    :return:
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    dct_code_and_name = dict()
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    # Создаем датафрейм для контроля дублирующихся кодов специальностей
    main_dupl_df = pd.DataFrame(columns=['Название файла','Полное наименование'])

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

            if '3. Целевики' not in lst_temp_sheets:  # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием 3. Целевики !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            # Считываем данные листа с общими данными
            df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='1. Форма сбора', skiprows=2, dtype=str)
            df.columns = list(map(str, df.columns))  # делаем названия колонок строковыми
            # создаем множество колонок наличие которых мы проверяем
            check_cols = {'1', '1.1', '1.2', '2', '3', '3.1', '3.2', '3.3',
                           '4', '4.1', '4.2', '4.3', '5', '6', '7', '8', '9', '10', '11', '12',
                           '13', '14', '15', '16', '17', '18'}
            diff_cols = check_cols.difference(list(df.columns))
            if len(diff_cols) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols}',
                                                    'На листе 1. Форма сбора не найдены указанные колонки. Проверьте соответствие файла шаблону с сайта, строка с номерами колонок должна быть на третьй строке ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            # Считываем данные листа с целевиками
            target_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name='3. Целевики', skiprows=2, dtype=str)
            target_df.columns = list(map(str, target_df.columns))  # делаем названия колонок строковыми
            # создаем множество колонок наличие которых мы проверяем
            check_cols_target = {'1', '2', '3', '4', '5', '6', '7', '8', '9', '10', '11',
                                   '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23',
                                   '24', '25', '26', '27'}
            diff_cols_target = check_cols_target.difference(list(target_df.columns))
            if len(diff_cols_target) != 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols_target}',
                                                    'На листе 3. Целевики не найдены указанные колонки. Проверьте соответствие файла шаблону с сайта, строка с номерами колонок должна быть на третьй строке ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            df = df[df['1'].notna()]  # убираем возможные наны из за лишних строк
            target_df = target_df[target_df['1'].notna()]  # убираем возможные наны из за лишних строк

            # Проверяем на заполнение лист с общими данными
            if len(df) == 0:
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'Отсутствуют коды специальностей в колонке 1',
                                                    'Лист 1. Форма сбора пустой. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue






    t = time.localtime()  # получаем текущее время
    current_time = time.strftime('%H_%M_%S', t)

        # Сохраняем ошибки
    wb = openpyxl.Workbook()
    for r in dataframe_to_rows(error_df, index=False, header=True):
        wb['Sheet'].append(r)

    wb['Sheet'].column_dimensions['A'].width = 30
    wb['Sheet'].column_dimensions['B'].width = 40
    wb['Sheet'].column_dimensions['C'].width = 50

    wb.save(f'{path_to_end_folder}/ОШИБКИ Май 2025 от {current_time}.xlsx')






if __name__ == '__main__':
    main_path_folder_data = 'data/Май Мониторинг'
    main_path_end_folder = 'data/РЕЗУЛЬТАТ'
    prepare_may_2025(main_path_folder_data,main_path_end_folder)
    print('Lindy Booth')



