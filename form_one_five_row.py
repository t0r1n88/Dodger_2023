# -*- coding: utf-8 -*-
"""
Скрипт для обработки данных Формы 1 пятистрочной мониторинга занятости выпускников
"""
from check_functions import * # импортируем функции проверки
from support_functions import * # импортируем вспомогательные функции и исключения
import pandas as pd
import numpy as np
import tkinter
import sys
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time
pd.options.mode.chained_assignment = None  # default='warn'
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', category=FutureWarning)
import copy
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import re
import random


def prepare_form_one_employment(path_folder_data:str,path_to_end_folder):
    """
    Фугкция для обработки данных формы 1 пять строк
    :return:
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    tup_correct = (6, 10)  # создаем кортеж  с поправками где 6 это первая строка с данными а 10 строка где заканчивается первый диапазон

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
            if 'Форма 1 пятистрочная' not in lst_temp_sheets: # проверяем наличие листа с названием в файле
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Не найден лист с названием Форма 1 пятистрочная !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            # temp_sheet_name_df = pd.DataFrame(data=[[f'{name_file}', f'{name_source_sheet}']],
            #                                   columns=['Название файла', 'Название листа откуда взяты данные'])
            # sheet_name_df = pd.concat([sheet_name_df, temp_sheet_name_df], axis=0, ignore_index=True)
            df = pd.read_excel(f'{path_folder_data}/{file}', skiprows=4, dtype=str)
            # создаем множество колонок наличие которых мы проверяем
            check_cols = ['01','02', '03','04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15',
                          '16', '17',
                          '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '30']
            if check_cols != list(df.columns):
                diff_cols = set(list(df.columns)).difference(set(check_cols))
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', f'{diff_cols}',
                                                    'Возможно старая версия формы сбора данных.Строка с номерами колонок (01,02,03,05 ... 28,30 как в исходной форме)\n должна находиться на 5 строке!\n'
                                                    ' указанные колонки являются лишними.Колонки с названимем Unnamed означаеют что на листе есть данные без заголовка в виде цифр на 5 строке  ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue











    print(error_df)
    error_df.to_excel('data/result/errro.xlsx',index=False,header=True)


if __name__ == '__main__':
    main_data_folder = 'data/example/testing'
    main_result_folder = 'data/result'
    prepare_form_one_employment(main_data_folder,main_result_folder)

    print('Lindy Booth')