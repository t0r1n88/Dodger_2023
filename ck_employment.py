# -*- coding: utf-8 -*-
"""
Скрипт для подсчета данных центров карьеры
"""
from cass_check_functions import * # импортируем функции проверки
from cass_support_functions import * # импортируем вспомогательные функции и исключения
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


def prepare_ck_employment(path_folder,path_to_end_folder):
    """
    Функция для обработки отчетов центров карьеры
    :return:
    """
    # создаем базовый датафрейм заполненный нулями
    base_df = pd.DataFrame(np.zeros((5, 27)))
    base_df = base_df.applymap(int)  # приводим его к инту
    cols_df = ['05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21',
               '22', '23', '24',
               '25', '26', '27', '28', '29', '30', '31']
    base_df.columns = cols_df

    # Создаем общую таблицы для проверки
    general_table = pd.DataFrame(columns=['Название файла'] + cols_df)

    # создаем датафрейм для регистрации ошибок
    base_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    # Создаем датафрейм для хранения строковых данных из колонки 32
    str_df = pd.DataFrame(index=range(5))

    try:
        for file in os.listdir(path_folder):
            if not file.startswith('~$') and file.endswith('.xls'):
                name_file = file.split('.xls')[0]
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Файл с расширением XLS (СТАРЫЙ ФОРМАТ EXCEL)!!! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                base_error_df = pd.concat([base_error_df, temp_error_df], axis=0, ignore_index=True)
                continue
            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                temp_df_ck = pd.read_excel(f'{path_folder}/{file}', skiprows=5, nrows=5)
                if temp_df_ck.shape[1] != 30:
                    temp_error_df = pd.DataFrame(data=[
                        [f'{name_file}', '',
                         'Количество колонок в таблице не равно 30 !!! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                        columns=['Название файла', 'Строка или колонка с ошибкой',
                                 'Описание ошибки'])
                    base_error_df = pd.concat([base_error_df, temp_error_df], axis=0, ignore_index=True)
                    continue
                # Создаем копию датафрейма только с числовыми колонками
                temp_df_int = temp_df_ck.iloc[:, 2:29].copy()
                # заменяем пустые ячейки нулями

                temp_df_int.fillna(0, inplace=True)
                temp_df_int = temp_df_int.applymap(int)

                # проверяем на ошибки
                temp_error_df = check_error_ck(temp_df_int.copy(), name_file)
                # Добавляем в итоговый датафрейм с ошибками
                base_error_df = pd.concat([base_error_df, temp_error_df], axis=0, ignore_index=True)
                # проверяем размер датафрейма с ошибками, если их нет то добавляем в результат.
                if base_error_df.shape[0] == 0:
                    base_df = base_df + temp_df_int  # складываем значения в таблицах
                    # делаем копию промежутчного датафрейма, так как мы будем добавлять новую колонку
                    temp_add_df = temp_df_int.copy()
                    temp_add_df.insert(0, 'Название файла', name_file)
                    temp_add_df['32'] = temp_df_ck.iloc[:, 29]
                    general_table = pd.concat([general_table, temp_add_df], axis=0,
                                              ignore_index=True)  # сохраняем в общую таблицу
                    # Добаввляем принимаемые меры
                    str_df = pd.concat([str_df, temp_df_ck.iloc[:, 29].to_frame().fillna('_')], axis=1,
                                       ignore_index=True)

                else:
                    continue
        # Объдиняем колонки с принимаемыми мерами в одну и добавляем в base df

        base_df['32'] = str_df.apply(lambda x: ';'.join(x), axis=1)
        # Добавляем колонки
        fourth = ['Всего (общая численность выпускников)',
                  'из общей численности выпускников (из строки 01): лица с ОВЗ',
                  'из числа лиц с ОВЗ (из строки 02): инвалиды и дети-инвалиды',
                  'Инвалиды и дети-инвалиды (кроме учтенных в строке 03)',
                  'Имеют договор о целевом обучении']
        three = ['01', '02', '03', '04', '05']
        base_df.insert(0, '03', three)
        base_df.insert(1, '04', fourth)
        # в общую таблицу
        miultipler = general_table.shape[0] // 5
        general_table.insert(1, '03', three * miultipler)
        general_table.insert(2, '04', fourth * miultipler)

        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        base_df.to_excel(f'{path_to_end_folder}/Отчет ЦК Общий результат от {current_time}.xlsx', index=False)
        base_error_df.to_excel(f'{path_to_end_folder}/Отчет ЦК Ошибки от {current_time}.xlsx', index=False)
        general_table.to_excel(f'{path_to_end_folder}/Отчет ЦК Данные из всех таблиц от {current_time}.xlsx',
                               index=False)
    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Не найдено значение {e.args}')
    except ValueError as e:
        foo_str = e.args[0].split(':')[1]

        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'В таблице с названием {name_file} в колонках: 05 -31 обнаружено НЕ числовое значение! В этих колонках не должно быть текста, пробелов или других символов, кроме чисел. \n'
                             f'Некорректное значение - {foo_str} !!!\n Исправьте и повторно запустите обработку')
    except FileNotFoundError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except PermissionError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Закройте открытые файлы Excel {e.args}')
    except:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'При обработке файла {name_file} возникла ошибка !!!\n'
                             f'Проверьте файл на соответствие шаблону.')

    else:
        if base_error_df.shape[0] != 0:
            messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 'Обнаружены ошибки в обрабатываемых файлах.\n'
                                 'Названия файлов с ошибками и ошибки вы можете найти в файле Отчет ЦК ошибки.\n'
                                 'Исправьте ошибки и запустите повторную обработку чтбы получить полный результат.')
        else:
            messagebox.showinfo('Кассандра Подсчет данных по трудоустройству выпускников',
                                'Данные успешно обработаны')

if __name__ == '__main__':
    path_data = 'data/example/Отчет ЦК'
    path_end = 'data/result/Отчет ЦК'
    prepare_ck_employment(path_data, path_end)
    print('Lindy Booth')