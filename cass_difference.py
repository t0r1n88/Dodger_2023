"""
Функция для подсчета разницы между страницами
"""
from cass_support_functions import *

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

def abs_diff(first_value, second_value):
    """
    Функция для подсчета абсолютной разницы между 2 значениями
    """
    try:
        return abs(float(first_value) - float(second_value))
    except:
        return None


def percent_diff(first_value, second_value):
    """
    функция для подсчета относительной разницы значений
    """
    try:
        # округляем до трех
        value = round(float(second_value) / float(first_value), 4) * 100
        return value
    except:
        return None


def change_perc_diff(first_value, second_value):
    """
    функция для подсчета процентного ихменения значений
    """
    try:
        value = (float(second_value) - float(first_value)) / float(first_value)
        return round(value, 4) * 100
    except:
        return None



def prepare_diffrence(data_first_diffrence,dif_first_sheet_name,data_second_diffrence,dif_second_sheet_name,path_to_end_folder_diffrence):
    """
    Функция для вычисления разницы между двумя таблицами
    """
    # загружаем датафреймы
    try:
        df1 = pd.read_excel(data_first_diffrence, sheet_name=dif_first_sheet_name, dtype=str)
        df2 = pd.read_excel(data_second_diffrence, sheet_name=dif_second_sheet_name, dtype=str)
        # проверяем на соответсвие размеров
        if df1.shape != df2.shape:
            raise ShapeDiffierence

        # Проверям на соответсвие колонок
        if list(df1.columns) != list(df2.columns):
            diff_columns = set(df1.columns).difference(set(df2.columns))  # получаем отличающиеся элементы
            raise ColumnsDifference

        df_cols = df1.compare(df2,
                              result_names=('Первая таблица', 'Вторая таблица'))  # датафрейм с разницей по колонкам
        df_cols.index = list(
            map(lambda x: x + 2, df_cols.index))  # добавляем к индексу +2 чтобы соответствовать нумерации в экселе
        df_cols.index.name = '№ строки'  # переименовываем индекс

        df_rows = df1.compare(df2, align_axis=0,
                              result_names=('Первая таблица', 'Вторая таблица'))  # датафрейм с разницей по строкам
        lst_mul_ind = list(map(lambda x: (x[0] + 2, x[1]),
                               df_rows.index))  # добавляем к индексу +2 чтобы соответствовать нумерации в экселе
        index = pd.MultiIndex.from_tuples(lst_mul_ind, names=['№ строки', 'Таблица'])  # создаем мультиндекс
        df_rows.index = index

        # Создаем датафрейм с подсчетом разниц
        df_diff_cols = df_cols.copy()

        # получаем список колонок первого уровня
        temp_first_level_column = list(map(lambda x: x[0], df_diff_cols.columns))
        first_level_column = []
        [first_level_column.append(value) for value in temp_first_level_column if value not in first_level_column]

        # Добавляем колонки с абсолютной и относительной разницей
        count_columns = 2
        for name_column in first_level_column:
            # высчитываем абсолютную разницу
            df_diff_cols.insert(count_columns, (name_column, 'Разница между первым и вторым значением'),
                                df_diff_cols.apply(lambda x: abs_diff(x[name_column]['Первая таблица'],
                                                                      x[name_column]['Вторая таблица']), axis=1))

            # высчитываем отношение второго значения от первого
            df_diff_cols.insert(count_columns + 1, (name_column, '% второго от первого значения'),
                                df_diff_cols.apply(lambda x: percent_diff(x[name_column]['Первая таблица'],
                                                                          x[name_column]['Вторая таблица']), axis=1))

            # высчитываем процентное изменение
            df_diff_cols.insert(count_columns + 2, (name_column, 'Изменение в процентах'),
                                df_diff_cols.apply(lambda x: change_perc_diff(x[name_column]['Первая таблица'],
                                                                              x[name_column]['Вторая таблица']),
                                                   axis=1))

            count_columns += 5

        # записываем
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # делаем так чтобы записать на разные листы
        with pd.ExcelWriter(f'{path_to_end_folder_diffrence}/Разница между 2 таблицами {current_time}.xlsx') as writer:
            df_cols.to_excel(writer, sheet_name='По колонкам')
            df_rows.to_excel(writer, sheet_name='По строкам')
            df_diff_cols.to_excel(writer, sheet_name='Значение разницы')
    except ShapeDiffierence:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Не совпадают размеры таблиц, В первой таблице {df1.shape[0]}-стр. и {df1.shape[1]}-кол.\n'
                             f'Во второй таблице {df2.shape[0]}-стр. и {df2.shape[1]}-кол.')

    except ColumnsDifference:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Названия колонок в сравниваемых таблицах отличаются\n'
                             f'Колонок:{diff_columns}  нет во второй таблице !!!\n'
                             f'Сделайте названия колонок одинаковыми.')

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except ValueError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'В файлах нет листа с таким названием!\n'
                             f'Проверьте написание названия листа')
    except FileNotFoundError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    # except:
    #     logging.exception('AN ERROR HAS OCCURRED')
    #     messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.34',
    #                          'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('Кассандра Подсчет данных по трудоустройству выпускников',
                            'Таблицы успешно обработаны')

if __name__ == '__main__':
    first_file = 'data/example/Разница между 2 таблицами/Форма №15 от 01.06.2023.xlsx'
    second_file = 'data/example/Разница между 2 таблицами/Форма №15 от 01.09.2023.xlsx'
    first_sheet_name = ''


    path_data = 'data/example/Форма ОПК по отраслям( 2 строки)'
    path_end = 'data/result/Форма ОПК по отраслям( 2 строки)'
    prepare_diffrence()
    print('Lindy Booth')