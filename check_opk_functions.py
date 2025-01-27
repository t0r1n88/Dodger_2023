"""
Проверки для ОПК
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

def check_error_opk(df1: pd.DataFrame, name_file, tup_correct):
    """
    Функция для проверки таблиц по ОПК
    :param df1: датафрейм форма 1
    :param df2: датафрейм форма 2
    :param name_file: название обрабатываемого файла
    :param tup_correct: значаение корректировки
    :return: датафрейм с найденными ошибками
    """
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # делаем датафреймы для простых проверок
    df1 = df1.iloc[:, 5:77]
    df1 = df1.applymap(check_data)

    # получаем количество датафреймов
    quantity = df1.shape[0] // 2
    # счетчик для обработанных строк
    border = 0
    for i in range(1, quantity + 1):
        temp_df = df1.iloc[border:border + 2, :]
        # Проверяем корректность заполнения формы 1
        first_error_opk = check_horizont_sum_opk_all(temp_df.copy(), name_file,
                                                     tup_correct)  # проверяем сумму по строкам
        error_df = pd.concat([error_df, first_error_opk], axis=0, ignore_index=True)

        # проверяем условие  по колонкам строка 02 не должна быть больше строки 01
        second_error_opk = check_vertical_opk_all(temp_df.copy(), border, name_file, tup_correct)
        error_df = pd.concat([error_df, second_error_opk], axis=0, ignore_index=True)

        # проверяем условие чтобы сумма по отраслям была равна колонке 08

        # список колонок которые нужно суммировать для проверки условия 07 = 08:31
        lst_07 = ['08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22', '23', '24',
                  '25', '26', '27', '28', '29', '30', '31']
        third_error_opk = check_horizont_chosen_sum_opk(temp_df.copy(), lst_07, '07', name_file)
        error_df = pd.concat([error_df, third_error_opk], axis=0, ignore_index=True)

        # считаем для будут трудоустроены по отраслям
        lst_038 = ['39', '40', '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53', '54', '55',
                   '56', '57', '58', '59', '60', '61', '62']
        fourth_error_opk = check_horizont_chosen_sum_opk(temp_df.copy(), lst_038, '38', name_file)
        error_df = pd.concat([error_df, fourth_error_opk], axis=0, ignore_index=True)

        border += 2

    return error_df


def check_cross_error_opk(df1: pd.DataFrame, df2: pd.DataFrame, name_file, tup_correct):
    """
    Функция для проверки значений между формой 1 и формой 2
    :param df1:
    :param df2:
    :param name_file:
    :param tup_correct:
    :return:
    """
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])

    df1['08'] = df1['08'].apply(check_data)
    df1['39'] = df1['39'].apply(check_data)  # приводим к инту

    group_df1 = df1.groupby(['03']).agg({'08': sum, '39': sum})  # группируем
    group_df1 = group_df1.reset_index()  # переносим индексы
    group_df1.columns = ['Специальность', 'Трудоустроено в ОПК', 'Будут трудоустроены в ОПК']

    # приводим колонку 2 формы с числов выпускников к инту
    df2['04'] = df2['04'].apply(check_data)

    # Проверяем заполенение 2 формы, есть ли там вообще хоть что то
    quantity_now = group_df1['Трудоустроено в ОПК'].sum()  # сколько трудоустроено сейчас
    quantity_future = group_df1['Будут трудоустроены в ОПК'].sum()  # сколько будут трудоустроены
    # проверяем заполнение формы 2
    if (quantity_now != 0 or quantity_future != 0) and df2.shape[0] == 0:
        temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                            'В форме 1 есть выпускники трудоустроенные или которые будут трудоустроены в ОПК,\n'
                                            ' при этом форма 2 не заполнена. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
        return error_df

    # проверка по трудоустроенным и трудоустроенным в будущем
    cross_first_error_df = check_cross_first_error_df(df1.copy(), df2.copy(), name_file)
    error_df = pd.concat([error_df, cross_first_error_df], axis=0, ignore_index=True)

    # # проверка по целевикам
    check_cross_second_error_df = check_cross_second_error(df1.copy(), df2.copy(), name_file)
    error_df = pd.concat([error_df, check_cross_second_error_df], axis=0, ignore_index=True)

    return error_df


def check_cross_first_error_df(df1: pd.DataFrame, df2: pd.DataFrame, name_file):
    """
    Функция для првоерки соответствия количества указанных в форме 1 трудоустроенных и списка в форме 2
    проверки 1 и 2

    :param df1:
    :param df2:
    :param name_file:
    :return:
    """
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # проверяем наличие незаполненных ячеек в колонках 05 06
    # отбираем значения первой строки
    df1 = df1[df1['04'] == '01']
    df1 = df1.groupby(['03']).agg({'08': sum, '39': sum})  # группируем
    df1 = df1.reset_index()  # переносим индексы
    df1.columns = ['Специальность', 'Трудоустроено в ОПК', 'Будут трудоустроены в ОПК']

    etalon_05 = {'уже трудоустроены', 'будут трудоустроены'}  # эталонный состав колонки 05
    etalon_06 = {'заключили договор о целевом обучении', 'нет'}  # эталонный состав колонки 05
    # получаем состав колонок
    st_05 = set(df2['05'].unique())
    st_06 = set(df2['06'].unique())

    if not (st_05.issubset(etalon_05) or st_06.issubset(etalon_06)):
        temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                            'В колонках 05 или 06 формы 2 есть незаполненные ячейки,\n'
                                            ' или значения отличающиеся от требуемых. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    # создаем 2 датафрейма, по колонке 05 трудоустроены и будут трудоустроены
    empl_now_df = df2[df2['05'] == 'уже трудоустроены']  # те что уже трудоустроены
    empl_future_df = df2[df2['05'] == 'будут трудоустроены']  # те что будут трудоустроены

    # проводим группировку
    empl_now_df_group = empl_now_df.groupby(['02']).agg({'04': sum})
    empl_future_df_group = empl_future_df.groupby(['02']).agg({'04': sum})

    df1_future = df1[df1[
                         'Будут трудоустроены в ОПК'] != 0]  # отбираем в форме 2 специальности по которым есть будущие трудоустроены выпускники
    check_df = empl_future_df_group.merge(df1_future, how='outer', left_on='02', right_on='Специальность')

    # находим строки где есть хотя бы один nan ,это значит что в формах есть разночтения по специальностям
    row_with_nan = check_df[check_df.isna().any(axis=1)]
    row_with_nan.fillna('Специальность не найдена', inplace=True)
    for row in row_with_nan.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'{row[2]} не совпадают данные !!! Отсутствуют данные по этой специальности либо в форме 1 либо в форме 2',
                                            'В форме 1 для этой специальности указаны выпускники которые будут трудоустроены, но в форме 2 такой специальности не найдено или наоборот. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    # отбираем все строки где нет нан
    check_df = check_df[~check_df.isna().any(axis=1)]

    check_df['Результат'] = check_df['04'] == check_df['Будут трудоустроены в ОПК']
    check_df = check_df[~check_df['Результат']]

    # записываем где есть ошибки
    for row in check_df.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'{row[2]} не совпадают данные !!! по форме 1 для этой специальности будут трудоустроено {row[4]} чел.'
                                            f' в форме 2 по этой специальности найдено {int(row[1])} чел.',
                                            'Несовпадает количество выпускников которые будут трудоустроены в форме 1 и в форме 2. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    # обрабатываем датафрейм с уже трудоустроенными
    df1_now = df1[df1[
                      'Трудоустроено в ОПК'] != 0]  # отбираем в форме 2 специальности по которым есть будущие трудоустроены выпускники
    check_df_now = empl_now_df_group.merge(df1_now, how='outer', left_on='02', right_on='Специальность')

    # находим строки где есть хотя бы один nan ,это значит что в формах есть разночтения по специальностям
    row_with_nan_now = check_df_now[check_df_now.isna().any(axis=1)]
    row_with_nan_now.fillna('Специальность не найдена', inplace=True)
    for row in row_with_nan_now.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'Перекрестная проверка трудоустроен/будет трудоустроен: {row[2]} не совпадают данные !!! Отсутствуют данные по этой специальности либо в форме 1 либо в форме 2',
                                            'В форме 1 для этой специальности указаны выпускники которые уже трудоустроены, но в форме 2 такой специальности не найдено или наоборот. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    # отбираем все строки где нет нан
    check_df_now = check_df_now[~check_df_now.isna().any(axis=1)]

    check_df_now['Результат'] = check_df_now['04'] == check_df_now['Трудоустроено в ОПК']
    check_df_now = check_df_now[~check_df_now['Результат']]

    # записываем где есть ошибки
    for row in check_df_now.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'Перекрестная проверка трудоустроен/будет трудоустроен: {row[2]} не совпадают данные !!! по форме 1 для этой специальности трудоустроено {row[3]} чел.'
                                            f' в форме 2 по этой специальности найдено {int(row[1])} чел.',
                                            'Несовпадает количество выпускников которые трудоустроены в форме 1 и в форме 2. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    return error_df


def check_cross_second_error(df1: pd.DataFrame, df2: pd.DataFrame, name_file):
    """
    Функция для проверки  корректности заполнения показателей по целевому приему.
    :param df1:
    :param df2:
    :param name_file:
    :return:
    """
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # проверяем наличие незаполненных ячеек в колонках 05 06
    # отбираем значения первой строки
    df1 = df1[df1['04'] == '02']
    df1 = df1.groupby(['03']).agg({'08': sum, '39': sum})  # группируем
    df1 = df1.reset_index()  # переносим индексы
    df1.columns = ['Специальность', 'Трудоустроено в ОПК', 'Будут трудоустроены в ОПК']
    # Создаем 2 датафрейма из формы 1 целевики будущие и настоящие
    df1_now = df1[['Специальность', 'Трудоустроено в ОПК']]  # целевики трудоустроенные форма 1
    df1_now = df1_now[df1_now['Трудоустроено в ОПК'] != 0]
    df1_future = df1[['Специальность', 'Будут трудоустроены в ОПК']]  # целевики трудоустроенные форма 1
    df1_future = df1_future[df1_future['Будут трудоустроены в ОПК'] != 0]

    # создаем 2 датафрейма целевики уже трудоустроенные и целевики в будущем трудоустроенные
    # создаем датафрейм, по колонке 06 заключили договор о целевом обучении и уже трудоустроенные
    target_df_now = df2[df2['06'] == 'заключили договор о целевом обучении']  # целевики
    target_df_now = target_df_now[target_df_now['05'] == 'уже трудоустроены']  # целевики

    # проводим группировку
    target_df_now = target_df_now.groupby(['02']).agg({'04': sum})
    target_df_now = target_df_now.reset_index()

    union_df_now = df1_now.merge(target_df_now, how='outer', left_on='Специальность', right_on='02')
    #
    #
    # # находим строки где есть хотя бы один nan ,это значит что в формах есть разночтения по специальностям
    row_with_nan = union_df_now[union_df_now.isna().any(axis=1)]
    row_with_nan.fillna('Специальность не найдена', inplace=True)

    for row in row_with_nan.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'Перекрестная проверка целевиков (трудоустроен/будет трудоустроен): в форме 1 - {row[1]} а в форме 2 - {row[3]} не совпадают данные !!! Отсутствуют данные по этой специальности либо в форме 1 либо в форме 2',
                                            'В форме 1 для специальности указаны выпускники целевики которые  трудоустроены, но в форме 2 такой специальности не найдено или наоборот. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
    # # отбираем все строки где нет нан
    check_df_now = union_df_now[~union_df_now.isna().any(axis=1)]

    check_df_now['Результат'] = check_df_now['04'] == check_df_now['Трудоустроено в ОПК']
    check_df_now = check_df_now[~check_df_now['Результат']]
    # записываем где есть ошибки
    for row in check_df_now.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'Перекрестная проверка целевиков (трудоустроен/будет трудоустроен):В форме 1 - {row[1]} не совпадают данные !!! по форме 1 для этой специальности трудоустроено {row[2]} чел.'
                                            f' в форме 2 по этой специальности найдено {int(row[4])} чел.',
                                            'Несовпадает количество выпускников целевиков которые трудоустроены в форме 1 и в форме 2. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
    # создаем 2 датафрейма целевики уже трудоустроенные и целевики в будущем трудоустроенные
    # создаем датафрейм, по колонке 06 заключили договор о целевом обучении и уже трудоустроенные
    target_df_future = df2[df2['06'] == 'заключили договор о целевом обучении']  # целевики
    target_df_future = target_df_future[target_df_future['05'] == 'будут трудоустроены']  # целевики

    # проводим группировку
    target_df_future = target_df_future.groupby(['02']).agg({'04': sum})
    target_df_future = target_df_future.reset_index()

    union_df_future = df1_future.merge(target_df_future, how='outer', left_on='Специальность', right_on='02')
    #
    #
    # # находим строки где есть хотя бы один nan ,это значит что в формах есть разночтения по специальностям
    row_with_nan = union_df_future[union_df_future.isna().any(axis=1)]
    row_with_nan.fillna('Специальность не найдена', inplace=True)
    for row in row_with_nan.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'Перекрестная проверка целевиков (трудоустроен/будет трудоустроен): В форме 1- {row[1]} а в форме 2- {row[3]} !!! Отсутствуют данные по этой специальности либо в форме 1 либо в форме 2',
                                            'В форме 1 для специальности указаны выпускники целевики которые БУДУТ трудоустроены, но в форме 2 выпускников целевиков с такой специальностью которые БУДУт трудоустроены не найдено ИЛИ наоборот. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
    # # отбираем все строки где нет нан
    check_df_future = union_df_future[~union_df_future.isna().any(axis=1)]

    check_df_future['Результат'] = check_df_future['04'] == check_df_future['Будут трудоустроены в ОПК']
    check_df_future = check_df_future[~check_df_future['Результат']]
    # записываем где есть ошибки
    for row in check_df_future.itertuples():
        temp_error_df = pd.DataFrame(data=[[f'{name_file}',
                                            f'{row[1]} не совпадают данные !!! по форме 1 для этой специальности трудоустроено {row[2]} чел.'
                                            f' в форме 2 по этой специальности найдено {int(row[4])} чел.',
                                            'Несовпадает количество выпускников целевиков которые БУДУТ трудоустроены в форме 1 и в форме 2. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                              'Описание ошибки'])
        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

    return error_df


def check_horizont_sum_opk_all(df: pd.DataFrame, name_file, tup_correct):
    """
    Функция для проверки простой горизонтальной суммы 06 = 07 +32,33,34,35,36,37,38,63,64,65,66:77
    :param df:
    :param name_file:
    :param tup_correct:
    :return:
    """
    # получаем строку диапазона
    first_correct = tup_correct[0]
    # конвертируем в инт
    drop_lst = ['08', '09', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
                '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '39', '40',
                '41', '42', '43', '44', '45', '46', '47', '48', '49', '50', '51', '52', '53',
                '54', '55', '56', '57', '58', '59', '60', '61', '62']

    # удаляем колонки лишние колонки
    df.drop(columns=drop_lst, inplace=True)

    # # получаем сумму колонок
    df['Сумма'] = df.iloc[:, 1:].sum(axis=1)
    # # Проводим проверку
    df['Результат'] = df['06'] == df['Сумма']
    # # заменяем булевые значения на понятные
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    # # получаем датафрейм с ошибками и извлекаем индекс
    df = df[df['Результат'] == 'Неправильно'].reset_index()
    # # создаем датафрейм дял добавления в ошибки
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # # обрабатываем индексы строк с ошибками чтобы строки совпадали с файлом excel
    raw_lst_index = df['index'].tolist()  # делаем список
    finish_lst_index = list(map(lambda x: x + first_correct, raw_lst_index))
    finish_lst_index = list(map(lambda x: f'Строка {str(x)}', finish_lst_index))
    temp_error_df['Строка или колонка с ошибкой'] = finish_lst_index
    temp_error_df['Название файла'] = name_file
    temp_error_df[
        'Описание ошибки'] = 'Не выполняется условие: гр. 06 = гр.07 + сумма(всех колонок за исключением распределения по отраслям)'
    return temp_error_df


def check_horizont_chosen_sum_opk(df: pd.DataFrame, tup_checked_cols: list, name_itog_cols, name_file):
    """
    Функция для проверки равенства одиночных или небольших групп колонок
    tup_checked_cols колонки сумму которых нужно сравнить с name_itog_cols чтобы она не превышала это значение
    """
    # Считаем проверяемые колонки
    df['Сумма'] = df[tup_checked_cols].sum(axis=1)
    # Проводим проверку
    df['Результат'] = df[name_itog_cols] == df['Сумма']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    # получаем датафрейм с ошибками и извлекаем индекс
    df = df[df['Результат'] == 'Неправильно'].reset_index()
    # создаем датафрейм дял добавления в ошибки
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # обрабатываем индексы строк с ошибками чтобы строки совпадали с файлом excel
    raw_lst_index = df['index'].tolist()  # делаем список
    finish_lst_index = list(map(lambda x: x + 1, raw_lst_index))
    finish_lst_index = list(map(lambda x: f'Строка {x+9}', finish_lst_index))
    temp_error_df['Строка или колонка с ошибкой'] = finish_lst_index
    temp_error_df['Название файла'] = name_file
    temp_error_df[
        'Описание ошибки'] = f'Не выполняется условие: гр. {name_itog_cols} == сумма {tup_checked_cols} ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!'
    return temp_error_df


def check_vertical_opk_all(df: pd.DataFrame, border, name_file, tup_correct):
    """
    Функция для проверки условия Количество целевиков не должно быть больше чем количество выпускников
    :param df:
    :param name_file:
    :param tup_correct:
    :return:
    """
    first_correct = tup_correct[0]
    second_correct = tup_correct[1]
    foo_df = pd.DataFrame(columns=['01', '02'])

    # Добавляем данные в датафрейм
    foo_df['01'] = df.iloc[0, :]
    foo_df['02'] = df.iloc[1, :]
    foo_df['Результат'] = foo_df['02'] <= foo_df['01']
    foo_df['Результат'] = foo_df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')

    foo_df = foo_df[foo_df['Результат'] == 'Неправильно'].reset_index()
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    # обрабатываем индексы строк с ошибками чтобы строки совпадали с файлом excel
    raw_lst_index = foo_df['index'].tolist()  # делаем список
    finish_lst_index = list(
        map(lambda x: f'Диапазон строк {border + first_correct} - {border + second_correct}, колонка {str(x)}',
            raw_lst_index))

    temp_error_df['Строка или колонка с ошибкой'] = finish_lst_index
    temp_error_df['Название файла'] = name_file
    temp_error_df['Описание ошибки'] = 'Не выполняется условие: стр. 02 <= стр. 01 '
    return temp_error_df


def create_check_tables_opk(high_level_dct: dict):
    """
        Функция для создания файла с данными по каждой специальности
        """
    # Создаем словарь в котором будут храниться словари по специальностям
    code_spec_dct = {}

    # инвертируем словарь так чтобы код специальности стал внешним ключом а названия файлов внутренними
    for poo, spec_data in high_level_dct.items():
        for code_spec, data in spec_data.items():
            if code_spec not in code_spec_dct:
                code_spec_dct[code_spec] = {f'{poo}': high_level_dct[poo][code_spec]}
            else:
                code_spec_dct[code_spec].update({f'{poo}': high_level_dct[poo][code_spec]})

    # Сортируем получившийся словарь по возрастанию для удобства использования
    sort_code_spec_dct = sorted(code_spec_dct.items())
    code_spec_dct = {dct[0]: dct[1] for dct in sort_code_spec_dct}

    # Создаем файл
    wb = openpyxl.Workbook()
    # Создаем листы
    for idx, code_spec in enumerate(code_spec_dct.keys()):
        if code_spec != 'nan':
            wb.create_sheet(title=code_spec, index=idx)

    for code_spec in code_spec_dct.keys():
        if code_spec != 'nan':
            temp_code_df = pd.DataFrame.from_dict(code_spec_dct[code_spec], orient='index')

            temp_code_df = temp_code_df.stack()
            # название такое выбрал потому что было лень заменять значения из блокнота юпитера
            temp_code_df = temp_code_df.to_frame()

            temp_code_df['Суммарный выпуск 2023 г.'] = temp_code_df[0].apply(lambda x: x.get('Колонка 6'))
            temp_code_df[
                'Трудоустроены (по трудовому договору, договору ГПХ в соответствии с трудовым законодательством, законодательством  об обязательном пенсионном страховании)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 7'))
            temp_code_df['Трудоустроены на предприятия оборонно-промышленного комплекса*'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 8'))
            temp_code_df['Трудоустроены на предприятия машиностроения (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 9'))
            temp_code_df['Трудоустроены на предприятия сельского хозяйства'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 10'))
            temp_code_df['Трудоустроены на предприятия металлургии'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 11'))
            temp_code_df['Трудоустроены на предприятия железнодорожного транспорта'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 12'))
            temp_code_df['Трудоустроены на предприятия легкой промышленности'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 13'))
            temp_code_df['Трудоустроены на предприятия химической отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 14'))
            temp_code_df['Трудоустроены на предприятия атомной отрасли (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 15'))
            temp_code_df['Трудоустроены на предприятия фармацевтической отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 16'))
            temp_code_df['Трудоустроены на предприятия отрасли информационных технологий'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 17'))
            temp_code_df['Трудоустроены на предприятия радиоэлектроники (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 18'))
            temp_code_df[
                'Трудоустроены на предприятия топливно-энергетического комплекса (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 19'))
            temp_code_df['Трудоустроены на предприятия транспортной отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 20'))
            temp_code_df['Трудоустроены на предприятия горнодобывающей отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 21'))
            temp_code_df[
                'Трудоустроены на предприятия отрасли электротехнической промышленности (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 22'))
            temp_code_df['Трудоустроены на предприятия лесной промышленности'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 23'))
            temp_code_df['Трудоустроены на предприятия строительной отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 24'))
            temp_code_df[
                'Трудоустроены на предприятия отрасли электронной промышленности (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 25'))
            temp_code_df['Трудоустроены на предприятия индустрии робототехники'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 26'))
            temp_code_df['Трудоустроены на предприятия в отрасли образования'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 27'))
            temp_code_df['Трудоустроены на предприятия в медицинской отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 28'))
            temp_code_df[
                'Трудоустроены на предприятия в отрасли сферы услуг, туризма, торговли, организациях финансового сектора, правоохранительной сферы и управления, средств массовой информации'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 29'))
            temp_code_df['Трудоустроены на предприятия в отрасли искусства'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 30'))
            temp_code_df['Трудоустроены на предприятия иная отрасль'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 31'))

            temp_code_df['Индивидуальные предприниматели'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 32'))
            temp_code_df['Самозанятые (перешедшие на специальный налоговый режим  - налог на профессиональный доход)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 33'))
            temp_code_df['Продолжили обучение'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 34'))
            temp_code_df['Проходят службу в армии по призыву'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 35'))
            temp_code_df[
                'Проходят службу в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации**'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 36'))
            temp_code_df['Находятся в отпуске по уходу за ребенком'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 37'))
            temp_code_df[
                'Будут трудоустроены (по трудовому договору, договору ГПХ в соответствии с трудовым законодательством, законодательством  об обязательном пенсионном страховании)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 38'))
            temp_code_df['Будут трудоустроены на предприятия оборонно-промышленного комплекса* '] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 39'))
            temp_code_df['Будут трудоустроены на предприятия машиностроения (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 40'))
            temp_code_df['Будут трудоустроены на предприятия сельского хозяйства'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 41'))
            temp_code_df['Будут трудоустроены на предприятия металлургии'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 42'))
            temp_code_df['Будут трудоустроены на предприятия железнодорожного транспорта'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 43'))
            temp_code_df['Будут трудоустроены на предприятия легкой промышленности'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 44'))
            temp_code_df['Будут трудоустроены на предприятия химической отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 45'))
            temp_code_df[
                'Будут трудоустроены на предприятия атомной отрасли (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 46'))
            temp_code_df['Будут трудоустроены на предприятия фармацевтической отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 47'))
            temp_code_df['Будут трудоустроены на предприятия отрасли информационных технологий'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 48'))
            temp_code_df[
                'Будут трудоустроены на предприятия радиоэлектроники (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 49'))
            temp_code_df[
                'Будут трудоустроены на предприятия топливно-энергетического комплекса (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 50'))

            temp_code_df['Будут трудоустроены на предприятия транспортной отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 51'))
            temp_code_df['Будут трудоустроены на предприятия горнодобывающей отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 52'))
            temp_code_df[
                'Будут трудоустроены на предприятия отрасли электротехнической промышленности (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 53'))
            temp_code_df['Будут трудоустроены на предприятия лесной промышленности'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 54'))
            temp_code_df['Будут трудоустроены на предприятия строительной отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 55'))
            temp_code_df[
                'Будут трудоустроены на предприятия отрасли электронной промышленности (кроме оборонно-промышленного комплекса)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 56'))
            temp_code_df['Будут трудоустроены на предприятия индустрии робототехники'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 57'))
            temp_code_df['Будут трудоустроены на предприятия в отрасли образования'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 58'))
            temp_code_df['Будут трудоустроены на предприятия в медицинской отрасли'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 59'))
            temp_code_df[
                'Будут трудоустроены на предприятия в отрасли сферы услуг, туризма, торговли, организациях финансового сектора, правоохранительной сферы и управления, средств массовой информации'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 60'))
            temp_code_df['Будут трудоустроены на предприятия в отрасли искусства'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 61'))
            temp_code_df['Будут трудоустроены на предприятия иная отрасль'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 62'))
            temp_code_df['Будут индивидуальными предпринимателями'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 63'))
            temp_code_df['Будут самозанятыми'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 64'))
            temp_code_df['Будут продолжать обучение'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 65'))
            temp_code_df['Будут призваны в армию'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 66'))
            temp_code_df[
                'будут в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации**'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 67'))
            temp_code_df['Будут находиться в отпуске по уходу за ребенком'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 68'))
            temp_code_df['Неформальная занятость (теневой сектор экономики)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 69'))
            temp_code_df[
                'Зарегистрированы в центрах занятости в качестве безработных (получают пособие по безработице) и не планируют трудоустраиваться'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 70'))
            temp_code_df[
                'Не имеют мотивации к трудоустройству (кроме зарегистрированных в качестве безработных) и не планируют трудоустраиваться, в том числе по причинам получения иных социальных льгот '] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 71'))
            temp_code_df['Иные причины нахождения под риском нетрудоустройства'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 72'))
            temp_code_df['Смерть, тяжелое состояние здоровья'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 73'))
            temp_code_df['Находятся под следствием, отбывают наказание'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 74'))
            temp_code_df[
                'Переезд за пределы Российской Федерации (кроме переезда в иные регионы - по ним регион должен располагать сведениями)'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 75'))
            temp_code_df[
                'Не могут трудоустраиваться в связи с уходом за больными родственниками, в связи с иными семейными обстоятельствами'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 76'))
            temp_code_df[
                'Иное в первую очередь выпускники распределяются по всем остальным графам. Данная графа предназначена для очень редких случаев. Если в нее включено более 0,5% выпускников - укажите причины в графе "принимаемые меры"'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 77'))
            temp_code_df[
                'Принимаемые меры по содействию занятости, в том числе по трудоустройству выпускников на предприятия оборонно-промышленного комплекса тезисно - вид меры, охват выпускников мерой'] = \
                temp_code_df[0].apply(lambda x: x.get('Колонка 78'))

            finish_df = temp_code_df.drop([0], axis=1)

            finish_df = finish_df.reset_index()

            finish_df.rename(
                columns={'level_0': 'Название файла', 'level_1': 'Наименование показателей (категория выпускников)'},
                inplace=True)

            dct = {'Строка 1': 'Всего (общая численность выпускников)',
                   'Строка 2': 'из строки 01: имеют договор о целевом обучении'}

            finish_df['Наименование показателей (категория выпускников)'] = finish_df[
                'Наименование показателей (категория выпускников)'].apply(lambda x: dct[x])

            for r in dataframe_to_rows(finish_df, index=False, header=True):
                wb[code_spec].append(r)
            wb[code_spec].column_dimensions['A'].width = 20
            wb[code_spec].column_dimensions['B'].width = 40

    return wb

def extract_code_full(value):
    """
    Функция для извлечения кода специальности из ячейки в которой соединены и код и название специалньости
    """
    value = str(value)
    re_code = re.compile('\d{2}?[.]\d{2}?[.]\d{2}')  # создаем выражение для поиска кода специальности
    result = re.search(re_code, value)
    if result:
        return result.group()
    else:
        return 'Не найден код специальности'