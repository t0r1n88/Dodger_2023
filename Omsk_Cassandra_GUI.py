# -*- coding: utf-8 -*-
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


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

class ShapeDiffierence(Exception):
    """
    Класс для обозначения несовпадения размеров таблицы
    """
    pass


class ColumnsDifference(Exception):
    """
    Класс для обозначения того что названия колонок не совпадают
    """
    pass

def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()


def select_file_data_xlsx():
    """
    Функция для выбора файла с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global path_data_omsk
    # Получаем путь к файлу
    path_data_omsk = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def check_data(cell):
    """
    Метод для проверки значения ячейки
    :param cell: значение ячейки
    :return: число в формате int
    """
    if cell is np.nan:
        return 0
    if cell.isdigit():
        return int(cell)
    else:
        return 0


def group_text_value(value):
    """
    функция для группировки текстовых данных, группирует только заполненные строки
    """
    tmp_set = set(value.tolist())  # создаем множество
    tmp_set.discard('Не заполнено')
    tmp_set.discard('0')
    return ';'.join(tmp_set)


def check_sameness_range(df: pd.DataFrame, names_columns: list, border, amendment):
    """
    Функция для проверки единнобразия значений в колонке по диапазонам
    """
    # получаем поправки для того чтобы диапазон ошибки указывался корректно
    first_correct = amendment[0]
    second_correct = amendment[1]
    _error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки'])  # датафрейм для ошибок
    for column in names_columns:
        _temp_set = set(df[column].tolist())
        if len(_temp_set) != 1:
            temp_error_df = pd.DataFrame(
                data=[[f'Диапазон строк {border + first_correct} - {border + second_correct},Колонка {column}',
                       'В указанном диапазоне и колонке  должны быть одинаковые значения во избежание ошибок при подсчете,если указан диапазон вне таблицы то удалите строки этого диапазона']],
                columns=['Строка или колонка с ошибкой',
                         'Описание ошибки'])
            _error_df = pd.concat([_error_df, temp_error_df], axis=0, ignore_index=True)
    return _error_df


def check_horizont_all_sum_error_omsk(df: pd.DataFrame, tup_exluded_cols: tuple, name_itog_cols):
    """
    Функция для проверки горизонтальных сумм по всей строке
    сумма в колонке 05 должна быть равна сумме всех колонок за исключением 07 и 15 Пример
    """
    # получаем список колонок
    all_sum_cols = list(df)
    # удаляем колонки
    for name_cols in tup_exluded_cols:
        all_sum_cols.remove(name_cols)
    # удаляем итоговую колонку
    all_sum_cols.remove(name_itog_cols)

    # получаем сумму колонок за вычетом исключаемых и итоговой колонки
    df['Сумма'] = df[all_sum_cols].sum(axis=1)
    # Проводим проверку
    df['Результат'] = df[name_itog_cols] == df['Сумма']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    # получаем датафрейм с ошибками и извлекаем индекс
    df = df[df['Результат'] == 'Неправильно'].reset_index()
    # создаем датафрейм дял добавления в ошибки
    temp_error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки', ])
    # обрабатываем индексы строк с ошибками чтобы строки совпадали с файлом excel
    raw_lst_index = df['index'].tolist()  # делаем список
    finish_lst_index = list(map(lambda x: x + 1, raw_lst_index))
    finish_lst_index = list(map(lambda x: f'Строка {x + 6}', finish_lst_index))
    temp_error_df['Строка или колонка с ошибкой'] = finish_lst_index
    temp_error_df[
        'Описание ошибки'] = f'Не выполняется условие: гр. {name_itog_cols} = сумма остальных гр. за искл.{tup_exluded_cols}!!!'
    return temp_error_df


def check_first_error_omsk(df: pd.DataFrame, tup_correct):
    """
    Функция для проверки гр. 8 и гр. 10 и 11 и  < гр. 7
    """
    # получаем строку диапазона
    first_correct = tup_correct[0]

    # Проводим проверку
    df['Результат'] = (df['7'] >= df['8']) & (df['7'] >= df['10']) & (df['7'] >= df['11'])
    # заменяем булевые значения на понятные
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    # получаем датафрейм с ошибками и извлекаем индекс
    df = df[df['Результат'] == 'Неправильно'].reset_index()
    # создаем датафрейм дял добавления в ошибки
    temp_error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки'])
    # обрабатываем индексы строк с ошибками чтобы строки совпадали с файлом excel
    raw_lst_index = df['index'].tolist()  # делаем список
    finish_lst_index = list(map(lambda x: x + first_correct, raw_lst_index))
    finish_lst_index = list(map(lambda x: f'Строка {str(x)}', finish_lst_index))
    temp_error_df['Строка или колонка с ошибкой'] = finish_lst_index
    temp_error_df['Описание ошибки'] = 'Не выполняется условие: гр. 8 <= гр.7  или гр. 10 <= гр. 7 или гр. 11<= 7 '
    return temp_error_df


def check_vertical_chosen_sum_omsk(df: pd.DataFrame, lst_checked_rows: list, itog_row, border, amendment):
    """
    Функция для проверки вертикальной суммы заданных строк сумма значений в tupl_checked_row должна быть равной ил меньше чем значение
    в itog_row
    """
    _error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки'])  # датафрейм для ошибок
    first_correct = amendment[0]
    second_correct = amendment[1]

    # обрабаотываем список строк чтобы привести его в читаемый вид
    lst_out_rows = list(map(lambda x: x + 1, lst_checked_rows))

    # делаем значения строковыми
    lst_out_rows = list(map(str, lst_out_rows))
    # Добавляем ноль в строки
    lst_out_rows = list(map(lambda x: '0' + x, lst_out_rows))
    # обрабатываем формат выходной строки
    out_itog_row = f'0{str(itog_row + 1)}'

    # создаем временный датафрейм
    foo_df = pd.DataFrame()
    # разворачиваем строки в колонки
    for idx_row in lst_checked_rows:
        foo_df[idx_row] = df.iloc[idx_row, :]
    # добавляем итоговую колонку
    foo_df[itog_row] = df.iloc[itog_row, :]

    # суммируем

    foo_df['Сумма'] = foo_df[lst_checked_rows].sum(axis=1)
    foo_df['Результат'] = foo_df[itog_row] >= foo_df['Сумма']

    foo_df['Результат'] = foo_df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    # получаем датафрейм с ошибками и извлекаем индекс
    error_df = foo_df[foo_df['Результат'] == 'Неправильно'].reset_index()
    # Добавляем слово колонка
    error_df['index'] = error_df['index'].apply(lambda x: 'Колонка ' + str(x))
    # создаем датафрейм дkz добавления в ошибки

    for row in error_df.itertuples():
        temp_error_df = pd.DataFrame(
            data=[[f'Диапазон строк {border + first_correct} - {border + second_correct} {row[1]}',
                   f'В указанном диапазоне и колонке  не соблюдается условие по вертикали']],
            columns=['Строка или колонка с ошибкой',
                     'Описание ошибки'])
        _error_df = pd.concat([_error_df, temp_error_df], axis=0, ignore_index=True)

    return _error_df


def check_error_omsk(df: pd.DataFrame, border, size_range, amendment):
    """
    Функция для проверки на ошибки
    """
    # колонки c с числами
    columns_to_apply = ['6', '7', '8', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22',
                        '23'
        , '24', '25', '26', '27', '28', '29', '30', '31', '32', '33']

    error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки'])  # датафрейм для ошибок
    quantity_range = df.shape[0] // size_range  # получаем количество проходов которые нужно пройти
    sameness_columns = ['1', '2', '3']  # список колонок которые нужно проверить на единообразие внутри диапазона

    # данные для проверки горизонтальной суммы по строкам колонка 6 должна быть равна сумме всех колонок за исключением excluded_cols
    itog_cols = '6'  # номер колонки по которой будет проверятся сумма по горизонтали
    excluded_cols = ('1', '2', '3', '4', '5', '8', '9', '10', '11', '34')  # колонки которые не будут учитываться

    # данные для проверки строк

    for i in range(quantity_range):
        temp_df = df.iloc[border:border + size_range, :]
        error_sameness_df = check_sameness_range(temp_df.copy(), sameness_columns, border,
                                                 amendment)  # проверка на единообразие
        error_df = pd.concat([error_df, error_sameness_df], axis=0, ignore_index=True)

        # Проверяем строки - стр 03 <= стр 02
        vertical_first_error_df = check_vertical_chosen_sum_omsk(temp_df[columns_to_apply], [2], 1, border, amendment)
        error_df = pd.concat([error_df, vertical_first_error_df], axis=0, ignore_index=True)
        # Проверяем строки - 01 >=02,04,05
        vertical_second_error_df = check_vertical_chosen_sum_omsk(temp_df[columns_to_apply], [1, 3, 4], 0, border,
                                                                  amendment)
        error_df = pd.concat([error_df, vertical_second_error_df], axis=0, ignore_index=True)

        border += 5  # смещаем диапазон

    # Проверяем сумму по всем строкам. Колонка Всего(6) должна быть равна остальным колонкам за исключением категорий трудоустроенных
    all_horizont_error_df = check_horizont_all_sum_error_omsk(df.copy(), excluded_cols, itog_cols)
    error_df = pd.concat([error_df, all_horizont_error_df], axis=0, ignore_index=True)

    # Проверяем сумму 7 колонки (Трудоустроено) и всех колонок подкатегорий
    chosen_horizint_error_df = check_first_error_omsk(df.copy(), amendment)
    error_df = pd.concat([error_df, chosen_horizint_error_df], axis=0, ignore_index=True)

    return error_df


def processing_svod_employment_omsk():
    """
    Фугкция для создания свода по трудоустройству
    :return:
    """
    try:
        # создаем датафрейм для регистрации ошибок
        error_df = pd.DataFrame(columns=['Строка или колонка с ошибкой', 'Описание ошибки'])
        # загружаем файл
        df = pd.read_excel(f'{path_data_omsk}', dtype=str, skiprows=6,header=None)

        df = df.iloc[:, :34]  # убираем строки проверки

        # заменяем названия колонок
        df.columns = list(map(str, range(1, df.shape[1] + 1)))

        # колонки к которым надо применить числовые суммирование
        columns_to_apply = ['6', '7', '8', '10', '11', '12', '13', '14', '15', '16', '17', '18', '19', '20', '21', '22',
                            '23'
            , '24', '25', '26', '27', '28', '29', '30', '31', '32', '33']
        df[columns_to_apply] = df[columns_to_apply].applymap(check_data)  # обрабатываем числовые колонки
        # Проверяем на количество строк, должно быть кратно 5
        if df.shape[0] % 5 != 0:
            temp_error_df = pd.DataFrame(data=[['',
                                                'Количество строк не кратно 5 !!! Возможно пропущены строки с данными или есть лишние строки в конце таблицы']],
                                         columns=['Строка или колонка с ошибкой',
                                                  'Описание ошибки'])
            error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)

        # Проверяем ошибки
        tup_correct = (7, 11)  # создаем кортеж  с поправками
        border = 0  # счетчик строк
        size_range = 5  # сколько строк занимает каждый диапазон
        check_error_df = check_error_omsk(df.copy(), border, size_range, tup_correct)
        error_df = pd.concat([error_df, check_error_df], axis=0, ignore_index=True)

        df.head(1)

        # заполняем нан в текстовых колонках
        df['9'] = df['9'].fillna('Не заполнено')
        df['34'] = df['34'].fillna('Не заполнено')

        # проводим группировку
        df = df.groupby(['2', '3', '4', '5']).agg(
            {'6': sum, '7': sum, '8': sum,'9': group_text_value,'10': sum, '11': sum, '12': sum, '13': sum, '14': sum, '15': sum,
             '16': sum,
             '17': sum, '18': sum, '19': sum, '20': sum, '21': sum, '22': sum, '23': sum, '24': sum, '25': sum, '26': sum,
             '27': sum,
             '28': sum, '29': sum, '30': sum, '31': sum, '32': sum, '33': sum,
             '34': group_text_value})

        # вытаскиваем индексы
        df = df.reset_index()

        all_sum_df = df.sum(axis=0).to_frame().transpose()  # создаем общую сумму по всем вычисляемым колонкам

        all_sum_df.drop(columns=['2', '3', '4', '5', '34'], inplace=True)  # удаляем лишние колонки

        all_sum_df['9'] = ''  # очищаем колонку 9

        # присваиваем названия колонкам
        df.columns = ['Код', 'Наименование', 'Номер строки', 'Наименование показателей', 'Суммарный выпуск (человек)',
                      'Трудоустроены (по трудовому договору, договору ГПХ в соответствии с трудовым законодательством, законодательством  об обязательном пенсионном страховании)',
                      'В том числе на предприятия ОПК', 'Наименование предприятий',
                      'В том числе (из трудоустроенных): в соответствии с освоенной профессией, специальностью (исходя из осуществляемой трудовой функции',
                      'В том числе (из трудоустроенных): работают на протяжении не менее 4-х месяцев на последнем месте работы',
                      'Индивидуальные предприниматели',
                      'Самозанятые (перешедшие на специальный налоговый режим  - налог на профессиональный доход)',
                      'Продолжили обучение', 'Проходят службу в армии по призыву',
                      'Проходят службу в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации*',
                      'Находятся в отпуске по уходу за ребенком', 'Неформальная занятость',
                      'Зарегистрированы в центрах занятости в качестве безработных (получают пособие по безработице) и не планируют трудоустраиваться',
                      'Не имеют мотивации к трудоустройству (кроме зарегистрированных в качестве безработных) и не планируют трудоустраиваться, в том числе по причинам получения иных социальных льгот',
                      'Иные причины нахождения под риском нетрудоустройства', 'Смерть, тяжелое состояние здоровья',
                      'Находятся под следствием, отбывают наказание',
                      'Переезд за пределы Российской Федерации (кроме переезда в иные регионы - по ним регион должен располагать сведениями)',
                      'Не могут трудоустраиваться в связи с уходом за больными родственниками, в связи с иными семейными обстоятельствами',
                      'Выпускники из числа иностранных граждан, которые не имеют СНИЛС',
                      'Иное (в первую очередь выпускники распределяются по всем остальным графам. Данная графа предназначена для очень редких случаев. Если в нее включено более 1 из 200 выпускников - укажите причины в гр. 33',
                      'будут трудоустроены', 'будут осуществлять предпринимательскую деятельность', 'будут самозанятыми',
                      'будут призваны в армию',
                      'будут в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации*',
                      'будут продолжать обучение',
                      'Принимаемые меры по содействию занятости (тезисно - вид меры, охват выпускников мерой)']

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        with pd.ExcelWriter(f'{path_to_end_folder}/Свод Трудоустройство от {current_time}.xlsx') as writer:
            df.to_excel(writer, sheet_name='СВОД', index=False)
            all_sum_df.to_excel(writer, sheet_name='Итоги по колонкам', index=False)

        # для ошибок
        # Создаем документ
        wb = openpyxl.Workbook()
        for r in dataframe_to_rows(error_df, index=False, header=True):
            wb['Sheet'].append(r)

        wb['Sheet'].column_dimensions['A'].width = 40
        wb['Sheet'].column_dimensions['B'].width = 70

        wb.save(f'{path_to_end_folder}/ОШИБКИ от {current_time}.xlsx')
    except NameError:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                         f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError as e:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Не найдено значение {e.args}')
    except FileNotFoundError:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except PermissionError as e:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Закройте открытые файлы Excel {e.args}')

    else:
        if error_df.shape[0] != 0:
            messagebox.showerror('ЦОПП Омская область ver 1.0',
                                 'Обнаружены ошибки в обрабатываемом файле.\n'
                                 'описания ошибок вы можете найти в файле Ошибки.\n'
                                 'Исправьте ошибки и запустите повторную обработку для того чтобы получить верный результат результат.')
        else:
            messagebox.showinfo('ЦОПП Омская область ver 1.0',
                                'Данные успешно обработаны, ошибок не обнаружено.')

# разница 2 датафреймов

"""
Функции для нахождения разницы между 2 таблицами
"""


def select_first_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_first_diffrence
    # Получаем путь к файлу
    data_first_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_second_diffrence
    # Получаем путь к файлу
    data_second_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_diffrence():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_diffrence
    path_to_end_folder_diffrence = filedialog.askdirectory()


def processing_diffrence():
    """
    Функция для вычисления разницы между двумя таблицами
    """
    # загружаем датафреймы
    try:
        dif_first_sheet_name = entry_first_sheet_name_diffrence.get()
        dif_second_sheet_name = entry_second_sheet_name_diffrence.get()
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

        # записываем
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        # делаем так чтобы записать на разные листы
        with pd.ExcelWriter(f'{path_to_end_folder_diffrence}/Разница между 2 таблицами {current_time}.xlsx') as writer:
            df_cols.to_excel(writer, sheet_name='По колонкам')
            df_rows.to_excel(writer, sheet_name='По строкам')

    except ShapeDiffierence:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Не совпадают размеры таблиц, В первой таблице {df1.shape[0]}-стр. и {df1.shape[1]}-кол.\n'
                             f'Во второй таблице {df2.shape[0]}-стр. и {df2.shape[1]}-кол.')

    except ColumnsDifference:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Названия колонок в сравниваемых таблицах отличаются\n'
                             f'Колонок:{diff_columns}  нет во второй таблице !!!\n'
                             f'Сделайте названия колонок одинаковыми.')

    except NameError:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except ValueError:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'В файлах нет листа с таким названием!\n'
                             f'Проверьте написание названия листа')
    except FileNotFoundError:
        messagebox.showerror('ЦОПП Омская область ver 1.0',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    # except:
    #     logging.exception('AN ERROR HAS OCCURRED')
    #     messagebox.showerror('Веста Обработка таблиц и создание документов ver 1.34',
    #                          'Возникла ошибка!!! Подробности ошибки в файле error.log')
    else:
        messagebox.showinfo('ЦОПП Омская область ver 1.0','Обработка завершена!')



if __name__ == '__main__':
    window = Tk()
    window.title('ЦОПП Омская область ver 1.0')
    window.geometry('700x860')
    window.resizable(False, False)


    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_svod_empl = ttk.Frame(tab_control)
    tab_control.add(tab_svod_empl, text='Трудоустройство СВОД')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_svod_empl,
                      text='Центр опережающей профессиональной подготовки Омской области\n'
                           'Скрипт для обработки выгрузки по трудоустройству с добавлением 2 колонок ОПК\n'
                           'Заголовок таблицы должен занимать 6 строк\n'
                           'Выгрузка должна находится на первом по порядку листе')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo-omsk.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_svod_empl,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data = Button(tab_svod_empl, text='1) Выберите файл с данными', font=('Arial Bold', 20),
                             command=select_file_data_xlsx
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_svod_empl, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    #Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_svod_empl, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_svod_employment_omsk
                                  )
    btn_proccessing_data.grid(column=0, row=4, padx=10, pady=10)

    # Создаем вкладку для нахождения разницы
    """
    Разница двух таблиц
    """
    tab_diffrence = ttk.Frame(tab_control)
    tab_control.add(tab_diffrence, text='Разница 2 таблиц')
    tab_control.pack(expand=1, fill='both')

    # Добавляем виджеты на вкладку разница 2 двух таблиц
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_diffrence,
                      text='Центр опережающей профессиональной подготовки Омской области\n'
                           'Количество строк и колонок в таблицах должно совпадать\n'
                           'Названия колонок в таблицах должны совпадать'
                           '\nДля корректной работы программмы уберите из таблицы объединенные ячейки')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_com = resource_path('logo-omsk.png')
    img_diffrence = PhotoImage(file=path_com)
    Label(tab_diffrence,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    frame_data_for_diffrence.grid(column=0, row=2, padx=10)

    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_diffrence = Button(frame_data_for_diffrence, text='1) Выберите файл с первой таблицей',
                                      font=('Arial Bold', 10),
                                      command=select_first_diffrence
                                      )
    btn_data_first_diffrence.grid(column=0, row=3, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_first_sheet_name_diffrence = StringVar()
    # Описание поля
    label_first_sheet_name_diffrence = Label(frame_data_for_diffrence,
                                             text='2) Введите название листа, где находится первая таблица')
    label_first_sheet_name_diffrence.grid(column=0, row=4, padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_first_sheet_name_diffrence,
                                             width=30)
    first_sheet_name_entry_diffrence.grid(column=0, row=5, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_diffrence = Button(frame_data_for_diffrence, text='3) Выберите файл со второй таблицей',
                                       font=('Arial Bold', 10),
                                       command=select_second_diffrence
                                       )
    btn_data_second_diffrence.grid(column=0, row=6, padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name_diffrence = StringVar()
    # Описание поля
    label_second_sheet_name_diffrence = Label(frame_data_for_diffrence,
                                              text='4) Введите название листа, где находится вторая таблица')
    label_second_sheet_name_diffrence.grid(column=0, row=7, padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_second_sheet_name_diffrence,
                                               width=30)
    second__sheet_name_entry_diffrence.grid(column=0, row=8, padx=5, pady=5, ipadx=15, ipady=10)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_diffrence = Button(frame_data_for_diffrence, text='5) Выберите конечную папку',
                                      font=('Arial Bold', 10),
                                      command=select_end_folder_diffrence
                                      )
    btn_select_end_diffrence.grid(column=0, row=10, padx=10, pady=10)

    # Создаем кнопку Обработать данные
    btn_data_do_diffrence = Button(tab_diffrence, text='6) Обработать таблицы', font=('Arial Bold', 20),
                                   command=processing_diffrence
                                   )
    btn_data_do_diffrence.grid(column=0, row=11, padx=10, pady=10)

    window.mainloop()