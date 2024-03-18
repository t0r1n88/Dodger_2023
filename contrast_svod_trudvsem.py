"""
Скрипт для подсчета изменения в сводных таблицах данных полученных с помощью Кассандры
"""
from support_functions import write_df_to_excel_color_selection # получаем функцию для записи в листы Excel
import pandas as pd
import numpy as np
import time
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill


class FirstNotSheets(Exception):
    """
    Класс для обработки случаев когда не хватает обязательных листов в первом файле
    """
    pass

class SecondNotSheets(Exception):
    """
    Класс для обработки случаев когда не хватает обязательных листов в втором файле
    """
    pass

class NotColumnVac(Exception):
    """
    Класс для проверки наличия колонки Количество вакансий
    """







def sum_category(row: pd.Series, user_sep: str):
    """
    Суммирование колонок для ID
    """
    out_id = f'{user_sep}'.join(row.tolist())  # создаем выходную строку

    return out_id


def processing_diff_svod(first_file:str,second_file:str,end_folder:str):
    """
    Функция для подсчета разницы между двумя сводами
    """
    t = time.localtime()  # получаем текущее время
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)
    # Список обязательных листов которые должны быть в файле
    lst_svod_sheets = ['Вакансии по отраслям', 'Вакансии по работодателям', 'Зарплата по отраслям',
                       'Зарплата по работодателям',
                       'Образование по отраслям', 'Образование по работодателям', 'График работы по отраслям',
                       'График работы по работодателям',
                       'Тип занятости по отраслям', 'Тип занятости по работодателям', 'Квоты по отраслям',
                       'Квоты по работодателям',
                       'Требуемый опыт по отраслям', 'Требуемый опыт по работодателям']

    lst_not_standard_sheets = ['Зарплата по отраслям',
                               'Зарплата по работодателям']  # список листов без колонки Количество вакансий

    dct_df = dict() # создаем словарь для хранения датафреймов

    #  получаем листы для проверки
    first_wb = openpyxl.load_workbook(first_file)
    first_wb_sheets = first_wb.sheetnames
    first_wb.close()
    second_wb = openpyxl.load_workbook(second_file)
    second_wb_sheets = second_wb.sheetnames
    second_wb.close()

    # Проводим проверки на наличие нужных листов
    diff_first_file_sheets = set(lst_svod_sheets).difference(set(first_wb_sheets))
    if len(diff_first_file_sheets) != 0:
        raise FirstNotSheets

    diff_second_file_sheets = set(lst_svod_sheets).difference(set(second_wb_sheets))
    if len(diff_second_file_sheets) != 0:
        raise SecondNotSheets

    # Обрабатываем листы
    for name_sheet in lst_svod_sheets:
        # Если свод стандартный по количеству вакансий
        if name_sheet not in lst_not_standard_sheets:
            first_df = pd.read_excel(first_file, sheet_name=name_sheet)  # первый файл для сравнения
            second_df = pd.read_excel(second_file, sheet_name=name_sheet)  # второй файл для сравнения

            # получаем колонки которые будем использовать для создания ID
            first_key_columns = list(first_df.columns)
            if 'Количество вакансий' not in first_key_columns: # проверяем наличие колонки с вакансиями
                raise NotColumnVac
            first_key_columns.remove('Количество вакансий')

            second_key_columns = list(second_df.columns)
            if 'Количество вакансий' not in second_key_columns: # проверяем наличие колонки с вакансиями
                raise NotColumnVac
            second_key_columns.remove('Количество вакансий')

            # Делаем колонки на основе котороых будут создаваться ID строковыми
            first_df[first_key_columns] = first_df[first_key_columns].astype(str)
            second_df[second_key_columns] = second_df[second_key_columns].astype(str)

            # Создаем колонки по которым будет вестись объединение
            first_df['ID'] = first_df[first_key_columns].apply(lambda x: sum_category(x, '-'), axis=1)
            second_df['ID'] = second_df[second_key_columns].apply(lambda x: sum_category(x, '-'), axis=1)

            # Проводим внешнее слияние
            merge_df = first_df.merge(second_df, how='outer', left_on=['ID'], right_on=['ID'], indicator=True)

            # удаляем лишние колонки
            del_columns = [column for column in merge_df.columns if ('вакансий' not in column) & ('ID' not in column)]
            merge_df.drop(columns=del_columns, inplace=True)

            # Упорядочиваем колонки
            merge_df = merge_df.reindex(
                columns=['ID', 'Количество вакансий_x', 'Количество вакансий_y'])  # меняем местами
            merge_df.columns = ['Показатель', 'Первая таблица', 'Вторая таблица']

            merge_df.fillna(0, inplace=True) # заполняем наны

            merge_df[['Первая таблица', 'Вторая таблица']] = merge_df[['Первая таблица', 'Вторая таблица']].astype(
                float, errors='ignore') # приводик колонки с числами к флоат

            # Создаем колонки с подчетом разниц
            merge_df['Разница'] = merge_df['Вторая таблица'] - merge_df['Первая таблица']
            merge_df['Абсолютная разница'] = abs(merge_df['Вторая таблица'] - merge_df['Первая таблица'])
            merge_df['Изменение в %'] = round(
                ((merge_df['Вторая таблица'] - merge_df['Первая таблица']) / merge_df['Первая таблица']) * 100, 2)

            merge_df['Отношение второй таблицы к первой %'] = round(
                (merge_df['Вторая таблица'] / merge_df['Первая таблица']) * 100, 2)

            merge_df.sort_values(by='Показатель', inplace=True)  # Сортируем по показателю



        else:
            print('Лист с зарплатой')

        dct_df[name_sheet] = merge_df # сохраняем в словарь

        # Создаем словарь с параметрами записи
        dct_change = {'number_column':3,'font':Font(color='FF000000'),
                      'fill':PatternFill(fill_type='solid', fgColor='ffa500'),
                      'find_value':'-'}
        change_wb = write_df_to_excel_color_selection(dct_df,False,[dct_change])
        change_wb.save(f'{end_folder}/Изменения от {current_time}.xlsx')
























if __name__ == '__main__':
    main_first_file = 'data/Аналитика по вакансиям региона/15_03_2024/Свод по региону Бурятия от 14 марта.xlsx'
    main_second_file = 'data/Аналитика по вакансиям региона/16_03_2024/Свод по региону Бурятия от 16 марта.xlsx'
    main_end_folder = 'data'

    processing_diff_svod(main_first_file,main_second_file,main_end_folder)

    print('Lindy Booth !!!')































