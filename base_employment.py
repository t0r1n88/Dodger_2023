# -*- coding: utf-8 -*-
"""
Скрипт для обработки базового мониторинга 5 строк и мониторинга нозоологий
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


def prepare_base_employment(path_folder_data:str,path_to_end_folder:str):
    """

    """


    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    tup_correct = (9, 23)  # создаем кортеж  с поправками

    try:
        for file in os.listdir(path_folder_data):
            if not file.startswith('~$') and file.endswith('.xls'):
                name_file = file.split('.xls')[0]
                temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                    'Файл с расширением XLS (СТАРЫЙ ФОРМАТ EXCEL)!!! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                             columns=['Название файла', 'Строка или колонка с ошибкой',
                                                      'Описание ошибки'])
                error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                continue

            if not file.startswith('~$') and file.endswith('.xlsx'):
                name_file = file.split('.xlsx')[0]
                print(name_file)
                df = pd.read_excel(f'{path_folder_data}/{file}', skiprows=7, dtype=str)
                # проверяем корректность заголовка
                # создаем множество колонок наличие которых мы проверяем
                check_cols = {'01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15',
                              '16', '17',
                              '18', '19', '20', '21', '22', '23', '24', '25', '26', '27', '28', '29', '30', '31', '32',
                              '33'}
                if not check_cols.issubset(set(df.columns)):
                    temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                        'Проверьте заголовок таблицы в файле.Строка с номерами колонок (01,02,03 и т.д. как в исходной форме)\n должна находиться на 8 строке! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                                 columns=['Название файла', 'Строка или колонка с ошибкой',
                                                          'Описание ошибки'])
                    error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                    continue
                df = df[df['05'] != '16']  # фильтруем строки с проверками
                # отсекаем возможный первый столбец с данными ПОО,начинаем датафрейм с колонки 01 и отсекаем колонки с примечаниями
                df = df.loc[:, '01':'33']

                # # получаем  часть с данными
                mask = pd.isna(df).all(axis=1)  # создаем маску для строк с пропущенными значениями
                # проверяем есть ли строка полностью состоящая из nan
                empty_row_index = np.where(df.isna().all(axis=1))
                if empty_row_index[0].tolist():
                    row_index = empty_row_index[0][0]
                    df = df.iloc[:row_index]
                #     # Проверка на размер таблицы, должно бьть кратно 15
                count_spec = df.shape[0] // 15  # количество специальностей
                df = df.iloc[:count_spec * 15, :]  # отбрасываем строки проверки
                check_code_lst = df['03'].tolist()  # получаем список кодов специальностей
                # Проверка на то чтобы в колонке 03 в первой строке не было пустой ячейки
                if True in mask.tolist():
                    if check_code_lst[0] is np.nan or check_code_lst[0] == ' ':
                        temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                            'В колонке 03 на первой строке не заполнен код специальности. ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!! ']],
                                                     columns=['Название файла', 'Строка или колонка с ошибкой',
                                                              'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                        continue
                # Проверка на непрерывность кода специальности, то есть на 15 строк должен быть только один код
                border_check_code = 0  # счетчик обработанных страниц
                quantity_check_code = len(check_code_lst) // 15  # получаем сколько специальностей в таблице
                correction = 0  # размер поправки
                sameness_error_df = check_sameness_column(check_code_lst, 15, border_check_code, quantity_check_code,
                                                          tup_correct, correction, name_file, 'Код и наименование')
                error_df = pd.concat([error_df, sameness_error_df], axis=0, ignore_index=True)

                blankness_error_df = check_blankness_column(check_code_lst, 15, border_check_code, quantity_check_code,
                                                            tup_correct, correction, name_file, 'Код и наименование')
                error_df = pd.concat([error_df, blankness_error_df], axis=0, ignore_index=True)

                df.columns = list(map(str, df.columns))
                # Заполняем пока пропуски в 15 ячейке для каждой специальности
                df['06'] = df['06'].fillna('15 ячейка')

                # Проводим проверку на корректность данных, отправляем копию датафрейма

                file_error_df = check_error_base_mon(df.copy(), name_file, tup_correct)
                error_df = pd.concat([error_df, file_error_df], axis=0, ignore_index=True)
                if file_error_df.shape[0] != 0:
                    temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                        'В файле обнаружены ошибки!!! ДАННЫЕ ФАЙЛА НЕ ОБРАБОТАНЫ !!!']],
                                                 columns=['Название файла', 'Строка или колонка с ошибкой',
                                                          'Описание ошибки'])
                    error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                    continue

                df['03'] = df['03'].apply(extract_code)  # очищаем от текста в кодах
                # Проверяем на наличие слова error что означает что там есть некорректные значения кодов специальности
                if 'error' in df['03'].values:
                    temp_error_df = pd.DataFrame(data=[[f'{name_file}', '',
                                                        'Некорректные значения в колонке 03 Код специальности.Вместо кода присутствует дата, вместе с кодом есть название,пробел перед кодом и т.п.!!!']],
                                                 columns=['Название файла', 'Строка или колонка с ошибкой',
                                                          'Описание ошибки'])
                    error_df = pd.concat([error_df, temp_error_df], axis=0, ignore_index=True)
                    continue

                code_spec = [spec for spec in df['03'].unique()]

                # Создаем список для строк
                row_cat = [f'Строка {i}' for i in range(1, 16)]
                # Создаем список для колонок
                column_cat = [f'Колонка {i}' for i in range(7, 34)]

                # Создаем словарь нижнего уровня содержащий в себе все данные для каждой специальности
                spec_dict = {}
                for row in row_cat:
                    spec_dict[row] = {key: 0 for key in column_cat}

                # Изменяем последний ключ на строковый поскольку там будут хранится примечания
                for row, value in spec_dict.items():
                    for col, data in value.items():
                        if col == 'Колонка 33':
                            spec_dict[row][col] = ''
                # Создаем словарь среднего уровня содержащй данные по всем специальностям
                poo_dct = {key: copy.deepcopy(spec_dict) for key in code_spec}

                high_level_dct[name_file] = copy.deepcopy(poo_dct)

                """
                В итоге получается такая структура
                {БРИТ:{13.01.10:{Строка 1:{Колонка 1:0}}},ТСИГХ:{22.01.10:{Строка 1:{Колонка 1:0}}}}

                """

                current_code = 'Ошибка проверьте правильность заполнения кодов специальностей'  # чекбокс для проверки заполнения кода специальности

                idx_row = 1  # счетчик обработанных строк

                # Итерируемся по полученному датафрейму через itertuples
                for row in df.itertuples():
                    # если счетчик колонок больше 15 то уменьшаем его до единицы
                    if idx_row > 15:
                        idx_row = 1

                    # Проверяем на незаполненные ячейки и ячейки заполненные пробелами
                    if (row[3] is not np.nan) and (row[3] != ' '):
                        # если значение ячейки отличается от текущего кода специальности то обновляем значение текущего кода
                        if row[3] != current_code:
                            current_code = row[3]

                    data_row = row[7:34]  # получаем срез с нужными данными

                    for idx_col, value in enumerate(data_row, start=1):
                        if idx_col + 6 == 33:
                            # сохраняем примечания в строке
                            high_level_dct[name_file][current_code][f'Строка {idx_row}'][
                                f'Колонка {idx_col + 6}'] = f'{name_file} {check_data_note(value)};'

                        else:
                            high_level_dct[name_file][current_code][f'Строка {idx_row}'][
                                f'Колонка {idx_col + 6}'] += check_data(value)

                    idx_row += 1

        t = time.localtime() # получаем текущее время
        current_time = time.strftime('%H_%M_%S', t)
        wb_check_tables = create_check_tables(high_level_dct) # проверяем данные по каждой специальности
        wb_check_tables.save(f'{path_to_end_folder}/Данные для проверки правильности заполнения файлов от {current_time}.xlsx')

        # получаем уникальные специальности
        all_spec_code = set()
        for poo, spec in high_level_dct.items():
            for code_spec, data in spec.items():
                all_spec_code.add(code_spec)

        itog_df = {key: copy.deepcopy(spec_dict) for key in all_spec_code}

        # Складываем результаты неочищенного словаря
        for poo, spec in high_level_dct.items():
            for code_spec, data in spec.items():
                for row, col_data in data.items():
                    for col, value in col_data.items():
                        if col == 'Колонка 33':
                            itog_df[code_spec][row][col] += check_data_note(value)
                        else:
                            itog_df[code_spec][row][col] += value

        # Сортируем получившийся словарь по возрастанию для удобства использования
        sort_itog_dct = sorted(itog_df.items())
        itog_df = {dct[0]: dct[1] for dct in sort_itog_dct}

        out_df = pd.DataFrame.from_dict(itog_df, orient='index')

        stack_df = out_df.stack()
        # название такое выбрал потому что было лень заменять значения из блокнота юпитера
        frame = stack_df.to_frame()

        frame['Всего'] = frame[0].apply(lambda x: x.get('Колонка 7'))
        frame[
            'Трудоустроены (по трудовому договору, договору ГПХ в соответствии с трудовым законодательством, законодательством  об обязательном пенсионном страховании)'] = \
            frame[0].apply(lambda x: x.get('Колонка 8'))
        frame[
            'В том числе (из трудоустроенных): в соответствии с освоенной профессией, специальностью (исходя из осуществляемой трудовой функции)'] = \
            frame[0].apply(lambda x: x.get('Колонка 9'))
        frame[
            'В том числе (из трудоустроенных): работают на протяжении не менее 4-х месяцев на последнем месте работы'] = \
            frame[0].apply(lambda x: x.get('Колонка 10'))
        frame['Индивидуальные предприниматели'] = frame[0].apply(lambda x: x.get('Колонка 11'))
        frame['Самозанятые (перешедшие на специальный налоговый режим  - налог на профессио-нальный доход)'] = frame[
            0].apply(lambda x: x.get('Колонка 12'))
        frame['Продолжили обучение'] = frame[0].apply(lambda x: x.get('Колонка 13'))
        frame['Проходят службу в армии по призыву'] = frame[0].apply(lambda x: x.get('Колонка 14'))
        frame[
            'Проходят службу в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации*'] = \
            frame[0].apply(lambda x: x.get('Колонка 15'))
        frame['Находятся в отпуске по уходу за ребенком'] = frame[0].apply(lambda x: x.get('Колонка 16'))
        frame['Неформальная занятость (нелегальная)'] = frame[0].apply(lambda x: x.get('Колонка 17'))
        frame[
            'Зарегистрированы в центрах занятости в качестве безработных (получают пособие по безработице) и не планируют трудоустраиваться'] = \
            frame[0].apply(lambda x: x.get('Колонка 18'))
        frame[
            'Не имеют мотивации к трудоустройству (кроме зарегистрированных в качестве безработных) и не планируют трудоустраиваться, в том числе по причинам получения иных социальных льгот '] = \
            frame[0].apply(lambda x: x.get('Колонка 19'))
        frame['Иные причины нахождения под риском нетрудоустройства'] = frame[0].apply(lambda x: x.get('Колонка 20'))
        frame['Смерть, тяжелое состояние здоровья'] = frame[0].apply(lambda x: x.get('Колонка 21'))
        frame['Находятся под следствием, отбывают наказание'] = frame[0].apply(lambda x: x.get('Колонка 22'))
        frame[
            'Переезд за пределы Российской Федерации (кроме переезда в иные регионы - по ним регион должен располагать сведениями)'] = \
            frame[0].apply(lambda x: x.get('Колонка 23'))
        frame[
            'Не могут трудоустраиваться в связи с уходом за больными родственниками, в связи с иными семейными обстоятельствами'] = \
            frame[0].apply(lambda x: x.get('Колонка 24'))
        frame['Выпускники из числа иностранных граждан, которые не имеют СНИЛС'] = frame[0].apply(
            lambda x: x.get('Колонка 25'))
        frame[
            'Иное (в первую очередь выпускники распределяются по всем остальным графам. Данная графа предназначена для очень редких случаев. Если в нее включено более 1 из 200 выпускников - укажите причины в гр. 33 '] = \
            frame[0].apply(lambda x: x.get('Колонка 26'))
        frame['будут трудоустроены'] = frame[0].apply(lambda x: x.get('Колонка 27'))
        frame['будут осуществлять предпринимательскую деятельность'] = frame[0].apply(lambda x: x.get('Колонка 28'))
        frame['будут самозанятыми'] = frame[0].apply(lambda x: x.get('Колонка 29'))
        frame['будут призваны в армию'] = frame[0].apply(lambda x: x.get('Колонка 30'))
        frame[
            'будут в армии на контрактной основе, в органах внутренних дел, Государственной противопожарной службе, органах по контролю за оборотом наркотических средств и психотропных веществ, учреждениях и органах уголовно-исполнительной системы, войсках национальной гвардии Российской Федерации, органах принудительного исполнения Российской Федерации*'] = \
            frame[0].apply(lambda x: x.get('Колонка 31'))
        frame['будут продолжать обучение'] = frame[0].apply(lambda x: x.get('Колонка 32'))
        frame['Принимаемые меры по содействию занятости (тезисно - вид меры, охват выпускников мерой)'] = frame[
            0].apply(lambda x: x.get('Колонка 33'))

        finish_df = frame.drop([0], axis=1)

        finish_df = finish_df.reset_index()

        finish_df.rename(
            columns={'level_0': 'Код специальности', 'level_1': 'Наименование показателей (категория выпускников)'},
            inplace=True)

        dct = {'Строка 1': 'Всего (общая численность выпускников)',
               'Строка 2': 'из общей численности выпускников (из строки 01): лица с ОВЗ',
               'Строка 3': 'из числа лиц с ОВЗ (из строки 02): инвалиды и дети-инвалиды',
               'Строка 4': 'Инвалиды и дети-инвалиды (кроме учтенных в строке 03)',
               'Строка 5': 'Имеют договор о целевом обучении',
               'Строка 6': 'Автосумма строк 02 и 04 - Всего (общая численность выпускников из числа лиц с ОВЗ, инвалидов и детей-инвалидов) '
            ,
               'Строка 7': 'из общей численности выпускников из числа лиц с ОВЗ, инвалидов и детей-инвалидов (из строки 06): с нарушениями: зрения',
               'Строка 8': 'слуха', 'Строка 9': 'опорно-двигательного аппарата',
               'Строка 10': 'тяжелыми нарушениями речи', 'Строка 11': 'задержкой психического развития',
               'Строка 12': 'расстройствами аутистического спектра',
               'Строка 13': 'с инвалидностью вследствие  других причин',
               'Строка 14': 'из общей численности выпускников из числа лиц с ОВЗ, инвалидов и детей-инвалидов (из строки 06): имеют договор о целевом обучении',
               'Строка 15': 'из общей численности выпускников из числа лиц с ОВЗ, инвалидов и детей-инвалидов (из строки 06): принимали участие в чемпионате «Абилимпикс»',
               }
        finish_df['Наименование показателей (категория выпускников)'] = finish_df[
            'Наименование показателей (категория выпускников)'].apply(lambda x: dct[x])
        # добавляем строки с проверкой
        count = 0
        for i in range(15, len(finish_df) + 1, 15):
            new_row = finish_df.iloc[i - 1 + count, :].to_frame().transpose().copy()
            new_row.iloc[:, 1] = 'Проверка (строка не редактируется)'
            new_row.iloc[:, 2:] = 'проверка пройдена'

            # Вставка новой строки через каждые 15 строк
            finish_df = pd.concat([finish_df.iloc[:i + count], new_row, finish_df.iloc[i + count:]]).reset_index(
                drop=True)
            count += 1
        lst_number_row = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12', '13', '14', '15',
                          '16']
        multipler = len(finish_df) // 16  # получаем количество специальностей/профессий
        # вставляем новую колонку
        finish_df.insert(1, 'Номер строки', pd.Series(lst_number_row * multipler))
        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)
        finish_df = finish_df[finish_df['Код специальности'] != 'nan']  # отбрасываем nan
        finish_df.to_excel(f'{path_to_end_folder}/Полная таблица  от {current_time}.xlsx', index=False)

        # Создаем файл с 5 строками
        small_finish_df = pd.DataFrame(columns=finish_df.columns)
        one_finish_df = pd.DataFrame(columns=finish_df.columns)

        lst_code_spec = finish_df['Код специальности'].unique()  # получаем список специальностей
        for code_spec in lst_code_spec:
            temp_df = finish_df[finish_df['Код специальности'] == code_spec]
            small_finish_df = pd.concat([small_finish_df, temp_df.iloc[:5, :]], axis=0, ignore_index=True)
            one_finish_df = pd.concat([one_finish_df, temp_df.iloc[:1, :]], axis=0, ignore_index=True)

        with pd.ExcelWriter(f'{path_to_end_folder}/5 строк Трудоустройство от {current_time}.xlsx') as writer:
            small_finish_df.to_excel(writer, sheet_name='5 строк', index=False)
            one_finish_df.to_excel(writer, sheet_name='1 строка (Всего выпускников)', index=False)

        # Создаем документ
        wb = openpyxl.Workbook()
        for r in dataframe_to_rows(error_df, index=False, header=True):
            wb['Sheet'].append(r)

        wb['Sheet'].column_dimensions['A'].width = 30
        wb['Sheet'].column_dimensions['B'].width = 40
        wb['Sheet'].column_dimensions['C'].width = 50

        wb.save(f'{path_to_end_folder}/ОШИБКИ от {current_time}.xlsx')

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Не найдено значение {e.args}')
    except FileNotFoundError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except PermissionError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Закройте открытые файлы Excel {e.args}')
    # except:
    #     messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
    #                          f'При обработке файла {name_file} возникла ошибка !!!\n'
    #                          f'Проверьте файл на соответствие шаблону')

    else:
        if error_df.shape[0] != 0:
            messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 'Обнаружены ошибки в обрабатываемых файлах.\n'
                                 'Названия файлов с ошибками и ошибки вы можете найти в файле Ошибки.\n'
                                 'Исправьте ошибки и запустите повторную обработку для того чтобы получить полный результат.')
        else:
            messagebox.showinfo('Кассандра Подсчет данных по трудоустройству выпускников',
                                'Данные успешно обработаны.')

if __name__ == '__main__':
    path_data = 'data/example/Базовый Мониторинг трудоустройства'
    path_end = 'data/result/Базовый Мониторинг трудоустройства'
    prepare_base_employment(path_data,path_end)

    print('Lindy Booth')





