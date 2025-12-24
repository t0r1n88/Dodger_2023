"""
Скрипт для создания временных рядов из сводных таблиц созданных Кассандрой
"""
import re

import openpyxl
import pandas as pd
import os
import time


class NotFile(Exception):
    """
    Обработка случаев когда нет файлов в папке
    """
    pass

class NotReqSheets(Exception):
    """
    Обработка случаев когда нет обязательных листов
    """
    pass

class NotReqColumns(Exception):
    """
    Обработка случаев когда нет обязательных колонок
    """
    pass



def preparing_data(data_folder:str,required_columns:dict,dct_index_svod:dict,error_df:pd.DataFrame,set_error_name_file:set):
    """
    Функция для проверки исходных файлов на базовые ошибки и создания списков встречающихся индексов (первой колонки)
    """
    for dirpath, dirnames, filenames in os.walk(data_folder):
        for file in filenames:
            if not file.startswith('~$') and (file.endswith('.xlsx') or file.endswith('.xlsm')):
                try:
                    if file.endswith('.xlsx'):
                        name_file = file.split('.xlsx')[0].strip()
                    else:
                        name_file = file.split('.xlsm')[0].strip()
                    # проверяем на правильность даты в названии
                    result_date = re.search(r'\d{2}_\d{2}_\d{4}', name_file)
                    if result_date:
                        file_date = result_date.group()
                        file_date = file_date.replace('_', '.')
                    else:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'В названии файла отсутствует дата в правильном формате. Требуется формат DD_MM_YYYY т.е. 25.12.2025'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        set_error_name_file.add(name_file)
                        continue
                    # открываем файл для проверки наличия листов и колонок
                    temp_wb = openpyxl.load_workbook(f'{dirpath}/{file}', read_only=True)
                    temp_wb_sheets = set(temp_wb.sheetnames)
                    diff_sheets = set(required_columns.keys()).difference(set(temp_wb_sheets))
                    if len(diff_sheets) != 0:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'Отсутствуют обязательные листы {diff_sheets}'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        set_error_name_file.add(name_file)
                        continue

                    # Собираем возможные индексы которые могут встретиться
                    for sheet, lst_cols in required_columns.items():
                        temp_req_df = pd.read_excel(f'{dirpath}/{file}', sheet_name=sheet)
                        diff_cols = set(lst_cols).difference(set(temp_req_df.columns))
                        if len(diff_cols) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}',
                                       f'На листе {sheet} отсутствуют обязательные колонки {diff_cols}'
                                       ]],
                                columns=['Название файла',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0,
                                                 ignore_index=True)
                            set_error_name_file.add(name_file)
                            continue

                        # Открываем файл для обработки
                        df = pd.read_excel(f'{dirpath}/{file}', sheet_name=sheet)  # открываем файл
                        dct_index_svod[sheet].update(df[df.columns[0]])
                except:
                    temp_error_df = pd.DataFrame(
                        data=[[f'{name_file}',
                               f'Не удалось обработать файл. Возможно файл поврежден'
                               ]],
                        columns=['Название файла',
                                 'Описание ошибки'])
                    error_df = pd.concat([error_df, temp_error_df], axis=0,
                                         ignore_index=True)
                    set_error_name_file.add(name_file)
                    continue
    return dct_index_svod,error_df,set_error_name_file













def processing_time_series(data_folder:str,end_folder:str):
    """
    Функция для формирования временных рядов
    """
    t = time.localtime()  # получаем текущее время и дату
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)
    # Обязательные листы
    error_df = pd.DataFrame(
        columns=['Название файла', 'Описание ошибки'])  # датафрейм для ошибок

    lst_files = []  # список для файлов
    for dirpath, dirnames, filenames in os.walk(data_folder):
        lst_files.extend(filenames)
    # отбираем файлы
    lst_xlsx = [file for file in lst_files if not file.startswith('~$') and file.endswith('.xlsx')]
    quantity_files = len(lst_xlsx)  # считаем сколько xlsx файлов в папке
    # Обрабатываем в зависимости от количества файлов в папке
    if quantity_files == 0:
        raise NotFile
    else:
        required_columns = {'Вакансии по отраслям':['Сфера деятельности','Количество вакансий'],
                            'Вакансии по муниципалитетам':['Муниципалитет','Количество вакансий'],
                            'Зарплата по отраслям':['Сфера деятельности','Средняя ариф. минимальная зп','Медианная минимальная зп']}
        dct_index_svod = {key:set() for key in required_columns.keys()} # словарь для хранения всех значений сводов которые могут встретиться в файлах
        set_error_name_file = set() # множество для хранения названий файлов с ошибками

        dct_index_svod,error_df,set_error_name_file = preparing_data(data_folder,required_columns,dct_index_svod,error_df,set_error_name_file) # Проверяем на ошибки

        # Создаем словарь с базовыми датафреймами
        dct_base_df = dict()

        for name_sheet,set_index in dct_index_svod.items():
            dct_base_df[name_sheet] = pd.DataFrame(index=sorted([value for value in set_index if value != 'Итого']))


        for dirpath, dirnames, filenames in os.walk(data_folder):
            for file in filenames:
                if not file.startswith('~$') and (file.endswith('.xlsx') or file.endswith('.xlsm')):
                    try:
                        if file.endswith('.xlsx'):
                            name_file = file.split('.xlsx')[0].strip()
                        else:
                            name_file = file.split('.xlsm')[0].strip()
                        if name_file not in set_error_name_file:
                            print(name_file)  # обрабатываемый файл
                            # ха повторяющийся код ну и ладно
                            result_date = re.search(r'\d{2}_\d{2}_\d{4}', name_file)
                            result_date = result_date.group().replace('_','.')
                            for sheet, lst_cols in required_columns.items():
                                temp_req_df = pd.read_excel(f'{dirpath}/{file}', sheet_name=sheet)
                                temp_req_df.set_index(temp_req_df.columns[0],inplace=True)
                                if temp_req_df.shape[1] == 1:
                                    temp_req_df.columns = [result_date]
                                    base_df = dct_base_df[sheet] # получаем базовый датафрейм
                                    base_df= base_df.join(temp_req_df)
                                    base_df.fillna(0,inplace=True)
                                    dct_base_df[sheet] = base_df
















                    except:
                        temp_error_df = pd.DataFrame(
                            data=[[f'{name_file}',
                                   f'Не удалось обработать файл. Возможно файл поврежден'
                                   ]],
                            columns=['Название файла',
                                     'Описание ошибки'])
                        error_df = pd.concat([error_df, temp_error_df], axis=0,
                                             ignore_index=True)
                        continue

        # Сохраняем в горизонтальном виде
        with pd.ExcelWriter(f'{end_folder}/ Горизонтальный вид.xlsx',engine='openpyxl') as writer:
            for sheet_name, df in dct_base_df.items():
                if sheet_name == 'Зарплата по отраслям':
                    continue
                # Преобразуем и сортируем колонки-даты
                date_cols = []

                for col in df.columns:
                    date_obj = pd.to_datetime(col, format='%d.%m.%Y')
                    date_cols.append((date_obj, col))

                # Сортируем даты
                date_cols.sort(key=lambda x: x[0])
                # Создаем новые названия колонок
                new_columns =[date for _, date in date_cols]

                # Переупорядочиваем DataFrame
                df = df[new_columns]

                # Преобразуем названия колонок-дат
                for date_obj, old_name in date_cols:
                    df = df.rename(columns={old_name: date_obj})
                df.columns = df.columns.strftime('%d.%m.%Y')
                df.to_excel(writer,sheet_name=sheet_name,index=True)



        error_df.to_excel(f'{end_folder}/Ошибки_{current_time}.xlsx',index=False)

















if __name__ == '__main__':
    main_data_folder = 'data/Своды'
    main_end_folder = 'data/РЕЗУЛЬТАТ'
    processing_time_series(main_data_folder,main_end_folder)
    print('Lindy Booth')







