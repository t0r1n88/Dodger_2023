"""
Скрипт для создания временных рядов из сводных таблиц созданных Кассандрой
"""
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













def processing_time_series(data_folder:str,end_folder:str):
    """
    Функция для формирования временных рядов
    """
    t = time.localtime()  # получаем текущее время и дату
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)
    # Обязательные листы
    required_sheets = {'Вакансии по отраслям','Вакансии по муниципалитетам','Зарплата по отраслям'}
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

        for dirpath, dirnames, filenames in os.walk(data_folder):
            for file in filenames:
                if not file.startswith('~$') and (file.endswith('.xlsx') or file.endswith('.xlsm')):
                    try:
                        if file.endswith('.xlsx'):
                            name_file = file.split('.xlsx')[0].strip()
                        else:
                            name_file = file.split('.xlsm')[0].strip()
                        print(name_file)  # обрабатываемый файл
                        temp_wb = openpyxl.load_workbook(f'{dirpath}/{file}',read_only=True)
                        temp_wb_sheets = set(temp_wb.sheetnames)
                        diff_sheets = required_sheets.difference(set(temp_wb_sheets))
                        if len(diff_sheets) != 0:
                            temp_error_df = pd.DataFrame(
                                data=[[f'{name_file}',
                                       f'Отсутствуют обязательные листы {diff_sheets}'
                                       ]],
                                columns=['Название файла',
                                         'Описание ошибки'])
                            error_df = pd.concat([error_df, temp_error_df], axis=0,
                                                 ignore_index=True)
                            continue


                        df = pd.read_excel(f'{dirpath}/{file}', dtype=str)  # открываем файл
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

        error_df.to_excel(f'{end_folder}/Ошибки_{current_time}.xlsx',index=False)

















if __name__ == '__main__':
    main_data_folder = 'data/Своды'
    main_end_folder = 'data/РЕЗУЛЬТАТ'
    processing_time_series(main_data_folder,main_end_folder)
    print('Lindy Booth')







