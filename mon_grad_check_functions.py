"""
Модуль для проверочных функций мониторинга занятости выпускников для сайта СССР
"""
from check_functions import extract_code_nose
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import pandas as pd



def create_check_tables_mon_grad(high_level_dct: dict):
    """
    Функция для создания файла с данными по каждой специальности чтобы можно было проверить правильность внесенных
    данных
    """
    # Создаем словарь в котором будут храниться словари по специальностям
    code_spec_dct = {}
    # инвертируем словарь так, чтобы код специальности стал внешним ключом а названия файлов внутренними
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
            code = extract_code_nose(code_spec)
            wb.create_sheet(title=code, index=idx)

    for code_spec in code_spec_dct.keys():
        if code_spec != 'nan':
            code = extract_code_nose(code_spec)
            # Конвертируем в датафрейм
            temp_code_df = pd.DataFrame.from_dict(code_spec_dct[code_spec], orient='index')
            # Удаляем колонки без названий
            temp_code_df.drop(columns=[name_column for name_column in temp_code_df.columns if 'Unnamed' in name_column],
                              inplace=True)

            temp_code_df = temp_code_df.reset_index()  # извлекаем название организации
            temp_code_df.rename(columns={'index': 'Наименование'},
                                inplace=True)  # переименовываем колонку с названием организации

            for r in dataframe_to_rows(temp_code_df, index=False, header=True):
                wb[code].append(r)
            wb[code].column_dimensions['A'].width = 20
            wb[code].column_dimensions['B'].width = 40

    # Удаляем лист Sheet
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']
    return wb