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

    used_name_sheet = set() # Множество для хранения названий листов
    # Создаем файл
    wb = openpyxl.Workbook()
    # Создаем листы
    for idx, code_spec in enumerate(code_spec_dct.keys()):
        if code_spec != 'nan':
            code = extract_code_nose(code_spec)
            # проверяем есть ли такой лист. На случай когда коды одинаковые а описание специальности разное
            if code not in used_name_sheet:
                wb.create_sheet(title=code, index=idx)
                used_name_sheet.add(code)

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


def check_first_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 2 = 3 + 31 + 32 + 60 + 61 + 62+ 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70.
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])

    check_sum_columns = ['3', '31', '32', '60', '61', '62', '63', '64', '65', '66', '67', '68', '69', '70']
    df['Сумма'] = df[check_sum_columns].sum(axis=1)
    df['Результат'] = df['2'] == df['Сумма']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]   # получаем значение проверки
    if df.iloc[0, -1] == 'Неправильно':
        # получаем значения из колонок сумма и проверочной колонки
        first_value = df['2'].tolist()[0]
        second_value = df['Сумма'].tolist()[0]

        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 2 = 3 + 31 + 32 + 60 + 61 + 62+ 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70.'
                                            f' Графа 2 = {first_value} ,сумма граф = {second_value}']])
        return temp_error_df

    return temp_error_df


def check_second_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 3 = сумма значений граф с 4 по 30
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])

    check_sum_columns = ['4', '5', '6', '7', '8', '9',
                         '10', '11', '12', '13', '14', '15',
                         '16', '17', '18', '19', '20', '21',
                         '22', '23', '24', '25', '26', '27', '28', '29', '30']
    df['Сумма'] = df[check_sum_columns].sum(axis=1)
    df['Результат'] = df['3'] == df['Сумма']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['3'].tolist()[0]
        second_value = df['Сумма'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 3 = сумма значений граф с 4 по 30. Графа 3 = {first_value}, сумма граф = {second_value}']])
        return temp_error_df

    return temp_error_df


def check_third_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 3 >= графы 3.1
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])
    df['Результат'] = df['3'] >= df['3.1']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['3'].tolist()[0]
        second_value = df['3.1'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 3 больше или равно графы 3.1. Графа 3 = {first_value}, графа 3.1 = {second_value}']])
        return temp_error_df

    return temp_error_df


def check_fourth_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 3 >= графы 3.2
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])
    df['Результат'] = df['3'] >= df['3.2']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['3'].tolist()[0]
        second_value = df['3.2'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 3 больше или равно графы 3.2. Графа 3 = {first_value}, графа 3.2 = {second_value}']])
        return temp_error_df

    return temp_error_df


def check_fifth_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 32 = сумма значений граф с 33 по 59.
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])

    check_sum_columns = ['33', '34', '35', '36', '37', '38',
                         '39', '40', '41', '42', '43', '44',
                         '45', '46', '47', '48', '49', '50',
                         '51', '52', '53', '54', '55', '56', '57', '58', '59']
    df['Сумма'] = df[check_sum_columns].sum(axis=1)
    df['Результат'] = df['32'] == df['Сумма']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['32'].tolist()[0]
        second_value = df['Сумма'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 32 = сумма значений граф с 33 по 59. Графа 32 = {first_value}, сумма граф = {second_value}']])
        return temp_error_df

    return temp_error_df



def check_six_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 32 >= графы 32.1
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])
    df['Результат'] = df['32'] >= df['32.1']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['32'].tolist()[0]
        second_value = df['32.1'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 32 больше или равно графы 32.1. Графа 32 = {first_value}, графа 32.1 = {second_value}']])
        return temp_error_df

    return temp_error_df

def check_seventh_error_grad(df: pd.DataFrame, name_file: str, number_row: int):
    """
    Функция для проверки условия Графа 32 >= графы 32.2
    """
    temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'])
    df['Результат'] = df['32'] >= df['32.2']
    df['Результат'] = df['Результат'].apply(lambda x: 'Правильно' if x else 'Неправильно')
    name_spec = df.iloc[0, 0]
    if df.iloc[0, -1] == 'Неправильно':
        first_value = df['32'].tolist()[0]
        second_value = df['32.2'].tolist()[0]
        temp_error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки'],
                                     data=[[name_file, f'Строка {number_row}- {name_spec}',
                                            f'Не выполняется условие: Графа 32 больше или равно графы 32.2. Графа 32 = {first_value}, графа 32.2 = {second_value}']])
        return temp_error_df

    return temp_error_df


def check_error_mon_grad_spo(df: pd.DataFrame, name_file: str,correction:int):
    """
    Точка входа для проверки датафрейма занятости выпускников на арифметические ошибки
    :param correction: число строк на которое нужно увеличить результат чтобы показывалась правильная строка у ошибки
    """
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    border = 0
    for i in range(1, len(df) + 1):
        row_df = df.iloc[border, :].to_frame().transpose()  # получаем датафрейм строку

        # Проводим проверку Графа 2 = 3 + 31 + 32 + 60 + 61 + 62+ 63 + 64 + 65 + 66 + 67 + 68 + 69 + 70.
        first_error_df_grad = check_first_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, first_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 3 = сумма значений граф с 4 по 30.
        second_error_df_grad = check_second_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, second_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 3.1 <= Графа 3
        third_error_df_grad = check_third_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, third_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 3.2 <= Графа 3
        fourth_error_df_grad = check_fourth_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, fourth_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 32 = сумма значений граф с 33 по 59
        fifth_error_df_grad = check_fifth_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, fifth_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 32 >= Графа 32.1
        six_error_df_grad = check_six_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, six_error_df_grad], axis=0, ignore_index=True)
        # Проводим проверку Графа 32 >= Графа 32.2
        seventh_error_df_grad = check_seventh_error_grad(row_df.copy(), name_file, correction+i)
        error_df = pd.concat([error_df, seventh_error_df_grad], axis=0, ignore_index=True)

        border += 1

    return error_df
