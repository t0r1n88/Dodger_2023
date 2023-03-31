import numpy as np
import pandas as pd
import warnings
warnings.simplefilter(action='ignore', category=UserWarning)

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


def check_error(df:pd.DataFrame,error_df:pd.DataFrame):
    """
    Функция для проверки правильности введеденных данных
    :param df: копия датафрейма с данными из файла поо
    :param error_df: датафрейм с ошибками
    :return:
    """
    # конвертируем в инт
    df = df.applymap(check_data)






# создаем словарь верхнего уровня для каждого поо
high_level_dct = {}
# создаем датафрейм для регистрации ошибок
error_df = pd.DataFrame(columns=['Название файла','Номер строки с ошибкой','Описание ошибки',])

df = pd.read_excel('data/poo/Брит - 2022.xlsx',skiprows=7,dtype=str)
df.columns = list(map(str, df.columns))
# Заполняем пока пропуски в 15 ячейке для каждой специальности
df['06'] = df['06'].fillna('15 ячейка')
# Проводим проверку на корректность данных, отправляем копию датафрейма
check_error(df.iloc[:,6:32], error_df)
