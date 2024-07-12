# -*- coding: utf-8 -*-
"""
Модуль для обработки таблиц мониторинга занятости выпускников используемого для загрузки на сайт СССР
"""
from check_functions import base_check_file

import pandas as pd
import numpy as np
import os
import warnings
from tkinter import messagebox
import time
pd.options.mode.chained_assignment = None  # default='warn'
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', category=FutureWarning)
import copy
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows


def prepare_graduate_employment(path_folder_data:str,result_folder:str):
    """
    Функция для обработки мониторинга занятости
    """
    # создаем словарь верхнего уровня для каждого поо
    high_level_dct = {}
    # создаем словарь верхнего уровня для хранения пары ключ значение где ключ это код специальности а значение- код и наименование
    dct_code_and_name = dict()
    # создаем датафрейм для регистрации ошибок
    error_df = pd.DataFrame(columns=['Название файла', 'Строка или колонка с ошибкой', 'Описание ошибки', ])
    check_required_dct = {'Выпуск-СПО':'В файле не найден лист с названием Выпуск-СПО', 'Выпуск-Целевое':'В файле не найден лист с названием Выпуск-Целевое',}

    try:
        for file in os.listdir(path_folder_data):
            print(file)
            error_df = base_check_file(file,error_df,path_folder_data,check_required_dct)

        print(error_df)
    except ZeroDivisionError:
        print('dssd')




if __name__ == '__main__':
    main_data_folder = 'data/Мониторинг занятости выпускников/Файлы'
    main_result_folder = 'data/Мониторинг занятости выпускников/Результат'
    prepare_graduate_employment(main_data_folder,main_result_folder)

    print('Lindy Booth')