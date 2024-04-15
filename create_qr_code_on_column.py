"""
Скрипт для создания QR кодов из значений выбранной колонки, коды распределяются по папкам созданным
по значениям из другой выбранной колонки
"""
"""
Скрипт для создания Qr кодов по папкам
"""
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.exceptions import IllegalCharacterError
import json
import ast
import qrcode
import re
import os
from tkinter import messagebox
import datetime
import time
from dateutil import parser
pd.options.mode.chained_assignment = None


class NotFolderColumn(Exception):
    """
    Класс исключение для случая когда не найдена колонка по которой будут создавать папки
    """
    pass

class NotQrColumn(Exception):
    """
    Класс исключение для случая когда не найдена колонка по которой будут создаваться QR
    """
    pass


def generate_qr_code(data_file:str,name_folder_column:str,name_qr_column:str,end_folder:str):
    """
    Функция для генерации QR кодов из значений колонки
    """
    t = time.localtime()  # получаем текущее время и дату
    current_time = time.strftime('%H_%M_%S', t)
    current_date = time.strftime('%d_%m_%Y', t)
    df = pd.read_excel(data_file,dtype=str) # открываем файл
    df.columns = list(map(str,df.columns)) # делаем все названия колонок строковыми
    # Проверяем наличие колонок
    if name_folder_column not in df.columns:
        raise NotFolderColumn
    if name_qr_column not in df.columns:
        raise NotQrColumn

    # получаем все значения по которым будем создавать папки
    lst_name_folders = df[name_folder_column].unique()



















if __name__ == '__main__':
    main_file = 'data/Вакансии по региону от 17_40_15.xlsx'
    main_name_folder_column = 'Сфера деятельности'
    main_name_qr_column = 'Ссылка на вакансию'
    main_end_folder = 'data/QR'
    generate_qr_code(main_file,main_name_folder_column,main_name_qr_column,main_end_folder)
    print('Lindy Booth')

