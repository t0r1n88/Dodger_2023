"""
Скрипт для создания куар кодов для киоска с программой
"""
import pandas as pd
import numpy as np
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




def create_file_qr(data_file:str, end_folder:str):
    """
    Скрипт для генерации куар кодов
    """
    df = pd.read_excel(data_file,dtype=str)
    lst_sphere = df['Сфера деятельности'].unique()
    lst_sphere = [value for value in lst_sphere if pd.notna(value)]

    df = df[['Вакансия','Краткое название работодателя','Зарплата','Минимальная зарплата','Сфера деятельности','Контактный телефон','Email контактного лица','Ссылка на вакансию',]]


    for sphere in lst_sphere:
        temp_df = df[df['Сфера деятельности'] == sphere]
        temp_df.sort_values(by='Минимальная зарплата',ascending=False,inplace=True)
        out_df = temp_df.head(50)
        out_df.drop(columns=['Минимальная зарплата','Сфера деятельности'],inplace=True)
        out_df.to_csv(f'{end_folder}/{sphere}.csv',encoding='utf-8',sep='|',index=False,header=False)










if __name__ == '__main__':
    main_file_data = 'data/data.xlsx'
    main_end_folder = 'data/Результат'

    create_file_qr(main_file_data,main_end_folder)

    print('Lindy Booth')


