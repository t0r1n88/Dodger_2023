# -*- coding: utf-8 -*-
"""
Скрипт для обработки формы сбора данных по потребности предприятий ОПК
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

def prepare_form_opk_2024(path_folder_data:str,path_to_end_folder):
    """
    Функция для обработки таблиц с мониторингом потребности ОПК версия 2024
    """










if __name__ == '__main__':
    main_data_folder = 'data/example/Форма ОПК 2024'
    main_result_folder = 'data/result/Форма ОПК 2024'
    prepare_form_opk_2024(main_data_folder,main_result_folder)

    print('Lindy Booth')


