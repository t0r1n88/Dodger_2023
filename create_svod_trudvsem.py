"""
Скрипт для обработки данных из файла с вакансиями с сайта Работа в России
"""
import pandas as pd
import numpy as np
from openpyxl.utils.exceptions import IllegalCharacterError
import json
import ast
import re
import os
from tkinter import messagebox
import datetime
import time
from dateutil import parser
pd.options.mode.chained_assignment = None
import warnings
warnings.filterwarnings("ignore")

class NotRegion(Exception):
    """
    Класс для отслеживания наличия региона в данных
    """
    pass

class MoreColumn(Exception):
    """
    Класс для отслеживания количества колонок, не более 3
    """
    pass


class NotColumn(Exception):
    """
    Класс для отслеживания наличия колонки указанной в параметрах в датафрейме
    """
    pass



def extract_data_from_list_cell(cell: str, lst_need_keys: list):
    """
    Функция для извлечения данных из словаря в ячейке датафрейма
    """
    try:

        lst_lang = ast.literal_eval(cell)  # превращаем в список
        if lst_lang:
            out_str_lst = []  # создаем список содержащий выходные строки
            for lang_dict in lst_lang:
                lst_lang_str = []  # список для хранения значений извлеченных из словаря
                for idx, key in enumerate(lst_need_keys):
                    lst_lang_str.append(lang_dict.get(key, None))

                lst_lang_str = [value for value in lst_lang_str if value]  # отбрасываем None
                single_lang_str = ','.join(lst_lang_str)  # создаем строку для одного языка
                out_str_lst.append(single_lang_str)
            return ';'.join(out_str_lst)

        else:
            return 'Не указано'


    except FileNotFoundError:
        return 'Не удалось обработать содержимое ячейки'


def clear_tag(cell):
    """
    Функция для очистки текста от тегов HTML
    """
    value = str(cell)
    if value != 'nan':
        result = re.sub(r'<.*?>', '', value)
        result = re.sub(r'&[a-z]*?;', '', result)
        return result
    else:
        return None


def clear_bonus_tag_br(cell):
    """
    Функция для очистки данных в колонке Бонусы
    """
    cell = str(cell)

    if cell != 'nan':
        value = str(cell)
        result = re.sub(r'<.*?>', '.', value)

        return result
    else:
        return None

def clean_text(cell):
    """
    Функция для очистки от незаписываемых символов
    """
    if isinstance(cell,str):
        return re.sub(r'[^\d\w ()=*+,.:;\"\'@-]','',cell)
    else:
        return cell

def clean_equal(cell:str):
    """
    Функция для очистки от знака равно в начале строки
    """
    if isinstance(cell,str):
        if cell.startswith('='):
            return f' {cell}'
        else:
            return cell

    else:
        return cell



def convert_date(cell):
    """
    Функция конвертации строки содержащей дату и время
    """
    value = str(cell)
    try:
        if value != 'nan':
            date_time = parser.parse(value).date()  # извлекаем дату
            date_time = datetime.datetime.strftime(date_time, '%d.%m.%Y')  # конфертируем в нужный формат
            return date_time

        else:
            return None
    except:
        return 'Не удалось обработать содержимое ячейки'

def convert_int(value):
    """
    Функция для конвертации в инт
    """
    try:
        return int(value)
    except:
        return 0

def extract_soc_category(df: pd.DataFrame, name_column: str, user_sep: str):
    """
    Функция для подсчета категорий в ячейке
    """
    # словарь для подсчета социальных категорий
    dct_value = {'беженцы': 0, 'лица, получившие временное убежище': 0, 'вынужденные переселенцы': 0, 'инвалиды': 0,
                 'лица, освобождаемые из мест лишения свободы': 0,
                 'матери и отцы, воспитывающие без супруга (супруги) детей в возрасте до пяти лет': 0,
                 'многодетные семьи': 0, 'несовершеннолетние работники': 0, 'работники, имеющие детей-инвалидов': 0,
                 'работники, осуществляющие уход за больными членами их семей в соответствии с медицинским заключением': 0}

    filter_value = df[name_column].notna()
    lst_soc_cat = df[filter_value][name_column].tolist()  # получаем список социальных категорий для вакансии
    lst_soc_cat = list(map(str.lower, lst_soc_cat))  # делаем буквы маленькими
    lst_soc_cat = list(map(str.strip, lst_soc_cat))  # очищаем от пробельных символов в начале и в конце
    # считаем
    for soc_cat in lst_soc_cat:
        for key in dct_value.keys():
            if key in soc_cat:
                dct_value[key] += 1

    return dct_value


def extract_id_company(cell):
    """
    Для извления айди компании
    """
    if isinstance(cell,str):
        lst_org = cell.split('/')
        return lst_org[-1]

def extract_municipality(cell):
    """
    Функция для извлечения муниципалитета, где расположена вакансия
    """
    # Если значение ячейки строковое
    if isinstance(cell, str):
        lst_value = cell.split(',')
        # Проверяем на длину
        if len(lst_value) >= 2:
            name_municipality = lst_value[1].strip()
            # проверяем на наличие слов город и район
            if 'город' not in name_municipality.lower() and 'район' not in name_municipality.lower():
                return 'Не определен'
            name_municipality = re.sub('\d','',name_municipality).strip() # очищаем от цифр
            return name_municipality
        else:
            return 'Не определен'

def extract_salary(cell):
    """
    Функция для извлечения значения зарплаты
    """
    if isinstance(cell,str):
        value = cell.replace('от ','')
        try:
            return int(value)
        except:
            result = re.search(r'\d+',value)
            if result:
                return result.group(0)
            else:
                return 0
    else:
        return cell


def extract_phone_number(value):
    """
    Фунция для извлечения номера телефона контактного лица
    """
    try:
        if value:
            data = json.loads(value)
            for item in data:
                if item.get('contactType') == 'Телефон':
                    phone_number = item.get('contactValue','Не указан')
                    break
            if phone_number:
                return phone_number
            else:
                return None
        else:
            return None
    except:
        return None

def extract_contact_email(value):
    """
    Фунция для извлечения email контактного лица
    """
    try:
        if value:
            data = json.loads(value)
            for item in data:
                if item.get('contactType') == 'Эл. почта':
                    email = item.get('contactValue','Не указан')
                    break
            if email:
                return email
            else:
                return None
        else:
            return None
    except:
        return None



def filtred_df(df:pd.DataFrame,params_filter:str):
    """
    Функция для фильтрации датафрейма по указанным значениям в колонке
    """
    params_df = pd.read_excel(params_filter,dtype=str) # считываем параметры фильтрации
    data_cols = [value for value in params_df.columns if 'Unnamed' not in value]
    params_df = params_df[data_cols]
    # проверяем длину колонок не более 3
    if len(params_df.columns) > 3:
        raise MoreColumn

    # проверяем наличие колонок в датафрейме
    diff_cols = set(params_df.columns).difference(set(df.columns))
    if len(diff_cols) != 0:
        raise NotColumn

    # Обрабатываем в зависимости от количества колонок
    if len(params_df.columns) == 1:
        name_filter_column = params_df.columns[0]
        lst_filter_values = params_df[name_filter_column].tolist() # делаем список значений
        lst_filter_values = [value for value in lst_filter_values if not pd.isna(value)] # очищаем от нанов
        lst_filter_values = list(map(str,lst_filter_values)) # делаем строковыми значения
        df[name_filter_column] = df[name_filter_column].astype(str) # делаем строковой колонку
        df = df[df[name_filter_column].str.contains('|'.join(lst_filter_values),case=False)]# фильтруем

        return df
    if len(params_df.columns) == 2:
        name_first_filter_column = params_df.columns[0] # первый фильтр
        name_second_filter_column = params_df.columns[1] # второй фильтр
        lst_filter_values = params_df[name_first_filter_column].tolist() # делаем список значений
        lst_filter_values = [value for value in lst_filter_values if not pd.isna(value)]  # очищаем от нанов
        lst_second_filter_values = params_df[name_second_filter_column].tolist() # делаем список значений второго фильтра
        lst_second_filter_values = [value for value in lst_second_filter_values if not pd.isna(value)]  # очищаем от нанов

        # первая фильтрация
        lst_filter_values = list(map(str,lst_filter_values)) # делаем строковыми значения
        df[name_first_filter_column] = df[name_first_filter_column].astype(str) # делаем строковой колонку

        first_df = df[df[name_first_filter_column].str.contains('|'.join(lst_filter_values),case=False)]# фильтруем
        # вторая фильтрация
        lst_second_filter_values = list(map(str,lst_second_filter_values)) # делаем строковыми значения
        first_df[name_second_filter_column] = first_df[name_second_filter_column].astype(str) # делаем строковой колонку

        second_df = first_df[first_df[name_second_filter_column].str.contains('|'.join(lst_second_filter_values),case=False)]# фильтруем

        return second_df

    if len(params_df.columns) == 3:
        name_first_filter_column = params_df.columns[0] # первый фильтр
        name_second_filter_column = params_df.columns[1] # второй фильтр
        name_third_filter_column = params_df.columns[2] # третий фильтр

        lst_filter_values = params_df[name_first_filter_column].tolist() # делаем список значений
        lst_filter_values = [value for value in lst_filter_values if not pd.isna(value)]  # очищаем от нанов
        lst_second_filter_values = params_df[name_second_filter_column].tolist() # делаем список значений второго фильтра
        lst_second_filter_values = [value for value in lst_second_filter_values if not pd.isna(value)]  # очищаем от нанов
        lst_third_filter_values = params_df[name_third_filter_column].tolist() # делаем список значений третьего фильтра
        lst_third_filter_values = [value for value in lst_third_filter_values if not pd.isna(value)]  # очищаем от нанов

        # первая фильтрация
        lst_filter_values = list(map(str,lst_filter_values)) # делаем строковыми значения
        df[name_first_filter_column] = df[name_first_filter_column].astype(str) # делаем строковой колонку

        first_df = df[df[name_first_filter_column].str.contains('|'.join(lst_filter_values),case=False)]# фильтруем
        # вторая фильтрация
        lst_second_filter_values = list(map(str,lst_second_filter_values)) # делаем строковыми значения
        first_df[name_second_filter_column] = first_df[name_second_filter_column].astype(str) # делаем строковой колонку

        second_df = first_df[first_df[name_second_filter_column].str.contains('|'.join(lst_second_filter_values),case=False)]# фильтруем

        # третья фильтрация
        lst_third_filter_values = list(map(str,lst_third_filter_values)) # делаем строковыми значения
        second_df[name_third_filter_column] = second_df[name_third_filter_column].astype(str) # делаем строковой колонку

        third_df = second_df[second_df[name_third_filter_column].str.contains('|'.join(lst_third_filter_values),case=False)]# фильтруем

        return third_df









def prepare_data_vacancy(df: pd.DataFrame, dct_name_columns: dict, lst_columns: list) -> pd.DataFrame:
    """
    Функция для обработки датафрейма с данными работы в России
    """
    dct_status_accommodationType = {'DORMITORY':'Общежитие','FLAT':'Квартира',
                      'HOUSE':'Дом','ROOM':'Комната'}

    dct_status_companyBusinessSize = {'SMALL':'Малая','MIDDLE':'Средняя',
                      'MICRO':'Микро'}

    # Словарь для аббревиатур
    dct_abbr = {'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕТСТВЕННОСТЬЮ':'ООО','КРАЕВОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ УЧРЕЖДЕНИЕ':'КГАУ',
                'КРАЕВОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'КГБУ','ОТКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО':'ОАО',
                'ИНДИВИДУАЛЬНЫЙ ПРЕДПРИНИМАТЕЛЬ':'ИП','МУНИЦИПАЛЬНОЕ УНИТАРНОЕ ПРЕДПРИЯТИЕ':'МУП',
                'ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ УЧРЕЖДЕНИЕ':'ГАУ','МУНИЦИПАЛЬНОЕ АВТОНОМНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МАОУ',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'ФГБУ',
                'ОБЛАСТНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'ОГКУ','УПРАВЛЕНИЕ ИМУЩЕСТВЕННЫХ ОТНОШЕНИЙ':'УИО',
                'МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ':'МБУДО','ФЕДЕРАЛЬНОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'ФБУ',
                'ЗАКРЫТОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО':'ЗАО','МУНИЦИПАЛЬНОЕ АВТОНОМНОЕ ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МАДОУ',
                'МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МБОУ','ГЛАВНОЕ УПРАВЛЕНИЕ ФЕДЕРАЛЬНОЙ СЛУЖБЫ СУДЕБНЫХ ПРИСТАВОВ':'ГУ ФССП',
                'ПУБЛИЧНОЕ АКЦИОНЕРНОЕ ОБЩЕСТВО':'ПАО',
                'ФИЛИАЛ АКЦИОНЕРНОГО ОБЩЕСТВА':'ФИЛИАЛ АО',
                'МУНИЦИПАЛЬНОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'МКУ','ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ УНИТАРНОЕ ПРЕДПРИЯТИЕ':'ФГУП',
                'МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МБДОУ',
                'МУНИЦИПАЛЬНОЕ АВТОНОМНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ':'МАОУ ДО',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ФГБДОУ',
                'МУНИЦИПАЛЬНОЕ КАЗЁННОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МКОУ','СЕЛЬСКОХОЗЯЙСТВЕННЫЙ ПРОИЗВОДСТВЕННЫЙ КООПЕРАТИВ':'СПК',
                'ОБЛАСТНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ УЧРЕЖДЕНИЕ':'ОГАУ','ГЛАВНОЕ УПРАВЛЕНИЕ':'ГУ','ОБЛАСТНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'ОГБУ',
                'МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'МБУ','ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'ФГКУ',
                'ГОСУДАРСТВЕННОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'ГКУ',
                'АВТОНОМНОЕ СТАЦИОНАРНОЕ УЧРЕЖДЕНИЕ':'АСУ','МУНИЦИПАЛЬНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МОУ',
                'МУНИЦИПАЛЬНОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МБОУ',
                'МУНИЦИПАЛЬНОЕ АВТОНОМНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МАОУ',
                'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ГБПОУ','БЮДЖЕТНОЕ ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'БДОУ',
                'ФЕДЕРАЛЬНОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'ФКУ','ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ':'ФГАОУ ВО',
                'ОБЛАСТНОЕ ГОСУДАРСТВЕННОЕ КАЗЁННОЕ УЧРЕЖДЕНИЕ':'ОГКУ','АВТОНОМНАЯ НЕКОММЕРЧЕСКАЯ ПРОФЕССИОНАЛЬНАЯ ОБРАЗОВАТЕЛЬНАЯ ОРГАНИЗАЦИЯ':'АНПОО',
                'ПРОИЗВОДСТВЕННЫЙ КООПЕРАТИВ':'ПК','ГОСУДАРСТВЕННОЕ УЧРЕЖДЕНИЕ ЗДРАВООХРАНЕНИЯ':'ГУЗ',
                'ФЕДЕРАЛЬНОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ ЗДРАВООХРАНЕНИЯ':'ФБУЗ',
                'БЮДЖЕТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'БПОУ','ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ГАПОУ',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ':'ФГБОУ ВО','МУНИЦИПАЛЬНОЕ КАЗЕННОЕ ПРЕДПРИЯТИ':'МКУ',
                'МУНИЦИПАЛЬНОЕ КАЗЁННОЕ УЧРЕЖДЕНИЕ':'МКУ',
                'ФИЛИАЛ ФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО БЮДЖЕТНОГО ОБРАЗОВАТЕЛЬНОГО УЧРЕЖДЕНИЯ ВЫСШЕГО ОБРАЗОВАНИЯ':'ФИЛИАЛ ФГБОУ ВО',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ ПРЕДПРИЯТИЕ':'ФГП',
                'АКЦИОНЕРНОЕ ОБЩЕСТВО': 'АО',
                'ГОСУДАРСТВЕННОЕ СПЕЦИАЛЬНОЕ УЧЕБНО-ВОСПИТАТЕЛЬНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ГСУВ ОО',
                'АВТОНОМНОЕ УЧРЕЖДЕНИЕ РЕСПУБЛИКИ':'АУР','Войсковая часть':'ВЧ',
                'ФИЛИАЛ ФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО БЮДЖЕТНОГО УЧРЕЖДЕНИЯ':'ФИЛИАЛ ФГБУ','МУНИЦИПАЛЬНОЕ УЧРЕЖДЕНИЕ':'МУ',
                'МУНИЦИПАЛЬНОГО ОБРАЗОВАНИЯ':'МО','РЕСПУБЛИКАНСКОЕ ГОСУДАРСТВЕННОЕ УЧРЕЖДЕНИЕ':'РГУ',
                'МУНИЦИПАЛЬНОЕ АВТОНОМНОЕ УЧРЕЖДЕНИЕ':'МАУ',
                'АВТОНОМНОЕ УЧРЕЖДЕНИЕ СОЦИАЛЬНОГО ОБСЛУЖИВАНИЯ':'АУ СО',
                'АВТОНОМНОЕ УЧРЕЖДЕНИЕ КУЛЬТУРЫ': 'АУ КУЛЬТУРЫ',
                'ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ ЧАСТНОЕ УЧРЕЖДЕНИЕ':'ПОЧУ',
                'МУНИЦИПАЛЬНОЕ ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МДОУ','ДОШКОЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ДОУ',
                'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ГБОУ','ЧАСТНОЕ УЧРЕЖДЕНИЕ ЗДРАВООХРАНЕНИЯ':'ЧУЗ',
                'МУНИЦИПАЛЬНОГО РАЙОНА':'МР','ТЕРРИТОРИАЛЬНОЕ УПРАВЛЕНИЕ':'ТУ',
                'ФИЛИАЛ ПУБЛИЧНОГО АКЦИОНЕРНОГО ОБЩЕСТВА':'ФИЛИАЛ ПАО','ВОЙСКОВАЯ ЧАСТЬ':'ВЧ',
                'ФИЛИАЛ ФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО АВТОНОМНОГО УЧРЕЖДЕНИЯ':'ФИЛИАЛ ФГАУ',
                'ФИЛИАЛ ФЕДЕРАЛЬНОГО ГОСУДАРСТВЕННОГО УНИТАРНОГО ПРЕДПРИЯТИЯ':'ФИЛИАЛ ФГУП',
                'СРЕДНЯЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ ШКОЛА':'СОШ',
                'ФИЛИАЛ ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'ФИЛИАЛ ФГБУ',
                'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ СОЦИАЛЬНОГО ОБСЛУЖИВАНИЯ':'ГБУ СО',
                'МИНИСТЕРСТВА ВНУТРЕННИХ ДЕЛ РОССИЙСКОЙ ФЕДЕРАЦИИ':'МВД РФ',
                'АВТОНОМНОЕ УЧРЕЖДЕНИЕ': 'АУ',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ АВТОНОМНОЕ': 'ФГА', 'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ': 'ГБУ',
                'БЮДЖЕТНОЕ УЧРЕЖДЕНИЕ':'БУ','ДОПОЛНИТЕЛЬНОГО ОБРАЗОВАНИЯ':'ДО',
                'АВТОНОМНАЯ НЕКОММЕРЧЕСКАЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ ОРГАНИЗАЦИЯ':'АНОО',
                'АВТОНОМНАЯ НЕКОММЕРЧЕСКАЯ ОРГАНИЗАЦИЯ':'АНО',
                'ГОСУДАРСТВЕННОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ КАЗЕННОЕ УЧРЕЖДЕНИЕ':'ГОКУ',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ КАЗЕННОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ФГКОУ',
                'ВЫСШЕГО ОБРАЗОВАНИЯ':'ВО',
                'ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ПОУ',
                'ЧАСТНОЕ ПРОФЕССИОНАЛЬНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ЧПОУ',
                'ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ НАУЧНОЕ УЧРЕЖДЕНИЕ':'ФГБНУ',
                'МУНИЦИПАЛЬНОЕ КАЗЁННОЕ':'МК',
                'СЕЛЬСКОЕ ПОТРЕБИТЕЛЬСКОЕ ОБЩЕСТВО':'СПО',
                'ЧАСТНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ЧОУ',
                'МУНИЦИПАЛЬНОЕ КАЗЕННОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'МКОУ',
                'ОСНОВНАЯ ОБЩЕОБРАЗОВАТЕЛЬНАЯ ШКОЛА':'ООШ',
                'ДЕТСКАЯ ШКОЛА ИСКУССТВ':'ДШИ',
                'ДЕТСКИЙ ОЗДОРОВИТЕЛЬНО-ОБРАЗОВАТЕЛЬНЫЙ ЦЕНТР':'ДООЦ',
                'ЦЕНТР ЗАНЯТОСТИ НАСЕЛЕНИЯ':'ЦЗН',
                'ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОЗДОРОВИТЕЛЬНОЕ ОБЩЕОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ':'ГБООУ',
                'МУНИЦИПАЛЬНАЯ БЮДЖЕТНАЯ ДОШКОЛЬНАЯ ОБРАЗОВАТЕЛЬНАЯ ОРГАНИЗАЦИЯ':'МБДОУ',
                'ОБЩЕСТВО С ОГРАНИЧЕННОЙ ОТВЕСТВЕННОСТЬЮ':'ООО',
                }


    df = df[dct_name_columns.keys()] # отбираем только нужные колонки
    df.rename(columns=dct_name_columns, inplace=True,errors='ignore')
    # Обрабатываем обычные колонки
    df['Сфера деятельности'] = df['Сфера деятельности'].fillna('Не указана сфера деятельности')
    df['График работы'] = df['График работы'].fillna('Не указан')
    df['Тип занятости'] = df['Тип занятости'].fillna('Не указан')
    df['Бонусы'] = df['Бонусы'].apply(clear_bonus_tag_br)
    df['Требования'] = df['Требования'].apply(clear_tag)
    df['Муниципалитет'] = df['Адрес вакансии'].apply(extract_municipality)
    # Создаем краткие наименование работодателей создавая аббревиатуры
    df['Полное название работодателя'] = df['Полное название работодателя'].fillna('Не заполнено название организации')
    df['Полное название работодателя'] = df['Полное название работодателя'].astype(str)
    # заменяем несколько пробелов на один
    df['Полное название работодателя'] = df['Полное название работодателя'].apply(lambda x:re.sub(r'\s+',' ',x))

    df['Краткое название работодателя'] = df['Полное название работодателя'].apply(lambda x:x.upper() if isinstance(x,str) else x).replace(dct_abbr,regex=True)

    # Числовые
    lst_number_columns = ['Требуемый опыт работы в годах','Количество рабочих мест']
    df[lst_number_columns] = df[lst_number_columns].fillna(0)
    df[lst_number_columns] = df[lst_number_columns].astype(int, errors='ignore')

    # Создаем числовую колонку для минимальной зарплаты извлекая цифры оттуда
    df['Минимальная зарплата'] = df['Зарплата'].fillna(0)
    df['Минимальная зарплата'] = df['Минимальная зарплата'].apply(extract_salary)


    # Создаем колонку с категориями минимальной зарплаты
    category = ['До 25 тысяч','25-50 тысяч','50-75 тысяч','75-100 тысяч','Свыше 100 тысяч']
    df['Категория минимальной зарплаты'] = pd.cut(df['Минимальная зарплата'],bins=[-1,24999,49999,74999,99999,float('inf')],
                                                  labels=category)

    # Временные

    df['Дата размещения вакансии'] = df['Дата размещения вакансии'].apply(convert_date)
    df['Дата изменения вакансии'] = df['Дата изменения вакансии'].apply(convert_date)
    # Категориальные
    df['Квотируемое место'] = df['Квотируемое место'].astype(str).apply(lambda x: 'Квотируемое место' if x == 'True' else None)


    df['Жилье от организации'] = df['Жилье от организации'].astype(str).apply(lambda x: 'Предоставляется жилье' if x == 'True' else None)
    df['Тип жилья'] = df['Тип жилья'].replace(dct_status_accommodationType)

    df['Размер организации'] = df['Размер организации'].replace(dct_status_companyBusinessSize)


    df['Для иностранцев'] = df['Для иностранцев'].astype(str).apply(lambda x: 'Возможно трудоустройство иностранных граждан' if x == 'True' else None)
    df['Программа трудовой мобильности'] = df['Программа трудовой мобильности'].astype(str).apply(lambda x: 'Вакансия по программе трудовой мобильности' if x == 'True' else None)
    df['Возможность переподготовки'] = df['Возможность переподготовки'].astype(str).apply(lambda x: 'Да' if x == 'True' else None)


    # Начинаем извлекать данные из сложных колонок с json
    # Данные по образованию
    df['Образование'] = df['Данные по образованию'].apply(lambda x: json.loads(x).get('educationType', 'Не указано'))
    df['Требуемая специализация'] = df['Данные по образованию'].apply(lambda x: json.loads(x).get('speciality', 'Не указано'))

    # Данные по широте и долготе
    df['Широта адрес вакансии'] = df['Геоданные'].apply(lambda x: json.loads(x).get('latitude', 'Не указано'))
    df['Долгота адрес вакансии'] = df['Геоданные'].apply(lambda x: json.loads(x).get('longitude', 'Не указано'))


    # данные по работодателю
    df['КПП работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('kpp', 'Не указано'))
    df['ОГРН работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('ogrn', 'Не указано'))
    df['ИНН работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('inn', 'Не указано'))
    df['Контактный телефон'] = df['Контактные данные'].apply(extract_phone_number)
    df['Email контактного лица'] = df['Контактные данные'].apply(extract_contact_email)

    df['Email работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('email', 'Не указано'))
    df['Профиль работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('url', 'Не указано'))
    df['Сайт работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('site', 'Не указано'))
    df['ID работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('companyCode', 'Не указано'))


    # Обрабатываем колонку с языками
    df['Требуемые языки'] = df['Данные по языкам'].apply(
        lambda x: extract_data_from_list_cell(x, ['code_language', 'level']))
    df['Требуемые хардскиллы'] = df['Данные по хардскиллам'].apply(
        lambda x: ','.join(ast.literal_eval(x)))
    df['Требуемые софтскиллы'] = df['Данные по софтскиллам'].apply(
        lambda x: ','.join(ast.literal_eval(x)))

    df.drop(columns=['Данные компании', 'Данные по языкам', 'Данные по хардскиллам', 'Данные по софтскиллам','Данные по образованию',
                     'Геоданные','Контактные данные'],
            inplace=True)

    df = df.reindex(columns=lst_columns)


    return df


def create_svod_for_df(prepared_df:pd.DataFrame,svod_region_folder:str,name_file:str,current_time):
    """
    Функция для создания аналитики по одинаковым по структуре датафреймам данных из Работы в России
    """
    svod_vac_reg_region_df = pd.pivot_table(prepared_df,
                                            index=['Сфера деятельности'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]})
    svod_vac_reg_region_df = svod_vac_reg_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_vac_reg_region_df) != 0:
        svod_vac_reg_region_df.sort_values(by=['sum'], ascending=False, inplace=True)
        svod_vac_reg_region_df.loc['Итого'] = svod_vac_reg_region_df['sum'].sum()
        svod_vac_reg_region_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
        svod_vac_reg_region_df = svod_vac_reg_region_df.reset_index()

    # Свод по муниципалитам
    svod_vac_mun_region_df = pd.pivot_table(prepared_df,
                                            index=['Муниципалитет'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]})
    svod_vac_mun_region_df = svod_vac_mun_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_vac_mun_region_df) != 0:
        svod_vac_mun_region_df.sort_values(by=['sum'], ascending=False, inplace=True)
        svod_vac_mun_region_df.loc['Итого'] = svod_vac_mun_region_df['sum'].sum()
        svod_vac_mun_region_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
        svod_vac_mun_region_df = svod_vac_mun_region_df.reset_index()

    # Свод по муниципалитетам и отраслям
    svod_vac_mun_sphere_region_df = pd.pivot_table(prepared_df,
                                            index=['Муниципалитет','Сфера деятельности'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]})
    svod_vac_mun_sphere_region_df = svod_vac_mun_sphere_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_vac_mun_sphere_region_df) != 0:
        svod_vac_mun_sphere_region_df = svod_vac_mun_sphere_region_df.reset_index()
        svod_vac_mun_sphere_region_df.rename(columns={'sum':'Количество вакансий'},inplace=True)

    # Свод по отраслям и муниципалитетам
    svod_vac_sphere_mun_region_df = pd.pivot_table(prepared_df,
                                            index=['Сфера деятельности','Муниципалитет'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]})
    svod_vac_sphere_mun_region_df = svod_vac_sphere_mun_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_vac_sphere_mun_region_df) != 0:
        svod_vac_sphere_mun_region_df = svod_vac_sphere_mun_region_df.reset_index()
        svod_vac_sphere_mun_region_df.rename(columns={'sum':'Количество вакансий'},inplace=True)




    # Свод по количеству рабочих мест по организациям
    svod_vac_org_region_df = pd.pivot_table(prepared_df,
                                            index=['Краткое название работодателя'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]})
    svod_vac_org_region_df = svod_vac_org_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_vac_org_region_df) != 0:
        svod_vac_org_region_df.sort_values(by=['sum'], ascending=False, inplace=True)
        svod_vac_org_region_df.loc['Итого'] = svod_vac_org_region_df['sum'].sum()
        svod_vac_org_region_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
        svod_vac_org_region_df = svod_vac_org_region_df.reset_index()

    # Список вакансий для последующего отслеживания динамики
    svod_vac_particular_org_region_df = prepared_df[
        ['Краткое название работодателя', 'Вакансия', 'Количество рабочих мест', 'ID вакансии', 'Ссылка на вакансию']]

    # Своды по минимальной зарплате
    prepared_df['Минимальная зарплата'] = prepared_df['Минимальная зарплата'].apply(convert_int)
    pay_df = prepared_df[prepared_df['Минимальная зарплата'] > 0]  # отбираем все вакансии с зп больше нуля

    # Свод по категориям минимальной заработной платы для сфеф деятельности
    svod_shpere_category_pay_region_df = pd.pivot_table(pay_df,
                                                        index=['Сфера деятельности', 'Категория минимальной зарплаты'],
                                                        values=['Количество рабочих мест'],
                                                        aggfunc={'Количество рабочих мест': 'sum'}
                                                        )

    if len(svod_shpere_category_pay_region_df) != 0:
        svod_shpere_category_pay_region_df = svod_shpere_category_pay_region_df.reset_index()
        svod_shpere_category_pay_region_df.rename(columns={'Количество рабочих мест': 'Количество вакансий'},
                                                  inplace=True)

    # Свод по категориям минимальной заработной платы для работодателей
    svod_org_category_pay_region_df = pd.pivot_table(pay_df,
                                                     index=['Краткое название работодателя',
                                                            'Категория минимальной зарплаты'],
                                                     values=['Количество рабочих мест'],
                                                     aggfunc={'Количество рабочих мест': 'sum'}
                                                     )

    if len(svod_org_category_pay_region_df) != 0:
        svod_org_category_pay_region_df = svod_org_category_pay_region_df.reset_index()
        svod_org_category_pay_region_df.rename(columns={'Количество рабочих мест': 'Количество вакансий'}, inplace=True)

    # Средняя и медианная зарплата по сфере деятельности
    svod_shpere_pay_region_df = pd.pivot_table(pay_df,
                                               index=['Сфера деятельности'],
                                               values=['Минимальная зарплата'],
                                               aggfunc={'Минимальная зарплата': [np.mean, np.median]}
                                               )
    svod_shpere_pay_region_df = svod_shpere_pay_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_shpere_pay_region_df) != 0:
        svod_shpere_pay_region_df = svod_shpere_pay_region_df.astype(int, errors='ignore')
        svod_shpere_pay_region_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
        svod_shpere_pay_region_df = svod_shpere_pay_region_df.reset_index()

    # Свод по средней и медианной минимальной зарплате для работодателей
    svod_org_pay_region_df = pd.pivot_table(pay_df,
                                            index=['Краткое название работодателя', 'Сфера деятельности'],
                                            values=['Минимальная зарплата'],
                                            aggfunc={'Минимальная зарплата': [np.mean, np.median]}
                                            )
    svod_org_pay_region_df = svod_org_pay_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_pay_region_df) != 0:
        svod_org_pay_region_df = svod_org_pay_region_df.astype(int, errors='ignore')
        svod_org_pay_region_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
        svod_org_pay_region_df = svod_org_pay_region_df.reset_index()

    # Свод по требуемому образованию для сфер деятельности
    svod_shpere_educ_region_df = pd.pivot_table(prepared_df,
                                                index=['Сфера деятельности', 'Образование'],
                                                values=['Количество рабочих мест'],
                                                aggfunc={'Количество рабочих мест': [np.sum]},
                                                fill_value=0)
    svod_shpere_educ_region_df = svod_shpere_educ_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_shpere_educ_region_df) != 0:
        svod_shpere_educ_region_df = svod_shpere_educ_region_df.astype(int, errors='ignore')
        svod_shpere_educ_region_df.columns = ['Количество вакансий']
        svod_shpere_educ_region_df = svod_shpere_educ_region_df.reset_index()

    # Свод по требуемому образованию для работодателей
    svod_org_educ_region_df = pd.pivot_table(prepared_df,
                                             index=['Краткое название работодателя', 'Образование'],
                                             values=['Количество рабочих мест'],
                                             aggfunc={'Количество рабочих мест': [np.sum]},
                                             fill_value=0)
    svod_org_educ_region_df = svod_org_educ_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_educ_region_df) != 0:
        svod_org_educ_region_df = svod_org_educ_region_df.astype(int, errors='ignore')
        svod_org_educ_region_df.columns = ['Количество вакансий']
        svod_org_educ_region_df = svod_org_educ_region_df.reset_index()

    # Свод по графику работы для сфер деятельности
    svod_shpere_schedule_region_df = pd.pivot_table(prepared_df,
                                                    index=['Сфера деятельности', 'График работы'],
                                                    values=['Количество рабочих мест'],
                                                    aggfunc={'Количество рабочих мест': [np.sum]},
                                                    fill_value=0)
    svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.droplevel(level=0,
                                                                              axis=1)  # убираем мультииндекс
    if len(svod_shpere_schedule_region_df) != 0:
        svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.astype(int, errors='ignore')
        svod_shpere_schedule_region_df.columns = ['Количество вакансий']
        svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.reset_index()

    # Свод по графику работы для работодателей
    svod_org_schedule_region_df = pd.pivot_table(prepared_df,
                                                 index=['Краткое название работодателя', 'График работы'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)
    svod_org_schedule_region_df = svod_org_schedule_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_schedule_region_df) != 0:
        svod_org_schedule_region_df = svod_org_schedule_region_df.astype(int, errors='ignore')
        svod_org_schedule_region_df.columns = ['Количество вакансий']
        svod_org_schedule_region_df = svod_org_schedule_region_df.reset_index()

    # Свод по типу занятости для сфер деятельности
    svod_shpere_type_job_region_df = pd.pivot_table(prepared_df,
                                                    index=['Сфера деятельности', 'Тип занятости'],
                                                    values=['Количество рабочих мест'],
                                                    aggfunc={'Количество рабочих мест': [np.sum]},
                                                    fill_value=0)
    svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.droplevel(level=0,
                                                                              axis=1)  # убираем мультииндекс
    if len(svod_shpere_type_job_region_df) != 0:
        svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.astype(int, errors='ignore')
        svod_shpere_type_job_region_df.columns = ['Количество вакансий']
        svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.reset_index()

    # Свод по типу занятости для работодателей
    svod_org_type_job_region_df = pd.pivot_table(prepared_df,
                                                 index=['Краткое название работодателя', 'Тип занятости'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)
    svod_org_type_job_region_df = svod_org_type_job_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_type_job_region_df) != 0:
        svod_org_type_job_region_df = svod_org_type_job_region_df.astype(int, errors='ignore')
        svod_org_type_job_region_df.columns = ['Количество вакансий']
        svod_org_type_job_region_df = svod_org_type_job_region_df.reset_index()

    # Свод по квотируемым местам для сфер деятельности
    svod_shpere_quote_region_df = pd.pivot_table(prepared_df,
                                                 index=['Сфера деятельности', 'Квотируемое место'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)

    svod_shpere_quote_region_df = svod_shpere_quote_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_shpere_quote_region_df) != 0:
        svod_shpere_quote_region_df = svod_shpere_quote_region_df.astype(int, errors='ignore')
        svod_shpere_quote_region_df.columns = ['Количество вакансий']
    svod_shpere_quote_region_df = svod_shpere_quote_region_df.reset_index()
    # Свод по квотируемым местам для работодателей
    svod_org_quote_region_df = pd.pivot_table(prepared_df,
                                              index=['Краткое название работодателя', 'Квотируемое место'],
                                              values=['Количество рабочих мест'],
                                              aggfunc={'Количество рабочих мест': [np.sum]},
                                              fill_value=0)
    svod_org_quote_region_df = svod_org_quote_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_quote_region_df) != 0:
        svod_org_quote_region_df = svod_org_quote_region_df.astype(int, errors='ignore')
        svod_org_quote_region_df.columns = ['Количество вакансий']
        svod_org_quote_region_df = svod_org_quote_region_df.reset_index()

    # Свод по вакансиям для социальных категорий
    dct_soc_cat_region = extract_soc_category(prepared_df, 'Социально защищенная категория',
                                              ';')  # считаем количество категорий
    svod_soc_region_df = pd.DataFrame.from_dict(dct_soc_cat_region, orient='index').reset_index()  # содаем датафрейм
    if len(svod_soc_region_df) != 0:
        svod_soc_region_df.columns = ['Категория', 'Количество вакансий']
        svod_soc_region_df.sort_values(by=['Количество вакансий'], ascending=False, inplace=True)

    # Свод по требуемому опыту для сфер деятельности
    svod_shpere_exp_region_df = pd.pivot_table(prepared_df,
                                               index=['Сфера деятельности', 'Требуемый опыт работы в годах'],
                                               values=['Количество рабочих мест'],
                                               aggfunc={'Количество рабочих мест': [np.sum]},
                                               fill_value=0)
    svod_shpere_exp_region_df = svod_shpere_exp_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_shpere_exp_region_df) != 0:
        svod_shpere_exp_region_df = svod_shpere_exp_region_df.astype(int, errors='ignore')
        svod_shpere_exp_region_df.columns = ['Количество вакансий']
        svod_shpere_exp_region_df = svod_shpere_exp_region_df.reset_index()

    # Свод по требуемому опыту для работодателей
    svod_org_exp_region_df = pd.pivot_table(prepared_df,
                                            index=['Краткое название работодателя', 'Требуемый опыт работы в годах'],
                                            values=['Количество рабочих мест'],
                                            aggfunc={'Количество рабочих мест': [np.sum]},
                                            fill_value=0)
    svod_org_exp_region_df = svod_org_exp_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
    if len(svod_org_exp_region_df) != 0:
        svod_org_exp_region_df = svod_org_exp_region_df.astype(int, errors='ignore')
        svod_org_exp_region_df.columns = ['Количество вакансий']
        svod_org_exp_region_df = svod_org_exp_region_df.reset_index()

    with pd.ExcelWriter(f'{svod_region_folder}/{name_file} от {current_time}.xlsx') as writer:
        svod_vac_reg_region_df.to_excel(writer, sheet_name='Вакансии по отраслям', index=False)
        svod_vac_mun_region_df.to_excel(writer, sheet_name='Вакансии по муниципалитетам', index=False)
        svod_vac_mun_sphere_region_df.to_excel(writer, sheet_name='Муниципалитеты отрасли', index=False)
        svod_vac_sphere_mun_region_df.to_excel(writer, sheet_name='Отрасли муниципалитеты', index=False)
        svod_vac_org_region_df.to_excel(writer, sheet_name='Вакансии по работодателям', index=False)
        svod_vac_particular_org_region_df.to_excel(writer, sheet_name='Вакансии для динамики', index=False)
        svod_shpere_pay_region_df.to_excel(writer, sheet_name='Зарплата по отраслям', index=False)
        svod_shpere_category_pay_region_df.to_excel(writer, sheet_name='Категории ЗП по отраслям', index=False)
        svod_org_pay_region_df.to_excel(writer, sheet_name='Зарплата по работодателям', index=False)
        svod_org_category_pay_region_df.to_excel(writer, sheet_name='Категории ЗП по работодателям', index=False)
        svod_shpere_educ_region_df.to_excel(writer, sheet_name='Образование по отраслям', index=False)
        svod_org_educ_region_df.to_excel(writer, sheet_name='Образование по работодателям', index=False)
        svod_shpere_schedule_region_df.to_excel(writer, sheet_name='График работы по отраслям', index=False)
        svod_org_schedule_region_df.to_excel(writer, sheet_name='График работы по работодателям', index=False)
        svod_shpere_type_job_region_df.to_excel(writer, sheet_name='Тип занятости по отраслям', index=False)
        svod_org_type_job_region_df.to_excel(writer, sheet_name='Тип занятости по работодателям', index=False)
        svod_shpere_quote_region_df.to_excel(writer, sheet_name='Квоты по отраслям', index=False)
        svod_org_quote_region_df.to_excel(writer, sheet_name='Квоты по работодателям', index=False)
        svod_soc_region_df.to_excel(writer, sheet_name='Вакансии для соц.кат.', index=False)
        svod_shpere_exp_region_df.to_excel(writer, sheet_name='Требуемый опыт по отраслям', index=False)
        svod_org_exp_region_df.to_excel(writer, sheet_name='Требуемый опыт по работодателям', index=False)





def processing_data_trudvsem(file_data:str,file_org,end_folder:str,region:str,param_filter):
    """
    Основная функция для обработки данных
    :param file_data: файл в формате csv с данными вакансий
    :param file_org: файл с данными организаций по которым нужно сделать отдельный свод
    :param region: регион вакансии которого нужно обработать
    :param param_filter: файл с параметрами по которым нужно отфильтровать датафрейм
    :param end_folder: конечная папка
    """
    # колонки которые нужно оставить и переименовать
    dct_name_columns = {'id':'ID вакансии','busyType': 'Тип занятости', 'contactPerson': 'Контактное лицо',
                        'creationDate': 'Дата размещения вакансии',
                        'dateModify': 'Дата изменения вакансии', 'educationRequirements': 'Данные по образованию',
                         'isQuoted': 'Квотируемое место',
                         'accommodationCapability': 'Жилье от организации',
                         'accommodationType': 'Тип жилья',
                        'needMedcard': 'Требуется медкнижка',
                        'otherVacancyBenefit': 'Бонусы', 'positionRequirements': 'Требования',
                        'regionName': 'Регион',
                        'foreignWorkersCapability': 'Для иностранцев',
                        'isMobilityProgram': 'Программа трудовой мобильности',
                        'experienceRequirements': 'Требуемый опыт работы в годах',
                        'retrainingCapability': 'Возможность переподготовки',
                        'requiredСertificates': 'Требуемые доп. документы',
                        'requiredDriveLicense': 'Требуемые водительские права',
                        'retrainingGrantValue': 'Размер стипендии', 'salary': 'Зарплата',
                        'scheduleType': 'График работы', 'socialProtecteds': 'Социально защищенная категория',
                        'sourceType': 'Источник вакансии', 'status': 'Статус проверки вакансии',
                        'transportCompensation': 'Компенсация транспорт',
                        'vacancyAddressAdditionalInfo': 'Доп информация по адресу вакансии',
                        'vacancyAddress': 'Адрес вакансии',
                         'vacancyName': 'Вакансия',
                        'workPlaces': 'Количество рабочих мест', 'professionalSphereName': 'Сфера деятельности',
                        'fullCompanyName': 'Полное название работодателя', 'companyBusinessSize': 'Размер организации',
                        'company': 'Данные компании',
                        'languageKnowledge': 'Данные по языкам', 'hardSkills': 'Данные по хардскиллам',
                        'softSkills': 'Данные по софтскиллам',
                        'vacancyUrl': 'Ссылка на вакансию',
                        'geo': 'Геоданные',
                        'contactList': 'Контактные данные',
                        }

    try:
        t = time.localtime()  # получаем текущее время и дату
        current_time = time.strftime('%H_%M_%S', t)
        current_date = time.strftime('%d_%m_%Y', t)
        # Получаем данные из csv
        main_df = pd.read_csv(file_data, encoding='UTF-8', sep='|', dtype=str, on_bad_lines='skip')
        if file_org == '' or file_org == 'Не выбрано':
            company_df = pd.DataFrame()
        else:
            company_df = pd.read_excel(file_org, dtype=str, usecols='A:B')  # получаем данные из файла с организациями

        company_df.dropna(inplace=True) # удаляем незаполненные строки
        # Список колонок итоговых таблиц с вакансиями
        lst_columns = ['Дата размещения вакансии','Дата изменения вакансии','Регион','Вакансия','Сфера деятельности','Количество рабочих мест',
                       'Зарплата','Минимальная зарплата','Категория минимальной зарплаты','График работы','Тип занятости','Образование','Требуемая специализация',
                       'Требования','Бонусы','Жилье от организации','Тип жилья','Возможность переподготовки','Размер стипендии','Компенсация транспорт',
                       'Квотируемое место','Социально защищенная категория',
                       'Для иностранцев','Программа трудовой мобильности',
                       'Требуемый опыт работы в годах','Требуется медкнижка','Требуемые доп. документы','Требуемые водительские права',
                       'Требуемые языки','Требуемые хардскиллы','Требуемые софтскиллы',
                       'Источник вакансии','Статус проверки вакансии','Размер организации','Полное название работодателя','Краткое название работодателя','Муниципалитет','Адрес вакансии','Доп информация по адресу вакансии',
                       'ИНН работодателя','КПП работодателя','ОГРН работодателя','Контактное лицо','Контактный телефон','Email контактного лица','Email работодателя',
                       'Профиль работодателя','Сайт работодателя','Широта адрес вакансии','Долгота адрес вакансии','ID вакансии','ID работодателя','Ссылка на вакансию']

        # Список колонок с текстом
        lst_text_columns = ['Вакансия', 'Требуемая специализация', 'Требования',
                            'Бонусы', 'Требуемые доп. документы',
                            'Требуемые хардскиллы', 'Требуемые софтскиллы', 'Полное название работодателя','Краткое название работодателя',
                            'Муниципалитет','Адрес вакансии', 'Доп информация по адресу вакансии', 'Email работодателя',
                            'Контактное лицо']
        lst_region = main_df['regionName'].unique() # Получаем список регионов
        # проверяем
        if region not in lst_region:
            raise NotRegion
        df = main_df[main_df['regionName'] == region] # Фильтруем данные по региону

        del main_df # очищаем память
        # получаем обработанный датафрейм со всеми статусами вакансий
        all_status_prepared_df = prepare_data_vacancy(df, dct_name_columns,lst_columns)

        if param_filter != '' and param_filter != 'Не выбрано':
            all_status_prepared_df = filtred_df(all_status_prepared_df,param_filter) # отфильтровываем датафрейм


        # получаем датафрейм только с подтвержденными вакансиями
        prepared_df = all_status_prepared_df[
            all_status_prepared_df['Статус проверки вакансии'] == 'Одобрено']
        union_company_df_columns = list(prepared_df.columns).insert(0, 'Организация')
        union_company_df = pd.DataFrame(columns=union_company_df_columns)

        # Создаем папку основную

        svod_region_folder = f'{end_folder}/Аналитика по вакансиям региона/{current_date}'  # создаем папку куда будем складывать аналитику по региону
        if not os.path.exists(svod_region_folder):
            os.makedirs(svod_region_folder)


            # Собираем датафреймы по ИНН

        if len(company_df) != 0:
            org_folder = f'{end_folder}/Вакансии по организациям/{current_date}'  # создаем папку куда будем складывать вакансии по организациям
            if not os.path.exists(org_folder):
                os.makedirs(org_folder)
            count_exists_file = 0 # счетчик для уже созданных файлов чтобы не затирались
            for idx, row in enumerate(company_df.itertuples()):
                name_company = row[1]  # название компании
                inn_company = row[2]  # инн компании
                temp_df = prepared_df[prepared_df['ИНН работодателя'] == inn_company]  # фильтруем по инн
                if len(temp_df) != 0:
                    temp_df.sort_values(by=['Вакансия'], inplace=True)  # сортируем
                    name_company = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_company)  # очищаем название от лишних символов
                    temp_df[lst_text_columns] = temp_df[lst_text_columns].applymap(clean_equal) # очищаем от знака равно в начале
                    temp_df[lst_text_columns] = temp_df[lst_text_columns].applymap(clean_text) # очищаем от неправильных символов
                    # считаем возможную длину названия файл с учетом слеша и расширения с точкой и порядковым номером файла
                    threshold_name = 200 - (len(org_folder)+10)
                    if threshold_name <= 0: # если путь к папке слшиком длинный вызываем исключение
                        raise OSError
                    name_company = name_company[:threshold_name] # ограничиваем название файла
                    # сохраняем файл с проверкой на существующие файлы с таким же именем
                    if not os.path.exists(f'{org_folder}/{name_company}.xlsx'):
                        temp_df.to_excel(f'{org_folder}/{name_company}.xlsx', index=False)  # сохраняем
                    else:
                        temp_df.to_excel(f'{org_folder}/{name_company}_{count_exists_file}.xlsx', index=False)  # сохраняем
                        count_exists_file += 1

                    # создаем отдельный файл в котором будут все вакансии по выбранным компаниям
                    temp_df.insert(0, 'Организация', name_company)
                    union_company_df = pd.concat([union_company_df, temp_df], ignore_index=True)

        # Сортируем по колонке Вакансия
        prepared_df.sort_values(by=['Вакансия'],inplace=True)
        all_status_prepared_df.sort_values(by=['Вакансия'],inplace=True)


        # Сохраняем общий файл с всеми вакансиями выбранных работодателей
        try:
            # очищаем текстовые колонки от возможного знака равно в начале ячейки, в таком случае возникает ошибка
            # потому что значение принимается за формулу
            prepared_df[lst_text_columns] = prepared_df[lst_text_columns].applymap(clean_equal)
            all_status_prepared_df[lst_text_columns] = all_status_prepared_df[lst_text_columns].applymap(clean_equal)
            # создаем 2 датафрейма для квотируемых и вакансий для соц категорий
            quote_df = prepared_df[prepared_df['Квотируемое место'] == 'Квотируемое место']
            soc_df = prepared_df[~prepared_df['Социально защищенная категория'].isna()]

            # Создаем 3 датафрейма по вакансием с предоставлением жилья, для мобильной занятости, для иностранных специалистов,
            accommodation_df = prepared_df[prepared_df['Жилье от организации'] == 'Предоставляется жилье']
            program_mobile_df = prepared_df[prepared_df['Программа трудовой мобильности'] == 'Вакансия по программе трудовой мобильности']
            migrant_mobile_df = prepared_df[prepared_df['Для иностранцев'] == 'Возможно трудоустройство иностранных граждан']
            if len(union_company_df) != 0:
                union_company_df.sort_values(by=['Вакансия'], inplace=True)
                union_company_df[lst_text_columns] = union_company_df[lst_text_columns].applymap(clean_equal)
                company_quote_df = union_company_df[union_company_df['Квотируемое место'] == 'Квотируемое место']
                company_soc_df = union_company_df[~union_company_df['Социально защищенная категория'].isna()]
                # Создаем 3 датафрейма по вакансием с предоставлением жилья, для мобильной занятости, для иностранных специалистов,
                company_accommodation_df = union_company_df[union_company_df['Жилье от организации'] == 'Предоставляется жилье']
                company_program_mobile_df = union_company_df[
                    union_company_df['Программа трудовой мобильности'] == 'Вакансия по программе трудовой мобильности']
                company_migrant_mobile_df = union_company_df[
                    union_company_df['Для иностранцев'] == 'Возможно трудоустройство иностранных граждан']
                with pd.ExcelWriter(f'{org_folder}/Общий файл.xlsx') as writer:
                    union_company_df.to_excel(writer, sheet_name='Общий список', index=False)
                    company_quote_df.to_excel(writer, sheet_name='Квотируемые', index=False)
                    company_soc_df.to_excel(writer, sheet_name='Для соц категорий', index=False)
                    company_accommodation_df.to_excel(writer, sheet_name='С предоставлением жилья', index=False)
                    company_program_mobile_df.to_excel(writer, sheet_name='Трудовая мобильность', index=False)
                    company_migrant_mobile_df.to_excel(writer, sheet_name='Для иностранцев', index=False)

            with pd.ExcelWriter(f'{end_folder}/Вакансии по региону от {current_date}.xlsx') as writer:
                prepared_df.to_excel(writer, sheet_name='Только подтвержденные вакансии', index=False)
                quote_df.to_excel(writer, sheet_name='Квотируемые', index=False)
                soc_df.to_excel(writer, sheet_name='Для соц категорий', index=False)
                accommodation_df.to_excel(writer, sheet_name='С предоставлением жилья', index=False)
                program_mobile_df.to_excel(writer, sheet_name='Трудовая мобильность', index=False)
                migrant_mobile_df.to_excel(writer, sheet_name='Для иностранцев', index=False)
                all_status_prepared_df.to_excel(writer, sheet_name='Вакансии со всеми статусами', index=False)
        except IllegalCharacterError:
            # Если в тексте есть ошибочные символы то очищаем данные
            # очищаем от неправильных символов
            prepared_df[lst_text_columns] = prepared_df[lst_text_columns].applymap(clean_text)
            all_status_prepared_df[lst_text_columns] = all_status_prepared_df[lst_text_columns].applymap(clean_text)

            quote_df = prepared_df[prepared_df['Квотируемое место'] == 'Квотируемое место']
            soc_df = prepared_df[~prepared_df['Социально защищенная категория'].isna()]
            # Создаем 3 датафрейма по вакансием с предоставлением жилья, для мобильной занятости, для иностранных специалистов,
            accommodation_df = prepared_df[prepared_df['Жилье от организации'] == 'Предоставляется жилье']
            program_mobile_df = prepared_df[prepared_df['Программа трудовой мобильности'] == 'Вакансия по программе трудовой мобильности']
            migrant_mobile_df = prepared_df[prepared_df['Для иностранцев'] == 'Возможно трудоустройство иностранных граждан']

            if len(union_company_df) != 0:
                union_company_df.sort_values(by=['Вакансия'], inplace=True)
                union_company_df[lst_text_columns] = union_company_df[lst_text_columns].applymap(clean_text)
                company_quote_df = union_company_df[union_company_df['Квотируемое место'] == 'Квотируемое место']
                company_soc_df = union_company_df[~union_company_df['Социально защищенная категория'].isna()]
                # Создаем 3 датафрейма по вакансием с предоставлением жилья, для мобильной занятости, для иностранных специалистов,
                company_accommodation_df = union_company_df[union_company_df['Жилье от организации'] == 'Предоставляется жилье']
                company_program_mobile_df = union_company_df[
                    union_company_df['Программа трудовой мобильности'] == 'Вакансия по программе трудовой мобильности']
                company_migrant_mobile_df = union_company_df[union_company_df['Для иностранцев'] == 'Возможно трудоустройство иностранных граждан']

                with pd.ExcelWriter(f'{org_folder}/Общий файл.xlsx') as writer:
                    union_company_df.to_excel(writer, sheet_name='Общий список', index=False)
                    company_quote_df.to_excel(writer, sheet_name='Квотируемые', index=False)
                    company_soc_df.to_excel(writer, sheet_name='Для соц категорий', index=False)
                    company_accommodation_df.to_excel(writer, sheet_name='С предоставлением жилья', index=False)
                    company_program_mobile_df.to_excel(writer, sheet_name='Трудовая мобильность', index=False)
                    company_migrant_mobile_df.to_excel(writer, sheet_name='Для иностранцев', index=False)

            with pd.ExcelWriter(f'{end_folder}/Вакансии по региону от {current_date}.xlsx') as writer:
                prepared_df.to_excel(writer, sheet_name='Только подтвержденные вакансии', index=False)
                quote_df.to_excel(writer, sheet_name='Квотируемые', index=False)
                soc_df.to_excel(writer, sheet_name='Для соц категорий', index=False)
                accommodation_df.to_excel(writer, sheet_name='С предоставлением жилья', index=False)
                program_mobile_df.to_excel(writer, sheet_name='Трудовая мобильность', index=False)
                migrant_mobile_df.to_excel(writer, sheet_name='Для иностранцев', index=False)
                all_status_prepared_df.to_excel(writer, sheet_name='Вакансии со всеми статусами', index=False)


        """
            Свод по региону
            """
        create_svod_for_df(prepared_df,svod_region_folder,'Свод по региону',current_date)
        create_svod_for_df(quote_df,svod_region_folder,'Квотируемые',current_date)
        create_svod_for_df(soc_df,svod_region_folder,'Соцкатегории',current_date)

        create_svod_for_df(accommodation_df,svod_region_folder,'С предоставлением жилья',current_date)
        create_svod_for_df(program_mobile_df,svod_region_folder,'Трудовая мобильность',current_date)
        create_svod_for_df(migrant_mobile_df,svod_region_folder,'Иностранцы',current_date)



        """
            Свод по выбранным работодателям
            """

        if len(union_company_df) != 0:
            svod_org_folder = f'{end_folder}/Аналитика по вакансиям выбранных работодателей/{current_date}'  # создаем папку куда будем складывать аналитику по выбранным работодателям
            if not os.path.exists(svod_org_folder):
                os.makedirs(svod_org_folder)

            # Свод по вакансиям выбранных работодателей

            # Свод по количеству рабочих мест по отраслям

            svod_vac_reg_org_df = pd.pivot_table(union_company_df,
                                                 index=['Сфера деятельности'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]})
            svod_vac_reg_org_df = svod_vac_reg_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_vac_reg_org_df) != 0:
                svod_vac_reg_org_df.sort_values(by=['sum'], ascending=False, inplace=True)
                svod_vac_reg_org_df.loc['Итого'] = svod_vac_reg_org_df['sum'].sum()
                svod_vac_reg_org_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
                svod_vac_reg_org_df = svod_vac_reg_org_df.reset_index()

            # Свод по муниципалитетам для выбранных работодателей
            svod_vac_mun_org_df = pd.pivot_table(union_company_df,
                                                 index=['Муниципалитет'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]})
            svod_vac_mun_org_df = svod_vac_mun_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_vac_mun_org_df) != 0:
                svod_vac_mun_org_df.sort_values(by=['sum'], ascending=False, inplace=True)
                svod_vac_mun_org_df.loc['Итого'] = svod_vac_mun_org_df['sum'].sum()
                svod_vac_mun_org_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
                svod_vac_mun_org_df = svod_vac_mun_org_df.reset_index()


            # Свод по количеству рабочих мест по организациям
            svod_vac_org_org_df = pd.pivot_table(union_company_df,
                                                 index=['Краткое название работодателя'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]})
            svod_vac_org_org_df = svod_vac_org_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_vac_org_org_df) != 0:
                svod_vac_org_org_df.sort_values(by=['sum'], ascending=False, inplace=True)
                svod_vac_org_org_df.loc['Итого'] = svod_vac_org_org_df['sum'].sum()
                svod_vac_org_org_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
                svod_vac_org_org_df = svod_vac_org_org_df.reset_index()

            # список вакансий выбранных работодателей для отслеживания динамики
            svod_vac_particular_org_org_df = union_company_df[
                ['Краткое название работодателя', 'Вакансия', 'Количество рабочих мест', 'ID вакансии',
                 'Ссылка на вакансию']]

            # Свод по средней и медианной минимальной зарплате для сфер деятельности
            union_company_df['Минимальная зарплата'] = union_company_df['Минимальная зарплата'].apply(convert_int)
            pay_union_df = union_company_df[union_company_df['Минимальная зарплата'] > 0]

            # Свод по категориям минимальной заработной платы для сфеф деятельности
            svod_shpere_category_pay_org_df = pd.pivot_table(pay_union_df,
                                                                index=['Сфера деятельности',
                                                                       'Категория минимальной зарплаты'],
                                                                values=['Количество рабочих мест'],
                                                                aggfunc={'Количество рабочих мест': 'sum'}
                                                                )

            if len(svod_shpere_category_pay_org_df) != 0:
                svod_shpere_category_pay_org_df = svod_shpere_category_pay_org_df.reset_index()
                svod_shpere_category_pay_org_df.rename(columns={'Количество рабочих мест': 'Количество вакансий'},
                                                       inplace=True)

            # Свод по категориям минимальной заработной платы для работодателей
            svod_org_category_pay_org_df = pd.pivot_table(pay_union_df,
                                                             index=['Краткое название работодателя',
                                                                    'Категория минимальной зарплаты'],
                                                             values=['Количество рабочих мест'],
                                                             aggfunc={'Количество рабочих мест': 'sum'}
                                                             )

            if len(svod_org_category_pay_org_df) != 0:
                svod_org_category_pay_org_df = svod_org_category_pay_org_df.reset_index()
                svod_org_category_pay_org_df.rename(columns={'Количество рабочих мест': 'Количество вакансий'},
                                                       inplace=True)

            svod_shpere_pay_org_df = pd.pivot_table(pay_union_df,
                                                    index=['Сфера деятельности'],
                                                    values=['Минимальная зарплата'],
                                                    aggfunc={'Минимальная зарплата': [np.mean, np.median]}
                                                   )
            svod_shpere_pay_org_df = svod_shpere_pay_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_pay_org_df) !=0:
                svod_shpere_pay_org_df = svod_shpere_pay_org_df.astype(int, errors='ignore')
                svod_shpere_pay_org_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
                svod_shpere_pay_org_df = svod_shpere_pay_org_df.reset_index()

            # Свод по средней и медианной минимальной зарплате для работодателей
            svod_org_pay_org_df = pd.pivot_table(pay_union_df,
                                                 index=['Краткое название работодателя', 'Сфера деятельности'],
                                                 values=['Минимальная зарплата'],
                                                 aggfunc={'Минимальная зарплата': [np.mean, np.median]}
                                                 )
            svod_org_pay_org_df = svod_org_pay_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_pay_org_df) != 0:
                svod_org_pay_org_df = svod_org_pay_org_df.astype(int, errors='ignore')
                svod_org_pay_org_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
                svod_org_pay_org_df = svod_org_pay_org_df.reset_index()

            # Свод по требуемому образованию для сфер деятельности
            svod_shpere_educ_org_df = pd.pivot_table(union_company_df,
                                                     index=['Сфера деятельности', 'Образование'],
                                                     values=['Количество рабочих мест'],
                                                     aggfunc={'Количество рабочих мест': [np.sum]},
                                                     fill_value=0)
            svod_shpere_educ_org_df = svod_shpere_educ_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_educ_org_df) != 0:
                svod_shpere_educ_org_df = svod_shpere_educ_org_df.astype(int, errors='ignore')
                svod_shpere_educ_org_df.columns = ['Количество вакансий']
                svod_shpere_educ_org_df = svod_shpere_educ_org_df.reset_index()

            # Свод по требуемому образованию для работодателей
            svod_org_educ_org_df = pd.pivot_table(union_company_df,
                                                  index=['Краткое название работодателя', 'Образование'],
                                                  values=['Количество рабочих мест'],
                                                  aggfunc={'Количество рабочих мест': [np.sum]},
                                                  fill_value=0)
            svod_org_educ_org_df = svod_org_educ_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_educ_org_df) != 0:
                svod_org_educ_org_df = svod_org_educ_org_df.astype(int, errors='ignore')
                svod_org_educ_org_df.columns = ['Количество вакансий']
                svod_org_educ_org_df = svod_org_educ_org_df.reset_index()

            # Свод по графику работы для сфер деятельности
            svod_shpere_schedule_org_df = pd.pivot_table(union_company_df,
                                                         index=['Сфера деятельности', 'График работы'],
                                                         values=['Количество рабочих мест'],
                                                         aggfunc={'Количество рабочих мест': [np.sum]},
                                                         fill_value=0)
            svod_shpere_schedule_org_df = svod_shpere_schedule_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_schedule_org_df) != 0:
                svod_shpere_schedule_org_df = svod_shpere_schedule_org_df.astype(int, errors='ignore')
                svod_shpere_schedule_org_df.columns = ['Количество вакансий']
                svod_shpere_schedule_org_df = svod_shpere_schedule_org_df.reset_index()

            # Свод по графику работы для работодателей
            svod_org_schedule_org_df = pd.pivot_table(union_company_df,
                                                      index=['Краткое название работодателя', 'График работы'],
                                                      values=['Количество рабочих мест'],
                                                      aggfunc={'Количество рабочих мест': [np.sum]},
                                                      fill_value=0)
            svod_org_schedule_org_df = svod_org_schedule_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_schedule_org_df) != 0:
                svod_org_schedule_org_df = svod_org_schedule_org_df.astype(int, errors='ignore')
                svod_org_schedule_org_df.columns = ['Количество вакансий']
                svod_org_schedule_org_df = svod_org_schedule_org_df.reset_index()

            # Свод по типу занятости для сфер деятельности
            svod_shpere_type_job_org_df = pd.pivot_table(union_company_df,
                                                         index=['Сфера деятельности', 'Тип занятости'],
                                                         values=['Количество рабочих мест'],
                                                         aggfunc={'Количество рабочих мест': [np.sum]},
                                                         fill_value=0)
            svod_shpere_type_job_org_df = svod_shpere_type_job_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_type_job_org_df) != 0:
                svod_shpere_type_job_org_df = svod_shpere_type_job_org_df.astype(int, errors='ignore')
                svod_shpere_type_job_org_df.columns = ['Количество вакансий']
                svod_shpere_type_job_org_df = svod_shpere_type_job_org_df.reset_index()

            # Свод по типу занятости для работодателей
            svod_org_type_job_org_df = pd.pivot_table(union_company_df,
                                                      index=['Краткое название работодателя', 'Тип занятости'],
                                                      values=['Количество рабочих мест'],
                                                      aggfunc={'Количество рабочих мест': [np.sum]},
                                                      fill_value=0)
            svod_org_type_job_org_df = svod_org_type_job_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_type_job_org_df) != 0:
                svod_org_type_job_org_df = svod_org_type_job_org_df.astype(int, errors='ignore')
                svod_org_type_job_org_df.columns = ['Количество вакансий']
                svod_org_type_job_org_df = svod_org_type_job_org_df.reset_index()

            # Свод по квотируемым местам для сфер деятельности
            svod_shpere_quote_org_df = pd.pivot_table(union_company_df,
                                                      index=['Сфера деятельности', 'Квотируемое место'],
                                                      values=['Количество рабочих мест'],
                                                      aggfunc={'Количество рабочих мест': [np.sum]},
                                                      fill_value=0)
            svod_shpere_quote_org_df = svod_shpere_quote_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_quote_org_df) != 0:
                svod_shpere_quote_org_df = svod_shpere_quote_org_df.astype(int, errors='ignore')
                svod_shpere_quote_org_df.columns = ['Количество вакансий']
                svod_shpere_quote_org_df = svod_shpere_quote_org_df.reset_index()

            # Свод по квотируемым местам для работодателей
            svod_org_quote_org_df = pd.pivot_table(union_company_df,
                                                   index=['Краткое название работодателя', 'Квотируемое место'],
                                                   values=['Количество рабочих мест'],
                                                   aggfunc={'Количество рабочих мест': [np.sum]},
                                                   fill_value=0)
            svod_org_quote_org_df = svod_org_quote_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_quote_org_df) != 0:
                svod_org_quote_org_df = svod_org_quote_org_df.astype(int, errors='ignore')
                svod_org_quote_org_df.columns = ['Количество вакансий']
                svod_org_quote_org_df = svod_org_quote_org_df.reset_index()

            # Свод по вакансиям для социальных категорий
            dct_soc_cat_org = extract_soc_category(union_company_df, 'Социально защищенная категория',
                                                   ';')  # считаем количество категорий
            svod_soc_org_df = pd.DataFrame.from_dict(dct_soc_cat_org, orient='index').reset_index()  # содаем датафрейм
            if len(svod_soc_org_df) !=0:
                svod_soc_org_df.columns = ['Категория', 'Количество вакансий']
                svod_soc_org_df.sort_values(by=['Количество вакансий'], ascending=False, inplace=True)

            # Свод по требуемому опыту для сфер деятельности
            svod_shpere_exp_org_df = pd.pivot_table(union_company_df,
                                                    index=['Сфера деятельности', 'Требуемый опыт работы в годах'],
                                                    values=['Количество рабочих мест'],
                                                    aggfunc={'Количество рабочих мест': [np.sum]},
                                                    fill_value=0)
            svod_shpere_exp_org_df = svod_shpere_exp_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_shpere_exp_org_df) != 0:
                svod_shpere_exp_org_df = svod_shpere_exp_org_df.astype(int, errors='ignore')
                svod_shpere_exp_org_df.columns = ['Количество вакансий']
                svod_shpere_exp_org_df = svod_shpere_exp_org_df.reset_index()

            # Свод по требуемому опыту для работодателей
            svod_org_exp_org_df = pd.pivot_table(union_company_df,
                                                 index=['Краткое название работодателя', 'Требуемый опыт работы в годах'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)
            svod_org_exp_org_df = svod_org_exp_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            if len(svod_org_exp_org_df) != 0:
                svod_org_exp_org_df = svod_org_exp_org_df.astype(int, errors='ignore')
                svod_org_exp_org_df.columns = ['Количество вакансий']
                svod_org_exp_org_df = svod_org_exp_org_df.reset_index()

            with pd.ExcelWriter(f'{svod_org_folder}/Свод по выбранным работодателям от {current_date}.xlsx') as writer:
                svod_vac_reg_org_df.to_excel(writer, sheet_name='Вакансии по отраслям', index=False)
                svod_vac_mun_org_df.to_excel(writer, sheet_name='Вакансии по муниципалитетам', index=False)
                svod_vac_org_org_df.to_excel(writer, sheet_name='Вакансии по работодателям', index=False)
                svod_vac_particular_org_org_df.to_excel(writer,sheet_name='Вакансии для динамики',index=False)
                svod_shpere_pay_org_df.to_excel(writer, sheet_name='Зарплата по отраслям', index=False)
                svod_shpere_category_pay_org_df.to_excel(writer, sheet_name='Категории ЗП по отраслям', index=False)
                svod_org_pay_org_df.to_excel(writer, sheet_name='Зарплата по работодателям', index=False)
                svod_org_category_pay_org_df.to_excel(writer, sheet_name='Категории ЗП по работодателям', index=False)
                svod_shpere_educ_org_df.to_excel(writer, sheet_name='Образование по отраслям', index=False)
                svod_org_educ_org_df.to_excel(writer, sheet_name='Образование по работодателям', index=False)
                svod_shpere_schedule_org_df.to_excel(writer, sheet_name='График работы по отраслям', index=False)
                svod_org_schedule_org_df.to_excel(writer, sheet_name='График работы по работодателям', index=False)
                svod_shpere_type_job_org_df.to_excel(writer, sheet_name='Тип занятости по отраслям', index=False)
                svod_org_type_job_org_df.to_excel(writer, sheet_name='Тип занятости по работодателям', index=False)
                svod_shpere_quote_org_df.to_excel(writer, sheet_name='Квоты по отраслям', index=False)
                svod_org_quote_org_df.to_excel(writer, sheet_name='Квоты по работодателям', index=False)
                svod_soc_org_df.to_excel(writer, sheet_name='Вакансии для соц.кат.', index=False)
                svod_shpere_exp_org_df.to_excel(writer, sheet_name='Требуемый опыт по отраслям', index=False)
                svod_org_exp_org_df.to_excel(writer, sheet_name='Требуемый опыт по работодателям', index=False)
    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except NotRegion:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 f'Не найден регион! Проверьте написание региона в соответствии с правилами сайта Работа в России')
    except MoreColumn:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'В файле параметров фильтрации найдено больше 3 колонок. Удалите лишние колонки')
    except NotColumn as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'В файле параметров фильтрации обнаружены названия колонок, которых нет в файле Вакансии по региону.\n'
                             f'Запустите программу без использования файла фильтрации и укажите корректные названия колонок которые есть в файле Вакансии по региону')

    except KeyError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Не найдено значение {e.args}')

    except PermissionError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Закройте открытые файлы Excel {e.args}')
    except OSError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Укажите в качестве конечной папки, папку в корне диска с коротким названием. Проблема может быть\n '
                             f'в слишком длинном пути к создаваемому файлу')

    else:
        messagebox.showinfo('Кассандра Подсчет данных по трудоустройству выпускников',
                            'Данные успешно обработаны.Ошибок не обнаружено')


if __name__ == '__main__':
    main_file_data = 'data/vacancy.csv'
    main_org_file = 'data/Организации Бурятия.xlsx'
    main_org_file = 'Не выбрано'
    main_region = 'Республика Бурятия'
    main_param_filter = 'data/Параметры отбора.xlsx'
    main_param_filter = 'Не выбрано'

    main_end_folder = 'c:/Users/1/PycharmProjects/Dodger_2023/data/Республика Бурятия'
    main_end_folder = 'c:/Users/1/PycharmProjects/Dodger_2023/data/РЕЗУЛЬТАТ'

    processing_data_trudvsem(main_file_data,main_org_file,main_end_folder,main_region,main_param_filter)

    print('Lindy Booth !!!')