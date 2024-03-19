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


class NotRegion(Exception):
    """
    Класс для отслеживания наличия региона в данных
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
    Функция для очистки данных в колонке Дополнительные бонусы
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
        return re.sub(r'[^\d\w ()=*+,.:;-]','',cell)
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


def prepare_data_vacancy(df: pd.DataFrame, dct_name_columns: dict, lst_columns: list) -> pd.DataFrame:
    """
    Функция для обработки датафрейма с данными работы в России
    """
    # Словарь для замены статусов подтверждения вакансии
    dct_status_vacancy = {'ACCEPTED':'Данные вакансии проверены работодателем','AUTOMODERATION':'Автомодерация',
                      'REJECTED':'Отклонено','CHANGED':'Статус вакансии изменен',
                      'WAITING':'Ожидает подтверждения',}
    df = df[dct_name_columns.keys()] # отбираем только нужные колонки
    df.rename(columns=dct_name_columns, inplace=True,errors='ignore')
    # Обрабатываем обычные колонки

    df['Дополнительные бонусы'] = df['Дополнительные бонусы'].apply(clear_bonus_tag_br)
    df['Требования'] = df['Требования'].apply(clear_tag)
    df['Обязанности'] = df['Обязанности'].apply(clear_tag)

    # Числовые
    lst_number_columns = ['Требуемый опыт работы в годах', 'Минимальная зарплата', 'Максимальная зарплата',
                          'Количество рабочих мест']
    df[lst_number_columns] = df[lst_number_columns].fillna(0)
    df[lst_number_columns] = df[lst_number_columns].astype(int, errors='ignore')

    # Временные

    df['Дата размещения вакансии'] = df['Дата размещения вакансии'].apply(convert_date)
    df['Дата изменения вакансии'] = df['Дата изменения вакансии'].apply(convert_date)

    # Категориальные
    df['Квотируемое место'] = df['Квотируемое место'].apply(lambda x: 'Квотируемое место' if x == 'true' else None)
    df['Требуется медкнижка'] = df['Требуется медкнижка'].apply(
        lambda x: 'Требуется медкнижка' if x == 'true' else None)
    df['Статус проверки вакансии'] = df['Статус проверки вакансии'].replace(dct_status_vacancy)
    # Начинаем извлекать данные из сложных колонок с json
    # данные по работодателю
    df['КПП работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('kpp', 'Не указано'))
    df['ОГРН работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('ogrn', 'Не указано'))
    df['Контактный телефон'] = df['Данные компании'].apply(lambda x: json.loads(x).get('phone', 'Не указано'))
    df['Email работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('email', 'Не указано'))
    df['Профиль работодателя'] = df['Данные компании'].apply(lambda x: json.loads(x).get('url', 'Не указано'))

    # Обрабатываем колонку с языками
    df['Требуемые языки'] = df['Данные по языкам'].apply(
        lambda x: extract_data_from_list_cell(x, ['code_language', 'level']))
    df['Требуемые хардскиллы'] = df['Данные по хардскиллам'].apply(
        lambda x: extract_data_from_list_cell(x, ['hard_skill_name']))
    df['Требуемые софтскиллы'] = df['Данные по софтскиллам'].apply(
        lambda x: extract_data_from_list_cell(x, ['soft_skill_name']))

    df.drop(columns=['Данные компании', 'Данные по языкам', 'Данные по хардскиллам', 'Данные по софтскиллам'],
            inplace=True)

    df = df.reindex(columns=lst_columns)

    return df


def processing_data_trudvsem(file_data:str,file_org:str,end_folder:str,region:str):
    """
    Основная функция для обработки данных
    :param file_data: файл в формате csv с данными вакансий
    :param file_org: файл с данными организаций по которым нужно сделать отдельный свод
    :param region: регион вакансии которого нужно обработать
    :param end_folder: конечная папка
    """
    # колонки которые нужно оставить и переименовать
    dct_name_columns = {'busy_type': 'Тип занятости', 'contact_person': 'Контактное лицо',
                        'date_create': 'Дата размещения вакансии',
                        'date_modify': 'Дата изменения вакансии', 'education': 'Образование',
                        'education_speciality': 'Требуемая специализация', 'is_quoted': 'Квотируемое место',
                        'need_medcard': 'Требуется медкнижка',
                        'other_vacancy_benefit': 'Дополнительные бонусы', 'position_requirements': 'Требования',
                        'position_responsibilities': 'Обязанности', 'regionName': 'Регион',
                        'required_experience': 'Требуемый опыт работы в годах',
                        'retraining_capability': 'Возможность переподготовки',
                        'required_certificates': 'Требуемые доп. документы',
                        'required_drive_license': 'Требуемые водительские права', 'retraining_grant': 'Стипендия',
                        'retraining_grant_value': 'Размер стипендии', 'salary': 'Зарплата',
                        'salary_min': 'Минимальная зарплата', 'salary_max': 'Максимальная зарплата',
                        'schedule_type': 'График работы', 'social_protected_ids': 'Социально защищенная категория',
                        'source_type': 'Источник вакансии', 'status': 'Статус проверки вакансии',
                        'transport_compensation': 'Компенсация транспорт',
                        'vacancy_address_additional_info': 'Доп информация по адресу вакансии',
                        'vacancy_address': 'Адрес вакансии',
                        'vacancy_address_latitude': 'Долгота адрес вакансии',
                        'vacancy_address_longitude': 'Широта адрес вакансии',
                        'vacancy_benefit_ids': 'Бонусы', 'vacancy_name': 'Вакансия',
                        'work_places': 'Количество рабочих мест', 'professionalSphereName': 'Сфера деятельности',
                        'full_company_name': 'Полное название работодателя', 'company_inn': 'ИНН работодателя',
                        'company': 'Данные компании',
                        'languageKnowledge': 'Данные по языкам', 'hardSkills': 'Данные по хардскиллам',
                        'softSkills': 'Данные по софтскиллам'}

    try:
        t = time.localtime()  # получаем текущее время и дату
        current_time = time.strftime('%H_%M_%S', t)
        current_date = time.strftime('%d_%m_%Y', t)
        # Получаем данные из csv
        df = pd.read_csv(file_data, encoding='UTF-8', sep='|', dtype=str, on_bad_lines='skip')
        company_df = pd.read_excel(file_org, dtype=str) # получаем данные из файла с организациями
        company_df.dropna(inplace=True) # удаляем незаполненные строки
        # Список колонок итоговых таблиц с вакансиями
        lst_columns = ['Дата размещения вакансии','Дата изменения вакансии','Регион','Вакансия','Сфера деятельности','Количество рабочих мест',
                       'Зарплата','Минимальная зарплата','Максимальная зарплата','График работы','Тип занятости','Образование','Требуемая специализация',
                       'Требования','Обязанности','Бонусы','Дополнительные бонусы','Возможность переподготовки','Стипендия','Размер стипендии','Компенсация транспорт',
                       'Квотируемое место','Социально защищенная категория',
                       'Требуемый опыт работы в годах','Требуется медкнижка','Требуемые доп. документы','Требуемые водительские права',
                       'Требуемые языки','Требуемые хардскиллы','Требуемые софтскиллы',
                       'Источник вакансии','Статус проверки вакансии','Полное название работодателя','Адрес вакансии','Доп информация по адресу вакансии',
                       'ИНН работодателя','КПП работодателя','ОГРН работодателя','Контактное лицо','Контактный телефон','Email работодателя',
                       'Профиль работодателя','Долгота адрес вакансии','Широта адрес вакансии']

        lst_region = df['regionName'].unique() # Получаем список регионов
        # проверяем
        if region not in lst_region:
            raise NotRegion
        print(region)
        df = df[df['regionName'] == region] # Фильтруем данные по региону

        # получаем обработанный датафрейм со всеми статусами вакансий
        all_status_prepared_df = prepare_data_vacancy(df, dct_name_columns,lst_columns)

        # получаем датафрейм только с подтвержденными вакансиями
        prepared_df = all_status_prepared_df[
            all_status_prepared_df['Статус проверки вакансии'] == 'Данные вакансии проверены работодателем']
        union_company_df_columns = list(prepared_df.columns).insert(0, 'Организация')
        union_company_df = pd.DataFrame(columns=union_company_df_columns)

        # Создаем нужные папки
        org_folder = f'{end_folder}/Вакансии по организациям/{current_date}'  # создаем папку куда будем складывать вакансии по организациям
        if not os.path.exists(org_folder):
            os.makedirs(org_folder)

        svod_region_folder = f'{end_folder}/Аналитика по вакансиям региона/{current_date}'  # создаем папку куда будем складывать аналитику по региону
        if not os.path.exists(svod_region_folder):
            os.makedirs(svod_region_folder)

        svod_org_folder = f'{end_folder}/Аналитика по вакансиям выбранных работодателей/{current_date}'  # создаем папку куда будем складывать аналитику по выбранным работодателям
        if not os.path.exists(svod_org_folder):
            os.makedirs(svod_org_folder)

            # Собираем датафреймы по ИНН
        for idx, row in enumerate(company_df.itertuples()):
            name_company = row[1]  # название компании
            inn_company = row[2]  # инн компании
            temp_df = prepared_df[prepared_df['ИНН работодателя'] == inn_company]  # фильтруем по инн
            temp_df.sort_values(by=['Вакансия'], inplace=True)  # сортируем
            name_company = re.sub(r'[\r\b\n\t<>:"?*|\\/]', '_', name_company)  # очищаем от лишних символов

            temp_df.to_excel(f'{org_folder}/{name_company}.xlsx', index=False)  # сохраняем
            # создаем отдельный файл в котором будут все вакансии по выбранным компаниям
            temp_df.insert(0, 'Организация', name_company)
            union_company_df = pd.concat([union_company_df, temp_df], ignore_index=True)

        # Сортируем по колонке Вакансия
        prepared_df.sort_values(by=['Вакансия'],inplace=True)
        all_status_prepared_df.sort_values(by=['Вакансия'],inplace=True)
        union_company_df.sort_values(by=['Вакансия'],inplace=True)

        # Сохраняем общий файл с всеми вакансиями выбранных работодателей
        try:
            # Список колонок с текстом
            lst_text_columns = ['Вакансия','Требуемая специализация','Требования','Обязанности',
                                'Бонусы','Дополнительные бонусы','Требуемые доп. документы',
                                'Требуемые хардскиллы','Требуемые софтскиллы','Полное название работодателя',
                                'Адрес вакансии','Доп информация по адресу вакансии','Email работодателя','Контактное лицо']
            # очищаем текстовые колонки от возможного знака равно в начале ячейки, в таком случае возникает ошибка
            # потому что значение принимается за формулу
            prepared_df[lst_text_columns] = prepared_df[lst_text_columns].applymap(clean_equal)
            all_status_prepared_df[lst_text_columns] = all_status_prepared_df[lst_text_columns].applymap(clean_equal)
            union_company_df[lst_text_columns] = union_company_df[lst_text_columns].applymap(clean_equal)

            union_company_df.to_excel(f'{org_folder}/Общий файл.xlsx', index=False)

            with pd.ExcelWriter(f'{end_folder}/Вакансии по региону от {current_time}.xlsx') as writer:
                prepared_df.to_excel(writer, sheet_name='Только подтвержденные вакансии', index=False)
                all_status_prepared_df.to_excel(writer, sheet_name='Вакансии со всеми статусами', index=False)
        except IllegalCharacterError:
            # Список колонок с текстом
            lst_text_columns = ['Вакансия','Требуемая специализация','Требования','Обязанности',
                                'Бонусы','Дополнительные бонусы','Требуемые доп. документы',
                                'Требуемые хардскиллы','Требуемые софтскиллы','Полное название работодателя',
                                'Адрес вакансии','Доп информация по адресу вакансии','Email работодателя','Контактное лицо']

            # очищаем от неправильных символов
            prepared_df[lst_text_columns] = prepared_df[lst_text_columns].applymap(clean_text)
            all_status_prepared_df[lst_text_columns] = all_status_prepared_df[lst_text_columns].applymap(clean_text)
            union_company_df[lst_text_columns] = union_company_df[lst_text_columns].applymap(clean_text)

            union_company_df.to_excel(f'{org_folder}/Общий файл.xlsx', index=False)

            with pd.ExcelWriter(f'{end_folder}/Вакансии по региону от {current_time}.xlsx') as writer:
                prepared_df.to_excel(writer, sheet_name='Только подтвержденные вакансии', index=False)
                all_status_prepared_df.to_excel(writer, sheet_name='Вакансии со всеми статусами', index=False)

        """
            Свод по региону
            """
        # Свод по количеству рабочих мест по отраслям

        svod_vac_reg_region_df = pd.pivot_table(prepared_df,
                                                index=['Сфера деятельности'],
                                                values=['Количество рабочих мест'],
                                                aggfunc={'Количество рабочих мест': [np.sum]})
        svod_vac_reg_region_df = svod_vac_reg_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_vac_reg_region_df.sort_values(by=['sum'], ascending=False, inplace=True)
        svod_vac_reg_region_df.loc['Итого'] = svod_vac_reg_region_df['sum'].sum()
        svod_vac_reg_region_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
        svod_vac_reg_region_df = svod_vac_reg_region_df.reset_index()

        # Свод по количеству рабочих мест по организациям
        svod_vac_org_region_df = pd.pivot_table(prepared_df,
                                                index=['Полное название работодателя'],
                                                values=['Количество рабочих мест'],
                                                aggfunc={'Количество рабочих мест': [np.sum]})
        svod_vac_org_region_df = svod_vac_org_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс
        svod_vac_org_region_df.sort_values(by=['sum'], ascending=False, inplace=True)
        svod_vac_org_region_df.loc['Итого'] = svod_vac_org_region_df['sum'].sum()
        svod_vac_org_region_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
        svod_vac_org_region_df = svod_vac_org_region_df.reset_index()

        # Свод по средней и медианной минимальной зарплате для сфер деятельности
        svod_shpere_pay_region_df = pd.pivot_table(prepared_df,
                                                   index=['Сфера деятельности'],
                                                   values=['Минимальная зарплата'],
                                                   aggfunc={'Минимальная зарплата': [np.mean, np.median]},
                                                   fill_value=0)
        svod_shpere_pay_region_df = svod_shpere_pay_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_shpere_pay_region_df = svod_shpere_pay_region_df.astype(int, errors='ignore')
        svod_shpere_pay_region_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
        svod_shpere_pay_region_df = svod_shpere_pay_region_df.reset_index()

        # Свод по средней и медианной минимальной зарплате для работодателей
        svod_org_pay_region_df = pd.pivot_table(prepared_df,
                                                index=['Полное название работодателя', 'Сфера деятельности'],
                                                values=['Минимальная зарплата'],
                                                aggfunc={'Минимальная зарплата': [np.mean, np.median]},
                                                fill_value=0)
        svod_org_pay_region_df = svod_org_pay_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

        svod_shpere_educ_region_df = svod_shpere_educ_region_df.astype(int, errors='ignore')
        svod_shpere_educ_region_df.columns = ['Количество вакансий']
        svod_shpere_educ_region_df = svod_shpere_educ_region_df.reset_index()

        # Свод по требуемому образованию для работодателей
        svod_org_educ_region_df = pd.pivot_table(prepared_df,
                                                 index=['Полное название работодателя', 'Образование'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)
        svod_org_educ_region_df = svod_org_educ_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_org_educ_region_df = svod_org_educ_region_df.astype(int, errors='ignore')
        svod_org_educ_region_df.columns = ['Количество вакансий']
        svod_org_educ_region_df = svod_org_educ_region_df.reset_index()

        # Свод по графику работы для сфер деятельности
        svod_shpere_schedule_region_df = pd.pivot_table(prepared_df,
                                                        index=['Сфера деятельности', 'График работы'],
                                                        values=['Количество рабочих мест'],
                                                        aggfunc={'Количество рабочих мест': [np.sum]},
                                                        fill_value=0)
        svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.astype(int, errors='ignore')
        svod_shpere_schedule_region_df.columns = ['Количество вакансий']
        svod_shpere_schedule_region_df = svod_shpere_schedule_region_df.reset_index()

        # Свод по графику работы для работодателей
        svod_org_schedule_region_df = pd.pivot_table(prepared_df,
                                                     index=['Полное название работодателя', 'График работы'],
                                                     values=['Количество рабочих мест'],
                                                     aggfunc={'Количество рабочих мест': [np.sum]},
                                                     fill_value=0)
        svod_org_schedule_region_df = svod_org_schedule_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_org_schedule_region_df = svod_org_schedule_region_df.astype(int, errors='ignore')
        svod_org_schedule_region_df.columns = ['Количество вакансий']
        svod_org_schedule_region_df = svod_org_schedule_region_df.reset_index()

        # Свод по типу занятости для сфер деятельности
        svod_shpere_type_job_region_df = pd.pivot_table(prepared_df,
                                                        index=['Сфера деятельности', 'Тип занятости'],
                                                        values=['Количество рабочих мест'],
                                                        aggfunc={'Количество рабочих мест': [np.sum]},
                                                        fill_value=0)
        svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.astype(int, errors='ignore')
        svod_shpere_type_job_region_df.columns = ['Количество вакансий']
        svod_shpere_type_job_region_df = svod_shpere_type_job_region_df.reset_index()

        # Свод по типу занятости для работодателей
        svod_org_type_job_region_df = pd.pivot_table(prepared_df,
                                                     index=['Полное название работодателя', 'Тип занятости'],
                                                     values=['Количество рабочих мест'],
                                                     aggfunc={'Количество рабочих мест': [np.sum]},
                                                     fill_value=0)
        svod_org_type_job_region_df = svod_org_type_job_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

        svod_shpere_quote_region_df = svod_shpere_quote_region_df.astype(int, errors='ignore')
        svod_shpere_quote_region_df.columns = ['Количество вакансий']
        svod_shpere_quote_region_df = svod_shpere_quote_region_df.reset_index()

        # Свод по квотируемым местам для работодателей
        svod_org_quote_region_df = pd.pivot_table(prepared_df,
                                                  index=['Полное название работодателя', 'Квотируемое место'],
                                                  values=['Количество рабочих мест'],
                                                  aggfunc={'Количество рабочих мест': [np.sum]},
                                                  fill_value=0)
        svod_org_quote_region_df = svod_org_quote_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_org_quote_region_df = svod_org_quote_region_df.astype(int, errors='ignore')
        svod_org_quote_region_df.columns = ['Количество вакансий']
        svod_org_quote_region_df = svod_org_quote_region_df.reset_index()

        # Свод по вакансиям для социальных категорий
        dct_soc_cat_region = extract_soc_category(prepared_df, 'Социально защищенная категория',
                                                  ';')  # считаем количество категорий
        svod_soc_region_df = pd.DataFrame.from_dict(dct_soc_cat_region, orient='index').reset_index()  # содаем датафрейм
        svod_soc_region_df.columns = ['Категория', 'Количество вакансий']
        svod_soc_region_df.sort_values(by=['Количество вакансий'], ascending=False, inplace=True)

        # Свод по требуемому опыту для сфер деятельности
        svod_shpere_exp_region_df = pd.pivot_table(prepared_df,
                                                   index=['Сфера деятельности', 'Требуемый опыт работы в годах'],
                                                   values=['Количество рабочих мест'],
                                                   aggfunc={'Количество рабочих мест': [np.sum]},
                                                   fill_value=0)
        svod_shpere_exp_region_df = svod_shpere_exp_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_shpere_exp_region_df = svod_shpere_exp_region_df.astype(int, errors='ignore')
        svod_shpere_exp_region_df.columns = ['Количество вакансий']
        svod_shpere_exp_region_df = svod_shpere_exp_region_df.reset_index()

        # Свод по требуемому опыту для работодателей
        svod_org_exp_region_df = pd.pivot_table(prepared_df,
                                                index=['Полное название работодателя', 'Требуемый опыт работы в годах'],
                                                values=['Количество рабочих мест'],
                                                aggfunc={'Количество рабочих мест': [np.sum]},
                                                fill_value=0)
        svod_org_exp_region_df = svod_org_exp_region_df.droplevel(level=0, axis=1)  # убираем мультииндекс

        svod_org_exp_region_df = svod_org_exp_region_df.astype(int, errors='ignore')
        svod_org_exp_region_df.columns = ['Количество вакансий']
        svod_org_exp_region_df = svod_org_exp_region_df.reset_index()

        with pd.ExcelWriter(f'{svod_region_folder}/Свод по региону от {current_time}.xlsx') as writer:
            svod_vac_reg_region_df.to_excel(writer, sheet_name='Вакансии по отраслям', index=False)
            svod_vac_org_region_df.to_excel(writer, sheet_name='Вакансии по работодателям', index=False)
            svod_shpere_pay_region_df.to_excel(writer, sheet_name='Зарплата по отраслям', index=False)
            svod_org_pay_region_df.to_excel(writer, sheet_name='Зарплата по работодателям', index=False)
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

        """
            Свод по выбранным работодателям
            """
        if len(union_company_df) != 0:

            # Свод по вакансиям выбранных работодателей

            # Свод по количеству рабочих мест по отраслям

            svod_vac_reg_org_df = pd.pivot_table(union_company_df,
                                                 index=['Сфера деятельности'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]})
            svod_vac_reg_org_df = svod_vac_reg_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

            svod_vac_reg_org_df.sort_values(by=['sum'], ascending=False, inplace=True)
            svod_vac_reg_org_df.loc['Итого'] = svod_vac_reg_org_df['sum'].sum()
            svod_vac_reg_org_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
            svod_vac_reg_org_df = svod_vac_reg_org_df.reset_index()

            # Свод по количеству рабочих мест по организациям
            svod_vac_org_org_df = pd.pivot_table(union_company_df,
                                                 index=['Полное название работодателя'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]})
            svod_vac_org_org_df = svod_vac_org_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс
            svod_vac_org_org_df.sort_values(by=['sum'], ascending=False, inplace=True)
            svod_vac_org_org_df.loc['Итого'] = svod_vac_org_org_df['sum'].sum()
            svod_vac_org_org_df.rename(columns={'sum': 'Количество вакансий'}, inplace=True)
            svod_vac_org_org_df = svod_vac_org_org_df.reset_index()

            # Свод по средней и медианной минимальной зарплате для сфер деятельности
            svod_shpere_pay_org_df = pd.pivot_table(union_company_df,
                                                    index=['Сфера деятельности'],
                                                    values=['Минимальная зарплата'],
                                                    aggfunc={'Минимальная зарплата': [np.mean, np.median]},
                                                    fill_value=0)
            svod_shpere_pay_org_df = svod_shpere_pay_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

            svod_shpere_pay_org_df = svod_shpere_pay_org_df.astype(int, errors='ignore')
            svod_shpere_pay_org_df.columns = ['Средняя ариф. минимальная зп', 'Медианная минимальная зп']
            svod_shpere_pay_org_df = svod_shpere_pay_org_df.reset_index()

            # Свод по средней и медианной минимальной зарплате для работодателей
            svod_org_pay_org_df = pd.pivot_table(union_company_df,
                                                 index=['Полное название работодателя', 'Сфера деятельности'],
                                                 values=['Минимальная зарплата'],
                                                 aggfunc={'Минимальная зарплата': [np.mean, np.median]},
                                                 fill_value=0)
            svod_org_pay_org_df = svod_org_pay_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

            svod_shpere_educ_org_df = svod_shpere_educ_org_df.astype(int, errors='ignore')
            svod_shpere_educ_org_df.columns = ['Количество вакансий']
            svod_shpere_educ_org_df = svod_shpere_educ_org_df.reset_index()

            # Свод по требуемому образованию для работодателей
            svod_org_educ_org_df = pd.pivot_table(union_company_df,
                                                  index=['Полное название работодателя', 'Образование'],
                                                  values=['Количество рабочих мест'],
                                                  aggfunc={'Количество рабочих мест': [np.sum]},
                                                  fill_value=0)
            svod_org_educ_org_df = svod_org_educ_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

            svod_shpere_schedule_org_df = svod_shpere_schedule_org_df.astype(int, errors='ignore')
            svod_shpere_schedule_org_df.columns = ['Количество вакансий']
            svod_shpere_schedule_org_df = svod_shpere_schedule_org_df.reset_index()

            # Свод по графику работы для работодателей
            svod_org_schedule_org_df = pd.pivot_table(union_company_df,
                                                      index=['Полное название работодателя', 'График работы'],
                                                      values=['Количество рабочих мест'],
                                                      aggfunc={'Количество рабочих мест': [np.sum]},
                                                      fill_value=0)
            svod_org_schedule_org_df = svod_org_schedule_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

            svod_shpere_type_job_org_df = svod_shpere_type_job_org_df.astype(int, errors='ignore')
            svod_shpere_type_job_org_df.columns = ['Количество вакансий']
            svod_shpere_type_job_org_df = svod_shpere_type_job_org_df.reset_index()

            # Свод по типу занятости для работодателей
            svod_org_type_job_org_df = pd.pivot_table(union_company_df,
                                                      index=['Полное название работодателя', 'Тип занятости'],
                                                      values=['Количество рабочих мест'],
                                                      aggfunc={'Количество рабочих мест': [np.sum]},
                                                      fill_value=0)
            svod_org_type_job_org_df = svod_org_type_job_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

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

            svod_shpere_quote_org_df = svod_shpere_quote_org_df.astype(int, errors='ignore')
            svod_shpere_quote_org_df.columns = ['Количество вакансий']
            svod_shpere_quote_org_df = svod_shpere_quote_org_df.reset_index()

            # Свод по квотируемым местам для работодателей
            svod_org_quote_org_df = pd.pivot_table(union_company_df,
                                                   index=['Полное название работодателя', 'Квотируемое место'],
                                                   values=['Количество рабочих мест'],
                                                   aggfunc={'Количество рабочих мест': [np.sum]},
                                                   fill_value=0)
            svod_org_quote_org_df = svod_org_quote_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

            svod_org_quote_org_df = svod_org_quote_org_df.astype(int, errors='ignore')
            svod_org_quote_org_df.columns = ['Количество вакансий']
            svod_org_quote_org_df = svod_org_quote_org_df.reset_index()

            # Свод по вакансиям для социальных категорий
            dct_soc_cat_org = extract_soc_category(union_company_df, 'Социально защищенная категория',
                                                   ';')  # считаем количество категорий
            svod_soc_org_df = pd.DataFrame.from_dict(dct_soc_cat_org, orient='index').reset_index()  # содаем датафрейм
            svod_soc_org_df.columns = ['Категория', 'Количество вакансий']
            svod_soc_org_df.sort_values(by=['Количество вакансий'], ascending=False, inplace=True)

            # Свод по требуемому опыту для сфер деятельности
            svod_shpere_exp_org_df = pd.pivot_table(union_company_df,
                                                    index=['Сфера деятельности', 'Требуемый опыт работы в годах'],
                                                    values=['Количество рабочих мест'],
                                                    aggfunc={'Количество рабочих мест': [np.sum]},
                                                    fill_value=0)
            svod_shpere_exp_org_df = svod_shpere_exp_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

            svod_shpere_exp_org_df = svod_shpere_exp_org_df.astype(int, errors='ignore')
            svod_shpere_exp_org_df.columns = ['Количество вакансий']
            svod_shpere_exp_org_df = svod_shpere_exp_org_df.reset_index()

            # Свод по требуемому опыту для работодателей
            svod_org_exp_org_df = pd.pivot_table(union_company_df,
                                                 index=['Полное название работодателя', 'Требуемый опыт работы в годах'],
                                                 values=['Количество рабочих мест'],
                                                 aggfunc={'Количество рабочих мест': [np.sum]},
                                                 fill_value=0)
            svod_org_exp_org_df = svod_org_exp_org_df.droplevel(level=0, axis=1)  # убираем мультииндекс

            svod_org_exp_org_df = svod_org_exp_org_df.astype(int, errors='ignore')
            svod_org_exp_org_df.columns = ['Количество вакансий']
            svod_org_exp_org_df = svod_org_exp_org_df.reset_index()

            with pd.ExcelWriter(f'{svod_org_folder}/Свод по выбранным работодателям от {current_time}.xlsx') as writer:
                svod_vac_reg_org_df.to_excel(writer, sheet_name='Вакансии по отраслям', index=False)
                svod_vac_org_org_df.to_excel(writer, sheet_name='Вакансии по работодателям', index=False)
                svod_shpere_pay_org_df.to_excel(writer, sheet_name='Зарплата по отраслям', index=False)
                svod_org_pay_org_df.to_excel(writer, sheet_name='Зарплата по работодателям', index=False)
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
        else:
            union_company_df.to_excel(f'{svod_org_folder}/Проверьте правильность ИНН организаций {current_time}.xlsx')
    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except NotRegion:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                                 f'Не найден регион! Проверьте написание региона в соответствии с правилами сайта Работа в России')
    # except KeyError as e:
    #     messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
    #                          f'Не найдено значение {e.args}')
    except FileNotFoundError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')

    except PermissionError as e:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Закройте открытые файлы Excel {e.args}')

    else:
        messagebox.showinfo('Кассандра Подсчет данных по трудоустройству выпускников',
                            'Данные успешно обработаны.Ошибок не обнаружено')


if __name__ == '__main__':
    main_file_data = 'data/vacancy.csv'
    main_file_data = 'data/vacancy_7 (4).csv'
    main_file_data = 'data/vacancy all.csv'
    main_org_file = 'data/company.xlsx'
    main_region = 'Республика Бурятия'
    main_region = 'Томская область'
    main_region = 'Иркутская область'
    main_region = 'Город Санкт-Петербург'
    main_region = 'Кемеровская область - Кузбасс'
    main_region = 'Город Москва'


    main_end_folder = 'data'

    processing_data_trudvsem(main_file_data,main_org_file,main_end_folder,main_region)

    print('Lindy Booth !!!')