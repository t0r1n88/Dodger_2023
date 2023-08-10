import tkinter
import sys
import pandas as pd
import openpyxl
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
import time

pd.options.mode.chained_assignment = None  # default='warn'
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.simplefilter(action='ignore', category=DeprecationWarning)
warnings.simplefilter(action='ignore', category=UserWarning)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def select_folder_data():
    """
    Функция для выбора папки c данными
    :return:
    """
    global path_folder_data
    path_folder_data = filedialog.askdirectory()


def select_end_folder():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder
    path_to_end_folder = filedialog.askdirectory()


def processing_data():
    """
    Фугкция для объединения данных из всех трех листов из файлов от каждого региона
    :return:
    """
    try:
        form1_df = pd.DataFrame(columns=range(22))
        form2_df = pd.DataFrame(columns=range(11))
        form3_df = pd.DataFrame(columns=range(8))

        for file in os.listdir(path_folder_data):
            if (file.endswith('.xlsx') and not file.startswith('~$')):
                print(file)
                # обрабатываем первый лист
                temp_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name=0, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:V')
                temp_df.columns = range(22)
                temp_df.dropna(thresh=4, inplace=True)
                form1_df = pd.concat([form1_df, temp_df], ignore_index=True)

                # обрабатываем второй лист
                temp_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name=1, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:K')
                temp_df.columns = range(11)
                temp_df.dropna(thresh=4, inplace=True)
                form2_df = pd.concat([form2_df, temp_df], ignore_index=True)

                # обрабатываем третий лист
                temp_df = pd.read_excel(f'{path_folder_data}/{file}', sheet_name=2, dtype={'гр.2': str}, skiprows=2,
                                        usecols='A:H')
                temp_df.columns = range(8)
                temp_df.dropna(thresh=3, inplace=True)
                form3_df = pd.concat([form3_df, temp_df], ignore_index=True)

        # генерируем текущее время
        t = time.localtime()
        current_time = time.strftime('%H_%M_%S', t)

        with pd.ExcelWriter(f'{path_to_end_folder}/Общий файл от {current_time}.xlsx') as writer:
            form1_df.to_excel(writer, sheet_name='Форма по мониторингу', index=False)
            form2_df.to_excel(writer, sheet_name='Форма по принимаемым мерам', index=False)
            form3_df.to_excel(writer, sheet_name='Форма по социальной поддежке', index=False)
    except NameError:
            messagebox.showerror('Морриган Объединение данных по ОПК ver 1.0',
                                 f'Выберите файлы с данными и папку куда будет генерироваться файл')
    except KeyError as e:
        messagebox.showerror('Морриган Объединение данных по ОПК ver 1.0',
                             f'Не найдено значение {e.args}')
    except FileNotFoundError:
        messagebox.showerror('Морриган Объединение данных по ОПК ver 1.0',
                             f'Перенесите файлы которые вы хотите обработать в корень диска. Проблема может быть\n '
                             f'в слишком длинном пути к обрабатываемым файлам')
    except PermissionError as e:
        messagebox.showerror('Морриган Объединение данных по ОПК ver 1.0',
                             f'Закройте открытые файлы Excel {e.args}')
    else:
        messagebox.showinfo('Морриган Объединение данных по ОПК ver 1.0',
                            'Данные успешно обработаны.')

if __name__ == '__main__':
    window = Tk()
    window.title('Морриган Объединение данных по ОПК ver 1.0')
    window.geometry('700x860')
    window.resizable(False, False)

    # Создаем объект вкладок

    tab_control = ttk.Notebook(window)

    # Создаем вкладку обработки данных для Приложения 6
    tab_union_1 = ttk.Frame(tab_control)
    tab_control.add(tab_union_1, text='Скрипт №1')
    tab_control.pack(expand=1, fill='both')
    # Добавляем виджеты на вкладку Создание образовательных программ
    # Создаем метку для описания назначения программы
    lbl_hello = Label(tab_union_1,
                      text='Центр опережающей профессиональной подготовки Республики Бурятия')
    lbl_hello.grid(column=0, row=0, padx=10, pady=25)

    # Картинка
    path_to_img = resource_path('logo.png')

    img = PhotoImage(file=path_to_img)
    Label(tab_union_1,
          image=img
          ).grid(column=1, row=0, padx=10, pady=25)

    # Создаем кнопку Выбрать файл с данными
    btn_choose_data = Button(tab_union_1, text='1) Выберите папку с данными', font=('Arial Bold', 20),
                             command=select_folder_data
                             )
    btn_choose_data.grid(column=0, row=2, padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder = Button(tab_union_1, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder
                                   )
    btn_choose_end_folder.grid(column=0, row=3, padx=10, pady=10)

    # Создаем кнопку обработки данных

    btn_proccessing_data = Button(tab_union_1, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_data
                                  )
    btn_proccessing_data.grid(column=0, row=4, padx=10, pady=10)

    window.mainloop()
