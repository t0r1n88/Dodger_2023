# -*- coding: utf-8 -*-
"""
Графический интерфейс для программы Кассандра
"""
from base_employment import prepare_base_employment  # импортируем функцию для обработки мониторинга 5 строк
from nose_employment import prepare_nose_employment  # импортируем функцию для обработки нозологий 15 строк
from ck_employment import prepare_ck_employment  # импортируем функцию для обработки данных для отчета центров карьеры
from opk_employment import prepare_opk_employment  # импортируем функцию для обработки данных по ОПК
from difference import prepare_diffrence  # импортируем функцию для нахождения разницы между двумя таблицами

import pandas as pd
import os
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk

pd.options.mode.chained_assignment = None  # default='warn'
import warnings

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
warnings.filterwarnings('ignore', category=DeprecationWarning)
warnings.filterwarnings('ignore', category=FutureWarning)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller
    Функция чтобы логотип отображался"""
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


def select_folder_data_base():
    """
    Функция для выбора папки c данными базового мониторинга 5 строк
    :return:
    """
    global path_folder_data_base
    path_folder_data_base = filedialog.askdirectory()


def select_end_folder_base():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы базового мониторинга 5 строк
    :return:
    """
    global path_to_end_folder_base
    path_to_end_folder_base = filedialog.askdirectory()

def select_folder_data_nose():
    """
    Функция для выбора папки c данными базового мониторинга 5 строк
    :return:
    """
    global path_folder_data_nose
    path_folder_data_nose = filedialog.askdirectory()


def select_end_folder_nose():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы базового мониторинга 5 строк
    :return:
    """
    global path_to_end_folder_nose
    path_to_end_folder_nose = filedialog.askdirectory()


# для обработки отчетов ЦК
def select_folder_data_ck():
    """
    Функция для выбора папки c данными
    :return:
    """
    global path_folder_data_ck
    path_folder_data_ck = filedialog.askdirectory()


def select_end_folder_ck():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_ck
    path_to_end_folder_ck = filedialog.askdirectory()


def select_folder_data_opk():
    """
    Функция для выбора папки c данными
    :return:
    """
    global path_folder_data_opk
    path_folder_data_opk = filedialog.askdirectory()


def select_end_folder_opk():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_opk
    path_to_end_folder_opk = filedialog.askdirectory()


def select_files_data_xlsx():
    """
    Функция для выбора нескоьких файлов с данными на основе которых будет генерироваться документ
    :return: Путь к файлу с данными
    """
    global files_data_xlsx
    # Получаем путь файлы
    files_data_xlsx = filedialog.askopenfilenames(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


"""
Функция обработки файлов базового мониторинга 5 строк
"""


def processing_base_employment():
    """
    Фугкция для обработки данных мониторинга 5 строк базовый мониторинг
    :return: файлы Excel  с результатами обработки
    """
    try:
        prepare_base_employment(path_folder_data_base, path_to_end_folder_base)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


"""
обработка таблицы нозологии 15 строк
"""


def processing_nose_employment():
    """
    Функция для обработки данных мониторинга нозологий 15 строк
    :return:
    """
    try:
        prepare_nose_employment(path_folder_data_nose, path_to_end_folder_nose)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


"""
Обработка отчетов ЦК
"""


def processing_ck_employment():
    """
    Функция для обработки отчетов центров карьеры
    :return:
    """
    try:
        prepare_ck_employment(path_folder_data_ck, path_to_end_folder_ck)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_opk_employment():
    """
    Функция для обработки полной таблицы занятости выпускников в ОПК
    """
    try:
        prepare_opk_employment(path_folder_data_opk, path_to_end_folder_opk)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


"""
Функции для нахождения разницы между 2 таблицами
"""


def select_first_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_first_diffrence
    # Получаем путь к файлу
    data_first_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_second_diffrence():
    """
    Функция для файла с данными
    :return: Путь к файлу с данными
    """
    global data_second_diffrence
    # Получаем путь к файлу
    data_second_diffrence = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_diffrence():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_diffrence
    path_to_end_folder_diffrence = filedialog.askdirectory()


def processing_diffrence():
    """
    Функция для вычисления разницы между двумя таблицами
    """
    try:
        dif_first_sheet_name = entry_first_sheet_name_diffrence.get()
        dif_second_sheet_name = entry_second_sheet_name_diffrence.get()

        prepare_diffrence(data_first_diffrence, dif_first_sheet_name, data_second_diffrence, dif_second_sheet_name,
                          path_to_end_folder_diffrence)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

def on_scroll(*args):
    canvas.yview(*args)

def set_window_size(window):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    # Устанавливаем размер окна в 80% от ширины и высоты экрана
    if screen_width >= 3840:
        width = int(screen_width * 0.2)
    elif screen_width >= 2560:
        width = int(screen_width * 0.31)
    elif screen_width >= 1920:
        width = int(screen_width * 0.41)
    elif screen_width >= 1600:
        width = int(screen_width * 0.5)
    elif screen_width >= 1280:
        width = int(screen_width * 0.62)
    elif screen_width >= 1024:
        width = int(screen_width * 0.77)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.8)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")



if __name__ == '__main__':
    window = Tk()
    window.title('Кассандра Подсчет данных по трудоустройству выпускников ver 3.6')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    # window.geometry('774x760')
    # window.geometry('980x910+700+100')
    window.resizable(True, True)

    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    """
    Создаем вкладку для обработки отчета 5 строк
    """
    tab_employment_five = ttk.Frame(tab_control)
    tab_control.add(tab_employment_five, text='Подсчет 5 строк')

    employment_five_frame_description = LabelFrame(tab_employment_five)
    employment_five_frame_description.pack()

    lbl_hello_employment_five = Label(employment_five_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Трудоустройство выпускников.\n Подсчет по специальностям/профессиям 5 строк', width=60)
    lbl_hello_employment_five.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_employment_five = resource_path('logo.png')
    img_employment_five = PhotoImage(file=path_to_img_employment_five)
    Label(employment_five_frame_description,
          image=img_employment_five, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_employment_five = LabelFrame(tab_employment_five, text='Подготовка')
    frame_data_employment_five.pack(padx=10, pady=10)

    btn_choose_data_employment_five = Button(frame_data_employment_five, text='1) Выберите папку с данными', font=('Arial Bold', 20),
                             command=select_folder_data_base
                             )
    btn_choose_data_employment_five.pack(padx=10, pady=10)


    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_employment_five = Button(frame_data_employment_five, text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                   command=select_end_folder_base
                                   )
    btn_choose_end_folder_employment_five.pack(padx=10, pady=10)
    #
    # Создаем кнопку обработки данных

    btn_proccessing_data_employment_five = Button(tab_employment_five, text='3) Обработать данные', font=('Arial Bold', 20),
                                  command=processing_base_employment
                                  )
    btn_proccessing_data_employment_five.pack(padx=10, pady=10)


    """
    Вкладка для обработки формы №15 нозологии
    """
    tab_employment_nose = ttk.Frame(tab_control)
    tab_control.add(tab_employment_nose, text='Подсчет нозология 15 строк')

    employment_nose_frame_description = LabelFrame(tab_employment_nose)
    employment_nose_frame_description.pack()

    lbl_hello_employment_nose = Label(employment_nose_frame_description,
                                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                           'Трудоустройство выпускников. Подсчет по специальностям/профессиям\n нозологии 15 строк',
                                      width=60)
    lbl_hello_employment_nose.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_employment_nose = resource_path('logo.png')
    img_employment_nose = PhotoImage(file=path_to_img_employment_nose)
    Label(employment_nose_frame_description,
          image=img_employment_nose, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_employment_nose = LabelFrame(tab_employment_nose, text='Подготовка')
    frame_data_employment_nose.pack(padx=10, pady=10)

    btn_choose_data_employment_nose = Button(frame_data_employment_nose, text='1) Выберите папку с данными',
                                             font=('Arial Bold', 20),
                                             command=select_folder_data_nose
                                             )
    btn_choose_data_employment_nose.pack(padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_employment_nose = Button(frame_data_employment_nose, text='2) Выберите конечную папку',
                                                   font=('Arial Bold', 20),
                                                   command=select_end_folder_nose
                                                   )
    btn_choose_end_folder_employment_nose.pack(padx=10, pady=10)
    #
    # Создаем кнопку обработки данных

    btn_proccessing_data_employment_nose = Button(tab_employment_nose, text='3) Обработать данные',
                                                  font=('Arial Bold', 20),
                                                  command=processing_nose_employment
                                                  )
    btn_proccessing_data_employment_nose.pack(padx=10, pady=10)
    #
    # """
    # Вкладка для обработки отчетов центров карьеры
    # """
    # # Создаем вкладку обработки отчетов центров карьеры
    # tab_ck_employment = ttk.Frame(tab_control)
    # tab_control.add(tab_ck_employment, text='Отчет ЦК')
    # tab_control.pack(expand=1, fill='both')
    # # Добавляем виджеты на вкладку
    # # Создаем метку для описания назначения программы
    # lbl_hello_ck = Label(tab_ck_employment,
    #                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
    #                           'Обработка данных центров карьеры по трудоустроенным выпускникам')
    # lbl_hello_ck.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Картинка
    # path_to_img_ck = resource_path('logo.png')
    #
    # img_ck = PhotoImage(file=path_to_img_ck)
    # Label(tab_ck_employment,
    #       image=img_ck
    #       ).grid(column=1, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_choose_ck_data = Button(tab_ck_employment, text='1) Выберите папку с данными', font=('Arial Bold', 20),
    #                             command=select_folder_data_ck
    #                             )
    # btn_choose_ck_data.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_ck_folder = Button(tab_ck_employment, text='2) Выберите конечную папку', font=('Arial Bold', 20),
    #                                   command=select_end_folder_ck
    #                                   )
    # btn_choose_end_ck_folder.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку обработки данных
    #
    # btn_proccessing_ck_data = Button(tab_ck_employment, text='3) Обработать данные', font=('Arial Bold', 20),
    #                                  command=processing_ck_employment
    #                                  )
    # btn_proccessing_ck_data.grid(column=0, row=4, padx=10, pady=10)
    #
    # """
    # Подсчет данных по трудоустройству ОПК
    # """
    # # Создаем вкладку обработки отчетов центров карьеры
    # tab_opk_employment = ttk.Frame(tab_control)
    # tab_control.add(tab_opk_employment, text='Отчет ОПК с отраслями')
    # tab_control.pack(expand=1, fill='both')
    # # Добавляем виджеты на вкладку
    # # Создаем метку для описания назначения программы
    # lbl_hello_opk = Label(tab_opk_employment,
    #                       text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
    #                            'Обработка данных по трудоустройству ОПК (по отраслям)\n'
    #                            'В обрабатываемых файлах должны быть листы Форма 1 и Форма 2,\n'
    #                            'В Форме 1 должно быть 80 колонок включая 2 колонки проверки\n'
    #                            ',внизу после окончания таблицы должна быть пустая строка.\n'
    #                            ' На 9 строке должна быть строка с номерами колонок.'
    #                            'В форме 2 должно быть 10 колонок')
    # lbl_hello_opk.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Картинка
    # path_to_img_opk = resource_path('logo.png')
    #
    # img_opk = PhotoImage(file=path_to_img_opk)
    # Label(tab_opk_employment,
    #       image=img_opk
    #       ).grid(column=1, row=0, padx=10, pady=25)
    #
    # # Создаем кнопку Выбрать файл с данными
    # btn_choose_opk_data = Button(tab_opk_employment, text='1) Выберите папку с данными', font=('Arial Bold', 20),
    #                              command=select_folder_data_opk
    #                              )
    # btn_choose_opk_data.grid(column=0, row=2, padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_opk_folder = Button(tab_opk_employment, text='2) Выберите конечную папку', font=('Arial Bold', 20),
    #                                    command=select_end_folder_opk
    #                                    )
    # btn_choose_end_opk_folder.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Создаем кнопку обработки данных
    #
    # btn_proccessing_opk_data = Button(tab_opk_employment, text='3) Обработать данные', font=('Arial Bold', 20),
    #                                   command=processing_opk_employment
    #                                   )
    # btn_proccessing_opk_data.grid(column=0, row=4, padx=10, pady=10)
    #
    # """
    # Разница двух таблиц
    # """
    # tab_diffrence = ttk.Frame(tab_control)
    # tab_control.add(tab_diffrence, text='Разница 2 таблиц')
    # tab_control.pack(expand=1, fill='both')
    #
    # # Добавляем виджеты на вкладку разница 2 двух таблиц
    # # Создаем метку для описания назначения программы
    # lbl_hello = Label(tab_diffrence,
    #                   text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
    #                        'Количество строк и колонок в таблицах должно совпадать\n'
    #                        'Названия колонок в таблицах должны совпадать'
    #                        '\nДля корректной работы программмы уберите из таблицы объединенные ячейки')
    # lbl_hello.grid(column=0, row=0, padx=10, pady=25)
    #
    # # Картинка
    # path_com = resource_path('logo.png')
    # img_diffrence = PhotoImage(file=path_com)
    # Label(tab_diffrence,
    #       image=img_diffrence
    #       ).grid(column=1, row=0, padx=10, pady=25)
    #
    # # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    # frame_data_for_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    # frame_data_for_diffrence.grid(column=0, row=2, padx=10)
    #
    # # Создаем кнопку Выбрать  первый файл с данными
    # btn_data_first_diffrence = Button(frame_data_for_diffrence, text='1) Выберите файл с первой таблицей',
    #                                   font=('Arial Bold', 10),
    #                                   command=select_first_diffrence
    #                                   )
    # btn_data_first_diffrence.grid(column=0, row=3, padx=10, pady=10)
    #
    # # Определяем текстовую переменную
    # entry_first_sheet_name_diffrence = StringVar()
    # # Описание поля
    # label_first_sheet_name_diffrence = Label(frame_data_for_diffrence,
    #                                          text='2) Введите название листа, где находится первая таблица')
    # label_first_sheet_name_diffrence.grid(column=0, row=4, padx=10, pady=10)
    # # поле ввода имени листа
    # first_sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_first_sheet_name_diffrence,
    #                                          width=30)
    # first_sheet_name_entry_diffrence.grid(column=0, row=5, padx=5, pady=5, ipadx=15, ipady=10)
    #
    # # Создаем кнопку Выбрать  второй файл с данными
    # btn_data_second_diffrence = Button(frame_data_for_diffrence, text='3) Выберите файл со второй таблицей',
    #                                    font=('Arial Bold', 10),
    #                                    command=select_second_diffrence
    #                                    )
    # btn_data_second_diffrence.grid(column=0, row=6, padx=10, pady=10)
    #
    # # Определяем текстовую переменную
    # entry_second_sheet_name_diffrence = StringVar()
    # # Описание поля
    # label_second_sheet_name_diffrence = Label(frame_data_for_diffrence,
    #                                           text='4) Введите название листа, где находится вторая таблица')
    # label_second_sheet_name_diffrence.grid(column=0, row=7, padx=10, pady=10)
    # # поле ввода
    # second__sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_second_sheet_name_diffrence,
    #                                            width=30)
    # second__sheet_name_entry_diffrence.grid(column=0, row=8, padx=5, pady=5, ipadx=15, ipady=10)
    #
    # # Создаем кнопку выбора папки куда будет генерироваьться файл
    # btn_select_end_diffrence = Button(frame_data_for_diffrence, text='5) Выберите конечную папку',
    #                                   font=('Arial Bold', 10),
    #                                   command=select_end_folder_diffrence
    #                                   )
    # btn_select_end_diffrence.grid(column=0, row=10, padx=10, pady=10)
    #
    # # Создаем кнопку Обработать данные
    # btn_data_do_diffrence = Button(tab_diffrence, text='6) Обработать таблицы', font=('Arial Bold', 20),
    #                                command=processing_diffrence
    #                                )
    # btn_data_do_diffrence.grid(column=0, row=11, padx=10, pady=10)


    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.mainloop()
