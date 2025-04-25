# -*- coding: utf-8 -*-
"""
Графический интерфейс для программы Кассандра
"""
from form_one_five_row import prepare_form_one_employment  # импортируем функцию для обработки Формы 1 (5 строк)
from form_two_fifteen_row_nose import prepare_form_two_employment  # импортируем функцию для обработки Форма 2 нозологии 15 строк
from form_three_expected_release import prepare_form_three_employment # импортируем функцию для обработки Формы 3 Ожидаемый выпуск
from monitoring_graduate_employment import prepare_graduate_employment
from ck_employment import prepare_ck_employment  # импортируем функцию для обработки данных для отчета центров карьеры
from opk_employment import prepare_opk_employment  # импортируем функцию для обработки данных по ОПК
from create_svod_trudvsem import processing_data_trudvsem # импортируем функцию для обработки данных с трудвсем
from contrast_svod_trudvsem import prepare_diff_svod_trudvsem # импортируем функцию для измерения разницы между двумя сводами
from cass_difference import prepare_diffrence  # импортируем функцию для нахождения разницы между двумя таблицами
import sys
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

def select_folder_data_form_three():
    """
    Функция для выбора папки c данными базового мониторинга 5 строк
    :return:
    """
    global path_folder_data_form_three
    path_folder_data_form_three = filedialog.askdirectory()


def select_end_folder_form_three():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы базового мониторинга 5 строк
    :return:
    """
    global path_to_end_folder_form_three
    path_to_end_folder_form_three = filedialog.askdirectory()


def select_folder_data_mon_grad():
    """
    Функция для выбора папки c данными мониторинга выпускников для СССР
    :return:
    """
    global path_folder_data_mon_grad
    path_folder_data_mon_grad = filedialog.askdirectory()


def select_end_folder_mon_grad():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы мониторинга выпускников для СССР
    :return:
    """
    global path_to_end_folder_mon_grad
    path_to_end_folder_mon_grad = filedialog.askdirectory()






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
Вспомогательные функции для обработки даннных с Работы в России
"""
def select_file_csv_trudvsem():
    """
    Функция для выбора файла csv
    """
    global file_csv_svod_trudvsem
    # Получаем путь к файлу
    file_csv_svod_trudvsem = filedialog.askopenfilename(filetypes=(('csv files', '*.csv'), ('all files', '*.*')))


def select_file_org_trudvsem():
    """
    Функция для выбора файла с организациями
    """
    global file_org_svod_trudvsem
    # Получаем путь к файлу
    file_org_svod_trudvsem = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_file_params():
    """
    Функция для выбора файла с параметрами обработки
    """
    global name_file_params
    # Получаем путь к файлу
    name_file_params = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_svod_trudvsem():
    """
    Функия для выбора папки.Определенно вот это когда нибудь я перепишу на ООП
    :return:
    """
    global path_to_end_folder_svod_trudvsem
    path_to_end_folder_svod_trudvsem = filedialog.askdirectory()


# Функции для изменений по сводам
def select_first_file_diff_svod_trudvsem():
    """
    Функция для выбора файла с организациями
    """
    global file_frist_diff_svod_trudvsem
    # Получаем путь к файлу
    file_frist_diff_svod_trudvsem = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))

def select_second_file_diff_svod_trudvsem():
    """
    Функция для выбора файла с организациями
    """
    global file_second_diff_svod_trudvsem
    # Получаем путь к файлу
    file_second_diff_svod_trudvsem = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


def select_end_folder_diff_svod_trudvsem():
    """
    Функия для выбора папки
    :return:
    """
    global path_to_end_folder_diff_svod_trudvsem
    path_to_end_folder_diff_svod_trudvsem = filedialog.askdirectory()






"""
Функция обработки формы 1 (пятистрочная)
"""


def processing_base_employment():
    """
    Фугкция для обработки данных мониторинга 5 строк базовый мониторинг
    :return: файлы Excel  с результатами обработки
    """
    try:
        prepare_form_one_employment(path_folder_data_base, path_to_end_folder_base)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


"""
обработка Формы 2 нозологии 15 строк
"""

def processing_nose_employment():
    """
    Функция для обработки данных мониторинга нозологий 15 строк
    :return:
    """
    try:
        prepare_form_two_employment(path_folder_data_nose, path_to_end_folder_nose)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_form_three_employment():
    """
    Функция для обработки данных мониторинга Форма 3 Ожидаемый выпуск
    :return:
    """
    try:
        prepare_form_three_employment(path_folder_data_form_three, path_to_end_folder_form_three)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')



def processing_mon_grad():
    """
    Функция для обработки данных мониторинга занятости выпускников 2024 для сайта СССР
    :return:
    """
    try:
        prepare_graduate_employment(path_folder_data_mon_grad, path_to_end_folder_mon_grad)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')

"""
Создание свода Работа в России
"""

def processing_svod_trudvsem():
    """
    Функция для обработки данных с сайта Работа в России
    :return:
    """
    try:
        name_region = str(entry_region.get()) # Получаем название региона
        processing_data_trudvsem(file_csv_svod_trudvsem, file_org_svod_trudvsem,path_to_end_folder_svod_trudvsem,name_region,name_file_params)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_diff_svod_trudvsem():
    """
    Функция для подсчета изменений
    """
    try:
        type_contrast = mode_region_value.get() # получаем значение режима сравнения
        prepare_diff_svod_trudvsem(file_frist_diff_svod_trudvsem, file_second_diff_svod_trudvsem,path_to_end_folder_diff_svod_trudvsem,type_contrast)

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

    height = int(screen_height * 0.6)

    # Рассчитываем координаты для центрирования окна
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2

    # Устанавливаем размер и положение окна
    window.geometry(f"{width}x{height}+{x}+{y}")

def make_textmenu(root):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # эта штука делает меню
    global the_menu
    the_menu = Menu(root, tearoff=0)
    the_menu.add_command(label="Вырезать")
    the_menu.add_command(label="Копировать")
    the_menu.add_command(label="Вставить")
    the_menu.add_separator()
    the_menu.add_command(label="Выбрать все")


def callback_select_all(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    # select text after 50ms
    window.after(50, lambda: event.widget.select_range(0, 'end'))


def show_textmenu(event):
    """
    Функции для контекстного меню( вырезать,копировать,вставить)
    взято отсюда https://gist.github.com/angeloped/91fb1bb00f1d9e0cd7a55307a801995f
    """
    e_widget = event.widget
    the_menu.entryconfigure("Вырезать", command=lambda: e_widget.event_generate("<<Cut>>"))
    the_menu.entryconfigure("Копировать", command=lambda: e_widget.event_generate("<<Copy>>"))
    the_menu.entryconfigure("Вставить", command=lambda: e_widget.event_generate("<<Paste>>"))
    the_menu.entryconfigure("Выбрать все", command=lambda: e_widget.select_range(0, 'end'))
    the_menu.tk.call("tk_popup", the_menu, event.x_root, event.y_root)


if __name__ == '__main__':
    window = Tk()
    window.title('Кассандра Подсчет данных по трудоустройству выпускников ver 6.5')
    # Устанавливаем размер и положение окна
    set_window_size(window)
    window.resizable(True, True)
    # Добавляем контекстное меню в поля ввода
    make_textmenu(window)
    # Создаем вертикальный скроллбар
    scrollbar = Scrollbar(window, orient="vertical")

    # Создаем холст
    canvas = Canvas(window, yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True)

    # Привязываем скроллбар к холсту
    scrollbar.config(command=canvas.yview)

    # Создаем ноутбук (вкладки)
    tab_control = ttk.Notebook(canvas)

    global name_file_params
    name_file_params = 'Не выбрано' # костыль конечно но что поделать

    global file_org_svod_trudvsem
    file_org_svod_trudvsem = 'Не выбрано' # костыль


    """
    Создаем вкладку для обработки формы 1 (пятистрочная)
    """
    tab_employment_five = ttk.Frame(tab_control)
    tab_control.add(tab_employment_five, text='Форма 1\n (5 строк)')

    employment_five_frame_description = LabelFrame(tab_employment_five)
    employment_five_frame_description.pack()

    lbl_hello_employment_five = Label(employment_five_frame_description,
                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                           'Трудоустройство выпускников. Форма 1 пятистрочная', width=60)
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
    Вкладка для обработки формы 2 нозологии 15 строк
    """
    tab_employment_nose = ttk.Frame(tab_control)
    tab_control.add(tab_employment_nose, text='Форма 2\n (нозологии)')

    employment_nose_frame_description = LabelFrame(tab_employment_nose)
    employment_nose_frame_description.pack()

    lbl_hello_employment_nose = Label(employment_nose_frame_description,
                                      text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                           'Трудоустройство выпускников. Форма 2 нозологии (15 строк)\n'
                                           'Форма мониторинга январь 2025',
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

    """
    Вкладка для обработки формы 3 Ожидаемый выпуск
    """
    tab_employment_expect = ttk.Frame(tab_control)
    tab_control.add(tab_employment_expect, text='Форма 3\n (ожидаемый выпуск)')

    employment_expect_frame_description = LabelFrame(tab_employment_expect)
    employment_expect_frame_description.pack()

    lbl_hello_employment_expect = Label(employment_expect_frame_description,
                                        text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                             'Трудоустройство выпускников. Форма 3 Ожидаемый выпуск',
                                        width=60)
    lbl_hello_employment_expect.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_employment_expect = resource_path('logo.png')
    img_employment_expect = PhotoImage(file=path_to_img_employment_expect)
    Label(employment_expect_frame_description,
          image=img_employment_expect, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_employment_expect = LabelFrame(tab_employment_expect, text='Подготовка')
    frame_data_employment_expect.pack(padx=10, pady=10)

    btn_choose_data_employment_expect = Button(frame_data_employment_expect, text='1) Выберите папку с данными',
                                               font=('Arial Bold', 20),
                                               command=select_folder_data_form_three
                                               )
    btn_choose_data_employment_expect.pack(padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_employment_expect = Button(frame_data_employment_expect, text='2) Выберите конечную папку',
                                                     font=('Arial Bold', 20),
                                                     command=select_end_folder_form_three
                                                     )
    btn_choose_end_folder_employment_expect.pack(padx=10, pady=10)
    #
    # Создаем кнопку обработки данных

    btn_proccessing_data_employment_expect = Button(tab_employment_expect, text='3) Обработать данные',
                                                    font=('Arial Bold', 20),
                                                    command=processing_form_three_employment
                                                    )
    btn_proccessing_data_employment_expect.pack(padx=10, pady=10)

    """
    Создаем вкладку для обработки мониторинга занятости выпускников для сайта СССР
    """


    tab_employment_grad_mon = ttk.Frame(tab_control)
    tab_control.add(tab_employment_grad_mon, text='Мониторинг занятости\nвыпускников')

    employment_grad_mon_frame_description = LabelFrame(tab_employment_grad_mon)
    employment_grad_mon_frame_description.pack()

    lbl_hello_employment_grad_mon = Label(employment_grad_mon_frame_description,
                                          text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                               'Мониторинг занятости выпускников 2024 для сервиса\n'
                                               '«Система сбора и синхронизации ресурсов» (https://data.firpo.ru).',
                                          width=60)
    lbl_hello_employment_grad_mon.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_employment_grad_mon = resource_path('logo.png')
    img_employment_grad_mon = PhotoImage(file=path_to_img_employment_grad_mon)
    Label(employment_grad_mon_frame_description,
          image=img_employment_grad_mon, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_employment_grad_mon = LabelFrame(tab_employment_grad_mon, text='Подготовка')
    frame_data_employment_grad_mon.pack(padx=10, pady=10)

    btn_choose_data_employment_grad_mon = Button(frame_data_employment_grad_mon, text='1) Выберите папку с данными',
                                                 font=('Arial Bold', 20),
                                                 command=select_folder_data_mon_grad
                                                 )
    btn_choose_data_employment_grad_mon.pack(padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_employment_grad_mon = Button(frame_data_employment_grad_mon,
                                                       text='2) Выберите конечную папку',
                                                       font=('Arial Bold', 20),
                                                       command=select_end_folder_mon_grad
                                                       )
    btn_choose_end_folder_employment_grad_mon.pack(padx=10, pady=10)
    #
    # Создаем кнопку обработки данных

    btn_proccessing_data_employment_grad_mon = Button(tab_employment_grad_mon, text='3) Обработать данные',
                                                      font=('Arial Bold', 20),
                                                      command=processing_mon_grad
                                                      )
    btn_proccessing_data_employment_grad_mon.pack(padx=10, pady=10)





    """
    Вкладка для создания свода из данных с сайта Работа в России
    """
    tab_svod_trudvsem = ttk.Frame(tab_control)
    tab_control.add(tab_svod_trudvsem, text='Свод Работа в России')

    svod_trudvsem_frame_description = LabelFrame(tab_svod_trudvsem)
    svod_trudvsem_frame_description.pack()

    lbl_hello_svod_trudvsem = Label(svod_trudvsem_frame_description,
                                    text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                         'Аналитика по кадровой ситуации в регионе на основании данных \n'
                                         'с сайта Работа в России (trudvsem.ru) https://trudvsem.ru/opendata/datasets',
                                    width=60)
    lbl_hello_svod_trudvsem.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_svod_trudvsem = resource_path('logo.png')
    img_svod_trudvsem = PhotoImage(file=path_to_img_svod_trudvsem)
    Label(svod_trudvsem_frame_description,
          image=img_svod_trudvsem, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_svod_trudvsem = LabelFrame(tab_svod_trudvsem, text='Подготовка')
    frame_data_svod_trudvsem.pack(padx=10, pady=10)

    # Кнопка для выбора файла csv
    btn_choose_data_svod_trudvsem = Button(frame_data_svod_trudvsem, text='1) Выберите файл скачанный с сайта\n'
                                                                          'Работа в России в формате csv',
                                           font=('Arial Bold', 15),
                                           command=select_file_csv_trudvsem
                                           )
    btn_choose_data_svod_trudvsem.pack(padx=10, pady=10)

    # Создаем кнопку для выбора файла с работодателями

    btn_choose_data_org_svod_trudvsem = Button(frame_data_svod_trudvsem, text='Необязательная опция\n Выберите файл с работодателями',
                                                 font=('Arial Bold', 10),
                                                 command=select_file_org_trudvsem
                                                 )
    btn_choose_data_org_svod_trudvsem.pack(padx=10, pady=10)
    #
    # Создаем поле для ввода региона

    # Определяем текстовую переменную
    entry_region = StringVar()
    # Описание поля
    label_svod_trudvsem = Label(frame_data_svod_trudvsem,
                                             text='2) Введите название региона')
    label_svod_trudvsem.pack(padx=10, pady=10)
    # поле ввода имени листа
    svod_trudvsem_entry = Entry(frame_data_svod_trudvsem, textvariable=entry_region,
                                             width=30)
    svod_trudvsem_entry.pack(padx=10, pady=10)

    # Кнопка для выбора файла с параметрами
    btn_choose_params_svod_trudvsem = Button(frame_data_svod_trudvsem, text='Необязательная опция\n Выберите файл с параметрами фильтрации',
                                                 font=('Arial Bold', 10),
                                                 command=select_file_params
                                                 )
    btn_choose_params_svod_trudvsem.pack(padx=10, pady=10)



    # Кнопка для выбора конечной папки
    btn_choose_end_folder_svod_trudvsem = Button(frame_data_svod_trudvsem, text='3) Выберите конечную папку',
                                                 font=('Arial Bold', 15),
                                                 command=select_end_folder_svod_trudvsem
                                                 )
    btn_choose_end_folder_svod_trudvsem.pack(padx=10, pady=10)



    btn_proccessing_data_svod_trudvsem = Button(tab_svod_trudvsem, text='4) Обработать данные',
                                                font=('Arial Bold', 20),
                                                command=processing_svod_trudvsem
                                                )
    btn_proccessing_data_svod_trudvsem.pack(padx=10, pady=10)


    """
    Вкладка для подсчета разницы в сводах
    """
    tab_diff_svod_trudvsem = ttk.Frame(tab_control)
    tab_control.add(tab_diff_svod_trudvsem, text='Динамика Работа в России')

    svod_trudvsem_frame_description = LabelFrame(tab_diff_svod_trudvsem)
    svod_trudvsem_frame_description.pack()

    lbl_hello_diff_svod_trudvsem = Label(svod_trudvsem_frame_description,
                                         text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                              'Подсчет изменений по кадровой ситуации на основании сводов \n'
                                              'полученных с помощью вкладки Свод Работа в России',
                                         width=60)
    lbl_hello_diff_svod_trudvsem.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_diff_svod_trudvsem = resource_path('logo.png')
    img_diff_svod_trudvsem = PhotoImage(file=path_to_img_diff_svod_trudvsem)
    Label(svod_trudvsem_frame_description,
          image=img_diff_svod_trudvsem, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_diff_svod_trudvsem = LabelFrame(tab_diff_svod_trudvsem, text='Подготовка')
    frame_data_diff_svod_trudvsem.pack(padx=10, pady=10)

    # Кнопка для выбора первого файла
    btn_choose_first_file_diff_svod_trudvsem = Button(frame_data_diff_svod_trudvsem,
                                                text='1) Выберите первый файл',
                                                font=('Arial Bold', 15),
                                                command=select_first_file_diff_svod_trudvsem
                                                )
    btn_choose_first_file_diff_svod_trudvsem.pack(padx=10, pady=10)

    # Создаем кнопку для выбора второго файла

    btn_choose_second_file_diff_svod_trudvsem = Button(frame_data_diff_svod_trudvsem,
                                                    text='2) Выберите второй файл',
                                                    font=('Arial Bold', 15),
                                                    command=select_second_file_diff_svod_trudvsem
                                                    )
    btn_choose_second_file_diff_svod_trudvsem.pack(padx=10, pady=10)


    # Кнопка для выбора конечной папки
    btn_choose_end_folder_diff_svod_trudvsem = Button(frame_data_diff_svod_trudvsem, text='3) Выберите конечную папку',
                                                      font=('Arial Bold', 15),
                                                      command=select_end_folder_diff_svod_trudvsem
                                                      )
    btn_choose_end_folder_diff_svod_trudvsem.pack(padx=10, pady=10)

    # Создаем чекбокс для режима подсчета изменений между регионами
    # Создаем переменную для хранения результа переключения чекбокса
    mode_region_value = StringVar()

    # Устанавливаем значение по умолчанию для этой переменной. По умолчанию будет вестись подсчет числовых данных
    mode_region_value.set('No')
    # Создаем чекбокс для выбора режима подсчета

    chbox_mode_contrast = Checkbutton(frame_data_diff_svod_trudvsem,
                                       text='Поставьте галочку, если вам нужно сравнить данные 2 регионов,\n'
                                            'сравнение будет идти только по отраслям',
                                       variable=mode_region_value,
                                       offvalue='No',
                                       onvalue='Yes')
    chbox_mode_contrast.pack(padx=10, pady=10)


    btn_proccessing_data_diff_svod_trudvsem = Button(tab_diff_svod_trudvsem, text='4) Обработать данные',
                                                     font=('Arial Bold', 20),
                                                     command=processing_diff_svod_trudvsem
                                                     )
    btn_proccessing_data_diff_svod_trudvsem.pack(padx=10, pady=10)











    # """
    # Вкладка для обработки отчетов центров карьеры
    # """
    # tab_employment_ck = ttk.Frame(tab_control)
    # tab_control.add(tab_employment_ck, text='Отчет ЦК')
    #
    # employment_ck_frame_description = LabelFrame(tab_employment_ck)
    # employment_ck_frame_description.pack()
    #
    # lbl_hello_employment_ck = Label(employment_ck_frame_description,
    #                                 text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
    #                                      'Обработка данных центров карьеры по трудоустроенным выпускникам',
    #                                 width=60)
    # lbl_hello_employment_ck.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    #
    # # Картинка
    # path_to_img_employment_ck = resource_path('logo.png')
    # img_employment_ck = PhotoImage(file=path_to_img_employment_ck)
    # Label(employment_ck_frame_description,
    #       image=img_employment_ck, padx=10, pady=10
    #       ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)
    #
    # # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    # frame_data_employment_ck = LabelFrame(tab_employment_ck, text='Подготовка')
    # frame_data_employment_ck.pack(padx=10, pady=10)
    #
    # btn_choose_data_employment_ck = Button(frame_data_employment_ck, text='1) Выберите папку с данными',
    #                                        font=('Arial Bold', 20),
    #                                        command=select_folder_data_ck
    #                                        )
    # btn_choose_data_employment_ck.pack(padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_employment_ck = Button(frame_data_employment_ck, text='2) Выберите конечную папку',
    #                                              font=('Arial Bold', 20),
    #                                              command=select_end_folder_ck
    #                                              )
    # btn_choose_end_folder_employment_ck.pack(padx=10, pady=10)
    # #
    # # Создаем кнопку обработки данных
    #
    # btn_proccessing_data_employment_ck = Button(tab_employment_ck, text='3) Обработать данные',
    #                                             font=('Arial Bold', 20),
    #                                             command=processing_ck_employment
    #                                             )
    # btn_proccessing_data_employment_ck.pack(padx=10, pady=10)
    #
    #
    #
    # """
    # Подсчет данных по трудоустройству ОПК
    # """
    # tab_employment_opk = ttk.Frame(tab_control)
    # tab_control.add(tab_employment_opk, text='Отчет ОПК с отраслями')
    #
    # employment_opk_frame_description = LabelFrame(tab_employment_opk)
    # employment_opk_frame_description.pack()
    #
    # lbl_hello_employment_opk = Label(employment_opk_frame_description,
    #                                  text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
    #                                       'Обработка данных по трудоустройству ОПК (по отраслям)\n'
    #                                       'В обрабатываемых файлах должны быть листы Форма 1 и Форма 2,\n'
    #                                       'В Форме 1 должно быть 80 колонок включая 2 колонки проверки\n'
    #                                       ',внизу после окончания таблицы должна быть пустая строка.\n'
    #                                       ' На 9 строке должна быть строка с номерами колонок.\n'
    #                                       'В форме 2 должно быть 10 колонок',
    #                                  width=60)
    # lbl_hello_employment_opk.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)
    #
    # # Картинка
    # path_to_img_employment_opk = resource_path('logo.png')
    # img_employment_opk = PhotoImage(file=path_to_img_employment_opk)
    # Label(employment_opk_frame_description,
    #       image=img_employment_opk, padx=10, pady=10
    #       ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)
    #
    # # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    # frame_data_employment_opk = LabelFrame(tab_employment_opk, text='Подготовка')
    # frame_data_employment_opk.pack(padx=10, pady=10)
    #
    # btn_choose_data_employment_opk = Button(frame_data_employment_opk, text='1) Выберите папку с данными',
    #                                         font=('Arial Bold', 20),
    #                                         command=select_folder_data_opk
    #                                         )
    # btn_choose_data_employment_opk.pack(padx=10, pady=10)
    #
    # # Создаем кнопку для выбора папки куда будут генерироваться файлы
    #
    # btn_choose_end_folder_employment_opk = Button(frame_data_employment_opk, text='2) Выберите конечную папку',
    #                                               font=('Arial Bold', 20),
    #                                               command=select_end_folder_opk
    #                                               )
    # btn_choose_end_folder_employment_opk.pack(padx=10, pady=10)
    # #
    # # Создаем кнопку обработки данных
    #
    # btn_proccessing_data_employment_opk = Button(tab_employment_opk, text='3) Обработать данные',
    #                                              font=('Arial Bold', 20),
    #                                              command=processing_opk_employment
    #                                              )
    # btn_proccessing_data_employment_opk.pack(padx=10, pady=10)



    """
    Разница двух таблиц
    """
    tab_diffrence = ttk.Frame(tab_control)
    tab_control.add(tab_diffrence, text='Разница 2 таблиц')

    diffrence_frame_description = LabelFrame(tab_diffrence)
    diffrence_frame_description.pack()

    lbl_hello_diffrence = Label(diffrence_frame_description,
                                text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                     'Количество строк и колонок в таблицах должно совпадать\n'
                                     'Названия колонок в таблицах должны совпадать'
                                     '\nДля корректной работы программмы уберите из таблицы объединенные ячейки',
                                width=60)
    lbl_hello_diffrence.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_diffrence = resource_path('logo.png')
    img_diffrence = PhotoImage(file=path_to_img_diffrence)
    Label(diffrence_frame_description,
          image=img_diffrence, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_for_diffrence = LabelFrame(tab_diffrence, text='Подготовка')
    frame_data_for_diffrence.pack(padx=10, pady=10)
    #
    # Создаем кнопку Выбрать  первый файл с данными
    btn_data_first_diffrence = Button(frame_data_for_diffrence, text='1) Выберите файл с первой таблицей',
                                      font=('Arial Bold', 10),
                                      command=select_first_diffrence
                                      )
    btn_data_first_diffrence.pack(padx=10, pady=10)
    #
    # Определяем текстовую переменную
    entry_first_sheet_name_diffrence = StringVar()
    # Описание поля
    label_first_sheet_name_diffrence = Label(frame_data_for_diffrence,
                                             text='2) Введите название листа, где находится первая таблица')
    label_first_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода имени листа
    first_sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_first_sheet_name_diffrence,
                                             width=30)
    first_sheet_name_entry_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку Выбрать  второй файл с данными
    btn_data_second_diffrence = Button(frame_data_for_diffrence, text='3) Выберите файл со второй таблицей',
                                       font=('Arial Bold', 10),
                                       command=select_second_diffrence
                                       )
    btn_data_second_diffrence.pack(padx=10, pady=10)

    # Определяем текстовую переменную
    entry_second_sheet_name_diffrence = StringVar()
    # Описание поля
    label_second_sheet_name_diffrence = Label(frame_data_for_diffrence,
                                              text='4) Введите название листа, где находится вторая таблица')
    label_second_sheet_name_diffrence.pack(padx=10, pady=10)
    # поле ввода
    second__sheet_name_entry_diffrence = Entry(frame_data_for_diffrence, textvariable=entry_second_sheet_name_diffrence,
                                               width=30)
    second__sheet_name_entry_diffrence.pack(padx=10, pady=10)

    # Создаем кнопку выбора папки куда будет генерироваьться файл
    btn_select_end_diffrence = Button(frame_data_for_diffrence, text='5) Выберите конечную папку',
                                      font=('Arial Bold', 10),
                                      command=select_end_folder_diffrence
                                      )
    btn_select_end_diffrence.pack(padx=10, pady=10)
    #
    # Создаем кнопку Обработать данные
    btn_data_do_diffrence = Button(tab_diffrence, text='6) Обработать таблицы', font=('Arial Bold', 20),
                                   command=processing_diffrence
                                   )
    btn_data_do_diffrence.pack(padx=10, pady=10)






    # Создаем виджет для управления полосой прокрутки
    canvas.create_window((0, 0), window=tab_control, anchor="nw")

    # Конфигурируем холст для обработки скроллинга
    canvas.config(yscrollcommand=scrollbar.set, scrollregion=canvas.bbox("all"))
    scrollbar.pack(side="right", fill="y")

    # Вешаем событие скроллинга
    canvas.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

    window.bind_class("Entry", "<Button-3><ButtonRelease-3>", show_textmenu)
    window.bind_class("Entry", "<Control-a>", callback_select_all)
    window.mainloop()
