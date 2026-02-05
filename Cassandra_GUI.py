# -*- coding: utf-8 -*-
"""
Графический интерфейс для программы Кассандра
"""
from form_one_five_row import prepare_form_one_employment  # импортируем функцию для обработки Формы 1 (5 строк)
from form_two_fifteen_row_nose import prepare_form_two_employment  # импортируем функцию для обработки Форма 2 нозологии 15 строк
from form_three_expected_release import prepare_form_three_employment # импортируем функцию для обработки Формы 3 Ожидаемый выпуск
from monitoring_may_2025 import prepare_may_2025 # мониториг май 2025
from monitoring_september_2025 import prepare_september_2025 # мониторинг сентябрь 2025
from create_svod_trudvsem import processing_data_trudvsem # импортируем функцию для обработки данных с трудвсем
from contrast_svod_trudvsem import prepare_diff_svod_trudvsem # импортируем функцию для измерения разницы между двумя сводами
from cass_difference import prepare_diffrence  # импортируем функцию для нахождения разницы между двумя таблицами
from create_time_series_svod import processing_time_series # импортируем функцию создания временных рядов
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



def select_folder_data_mon_grad_sept():
    """
    Функция для выбора папки c данными мониторинга выпускников для СССР
    :return:
    """
    global path_folder_data_mon_grad_sept
    path_folder_data_mon_grad_sept = filedialog.askdirectory()


def select_end_folder_mon_grad_sept():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы мониторинга выпускников для СССР
    :return:
    """
    global path_to_end_folder_mon_grad_sept
    path_to_end_folder_mon_grad_sept = filedialog.askdirectory()






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

def select_file_params_chosen_vac():
    """
    Функция для выбора файла с параметрами обработки
    """
    global name_file_params_chosen_vac
    # Получаем путь к файлу
    name_file_params_chosen_vac = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))


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
Функции для общей функции мониторинга
"""

def select_folder_data_monitoring():
    """
    Функция для выбора папки c данными мониторингов
    :return:
    """
    global path_folder_data_monitoring
    path_folder_data_monitoring = filedialog.askdirectory()


def select_end_folder_monitoring():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы мониторинга
    :return:
    """
    global path_to_end_folder_monitoring
    path_to_end_folder_monitoring = filedialog.askdirectory()


def processing_monintoring():
    """
    Функция для обработки данных мониторингов. Основная функция
    :return:
    """
    type_monitoring = group_rb_type_monitoring.get() # получаем тип мониторинга

    try:
        if type_monitoring == 0:
            prepare_form_one_employment(path_folder_data_monitoring, path_to_end_folder_monitoring)
        elif type_monitoring == 1:
            prepare_form_two_employment(path_folder_data_monitoring, path_to_end_folder_monitoring)
        elif type_monitoring == 2:
            prepare_form_three_employment(path_folder_data_monitoring, path_to_end_folder_monitoring)
        elif type_monitoring == 3:
            prepare_may_2025(path_folder_data_monitoring, path_to_end_folder_monitoring)
        elif type_monitoring == 4:
            prepare_september_2025(path_folder_data_monitoring, path_to_end_folder_monitoring)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


"""
Функции для построения временных рядов по сводам
"""

def select_folder_data_svod_vacance():
    """
    Функция для выбора папки cодержащей своды
    :return:
    """
    global path_folder_data_svod_vacance
    path_folder_data_svod_vacance = filedialog.askdirectory()


def select_file_svod_vacance():
    """
    Функция для выбора файла с организациями
    """
    global file_svod_vacance
    # Получаем путь к файлу
    file_svod_vacance = filedialog.askopenfilename(filetypes=(('Excel files', '*.xlsx'), ('all files', '*.*')))



def select_end_folder_svod_vacance():
    """
    Функция для выбора конечной папки куда будут складываться итоговые файлы
    :return:
    """
    global path_to_end_folder_svod_vacance
    path_to_end_folder_svod_vacance = filedialog.askdirectory()


def processing_svod_vacance():
    """
    Фугкция для обработки данных мониторинга 5 строк базовый мониторинг
    :return: файлы Excel  с результатами обработки
    """
    try:
        processing_time_series(path_folder_data_svod_vacance, path_to_end_folder_svod_vacance,file_svod_vacance)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')








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



def processing_mon_may_2025():
    """
    Функция для обработки данных мониторинга занятости выпускников 2024 для сайта СССР
    :return:
    """
    try:
        prepare_may_2025(path_folder_data_mon_grad, path_to_end_folder_mon_grad)

    except NameError:
        messagebox.showerror('Кассандра Подсчет данных по трудоустройству выпускников',
                             f'Выберите файлы с данными и папку куда будет генерироваться файл')


def processing_mon_sept_2025():
    """
    Функция для обработки данных мониторинга занятости выпускников 2024 для сайта СССР
    :return:
    """
    try:
        prepare_september_2025(path_folder_data_mon_grad_sept, path_to_end_folder_mon_grad_sept)

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
        processing_data_trudvsem(file_csv_svod_trudvsem, file_org_svod_trudvsem,path_to_end_folder_svod_trudvsem,name_region,name_file_params,name_file_params_chosen_vac)

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
        width = int(screen_width * 0.3)
    elif screen_width >= 2560:
        width = int(screen_width * 0.41)
    elif screen_width >= 1920:
        width = int(screen_width * 0.47)
    elif screen_width >= 1600:
        width = int(screen_width * 0.5)
    elif screen_width >= 1280:
        width = int(screen_width * 0.62)
    elif screen_width >= 1024:
        width = int(screen_width * 0.77)
    else:
        width = int(screen_width * 1)

    height = int(screen_height * 0.7)

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
    window.title('Кассандра Обработка данных Работа в России и мониторингов трудоустройства выпускников ver 7.0')
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

    global name_file_params_chosen_vac
    name_file_params_chosen_vac = 'Не выбрано' # костыль

    global file_svod_vacance
    file_svod_vacance = 'Не выбрано' # костыль





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


    # Кнопка для выбора файла с вакансиями
    btn_choose_params_chosen_vac_svod_trudvsem = Button(frame_data_svod_trudvsem, text='Необязательная опция\n Выберите файл с отслеживаемыми вакансиями',
                                                 font=('Arial Bold', 10),
                                                 command=select_file_params_chosen_vac
                                                 )
    btn_choose_params_chosen_vac_svod_trudvsem.pack(padx=10, pady=10)



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




    tab_extract_svod = ttk.Frame(tab_control)
    tab_control.add(tab_extract_svod, text='Создание временных рядов\nпо сводам')

    extract_svod_frame_description = LabelFrame(tab_extract_svod)
    extract_svod_frame_description.pack()

    lbl_hello_extract_svod = Label(extract_svod_frame_description,
                                   text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                        'Построение временных рядов по основным показателям\n'
                                        'на основе файлов сводов созданных функцией Свод Работа в России',
                                   width=60)
    lbl_hello_extract_svod.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_extract_svod = resource_path('logo.png')
    img_path_to_img_extract_svod = PhotoImage(file=path_to_img_extract_svod)
    Label(extract_svod_frame_description,
          image=img_path_to_img_extract_svod, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_extract_svod = LabelFrame(tab_extract_svod, text='Подготовка')
    frame_data_extract_svod.pack(padx=10, pady=10)

    btn_choose_data_extract_svod = Button(frame_data_extract_svod,
                                          text='1) Выберите папку с данными', font=('Arial Bold', 20),
                                          command=select_folder_data_svod_vacance
                                          )
    btn_choose_data_extract_svod.pack(padx=10, pady=10)


    btn_choose_filter_extract_svod = Button(frame_data_extract_svod, text='Необязательная опция\n Выберите файл с отслеживаемыми вакансиями',
                                                 font=('Arial Bold', 10),
                                                 command=select_file_svod_vacance
                                                 )
    btn_choose_filter_extract_svod.pack(padx=10, pady=10)



    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_extract_svod = Button(frame_data_extract_svod,
                                                text='3) Выберите конечную папку', font=('Arial Bold', 20),
                                                command=select_end_folder_svod_vacance
                                                )
    btn_choose_end_folder_extract_svod.pack(padx=10, pady=10)

    btn_proccessing_data_extract_svod = Button(tab_extract_svod, text='4) Обработать данные',
                                               font=('Arial Bold', 20),
                                               command=processing_svod_vacance
                                               )
    btn_proccessing_data_extract_svod.pack(padx=10, pady=10)






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

    """
        Создаем вкладку для обработки мониторингов
        """
    tab_monitorings_employment = ttk.Frame(tab_control)
    tab_control.add(tab_monitorings_employment, text='Обработка мониторингов\n трудоустройства выпускников')

    monitorings_employment_frame_description = LabelFrame(tab_monitorings_employment)
    monitorings_employment_frame_description.pack()

    lbl_hello_monitorings_employment = Label(monitorings_employment_frame_description,
                                             text='Центр опережающей профессиональной подготовки Республики Бурятия\n'
                                                  'Обработка различных мониторингов трудоустройства выпускников',
                                             width=60)
    lbl_hello_monitorings_employment.pack(side=LEFT, anchor=N, ipadx=25, ipady=10)

    # Картинка
    path_to_img_monitorings_employment = resource_path('logo.png')
    img_path_to_img_monitorings_employment = PhotoImage(file=path_to_img_monitorings_employment)
    Label(monitorings_employment_frame_description,
          image=img_path_to_img_monitorings_employment, padx=10, pady=10
          ).pack(side=LEFT, anchor=E, ipadx=5, ipady=5)

    # Создаем область для того чтобы поместить туда подготовительные кнопки(выбрать файл,выбрать папку и т.п.)
    frame_data_monitorings_employment = LabelFrame(tab_monitorings_employment, text='Подготовка')
    frame_data_monitorings_employment.pack(padx=10, pady=10)

    # Создаем переключатель
    group_rb_type_monitoring = IntVar()
    # Создаем фрейм для размещения переключателей(pack и грид не используются в одном контейнере)
    frame_rb_type_monitoring = LabelFrame(frame_data_monitorings_employment, text='1) Выберите вид мониторинга')
    frame_rb_type_monitoring.pack(padx=10, pady=10)
    #
    Radiobutton(frame_rb_type_monitoring, text='Форма 1 (5 строк)', variable=group_rb_type_monitoring, value=0).pack(anchor='w', padx=5)
    Radiobutton(frame_rb_type_monitoring, text='Форма 2 (нозологии)', variable=group_rb_type_monitoring, value=1).pack(anchor='w', padx=5)
    Radiobutton(frame_rb_type_monitoring, text='Форма 3 (ожидаемый выпуск)', variable=group_rb_type_monitoring,
                value=2).pack(anchor='w', padx=5)
    Radiobutton(frame_rb_type_monitoring, text='Мониторинг занятости выпускников Май 2025',
                variable=group_rb_type_monitoring, value=3).pack(anchor='w', padx=5)
    Radiobutton(frame_rb_type_monitoring, text='Мониторинг занятости выпускников Сентябрь 2025',
                variable=group_rb_type_monitoring, value=4).pack(anchor='w', padx=5)

    btn_choose_data_monitorings_employment = Button(frame_data_monitorings_employment,
                                                    text='1) Выберите папку с данными', font=('Arial Bold', 20),
                                                    command=select_folder_data_monitoring
                                                    )
    btn_choose_data_monitorings_employment.pack(padx=10, pady=10)

    # Создаем кнопку для выбора папки куда будут генерироваться файлы

    btn_choose_end_folder_monitorings_employment = Button(frame_data_monitorings_employment,
                                                          text='2) Выберите конечную папку', font=('Arial Bold', 20),
                                                          command=select_end_folder_monitoring
                                                          )
    btn_choose_end_folder_monitorings_employment.pack(padx=10, pady=10)

    btn_proccessing_data_monitorings_employment = Button(tab_monitorings_employment, text='3) Обработать данные',
                                                         font=('Arial Bold', 20),
                                                         command=processing_monintoring
                                                         )
    btn_proccessing_data_monitorings_employment.pack(padx=10, pady=10)







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
