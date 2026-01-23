"""
Вспомогательный скрипт Извлечение в одну папку сводов по одному региону
"""

import os
from pathlib import Path
import shutil
import re
import time
from datetime import datetime, timedelta



def collecting_svod_one_region(data_folder:str,end_folder:str):
    """
    Функция для сбора из папок файлов сводов одного региона и сбор их в одну папку
    """
    source = Path(data_folder)
    target = Path(end_folder)

    target.mkdir(parents=True, exist_ok=True)
    for dir_date in source.iterdir():
        if dir_date.is_dir():
            result = re.search(r'\d{2}_\d{2}_\d{4}',str(dir_date))
            if result:
                date_str = result.group()
                # Конвертация строки в дату
                date_obj = datetime.strptime(date_str, "%d_%m_%Y")
                # Отнимаем один день
                previous_day = date_obj - timedelta(days=1)
                # Конвертируем обратно в строку
                date_str = previous_day.strftime("%d_%m_%Y")
                for file in dir_date.iterdir():
                    if file.is_file() and 'Свод по региону' in file.name:
                        target_file = target / f'Свод по региону от {date_str}.xlsx'
                        shutil.copy2(file, target_file)  # copy2 сохраняет метаданные



















if __name__ == '__main__':
    main_data_folder = 'data/Республика Бурятия/Аналитика по вакансиям региона'
    main_end_folder = 'data/СВОД Бурятия'
    collecting_svod_one_region(main_data_folder,main_end_folder)

    print('Lindy Booth')
