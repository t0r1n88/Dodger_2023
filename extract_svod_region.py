"""
Вспомогательный скрипт для сбора файлов сводов по всем регионам в соответствующие папки
"""
import os
from pathlib import Path
import shutil
import re
import time





def collecting_svod_region(data_folder:str,end_folder:str,lst_region:list):
    """
    Функция для сбора из папок файлов сводов и сохранения их в папки по названию региона
    """
    source = Path(data_folder)
    target = Path(end_folder)

    target.mkdir(parents=True, exist_ok=True)
    if not lst_region:
        for dir_date in source.iterdir():
            if dir_date.is_dir():
                result = re.search(r'\d{2}.\d{2}.\d{4}',str(dir_date))
                if result:
                    date_for_file = result.group()
                    date_for_file = date_for_file.replace('.','_')
                    print(date_for_file)
                else:
                    continue

                for region_dir in dir_date.iterdir():
                    if region_dir.is_dir():
                        region_name = region_dir.name
                        # создаем папку в конечной папке
                        name_end_folder = 'Своды'
                        target_region_path = target / region_name / name_end_folder
                        target_region_path.mkdir(exist_ok=True, parents=True)

                        for prom_dir in region_dir.iterdir():
                            if prom_dir.is_dir():
                                for prom_date_dir in prom_dir.iterdir():
                                    if prom_date_dir.is_dir():
                                        for file in prom_date_dir.iterdir():
                                            if file.is_file() and 'Свод по региону' in file.name:
                                                target_file = target_region_path / f'Свод по региону от {date_for_file}.xlsx'
                                                shutil.copy2(file, target_file)  # copy2 сохраняет метаданные
                for region_dir in dir_date.iterdir():
                    if region_dir.is_dir():
                        region_name = region_dir.name
                        # создаем папку в конечной папке
                        name_end_folder = 'Списки'
                        target_region_path = target / region_name / name_end_folder
                        target_region_path.mkdir(exist_ok=True, parents=True)
                        for file in region_dir.iterdir():
                            if file.is_file() and 'Вакансии по региону' in file.name:
                                target_file = target_region_path / f'Вакансии по региону от {date_for_file}.xlsx'
                                shutil.copy2(file, target_file)  # copy2 сохраняет метаданные

    else:
        for dir_date in source.iterdir():
            if dir_date.is_dir():
                result = re.search(r'\d{2}.\d{2}.\d{4}', str(dir_date))
                if result:
                    date_for_file = result.group()
                    date_for_file = date_for_file.replace('.', '_')
                    print(date_for_file)
                else:
                    continue

                for region_dir in dir_date.iterdir():
                    if region_dir.is_dir():
                        region_name = region_dir.name
                        if region_name in lst_region:

                            # создаем папку в конечной папке
                            name_end_folder = 'Своды'
                            target_region_path = target / region_name / name_end_folder
                            target_region_path.mkdir(exist_ok=True,parents=True)


                            for prom_dir in region_dir.iterdir():
                                if prom_dir.is_dir():
                                    for prom_date_dir in prom_dir.iterdir():
                                        if prom_date_dir.is_dir():
                                            for file in prom_date_dir.iterdir():
                                                if file.is_file() and 'Свод по региону' in file.name:
                                                    target_file = target_region_path / f'Свод по региону от {date_for_file}.xlsx'
                                                    shutil.copy2(file, target_file)  # copy2 сохраняет метаданные
                for region_dir in dir_date.iterdir():
                    if region_dir.is_dir():
                        region_name = region_dir.name
                        if region_name in lst_region:
                            # создаем папку в конечной папке
                            name_end_folder = 'Списки'
                            target_region_path = target / region_name / name_end_folder
                            target_region_path.mkdir(exist_ok=True,parents=True)
                            for file in region_dir.iterdir():
                                if file.is_file() and 'Вакансии по региону' in file.name:
                                    target_file = target_region_path / f'Вакансии по региону от {date_for_file}.xlsx'
                                    shutil.copy2(file, target_file)  # copy2 сохраняет метаданные









if __name__ == '__main__':
    main_data_folder = 'e:/Тест/'
    main_data_folder = 'e:/Работа в России/'
    main_end_folder ='e:/Все регионы'
    main_end_folder ='e:/Отдельные регионы'
    main_lst_region = ['Иркутская область','Республика Бурятия']
    main_lst_region = []

    start_time = time.time()
    collecting_svod_region(main_data_folder,main_end_folder,main_lst_region)
    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Время выполнения: {execution_time} секунд")
    print('Lindy Booth')




