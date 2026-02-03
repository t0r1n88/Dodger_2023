"""
Скрипт для массового создания временных рядов по сводам данных по вакансиям
"""
from create_time_series_svod import processing_time_series
import os
from pathlib import Path
import shutil

def processing_all_region_time_series(data_folder:str, end_folder:str):
    """

    """
    source = Path(data_folder)
    target = Path(end_folder)
    target.mkdir(parents=True, exist_ok=True)


    for region_dir in source.iterdir():
        if region_dir.is_dir():
            target_region_path = target / region_dir.name
            target_region_path.mkdir(exist_ok=True)
            print(region_dir)
            processing_time_series(region_dir,target_region_path)








if __name__ == '__main__':
    main_data_folder = 'data/Своды'
    main_end_folder = 'data/РЕЗУЛЬТАТ'
    processing_all_region_time_series(main_data_folder,main_end_folder)
    print('Lindy Booth')

