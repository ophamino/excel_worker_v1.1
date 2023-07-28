import os
from datetime import datetime

from logic.const import MAIN_DIR, MONTH_LIST


class TreeDir:
    """Класс для создания деревыа каталогов"""

    def __init__(self) -> None:
        self.main_dir = MAIN_DIR
        self.validate_dir(self.main_dir)
        
    def validate_dir(self, path: str) -> None:
        
        """Функция для проверки сущесьвует ли папка"""
        if not os.path.exists(path):
            os.makedirs(path)
    
    def __create_main_folders(self) -> None:
        """Функия для создания главных папок"""
        folders = [
            "Сводный баланс энергопотребления", "Шаблоны расчетных ведомостей",
            "Аналитика баланса электроэнергии", "Реестровая база данных"
        ]
        for folder in folders:
            path = f"{self.main_dir}/{folder}"
            self.validate_dir(path)
    
    def __create_templates_folders(self) -> None:
        """Функция для создания содержимого папки 'Шаблоны расчетных ведомостей'"""
        folders = ["РВ БИКУ", "РВ Бытовых потребителей", "РВ Коммерческих потребителей"]
        path = f"{self.main_dir}/Шаблоны расчетных ведомостей"
        for folder in folders:
            path = f"{self.main_dir}/Шаблоны расчетных ведомостей/{folder}"
            self.validate_dir(path)
        
    def __create_base_folders(self) -> None:
        """Функция для создания содержимого папки 'Реестровая база данных'"""
        folders = ["Реестр потребителей", "Реестр БИКУ", "Структура электросети"]
        for folder in folders:
            path = f"{self.main_dir}/Реестровая база данных/{folder}"
            self.validate_dir(path)
    
    def __create_bicu(self, path: str) -> None:
        """Функуия для создания содержимого 'Сводная ведомость БИКУ'"""
        path += "/Сводная ведомость БИКУ"
        self.validate_dir(path)
        [self.validate_dir(f"{path}/{folder}") for folder in MONTH_LIST]
    
    def __create_comsumers(self, path: str) -> None:
        """Функуия для создания содержимого 'Сводная ведомость потребителей'"""
        path += "/Сводная ведомость потребителей"
        [self.validate_dir(f"{path}/{folder}") for folder in MONTH_LIST]
        [self.validate_dir(f"{path}/{folder}/РВ Бытовых потребителей") for folder in MONTH_LIST]
        [self.validate_dir(f"{path}/{folder}/РВ Коммерческих потребителей") for folder in MONTH_LIST]
    
    def __create_balance_folders(self) -> None:
        """Функция для создания содержимого папки 'Сводный баланс энергопотребления'"""
        this_yesr_balance = f"{self.main_dir}/Сводный баланс энергопотребления/Сводный баланс {datetime.now().year}"
        self.validate_dir(this_yesr_balance)
        self.__create_bicu(this_yesr_balance)
        self.__create_comsumers(this_yesr_balance)
    
    
    def create_tree_dirs(self):
        self.__create_main_folders()
        self.__create_templates_folders()
        self.__create_base_folders()
        self.__create_balance_folders()
