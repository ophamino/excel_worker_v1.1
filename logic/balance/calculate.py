from typing import Any
import os
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet

from logic.const import MAIN_DIR, MONTH_LIST
from logic.utils.excel_extends import open_excel, open_sheet


class Balance:
    """
    Класс для формирования сводного баланса
    """

    def serialize_network(self, month: str | int) -> dict[str, dict[str, str]]:
        """
        Функция для формирования хэш-таблицы сети из файлы "Структура сети.xlsx"
        Args:
            month (str | int): Номер месяца

        Returns:
            dict[str, dict[str, str]]: Данные структуры сети
        """
        data = {}
        
        path = f"{MAIN_DIR}/Реестровая база данных/Структура электросети/Свод ОЭСХ.xlsx"
        file = open_excel(path, data_only=True)
        sheet = file.worksheets[0]
        for row in range(2, sheet.max_row + 1):
            data[sheet.cell(row=row, column=1).value] = {
                "name": sheet.cell(row=row, column=2).value,
                "consumption": 0,
                "reception": 0,
                "transmission": 0,
                "balance": 0,
                "waste": 0,
                "foreign_key": sheet.cell(row=row, column=3).value,
                "foreign_key_name": sheet.cell(row=row, column=4).value,
            }

        return data
        
    
    def serialize_bicu(self, month: str | int) -> dict[str, dict[str, str]]:
        """
        Функция для формирования хэш-таблицы потребителей из файлы "Сводная ведомость БИКУ.xlsx"
        Args:
            month (str | int): Номер или название месяца

        Returns:
            dict[str, dict[str, str]]: Данные Бику
        """
        data = {}
        year = datetime.now().year
        
        path = f"{MAIN_DIR}/Сводный баланс энергопотребления/Сводный баланс {year}/Сводная ведомость БИКУ/Сводная ведомость БИКУ.xlsx"
        file = open_excel(path, data_only=True)
        sheet = open_sheet(file, month)
        
        for row in range(8, sheet.max_row + 1):
            data[sheet.cell(row=row, column=3).value] = {
                "ID": sheet.cell(row=row, column=3).value,
                "name": sheet.cell(row=row, column=6).value,
                "status": sheet.cell(row=row, column=5).value,
                "expenses": sheet.cell(row=row, column=23).value,
                "foreign_key": sheet.cell(row=row, column=29).value,
            }
        return data
    
    
    def serialize_consumers(self, month: str | int) -> dict[str, dict[str, str]]:
        """
        Функция для формирования хэш-таблицы потребителей из файлы "Сводная ведомость.xlsx"
        Args:
            month (str | int): Номер или название месяца

        Returns:
            dict[str, dict[str, str]]: Данные потребителей
        """        
        data = {}
        year = datetime.now().year
        
        path = f"{MAIN_DIR}/Сводный баланс энергопотребления/Сводный баланс {year}/Сводная ведомость потребителей/Сводная ведомость потребителей.xlsx"
        file = load_workbook(path, data_only=True)
        sheet = file[MONTH_LIST[month - 1]]
        
        for row in range(6, sheet.max_row + 1):
            data[sheet.cell(row=row, column=3).value] = {
                "ID": sheet.cell(row=row, column=3).value,
                "name": sheet.cell(row=row, column=8).value,
                "expenses": sheet.cell(row=row, column=29).value,
                "foreign_key": sheet.cell(row=row, column=39).value,
            }
            print(sheet.cell(row=row, column=29).value)
        return data
        

    def serialize_balance(self, month: str | int) -> dict[str, dict[str, str]]:
        """
        Функция для формирования данных сводного баланса из файлов "Структура сети.xlsx" и "Сводная ведомость.xlsx"



        Args:
            month (str | int): Номер или название месяцау

        Returns:
            dict[str, dict[str, str]]: Хэш-таблица баланса
        """
        consumers_data = self.serialize_consumers(month)
        network_data = self.serialize_network(month)
        bicu_data = self.serialize_bicu(month)
        
        balance = network_data

        for value in consumers_data.values():
            foreign_key = value["foreign_key"]
            if foreign_key in balance.keys():
                if value["expenses"]:
                    balance[foreign_key]["consumption"] += value["expenses"]

        for value in bicu_data.values():
            foreign_key = value["foreign_key"]
            if foreign_key:
                if value["status"] == "Прием электроэнергии":
                    if value["expenses"]:
                        balance[foreign_key]["reception"] += value["expenses"]
                if value["status"] == "Передача электроэнергии":
                    if value["expenses"]:
                        balance[foreign_key]["transmission"] += value["expenses"]
                if value["status"] not in ("Прием электроэнергии", "Передача электроэнергии"):
                    raise ValueError(
                        f'Неверные данные в Сводной ведомости БИКУ, - ID {value["ID"]} Должно быть "Прием электроэнергии" или "Передача электроэнергии"'
                        )
                    
        for value in network_data.values():
            foreign_key = value["foreign_key"]
            if foreign_key:
                if value["consumption"]:
                    balance[foreign_key]["consumption"] += value["consumption"]
                if value["reception"]:
                    balance[foreign_key]["reception"] += value["reception"]
                if value["transmission"]:
                    balance[foreign_key]["transmission"] += value["transmission"]
                
        for key in balance.keys():
            balance[key]["balance"] = balance[key]["reception"] - balance[key]["transmission"]
            balance[key]["waste"] = balance[key]["balance"] - balance[key]["consumption"]
        
        return balance
        
    def create_balance(self, month: str | int) -> None:
        """
        Функция для вставки итоговых данных в файл "Сводный баланс"
        """
        year = datetime.now().year
        path = f"{MAIN_DIR}/Сводный баланс энергопотребления/Сводный баланс {year}/Сводный баланс.xlsx"
        file = open_excel(path)
        sheet = open_sheet(file, month)
        
        data = self.serialize_balance(month)
        
        sheet.append(("№", "Идентификатор", "Наименование", "Вход", "Выход", "Сальдо переток", "Полезный отпуск", "Потери"))
        number = 1
        for key, value in data.items():
            sheet.append(
                (
                    number, key, value["name"], value["reception"],
                    value["transmission"], value["balance"], value["consumption"], value["waste"]
                )
            )
            number = number + 1
        
        file.save(path)


    def __call__(self, *args: Any, **kwds: Any) -> Any: 
        pass
