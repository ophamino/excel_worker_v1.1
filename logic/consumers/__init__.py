from .calculate import ConsumersCalculate
from .comparer import ConsumersComparer
from logic.utils.log import Log

from logic.const import MONTH_LIST

def start_consumer():
    menu = [
         "",
        "----------------Потребители-------------------",
        "Выберите действие, которое хотите совершить: ",
        "1. Сверить статические данные",
        "2. Сформировать сводную ведомость",
        "3. Сформировать расчетную ведомость",
        "0. Главное меню",
        "________________________________________________",
        ""
    ]
    print(*menu, sep='\n')
    action = int(input("Выберите порядковый номер действия: " ))
    while True:
        if action == 0:
            break
        
        if action == 1:
            Log().search_changes()
        
        if action == 2:
            [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
            month = int(input('Введите номер месяца: '))
            print("1. Бытовое потребление", '2. Комерческое потребление', "3. Общая", sep='\n')
            status = int(input("Выберите статус сводной ведомости: "))
            if status == 1:
                ConsumersComparer().collect_files(month, "Бытовых")
            if status == 2:
                ConsumersComparer().collect_files(month, "Коммерческих")
            if status == 3:
                ConsumersComparer().collect_total_files(month)
        
        if action == 3:
            [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
            month = int(input('Введите номер месяца: '))
            print("1. Бытовое потребление", '2. Комерческое потребление', sep='\n')
            status = int(input("Выберите статус сводной ведомости: "))
            
            if status == 1:
                ConsumersCalculate().format_data(month, "Бытовых")
            if status == 2:
                ConsumersCalculate().format_data(month, "Коммерческих")
        break