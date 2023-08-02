from .calculate import BicuCalculate
from .comparer import BicuComparer

from logic.const import MONTH_LIST

def start_bicu():
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
            pass
        
        if action == 2:
            [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
            month = int(input('Введите номер месяца: '))
            BicuComparer().collect_files(month)
        
        if action == 3:
            [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
            month = int(input('Введите номер месяца: '))
            BicuCalculate().format_data(month)
        break