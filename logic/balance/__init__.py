from logic.const import MONTH_LIST
from .calculate import Balance


def start_balance():
    [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
    month = int(input('Введите номер месяца: '))
    Balance().create_balance(month)
    