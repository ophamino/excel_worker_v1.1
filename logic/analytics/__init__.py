from logic.const import MONTH_LIST
from .balance import BalanceAnalytics

def start_analytics():
    [print(f"{number}. {month}") for number, month in enumerate(MONTH_LIST, 1)]
    month = int(input('Введите номер месяца: '))
    BalanceAnalytics().create_analytics(month)