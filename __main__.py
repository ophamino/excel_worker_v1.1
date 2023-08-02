from logic.utils.tree_dirs import TreeDir
from logic.consumers import start_consumer
from logic.bicu import start_bicu
from logic.balance import start_balance
from logic.analytics import start_analytics



def main():
    TreeDir().create_tree_dirs()
    print()
    print("Welcome!")
    main_action = [
        "",
        "----------------Меню---------------------------",
        "Выберите действие, которое хотите совершить: ",
        "1. Потребители",
        "2. БИКУ",
        "3. Сформировать сводный баланс",
        "4. Сформировать аналитику"
        "________________________________________________"
        ""
    ]
    while True:
        print(*main_action, sep='\n')
        action = int(input("Выберите порядковый номер действия: "))
        if action == 1:
            start_consumer()
        if action == 2:
            start_bicu()
        if action == 3:
            start_balance()
        if action == 4:
            start_analytics()


if __name__ == "__main__":
    main()