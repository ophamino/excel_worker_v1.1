import os
from datetime import datetime

from openpyxl.worksheet.worksheet import Worksheet

from logic.utils.excel_extends import open_sheet, load_workbook
from logic.const import MAIN_DIR, MONTH_LIST


class BicuComparer:
    
    def collect_files(self, month: int):
        path = f"{MAIN_DIR}/Сводный баланс энергопотребления/Сводный баланс {datetime.now().year}/Сводная ведомость БИКУ/"
        file_path = f"{MAIN_DIR}/Сводный баланс энергопотребления/Сводный баланс {datetime.now().year}/Сводная ведомость БИКУ/Сводная ведомость БИКУ.xlsx"
        
        if not os.path.exists(file_path):
            file = load_workbook('./template/bicu.xlsx')
            file.save(file_path)
        
        file = load_workbook(file_path)
        sheet = open_sheet(file, month)
        
        try:
            file_names = os.listdir(f"{path}{MONTH_LIST[month - 1]}")
            for file_name in file_names:
                rv_file = load_workbook(f"{path}{MONTH_LIST[month - 1]}/{file_name}")
                rv_sheet = rv_file.worksheets[0]
                flag = 8
                for row in rv_sheet.iter_rows(min_row=8, values_only=True):
                    sheet.append(row)
                    sheet.cell(row=flag, column=14).value = str(sheet.cell(row=flag, column=14).value)
                    flag += 1
            self.insert_formula(sheet)
            file.save(file_path)
        except Exception as error:
            print(f"Файл отсутсвует, проверьте директорию {path}{MONTH_LIST[month - 1]}>")
        
    def insert_formula(self, sheet: Worksheet):
        letter_list = ["R", "S", "T", "U", "V", "W"]
            
        for letter in letter_list:
            sheet[f"{letter}4"] = f'=SUMIFS({letter}8:{letter}{sheet.max_row + 1},$E8:$E{sheet.max_row + 1},"Прием электроэнергии")'
            sheet[f"{letter}5"] = f'=SUMIFS({letter}8:{letter}{sheet.max_row + 1},$E8:$E{sheet.max_row + 1},"Передача электроэнергии")'
            sheet[f"{letter}6"] = f'={letter}4-{letter}5'
