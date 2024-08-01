import pandas as pd
import json
import tkinter as tk
from tkinter import filedialog
from datetime import datetime
from colorama import Fore, Style, init

# Инициализация colorama
init(autoreset=True)

def log_message(message, color=Fore.WHITE):
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    print(f"{color}{timestamp} - {message}{Style.RESET_ALL}")

def select_json_file():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
    return file_path

def main():
    log_message("Запуск скрипта...", Fore.CYAN)
    json_file = select_json_file()
    if json_file:
        log_message(f"Выбран файл: {json_file}", Fore.GREEN)
        try:
            with open(json_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            log_message("Файл успешно прочитан", Fore.GREEN)

            df = pd.DataFrame(data)
            output_file = json_file.replace('.json', '.xlsx')
            df.to_excel(output_file, index=False)
            log_message(f"Данные успешно преобразованы и сохранены в {output_file}", Fore.GREEN)
        except Exception as e:
            log_message(f"Ошибка при обработке файла: {e}", Fore.RED)
    else:
        log_message("Файл не был выбран", Fore.YELLOW)

if __name__ == "__main__":
    main()
