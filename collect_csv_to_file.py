import pandas as pd
import os
from tkinter import Tk, filedialog
from colorama import init, Fore, Style
import time
import gc

# Инициализация colorama
init(autoreset=True)

# Функция для выбора файлов CSV
def select_csv_files():
    root = Tk()
    root.withdraw()  # Скрыть основное окно
    # Открыть диалог выбора файлов CSV
    file_paths = filedialog.askopenfilenames(
        title="Выберите CSV файлы",
        filetypes=[("CSV files", "*.csv")],
        initialdir=os.path.expanduser("~")  # Начальная директория
    )
    return file_paths

# Функция для выбора места сохранения конечного файла
def select_save_location():
    root = Tk()
    root.withdraw()  # Скрыть основное окно
    # Открыть диалог выбора места для сохранения
    file_path = filedialog.asksaveasfilename(
        title="Выберите место для сохранения Excel файла",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialdir=os.path.expanduser("~")  # Начальная директория
    )
    return file_path

def print_with_time(message, color=Fore.WHITE):
    current_time = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
    print(f"{color}{current_time} - {message}{Style.RESET_ALL}")

def main():
    print_with_time("Запуск процесса...", Fore.CYAN)
    
    # Выбор CSV файлов
    print_with_time("Выберите CSV файлы для импорта...", Fore.YELLOW)
    csv_files = select_csv_files()
    if not csv_files:
        print_with_time("CSV файлы не выбраны. Программа завершена.", Fore.RED)
        return

    # Выбор места для сохранения конечного Excel файла
    print_with_time("Выберите место для сохранения Excel файла...", Fore.YELLOW)
    excel_file_path = select_save_location()
    if not excel_file_path:
        print_with_time("Место для сохранения Excel файла не выбрано. Программа завершена.", Fore.RED)
        return

    # Создаем ExcelWriter объект для записи в Excel файл
    print_with_time("Создание Excel файла...", Fore.GREEN)
    with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
        for csv_file_path in csv_files:
            try:
                # Чтение данных из CSV файла в DataFrame с кодировкой UTF-8
                print_with_time(f"Начало импорта данных из: {csv_file_path}", Fore.BLUE)
                df = pd.read_csv(csv_file_path, encoding='utf-8')
                
                # Удаление пробелов в именах столбцов и значениях
                print_with_time("Удаление пробелов в именах столбцов и значениях...", Fore.MAGENTA)
                df.columns = [col.strip() for col in df.columns]  # Удаление пробелов в именах столбцов
                df = df.applymap(lambda x: x.strip() if isinstance(x, str) else x)  # Удаление пробелов в значениях
                
                # Имя листа в Excel файле будет соответствовать имени CSV файла без расширения
                sheet_name = os.path.splitext(os.path.basename(csv_file_path))[0]
                
                # Если имя листа слишком длинное для Excel (макс 31 символ), обрезаем его
                if len(sheet_name) > 31:
                    print_with_time(f"Имя листа '{sheet_name}' слишком длинное, обрезаем до 31 символа...", Fore.MAGENTA)
                    sheet_name = sheet_name[:31]
                
                # Записываем данные в Excel файл на отдельный лист
                df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                print_with_time(f"Успешный импорт данных из: {csv_file_path}", Fore.GREEN)

                # Очистка памяти после обработки каждого файла
                del df
                gc.collect()

            except Exception as e:
                print_with_time(f"Ошибка при импорте данных из: {csv_file_path} - {str(e)}", Fore.RED)

    print_with_time(f"Файл успешно сохранен как: {excel_file_path}", Fore.GREEN)

if __name__ == "__main__":
    main()
