import pandas as pd
from tqdm import tqdm
from time import strftime, localtime
from tkinter import Tk, filedialog, simpledialog, ttk
from colorama import init, Fore, Style
import re
import os
import string

# Инициализация colorama
init(autoreset=True)

def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

def select_files():
    root = Tk()
    root.withdraw()
    file_paths = filedialog.askopenfilenames(title="Выберите исходные файлы",
                                             filetypes=[("Excel files", "*.xlsx *.xls"), ("CSV files", "*.csv"), ("JSON files", "*.json"), ("Text files", "*.txt")])
    return list(file_paths)

def get_column_letter(index):
    """Возвращает букву столбца для указанного индекса."""
    return string.ascii_uppercase[index]

def select_column(headers, sheet_name):
    def on_select(event):
        root.quit()

    root = Tk()
    root.title(f"Выбор столбца для листа: {sheet_name}")
    label = ttk.Label(root, text=f"Выберите столбец для листа: {sheet_name}:")
    label.pack()

    # Создаём список с именами столбцов
    column_options = [f"{get_column_letter(idx)}1: {header}" for idx, header in enumerate(headers)]
    
    combo = ttk.Combobox(root, values=column_options)
    combo.bind("<<ComboboxSelected>>", on_select)
    combo.pack()
    center_window(root)
    root.mainloop()
    selected_column = combo.get().split(":")[0]  # Берём только имя столбика (например, A1)
    root.destroy()
    return selected_column

def generate_save_path(source_file):
    """Генерация пути для сохранения файла в той же директории с суффиксом _filtered."""
    dir_name = os.path.dirname(source_file)
    base_name = os.path.basename(source_file).split('.')[0]  # Имя файла без расширения
    new_file_name = f"{base_name}_filtered.xlsx"  # Новое имя файла с суффиксом _filtered
    return os.path.join(dir_name, new_file_name)

def print_with_time(message, color=Fore.WHITE):
    current_time = strftime("%Y-%m-%d %H:%M:%S", localtime())
    print(f"{color}{current_time} - {message}{Style.RESET_ALL}")

# Массив номеров для поиска
phone_numbers = [
    "920", "980", "951", "900", "930", "952", "9601", "910", "919", "9507", 
    "906", "903", "9081", "9155", "961", "995", "905", "958", "939", "969", 
    "9623", "992", "993", "999", "991", "90921", "901", "996", "9296", 
    "904", "933", "9675", "97761", "953", "984", "966778", "994", "9861", 
    "98186", "932506", "985657", "923891", "934477", "9821250", "95510006", 
    "941887530"
]

def clean_phone_number(number):
    """Удаляет все нецифровые символы из строки и начинает с первой '9'."""
    cleaned = re.sub(r'\D', '', str(number))  # Убираем все символы, кроме цифр
    start_index = cleaned.find('9')  # Находим первую '9'
    return cleaned[start_index:] if start_index != -1 else cleaned

def filter_rows_by_phone_numbers(df, column_name, phone_numbers):
    filtered_rows = []
    total_rows = len(df)
    
    # Получаем индекс столбца по его буквенной нотации
    column_index = string.ascii_uppercase.index(column_name[0])

    for index, row in tqdm(df.iterrows(), total=total_rows, desc="Обработка строк"):
        value = clean_phone_number(row[column_index])
        if any(value.startswith(num) for num in phone_numbers):
            filtered_rows.append(row)

    if filtered_rows:
        return pd.DataFrame(filtered_rows)  # Создаем DataFrame с отфильтрованными строками
    else:
        print_with_time("Нет строк, соответствующих номерам из списка.", color=Fore.RED)
        return None

if __name__ == "__main__":
    # Выбор исходных файлов
    source_files = select_files()
    
    if not source_files:
        print_with_time("Не выбран ни один файл.", color=Fore.RED)
        exit()
    
    if any(file.endswith('.csv') or file.endswith('.txt') for file in source_files):
        separator_dialog = Tk()
        separator_dialog.withdraw()
        csv_separator = simpledialog.askstring("Сепаратор CSV/TXT", "Введите сепаратор для CSV/TXT файлов:", initialvalue=";")
    
    # Обработка первого файла для выбора столбца
    first_file = source_files[0]
    print_with_time(f"Обработка первого файла для выбора столбца: {first_file}", color=Fore.BLUE)

    # Определение формата файла
    if first_file.endswith(('.xlsx', '.xls')):
        file_format = 'excel'
    elif first_file.endswith('.csv'):
        file_format = 'csv'
    elif first_file.endswith('.json'):
        file_format = 'json'
    elif first_file.endswith('.txt'):
        file_format = 'txt'
    else:
        print_with_time(f"Неподдерживаемый формат файла: {first_file}", color=Fore.RED)
        exit()

    # Чтение всех листов (если формат Excel)
    if file_format == 'excel':
        xls = pd.ExcelFile(first_file)
        sheets = xls.sheet_names
    else:
        sheets = [None]  # Для CSV, JSON и TXT используем псевдолист

    # Выбор столбца из первого листа или файла
    sheet_columns = {}
    
    for sheet_name in sheets:
        # Чтение первой НЕ пустой строки для выбора столбца
        if file_format == 'excel':
            df_headers = pd.read_excel(first_file, sheet_name=sheet_name, header=None)
            df_headers = df_headers.dropna(how='all').reset_index(drop=True)
            
            if df_headers.empty:
                print_with_time(f"Лист {sheet_name} содержит только пустые строки. Невозможно выбрать столбец.", color=Fore.RED)
                exit()
            
            headers = df_headers.iloc[0].astype(str).tolist()
            column_name = select_column(headers, sheet_name)
            sheet_columns[sheet_name] = column_name
        elif file_format in ('csv', 'txt'):
            df_headers = pd.read_csv(first_file, sep=csv_separator, nrows=1, encoding='utf-8')
            headers = df_headers.columns.astype(str).tolist()
            column_name = select_column(headers, 'CSV/TXT файл')
            sheet_columns['Sheet1'] = column_name
        elif file_format == 'json':
            df_headers = pd.read_json(first_file, lines=True, nrows=1)
            headers = df_headers.columns.astype(str).tolist()
            column_name = select_column(headers, 'JSON файл')
            sheet_columns['Sheet1'] = column_name

    if not sheet_columns:
        print_with_time("Не удалось выбрать столбец. Завершение работы.", color=Fore.RED)
        exit()

    # Применение выбранного столбца ко всем файлам
    for i, source_file in enumerate(source_files, 1):
        print_with_time(f"Обработка файла {i} из {len(source_files)}: {source_file}", color=Fore.BLUE)

        # Определение формата файла
        if source_file.endswith(('.xlsx', '.xls')):
            file_format = 'excel'
        elif source_file.endswith('.csv'):
            file_format = 'csv'
        elif source_file.endswith('.json'):
            file_format = 'json'
        elif source_file.endswith('.txt'):
            file_format = 'txt'
        else:
            print_with_time(f"Неподдерживаемый формат файла: {source_file}", color=Fore.RED)
            continue

        # Копирование строк и сохранение в новый файл
        for sheet_name, column_name in sheet_columns.items():
            if file_format == 'excel':
                df = pd.read_excel(source_file, sheet_name=sheet_name)
            elif file_format == 'csv':
                df = pd.read_csv(source_file, sep=csv_separator, encoding='utf-8')
            elif file_format == 'json':
                df = pd.read_json(source_file, lines=True)
            elif file_format == 'txt':
                df = pd.read_csv(source_file, sep=csv_separator, encoding='utf-8')

            filtered_df = filter_rows_by_phone_numbers(df, column_name, phone_numbers)

            if filtered_df is not None:
                save_path = generate_save_path(source_file)
                if file_format == 'excel':
                    with pd.ExcelWriter(save_path) as writer:
                        filtered_df.to_excel(writer, sheet_name=sheet_name, index=False)
                else:
                    filtered_df.to_csv(save_path, sep=csv_separator, index=False, encoding='utf-8')
                print_with_time(f"Файл сохранен: {save_path}", color=Fore.GREEN)
            else:
                print_with_time(f"Нет строк для сохранения в листе {sheet_name}.", color=Fore.RED)

    print_with_time("Обработка завершена.", color=Fore.GREEN)
