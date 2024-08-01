import pandas as pd
from tqdm import tqdm
from time import time, strftime, localtime
from tkinter import Tk, filedialog, simpledialog, ttk
from colorama import init, Fore, Style

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

def select_column(headers, sheet_name):
    def on_select(event):
        root.quit()

    root = Tk()
    root.title(f"Выбор столбца для листа: {sheet_name}")
    label = ttk.Label(root, text=f"Выберите столбец для листа: {sheet_name}:")
    label.pack()
    combo = ttk.Combobox(root, values=headers)
    combo.bind("<<ComboboxSelected>>", on_select)
    combo.pack()
    center_window(root)
    root.mainloop()
    selected_column = combo.get()
    root.destroy()
    return selected_column

def save_file(source_file):
    root = Tk()
    root.withdraw()
    
    # Генерация имени файла
    base_name = source_file.split('/')[-1].split('.')[0]  # Получаем имя исходного файла без расширения
    file_name = f"{base_name}_Filtered.xlsx"
    
    # Выбор пути и имени файла для сохранения
    file_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                            initialfile=file_name,
                                            filetypes=[("Excel files", "*.xlsx *.xls")],
                                            title="Сохранить файл как")
    return file_path

def print_with_time(message, color=Fore.WHITE):
    current_time = strftime("%Y-%m-%d %H:%M:%S", localtime())
    print(f"{color}{current_time} - {message}{Style.RESET_ALL}")

def copy_rows_by_city(source_file, sheet_name, column_name, city, file_format):
    print_with_time(f"Чтение исходного файла {file_format} (лист: {sheet_name})...", color=Fore.BLUE)
    
    # Чтение исходного файла
    if file_format == 'excel':
        df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
    elif file_format == 'csv':
        df = pd.read_csv(source_file, header=None, encoding='utf-8')
    elif file_format == 'json':
        df = pd.read_json(source_file)
    elif file_format == 'txt':
        df = pd.read_csv(source_file, sep=txt_separator, header=None, encoding='utf-8')
    
    if df.empty:
        print_with_time(f"Лист {sheet_name} пустой. Нет данных для обработки.", color=Fore.RED)
        return None

    # Ищем первую НЕ пустую строку
    df = df.dropna(how='all').reset_index(drop=True)
    
    if df.empty:
        print_with_time(f"Лист {sheet_name} содержит только пустые строки.", color=Fore.RED)
        return None
    
    print_with_time("Файл прочитан успешно.", color=Fore.GREEN)
    
    # Получение индекса столбца
    headers = df.iloc[0]
    if column_name not in headers.values:
        print_with_time(f"Столбец {column_name} не найден в листе {sheet_name}.", color=Fore.RED)
        return None

    column_index = headers[headers == column_name].index[0]
    
    # Инициализация времени начала
    start_time = time()
    
    # Проход по строкам и копирование соответствующих строк
    print_with_time("Начало обработки строк...", color=Fore.BLUE)
    filtered_rows = []
    total_rows = len(df)
    for index, row in tqdm(df.iterrows(), total=total_rows, desc="Обработка строк"):
        value = row[column_index]
        if pd.notna(value) and isinstance(value, str) and city in value:
            filtered_rows.append(row)
        
        # Обновление прогресса
        progress = (index + 1) / total_rows * 100
        print_with_time(f'Обработано строк: {index + 1} из {total_rows} ({progress:.2f}%)', color=Fore.CYAN)
    
    if filtered_rows:
        filtered_df = pd.DataFrame(filtered_rows)  # Устанавливаем правильные столбцы
        return filtered_df
    else:
        print_with_time(f"Нет строк, соответствующих городу {city} в листе {sheet_name}.", color=Fore.RED)
        return None

def copy_rows_by_city_csv(source_file, column_name, city, sep):
    print_with_time(f"Чтение исходного файла CSV...", color=Fore.BLUE)
    
    # Чтение исходного файла
    df = pd.read_csv(source_file, sep=sep, encoding='utf-8')

    if df.empty:
        print_with_time(f"Файл {source_file} пустой. Нет данных для обработки.", color=Fore.RED)
        return None

    print_with_time("Файл прочитан успешно.", color=Fore.GREEN)
    
    # Проверка наличия столбца
    if column_name not in df.columns:
        print_with_time(f"Столбец {column_name} не найден в файле {source_file}.", color=Fore.RED)
        return None
    
    # Инициализация времени начала
    start_time = time()
    
    # Проход по строкам и копирование соответствующих строк
    print_with_time("Начало обработки строк...", color=Fore.BLUE)
    filtered_rows = df[df[column_name].apply(lambda x: pd.notna(x) and isinstance(x, str) and city in x)]
    
    if not filtered_rows.empty:
        return filtered_rows
    else:
        print_with_time(f"Нет строк, соответствующих городу {city}.", color=Fore.RED)
        return None

def copy_rows_by_city_txt(source_file, column_name, city, sep):
    print_with_time(f"Чтение исходного файла TXT...", color=Fore.BLUE)
    
    # Чтение исходного файла
    df = pd.read_csv(source_file, sep=sep, encoding='utf-8')

    if df.empty:
        print_with_time(f"Файл {source_file} пустой. Нет данных для обработки.", color=Fore.RED)
        return None

    print_with_time("Файл прочитан успешно.", color=Fore.GREEN)
    
    # Проверка наличия столбца
    if column_name not in df.columns:
        print_with_time(f"Столбец {column_name} не найден в файле {source_file}.", color=Fore.RED)
        return None
    
    # Инициализация времени начала
    start_time = time()
    
    # Проход по строкам и копирование соответствующих строк
    print_with_time("Начало обработки строк...", color=Fore.BLUE)
    filtered_rows = df[df[column_name].apply(lambda x: pd.notna(x) and isinstance(x, str) and city in x)]
    
    if not filtered_rows.empty:
        return filtered_rows
    else:
        print_with_time(f"Нет строк, соответствующих городу {city}.", color=Fore.RED)
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

        # Чтение всех листов (если формат Excel)
        if file_format == 'excel':
            xls = pd.ExcelFile(source_file)
            sheets = xls.sheet_names
        else:
            sheets = [None]  # Для CSV, JSON и TXT используем псевдолист

        sheet_columns = {}
        
        for sheet_name in sheets:
            # Чтение первой НЕ пустой строки для выбора столбца
            if file_format == 'excel':
                df_headers = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
                df_headers = df_headers.dropna(how='all').reset_index(drop=True)
                
                if df_headers.empty:
                    print_with_time(f"Лист {sheet_name} содержит только пустые строки. Невозможно выбрать столбец.", color=Fore.RED)
                    continue
                
                headers = df_headers.iloc[0].astype(str).tolist()
                column_name = select_column(headers, sheet_name)
                sheet_columns[sheet_name] = column_name
            elif file_format in ('csv', 'txt'):
                df_headers = pd.read_csv(source_file, sep=csv_separator, nrows=1, encoding='utf-8')
                headers = df_headers.columns.astype(str).tolist()
                column_name = select_column(headers, 'CSV/TXT файл')
                sheet_columns['Sheet1'] = column_name
            elif file_format == 'json':
                df_headers = pd.read_json(source_file, lines=True, nrows=1)
                headers = df_headers.columns.astype(str).tolist()
                column_name = select_column(headers, 'JSON файл')
                sheet_columns['Sheet1'] = column_name

        if not sheet_columns:
            print_with_time(f"Нет доступных столбцов для обработки в файле {source_file}.", color=Fore.RED)
            continue

        # Ввод города
        root = Tk()
        root.withdraw()
        city = simpledialog.askstring("Введите город", "Введите город для фильтрации:")
        root.update_idletasks()
        center_window(root)
        root.destroy()
        
        if not city:
            print_with_time("Город не был введен. Пропуск файла.", color=Fore.RED)
            continue
        
        # Копирование строк и сохранение в новый файл
        for sheet_name, column_name in sheet_columns.items():
            if file_format == 'excel':
                filtered = copy_rows_by_city(source_file, sheet_name, column_name, city, file_format)
            elif file_format == 'csv':
                filtered = copy_rows_by_city_csv(source_file, column_name, city, csv_separator)
            elif file_format == 'json':
                filtered = copy_rows_by_city(source_file, sheet_name, column_name, city, file_format)
            elif file_format == 'txt':
                filtered = copy_rows_by_city_txt(source_file, column_name, city, csv_separator)
            
            if filtered is not None and not filtered.empty:
                save_path = save_file(source_file)
                filtered.to_excel(save_path, index=False)
                print_with_time(f"Файл сохранен: {save_path}", color=Fore.GREEN)
            else:
                print_with_time(f"Нет строк для сохранения в листе {sheet_name}.", color=Fore.RED)

    print_with_time("Обработка завершена.", color=Fore.GREEN)
