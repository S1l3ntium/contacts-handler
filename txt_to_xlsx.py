import os
import pandas as pd
import datetime
from termcolor import colored
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import gc

def print_colored_message(message):
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(colored(f"{timestamp} - {message}", 'green'))

def choose_file():
    print_colored_message("Выбор файла.")
    Tk().withdraw()  # Скрыть окно Tkinter
    file_path = askopenfilename()
    print_colored_message(f"Файл выбран: {file_path}")
    return file_path

def process_and_write_chunk(file, chunk_size, delimiter, file_index, base_output_path):
    data = []
    for _ in range(chunk_size):
        line = file.readline()
        if not line:
            break
        data.append(line.strip().split(delimiter))
    
    if data:
        df = pd.DataFrame(data)
        output_file = f"{base_output_path}_part{file_index}.xlsx"
        df.to_excel(output_file, index=False, header=False)
        print_colored_message(f"Файл {output_file} создан с {len(data)} строками.")
        del data  # Удаление данных после записи
        gc.collect()  # Явный вызов сборщика мусора

def read_and_write_in_chunks(file_path, delimiter, chunk_size):
    print_colored_message("Чтение файла и запись в отдельные xlsx файлы порциями.")
    base_output_path = os.path.splitext(file_path)[0]
    
    with open(file_path, 'r', encoding='utf-8') as file:
        file_index = 1
        while True:
            current_position = file.tell()
            process_and_write_chunk(file, chunk_size, delimiter, file_index, base_output_path)
            if current_position == file.tell():
                break
            file_index += 1
            print_colored_message(f"Чанк {file_index} обработан и записан в новый файл.")
    
    print_colored_message("Все части успешно сохранены.")

def main():
    file_path = choose_file()
    delimiter = input("Введите делитель (по умолчанию табуляция): ") or "\t"
    
    chunk_size = 100000  # Количество строк в одном файле Excel
    read_and_write_in_chunks(file_path, delimiter, chunk_size)
    
    print_colored_message("Конвертация завершена.")

if __name__ == "__main__":
    main()
