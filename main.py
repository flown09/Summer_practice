import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

# Глобальные переменные для хранения путей к файлам
file_paths = [None, None]


def load_file(button_number):
    filetypes = [
        ("Excel files", "*.xlsx *.xls"),
        ("CSV files", "*.csv"),
        ("ODS files", "*.ods"),
        ("All files", "*.*")
    ]
    filename = filedialog.askopenfilename(title=f"Выберите файл {button_number}", filetypes=filetypes)
    if filename:
        file_paths[button_number - 1] = filename
        #messagebox.showinfo("Файл выбран", f"Файл {button_number}:\n{filename}")


def read_data(file_path):
    # Определяем тип файла по расширению и читаем его
    if not file_path:
        return None
    ext = file_path.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls']:
            df = pd.read_excel(file_path)
        elif ext == 'csv':
            df = pd.read_csv(file_path)
        elif ext == 'ods':
            df = pd.read_excel(file_path, engine='odf')
        else:
            messagebox.showerror("Ошибка", f"Не поддерживаемый формат файла: {ext}")
            return None
        return df
    except Exception as e:
        messagebox.showerror("Ошибка чтения файла", str(e))
        return None


def compare_files():
    if None in file_paths:
        messagebox.showwarning("Внимание", "Пожалуйста, загрузите оба файла.")
        return

    df1 = read_data(file_paths[0])
    df2 = read_data(file_paths[1])

    if df1 is None or df2 is None:
        return

    # Пример сравнения: ищем общие строки по всем столбцам
    common_rows = pd.merge(df1, df2, how='inner')

    # Записываем результат в файл
    output_file = "comparison_result.xlsx"
    try:
        common_rows.to_excel(output_file, index=False)
        messagebox.showinfo("Готово", f"Результат сохранен в {output_file}")
    except Exception as e:
        messagebox.showerror("Ошибка записи файла", str(e))


# Создаем интерфейс
root = tk.Tk()
root.title("Сравнение файлов")
root.geometry("400x200")

btn_load1 = tk.Button(root, text="Загрузить 1 файл", command=lambda: load_file(1))
btn_load2 = tk.Button(root, text="Загрузить 2 файла", command=lambda: load_file(2))
btn_compare = tk.Button(root, text="Сравнить файлы и сохранить результат", command=compare_files)

btn_load1.pack(pady=10)
btn_load2.pack(pady=10)
btn_compare.pack(pady=20)

root.mainloop()