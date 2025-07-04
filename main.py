import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os

# Глобальные переменные для хранения данных
file_paths = [None, None]
cmb_box_field = None
dfs = [None, None]  # Здесь будем хранить загруженные DataFrame
labels = [None, None]
filetypes = [
    ("Excel files", "*.xlsx *.xls"),
    ("CSV files", "*.csv"),
    ("ODS files", "*.ods"),
    ("All files", "*.*")
]


def load_file(button_number):
    filename = filedialog.askopenfilename(title=f"Выберите файл {button_number}", filetypes=filetypes)
    if filename:
        file_paths[button_number - 1] = filename
        file_name_only = os.path.basename(filename)
        labels[button_number - 1].config(text=f"{file_name_only}", fg="green")

        # Сохраняем загруженные данные в dfs
        dfs[button_number - 1] = read_data(filename)

        # Обновляем информацию о столбцах (если нужно)
        update_columns_info()


def read_data(file_path):
    if not file_path:
        return None
    ext = file_path.split('.')[-1].lower()
    try:
        if ext in ['xlsx', 'xls']:
            return pd.read_excel(file_path)
        elif ext == 'csv':
            return pd.read_csv(file_path)
        elif ext == 'ods':
            return pd.read_excel(file_path, engine='odf')
        else:
            messagebox.showerror("Ошибка", f"Не поддерживаемый формат файла: {ext}")
            return None
    except Exception as e:
        messagebox.showerror("Ошибка чтения файла", str(e))
        return None


def compare_files():
    # Используем глобальные dfs
    global dfs

    if None in file_paths:
        messagebox.showwarning("Внимание", "Пожалуйста, загрузите оба файла.")
        return

    # Если данные ещё не загружены в dfs (на всякий случай)
    if dfs[0] is None:
        dfs[0] = read_data(file_paths[0])
    if dfs[1] is None:
        dfs[1] = read_data(file_paths[1])

    if dfs[0] is None or dfs[1] is None:
        return

    if set(dfs[0].columns) != set(dfs[1].columns):
        messagebox.showwarning("Внимание", "Названия или количество столбцов файлов не совпадают!")
        return

    common_rows = pd.merge(dfs[0], dfs[1], how='inner')

    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=filetypes,
        title="Сохранить результат как"
    )
    if output_file:  # Проверяем, что пользователь не отменил сохранение
        try:
            common_rows.to_excel(output_file, index=False)
            messagebox.showinfo("Готово", f"Результат сохранен в {output_file}")
        except Exception as e:
            messagebox.showerror("Ошибка записи файла", str(e))


def update_columns_info():
    """Обновляем информацию о столбцах и выводим в Combobox"""
    global cmb_box_field  # Используем глобальную переменную

    if dfs[0] is not None:
        print("Столбцы первого файла:", list(dfs[0].columns))

        # Если Combobox ещё не создан — создаём
        if cmb_box_field is None:
            cmb_box_field = ttk.Combobox(middle_frame, values=list(dfs[0].columns), state='readonly')
            cmb_box_field.grid(row=0, column=1)
        else:  # Если уже создан — обновляем значения
            cmb_box_field['values'] = list(dfs[0].columns)

    if dfs[1] is not None:
        print("Столбцы второго файла:", list(dfs[1].columns))

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Сравнение файлов")
    root.geometry("400x400+400+200")
    root.resizable(False, True)

    # --- Верхний слой ---
    top_frame = tk.LabelFrame(root, text=" 1. Загрузка данных ", font=('Arial', 10, 'bold'),
                              padx=10, pady=10, relief=tk.GROOVE, bd=2)
    top_frame.pack(fill="x", padx=10, pady=5)

    btn_load1 = tk.Button(top_frame, text="Загрузить реестр 1", command=lambda: load_file(1))
    btn_load2 = tk.Button(top_frame, text="Загрузить реестр 2", command=lambda: load_file(2))
    btn_load1.grid(row=0, column=0, padx=20, pady=5)
    btn_load2.grid(row=0, column=1, padx=20, pady=5)

    labels[0] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
    labels[1] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
    labels[0].grid(row=1, column=0)
    labels[1].grid(row=1, column=1)

    # --- Средний слой ---
    middle_frame = tk.LabelFrame(root, text=" 2. Условия сравнения ", font=('Arial', 10, 'bold'),
                                 padx=10, pady=10, relief=tk.GROOVE, bd=2)
    middle_frame.pack(fill="x", padx=10, pady=5)
    options = ["Совпадают", "Не совпадают"]
    cmb_box_cond1 = ttk.Combobox(middle_frame, values=options, state='readonly', width=20)
    cmb_box_cond1.set(options[0])
    cmb_box_cond1.grid(row=0, column=0)

    # --- Нижний слой ---
    bottom_frame = tk.LabelFrame(root, text=" 3. Выгрузка результатов ", font=('Arial', 10, 'bold'),
                                 padx=10, pady=10, relief=tk.GROOVE, bd=2)
    bottom_frame.pack(fill="x", padx=10, pady=5)

    # Передаем только имя функции без вызова и параметров
    btn_compare = tk.Button(bottom_frame, text="Сравнить и сохранить",
                            command=compare_files)
    btn_compare.pack(pady=10)

    print(dfs)


    root.mainloop()