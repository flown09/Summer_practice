import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os

file_paths = [None, None]
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
        labels[button_number - 1].config(
            text=f"Загружен файл: {file_name_only}", fg="green"
        )

def read_data(file_path):
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

    common_rows = pd.merge(df1, df2, how='inner')

    output_file = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=filetypes,
        title="Сохранить результат как"
    )
    try:
        common_rows.to_excel(output_file, index=False)
        messagebox.showinfo("Готово", f"Результат сохранен в {output_file}")
    except Exception as e:
        messagebox.showerror("Ошибка записи файла", str(e))


if __name__ == "__main__":

    root = tk.Tk()
    root.title("Сравнение файлов")
    root.geometry("400x200+700+400")
    root.resizable(False, True)

    top_frame = tk.Frame(root)

    top_frame.pack(side="top", fill="x", pady=10)

    btn_load1 = tk.Button(top_frame, text="Загрузить реестр 1", command=lambda: load_file(1))
    btn_load2 = tk.Button(top_frame, text="Загрузить реестр 2", command=lambda: load_file(2))

    btn_load1.grid(row=0, column=0, padx=30)
    btn_load2.grid(row=0, column=1, padx=30)

    labels[0] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140, anchor="w", justify="left")
    labels[1] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140, anchor="w", justify="left")

    labels[0].grid(row=1, column=0, pady=(5, 0))
    labels[1].grid(row=1, column=1, pady=(5, 0))


    middle_frame = tk.Frame(root)


    bottom_frame = tk.Frame(root)
    bottom_frame.pack(side="bottom", fill="x", pady=10)

    btn_compare = tk.Button(bottom_frame, text="Сравнить файлы и сохранить результат", command=compare_files)
    btn_compare.pack(side="bottom", pady=20, ipadx=10, ipady=10)

    bottom_frame.pack_propagate(False)
    bottom_frame.configure(height=60)

    root.mainloop()