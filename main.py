import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd

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

    output_file = filedialog.asksaveasfilename() #"comparison_result.xlsx"
    try:
        common_rows.to_excel(output_file, index=False)
        messagebox.showinfo("Готово", f"Результат сохранен в {output_file}")
    except Exception as e:
        messagebox.showerror("Ошибка записи файла", str(e))


if __name__ == "__main__":

    root = tk.Tk()
    root.title("Сравнение файлов")
    root.geometry("500x200+700+400")

    top_frame = tk.Frame(root)
    middle_frame = tk.Frame(root)
    bottom_frame = tk.Frame(root)

    top_frame.pack(side="top", fill="x", pady=10)
    bottom_frame.pack(side="bottom", fill="x", pady=10)

    btn_load1 = tk.Button(top_frame, text="Загрузить реестр 1", command=lambda: load_file(1))
    btn_load2 = tk.Button(top_frame, text="Загрузить реестр 2", command=lambda: load_file(2))
    btn_compare = tk.Button(bottom_frame, text="Сравнить файлы и сохранить результат", command=compare_files)

    btn_load1.pack(side='left', padx=40, ipadx=10, ipady=10)
    btn_load2.pack(side='right', padx=40, ipadx=10, ipady=10)
    btn_compare.pack(side="bottom", fill="x", pady=20)

    bottom_frame.pack_propagate(False)
    bottom_frame.configure(height=60)

    root.mainloop()