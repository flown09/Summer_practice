import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os


class FileComparator:
    def __init__(self, root):
        self.root = root
        self.file_paths = [None, None]
        self.dfs = [None, None]
        self.labels = [None, None]
        self.cmb_box_field = None
        self.condition_cb = None

        self.filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("ODS files", "*.ods"),
            ("All files", "*.*")
        ]

        self.setup_ui()

    def setup_ui(self):
        self.root.title("Сравнение файлов")
        self.root.geometry("400x270+400+200")
        self.root.resizable(False, True)
        self.root.minsize(400, 270)

        main_menu = tk.Menu(root)

        file_menu = tk.Menu(main_menu, tearoff=0)

        load_submenu = tk.Menu(file_menu, tearoff=0)
        load_submenu.add_command(label="Загрузить реестр 1", command=lambda: self.load_file(1))
        load_submenu.add_command(label="Загрузить реестр 2", command=lambda: self.load_file(2))

        file_menu.add_cascade(label="Загрузить", menu=load_submenu)

        file_menu.add_command(label="Очистить", command=self.clear_data)
        file_menu.add_command(label="Скачать шаблон", command=self.download_template)
        file_menu.add_command(label="Сравнить", command=self.compare_files)
        file_menu.add_separator()
        file_menu.add_command(label="Выход", command=root.quit)

        main_menu.add_cascade(label="Файл", menu=file_menu)
        main_menu.add_cascade(label="?")

        root.config(menu=main_menu)

        # --- Верхний слой: Загрузка данных ---
        top_frame = tk.LabelFrame(self.root, text=" 1. Загрузка данных ", font=('Arial', 10, 'bold'),
                                  padx=10, pady=10, relief=tk.GROOVE, bd=2)
        top_frame.pack(fill="x", padx=10, pady=5)

        btn_load1 = tk.Button(top_frame, text="Загрузить реестр 1", command=lambda: self.load_file(1),
                              width=20, height=2)
        btn_load2 = tk.Button(top_frame, text="Загрузить реестр 2", command=lambda: self.load_file(2),
                              width=20, height=2)
        btn_load1.grid(row=0, column=0, padx=14, pady=5)
        btn_load2.grid(row=0, column=1, padx=14, pady=5)

        self.labels[0] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
        self.labels[1] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
        self.labels[0].grid(row=1, column=0)
        self.labels[1].grid(row=1, column=1)

        # --- Средний слой: Условия сравнения ---
        self.middle_frame = tk.LabelFrame(self.root, text=" 2. Условия сравнения ", font=('Arial', 10, 'bold'),
                                          padx=10, pady=10, relief=tk.GROOVE, bd=2)
        self.middle_frame.pack(fill="x", padx=10, pady=5)

        options = ["Совпадают", "Не совпадают"]
        self.cmb_box_cond1 = ttk.Combobox(self.middle_frame, values=options, state='readonly', width=20)
        self.cmb_box_cond1.set(options[0])
        self.cmb_box_cond1.grid(row=0, column=0)

        # --- Нижний слой: Выгрузка результатов ---
        bottom_frame = tk.LabelFrame(self.root, text=" 3. Выгрузка результатов ", font=('Arial', 10, 'bold'),
                                     padx=10, pady=10, relief=tk.GROOVE, bd=2)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        btn_template = tk.Button(bottom_frame, text="Скачать шаблон", command=self.download_template)
        btn_template.grid(row=0, column=0, padx=0)

        btn_compare = tk.Button(bottom_frame, text="Сравнить и сохранить",
                                command=self.compare_files)
        btn_compare.grid(row=0, column=1, padx=(50, 0), sticky='w')

        btn_exit = tk.Button(bottom_frame, text="Закрыть", command=root.quit)
        btn_exit.grid(row=0, column=2, padx=10)

    def clear_data(self):
        self.dfs.clear()
        self.file_paths.clear()
        self.labels[0].config(text="Файл не загружен", fg="red")
        self.labels[1].config(text="Файл не загружен", fg="red")

    def download_template(self):
        data = {
            'Фамилия': ['Иванов', 'Петрова'],
            'Имя': ['Иван', 'Мария'],
            'Отчество': ['Иванович', 'Петровна'],
            'СНИЛС': ['123-456-789 01', '987-654-321 00'],
            'ИНН': ['123456789012', '098765432109'],
            'Серия и номер паспорта': ['1234 567890', '5678 901234']
        }

        template_df = pd.DataFrame(data)

        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=self.filetypes,
            title="Сохранить результат как",
            initialfile="Шаблон_реестра.xlsx"
        )

        if output_file:
            try:
                template_df.to_excel(output_file, index=False)
            except Exception as e:
                messagebox.showerror("Ошибка записи файла", str(e))

    def load_file(self, button_number):
        filename = filedialog.askopenfilename(title=f"Выберите файл {button_number}", filetypes=self.filetypes)
        if filename:
            self.file_paths[button_number - 1] = filename
            file_name_only = os.path.basename(filename)
            self.labels[button_number - 1].config(text=f"{file_name_only}", fg="green")

            # Сохраняем загруженные данные
            self.dfs[button_number - 1] = self.read_data(filename)

            # Обновляем информацию о столбцах
            self.update_columns_info()

    def read_data(self, file_path):
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

    def update_columns_info(self):
        """Обновляем информацию о столбцах и выводим в Combobox"""
        if self.dfs[0] is not None:
            # Если Combobox ещё не создан — создаём
            if self.cmb_box_field is None:
                self.cmb_box_field = ttk.Combobox(
                    self.middle_frame,
                    values=list(self.dfs[0].columns),
                    state='readonly',
                    width=20
                )
                self.cmb_box_field.grid(row=0, column=1, padx=20)
            else:  # Если уже создан — обновляем значения
                self.cmb_box_field['values'] = list(self.dfs[0].columns)

    def compare_files(self):
        if None in self.file_paths:
            messagebox.showwarning("Внимание", "Пожалуйста, загрузите оба файла.")
            return

        if self.dfs[0] is None or self.dfs[1] is None:
            return

        # Проверяем, что выбрано поле для сравнения
        if not self.cmb_box_field.get():
            messagebox.showwarning("Внимание", "Выберите поле для сравнения!")
            return

        # Получаем выбранное поле
        field = self.cmb_box_field.get()

        if field not in self.dfs[0].columns or field not in self.dfs[1].columns:
            messagebox.showwarning("Внимание", f"Поле '{field}' отсутствует в одном из файлов!")
            return

        # Получаем выбранное условие
        condition = self.cmb_box_cond1.get()

        # Выполняем сравнение в зависимости от выбранного условия
        if condition == "Совпадают":
            # Находим значения, которые есть в обоих файлах
            mask = self.dfs[0][field].isin(self.dfs[1][field])
            # Берем только данные из первого файла
            common_rows = self.dfs[0][mask]

        elif condition == "Не совпадают":
            mask1 = ~self.dfs[0][field].isin(self.dfs[1][field])
            unique_from_file1 = self.dfs[0][mask1]

            # Находим уникальные значения из второго файла
            mask2 = ~self.dfs[1][field].isin(self.dfs[0][field])
            unique_from_file2 = self.dfs[1][mask2]

            # Объединяем оба набора уникальных значений
            common_rows = pd.concat([unique_from_file1, unique_from_file2], ignore_index=True)
        else:
            messagebox.showwarning("Ошибка", "Неизвестное условие сравнения")
            return


        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=self.filetypes,
            title="Сохранить результат как",
            initialfile="Результат_сравнения.xlsx"  # Добавлено имя по умолчанию
        )

        if output_file:
            try:
                common_rows.to_excel(output_file, index=False)
            except Exception as e:
                messagebox.showerror("Ошибка записи файла", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = FileComparator(root)
    root.mainloop()