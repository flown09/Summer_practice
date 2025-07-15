import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
from functools import partial
import subprocess


class FileComparator:
    def __init__(self, root):
        self.root = root
        self.file_paths = [None, None]
        self.dfs = [None, None]
        self.labels = [None, None]
        self.condition_rows = []
        self.condition_frame = None

        # Константы для управления размерами
        self.ROW_HEIGHT = 40  # Высота одной строки условия
        self.MAX_VISIBLE_ROWS = 5  # Максимальное количество видимых строк перед включением скролла
        self.BASE_HEIGHT = 330  # Базовая высота окна

        self.filetypes = [
            ("All files", "*.*"),
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("ODS files", "*.ods")
        ]

        self.setup_ui()

    def setup_ui(self):
        # --- принудительная светлая тема для ttk ---
        style = ttk.Style(self.root)
        # выбор светлого базового стиля
        style.theme_use('clam')
        # глобально белый фон и чёрный текст для всех ttk‑виджетов
        style.configure('.', background='white', foreground='black')
        # специфическая настройка кнопок, лейблов, комбобоксов
        style.configure('TButton', background='#f0f0f0')
        style.configure('TLabel', background='white')
        style.configure('TFrame', background='white')
        style.configure('TEntry', fieldbackground='white', foreground='black')
        style.configure('TCombobox', fieldbackground='white', foreground='black')
        # убираем выделение у активных элементов (чтобы было неярко)
        style.map('TButton',
                  background=[('active', '#e0e0e0')],
                  foreground=[('disabled', 'grey')],
                  relief=[('pressed', 'sunken'), ('!pressed', 'raised')])

        # светлый фон для чистых tk‑фреймов/текстов
        self.root.configure(bg='white')
        self.root.title("Сравнение файлов")
        #self.root.geometry(f"450x{self.BASE_HEIGHT}+400+200")
        self.root.update_idletasks()
        self.root.minsize(530, 370)
        #self.root.resizable(False, True)
        #self.root.minsize(450, self.BASE_HEIGHT)

        # Главное меню
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
        main_menu.add_command(label="?", command=self.show_help)
        root.config(menu=main_menu)

        # --- Верхний слой: Загрузка данных ---
        top_frame = tk.LabelFrame(self.root, text=" 1. Загрузка данных ", font=('Arial', 10, 'bold'),
                                  padx=10, pady=10, relief=tk.GROOVE, bd=2)
        top_frame.pack(fill="x", padx=10, pady=5)

        btn_load1 = tk.Button(top_frame, text="Загрузить реестр 1", command=lambda: self.load_file(1),
                              width=22, height=2)
        btn_load2 = tk.Button(top_frame, text="Загрузить реестр 2", command=lambda: self.load_file(2),
                              width=22, height=2)
        btn_load1.grid(row=0, column=0, padx=19, pady=5)
        btn_load2.grid(row=0, column=1, padx=19, pady=5)

        self.labels[0] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
        self.labels[1] = tk.Label(top_frame, text="Файл не загружен", fg="red", wraplength=140)
        self.labels[0].grid(row=1, column=0)
        self.labels[1].grid(row=1, column=1)

        # --- Средний слой: Условия сравнения ---
        self.middle_frame = tk.LabelFrame(self.root, text=" 2. Условия сравнения ", font=('Arial', 10, 'bold'),
                                          padx=10, pady=10, relief=tk.GROOVE, bd=2)
        self.middle_frame.pack(fill="x", padx=10, pady=5)

        # Создаем контейнер для условий с возможностью прокрутки
        self.create_conditions_container()

        # Кнопка добавления нового условия
        btn_add_condition = tk.Button(
            self.middle_frame,
            text="+ Добавить условие",
            command=self.add_condition_row,
            bg="#e0e0e0"
        )
        btn_add_condition.pack(pady=(5, 0))

        # --- Нижний слой: Выгрузка результатов ---
        bottom_frame = tk.LabelFrame(self.root, text=" 3. Выгрузка результатов ", font=('Arial', 10, 'bold'),
                                     padx=10, pady=10, relief=tk.GROOVE, bd=2)
        bottom_frame.pack(fill="x", padx=10, pady=5)

        btn_template = tk.Button(bottom_frame, text="Скачать шаблон", command=self.download_template)
        btn_template.grid(row=0, column=0, padx=5)

        btn_compare = tk.Button(bottom_frame, text="Сравнить", command=self.confirm_comparison)
        btn_compare.grid(row=0, column=1, padx=(155, 0), sticky='w')

        btn_exit = tk.Button(bottom_frame, text="Закрыть", command=root.quit)
        btn_exit.grid(row=0, column=2, padx=10)

        # Добавляем первое условие
        self.add_condition_row()

    def show_help(self):
        #help_window = tk.Toplevel(self.root)
        #help_window.title("Инструкция по использованию")
        #help_window.geometry("600x400")
        #help_window.transient(self.root)
        #help_window.grab_set()

        #text = tk.Text(help_window, wrap="word", padx=10, pady=10)
        messagebox.showinfo("Инструкция",
                            """\
    Инструкция по использованию приложения:

    1. Загрузка файлов:
       - Нажмите «Загрузить реестр 1» и выберите первый файл.
       - Нажмите «Загрузить реестр 2» и выберите второй файл.
       Поддерживаются форматы: Excel (.xlsx, .xls), CSV и ODS.

    2. Добавление условий сравнения:
       - Нажмите кнопку «+ Добавить условие».
       - Выберите поле для сравнения (например, СНИЛС).
       - Укажите тип условия: «Совпадают» или «Не совпадают».
       - Если добавляется несколько условий — выберите логику (И / ИЛИ).

    3. Сравнение файлов:
       - Нажмите «Сравнить».
       - Появится окно с количеством найденных строк.
       - Подтвердите сохранение результата.

    4. Скачивание шаблона:
       - Кнопка «Скачать шаблон» позволяет загрузить пример Excel-файла для заполнения.

    5. Очистка данных:
       - Вы можете сбросить загруженные файлы и условия через пункт меню «Очистить».

    Результаты сравнения сохраняются в файл.
    """)
        #text.config(state="disabled")
        #text.pack(fill="both", expand=True)

        # Кнопка закрытия
        #close_btn = tk.Button(help_window, text="Закрыть", command=help_window.destroy)
        #close_btn.pack(pady=5)

    def confirm_comparison(self):
        """Показывает модальное окно с количеством строк перед сравнением"""
        if None in self.file_paths or self.dfs[0] is None or self.dfs[1] is None:
            messagebox.showwarning("Внимание", "Пожалуйста, загрузите оба файла.")
            return

        # Новая проверка: совпадение названий столбцов
        cols1 = list(self.dfs[0].columns)
        cols2 = list(self.dfs[1].columns)
        set1, set2 = set(cols1), set(cols2)
        if set1 != set2:
            missing_in_1 = set2 - set1
            missing_in_2 = set1 - set2
            msg = "Набор столбцов в файлах не совпадает!\n"
            if missing_in_1:
                msg += f"Столбцы, отсутствующие в файле 1: {', '.join(missing_in_1)}\n"
            if missing_in_2:
                msg += f"Столбцы, отсутствующие в файле 2: {', '.join(missing_in_2)}"
            messagebox.showerror("Ошибка", msg)
            return

        conditions = []
        for i, row in enumerate(self.condition_rows):
            logic = row.get("logic_cb").get() if i > 0 else "И"
            condition_type = row["cond_cb"].get()
            field = row["field_cb"].get()

            if not field:
                messagebox.showwarning("Внимание", f"Выберите поле в условии #{i + 1}!")
                return
            if field not in self.dfs[0].columns or field not in self.dfs[1].columns:
                messagebox.showwarning("Внимание", f"Поле '{field}' отсутствует в одном из файлов!")
                return

            conditions.append((field, condition_type, logic))

        result_df = self.apply_conditions(conditions)
        row_count = len(result_df)

        popup = tk.Toplevel(self.root)
        popup.title("Результат сравнения")
        popup.resizable(False, False)
        popup.transient(self.root)
        popup.grab_set()

        # Центрирование относительно главного окна
        self.root.update_idletasks()
        root_x = self.root.winfo_rootx()
        root_y = self.root.winfo_rooty()
        root_width = self.root.winfo_width()
        root_height = self.root.winfo_height()

        popup_width = 300
        popup_height = 120
        center_x = root_x + (root_width - popup_width) // 2
        center_y = root_y + (root_height - popup_height) // 2
        popup.geometry(f"{popup_width}x{popup_height}+{center_x}+{center_y}")

        # Виджеты
        tk.Label(popup, text=f"Найдено строк: {row_count}", font=("Arial", 11)).pack(pady=10)

        btn_frame = tk.Frame(popup)
        btn_frame.pack(pady=10)

        def proceed():
            popup.destroy()
            self.compare_files()

        tk.Button(btn_frame, text="Сохранить", command=proceed).pack(side="left", padx=10)
        tk.Button(btn_frame, text="Отмена", command=popup.destroy).pack(side="left", padx=10)

    def create_conditions_container(self):
        """Создает контейнер для условий с возможностью прокрутки"""
        # Основной фрейм для контейнера
        container_frame = tk.Frame(self.middle_frame)
        container_frame.pack(fill="both", expand=True)

        # Создаем холст (Canvas) и скроллбар
        self.canvas = tk.Canvas(container_frame)
        scrollbar = ttk.Scrollbar(container_frame, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)

        # Обёртка для условий
        self.condition_frame = tk.Frame(self.canvas)
        self.condition_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.condition_frame, anchor="nw")

        # Добавляем прокрутку мышью
        self.canvas.bind_all("<MouseWheel>", self._on_mousewheel)

    def _on_mousewheel(self, event):
        self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def add_condition_row(self):
        # Создаем фрейм для строки условия
        row_frame = tk.Frame(self.condition_frame)
        row_frame.pack(fill="x", pady=2)

        logic_cb = ttk.Combobox(row_frame, values=["И", "ИЛИ"], state='readonly', width=5)

        if len(self.condition_rows) == 0:
            # Первая строка: делаем пустой и неактивный
            logic_cb.set("")  # Пустой текст
            logic_cb.configure(state="disabled")  # Деактивировать выбор
        else:
            logic_cb.set("И")  # По умолчанию можно "И" или "ИЛИ"

        logic_cb.pack(side="left", padx=5)

        # Выбор типа условия
        cond_options = ["Совпадают", "Не совпадают"]
        cond_cb = ttk.Combobox(row_frame, values=cond_options, state='readonly', width=15)
        cond_cb.set(cond_options[0])
        cond_cb.pack(side="left", padx=5)

        # Выбор поля для сравнения
        field_cb = ttk.Combobox(row_frame, state='readonly', width=20)
        field_cb.pack(side="left", padx=5)

        # Если данные уже загружены, обновляем значения
        if self.dfs[0] is not None:
            field_cb['values'] = list(self.dfs[0].columns)
            if self.dfs[0].columns.size > 0:
                field_cb.set(self.dfs[0].columns[0])

        # Кнопка удаления условия
        if len(self.condition_rows) > 0:  # Только если это не первая строка
            btn_remove = tk.Button(
                row_frame,
                text="×",
                fg="red",
                font=("Arial", 10, "bold"),
                command=partial(self.remove_condition_row, row_frame),
                width=2
            )
            btn_remove.pack(side="left", padx=5)

        # Сохраняем информацию о строке
        self.condition_rows.append({
            "frame": row_frame,
            "logic_cb": logic_cb,
            "cond_cb": cond_cb,
            "field_cb": field_cb
        })

        # Обновляем размеры окна
        self.update_window_size()

    def remove_condition_row(self, row_frame):
        """Удаляет строку с условием"""
        # Находим и удаляем строку
        for i, row in enumerate(self.condition_rows):
            if row["frame"] == row_frame:
                row["frame"].destroy()
                self.condition_rows.pop(i)
                break

        # Если это была последняя строка, добавляем новую пустую
        if len(self.condition_rows) == 0:
            self.add_condition_row()

        # Обновляем размеры окна
        self.update_window_size()

    def update_window_size(self):
        """Обновляет размер окна в зависимости от количества условий"""
        # Вычисляем новую высоту
        # visible_rows = min(len(self.condition_rows), self.MAX_VISIBLE_ROWS)
        # extra_height = visible_rows * self.ROW_HEIGHT
        # new_height = self.BASE_HEIGHT + extra_height
        visible_rows = len(self.condition_rows)
        if visible_rows > self.MAX_VISIBLE_ROWS:
            extra_height = self.MAX_VISIBLE_ROWS * self.ROW_HEIGHT
        else:
            extra_height = visible_rows * self.ROW_HEIGHT

        new_height = self.BASE_HEIGHT + extra_height - 40

        # Устанавливаем новую высоту
        self.root.geometry(f"400x{new_height}")

        # Настраиваем высоту холста
        self.canvas.configure(height=min(visible_rows * self.ROW_HEIGHT, self.MAX_VISIBLE_ROWS * self.ROW_HEIGHT))

    def clear_data(self):
        """Очищает все данные и сбрасывает интерфейс"""
        self.dfs = [None, None]
        self.file_paths = [None, None]
        self.labels[0].config(text="Файл не загружен", fg="red")
        self.labels[1].config(text="Файл не загружен", fg="red")

        # Удаляем все условия, кроме одного
        while len(self.condition_rows) > 1:
            self.remove_condition_row(self.condition_rows[-1]["frame"])

        # Сбрасываем первое условие
        if self.condition_rows:
            row = self.condition_rows[0]
            row["cond_cb"].set("Совпадают")
            row["field_cb"].set('')
            row["logic_cb"].set("")

    def download_template(self):
        """Скачивает шаблон файла"""
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
            title="Сохранить шаблон как",
            initialfile="Шаблон_реестра.xlsx"
        )

        if output_file:
            try:
                template_df.to_excel(output_file, index=False)
                messagebox.showinfo("Успех", "Шаблон успешно сохранён!")
            except Exception as e:
                messagebox.showerror("Ошибка записи файла", str(e))

    def load_file(self, button_number):
        """Загружает файл"""
        filename = filedialog.askopenfilename(title=f"Выберите файл {button_number}", filetypes=self.filetypes)
        if filename:
            self.file_paths[button_number - 1] = filename
            file_name_only = os.path.basename(filename)
            self.labels[button_number - 1].config(text=f"{file_name_only}", fg="green")

            # Сохраняем загруженные данные
            self.dfs[button_number - 1] = self.read_data(filename)
            self.update_field_comboboxes()

    def read_data(self, file_path):
        """Читает данные из файла"""
        if not file_path:
            return None
        ext = file_path.split('.')[-1].lower()
        try:
            if ext in ['xlsx', 'xls']:
                # Читаем все листы как словарь DataFrame
                sheets = pd.read_excel(file_path, sheet_name=None)
                # Склеиваем все листы по строкам, добавляем колонку 'Лист' при необходимости
                df = pd.concat(
                    sheets.values(),
                    ignore_index=True
                )
                return df
            elif ext == 'csv':
                return pd.read_csv(file_path)
            elif ext == 'ods':
                # Аналогично для ODS
                sheets = pd.read_excel(file_path, engine='odf', sheet_name=None)
                df = pd.concat(sheets.values(), ignore_index=True)
                return df
            else:
                messagebox.showerror("Ошибка", f"Не поддерживаемый формат файла: {ext}")
                return None
        except Exception as e:
            messagebox.showerror("Ошибка чтения файла", str(e))
            return None

    def update_field_comboboxes(self):
        """Обновляет выпадающие списки с полями"""
        if self.dfs[0] is not None:
            columns = list(self.dfs[0].columns)
            for row in self.condition_rows:
                field_cb = row["field_cb"]
                field_cb['values'] = columns
                if columns:
                    field_cb.set(columns[0])

    def compare_files(self):
        """Выполняет сравнение файлов по заданным условиям"""
        if None in self.file_paths:
            messagebox.showwarning("Внимание", "Пожалуйста, загрузите оба файла.")
            return

        if self.dfs[0] is None or self.dfs[1] is None:
            return

        # Проверяем условия
        conditions = []
        for i, row in enumerate(self.condition_rows):
            logic = row["logic_cb"].get() if i > 0 else "И"  # Первый — всегда "И"
            condition_type = row["cond_cb"].get()
            field = row["field_cb"].get()

            if not field:
                messagebox.showwarning("Внимание", f"Выберите поле в условии #{i + 1}!")
                return

            conditions.append((field, condition_type, logic))

        # Применяем условия
        result_df = self.apply_conditions(conditions)

        # Сохраняем результат
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=self.filetypes,
            title="Сохранить результат как",
            initialfile="Результат_сравнения.xlsx"
        )

        if output_file:
            try:
                result_df.to_excel(output_file, index=False)
                #messagebox.showinfo("Успех", "Результат успешно сохранён!")
                popup = tk.Toplevel(self.root)
                popup.title("Успех")
                popup.resizable(False, False)
                popup.transient(self.root)
                popup.grab_set()

                # Центрирование
                self.root.update_idletasks()
                root_x = self.root.winfo_rootx()
                root_y = self.root.winfo_rooty()
                root_width = self.root.winfo_width()
                root_height = self.root.winfo_height()

                popup_width = 320
                popup_height = 120
                center_x = root_x + (root_width - popup_width) // 2
                center_y = root_y + (root_height - popup_height) // 2
                popup.geometry(f"{popup_width}x{popup_height}+{center_x}+{center_y}")

                # Текст
                tk.Label(popup, text="Результат успешно сохранён!", font=("Arial", 11)).pack(pady=10)

                # Кнопки
                btn_frame = tk.Frame(popup)
                btn_frame.pack(pady=10)

                def open_directory():
                    folder = os.path.dirname(output_file)
                    try:
                        if os.name == 'nt':  # Windows
                            os.startfile(folder)
                        elif os.name == 'posix':  # Linux, macOS
                            subprocess.call(['xdg-open', folder])
                        else:
                            messagebox.showinfo("Информация", f"Путь к папке: {folder}")
                    except Exception as e:
                        messagebox.showerror("Ошибка", f"Не удалось открыть папку:\n{e}")

                tk.Button(btn_frame, text="Открыть папку", command=open_directory).pack(side="left", padx=10)
                tk.Button(btn_frame, text="ОК", command=popup.destroy).pack(side="left", padx=10)
            except Exception as e:
                messagebox.showerror("Ошибка записи файла", str(e))

    def apply_conditions(self, conditions):
        """Применяет все условия к объединённым данным с логикой И / ИЛИ"""
        df1 = self.dfs[0]
        df2 = self.dfs[1]
        combined = pd.concat([df1, df2], ignore_index=True)

        result_mask = pd.Series([True] * len(combined))

        for i, (field, condition_type, logic) in enumerate(conditions):
            # Приводим значения к строке и нижнему регистру
            df1_vals = df1[field].astype(str).str.lower()
            df2_vals = df2[field].astype(str).str.lower()
            combined_vals = combined[field].astype(str).str.lower()

            if condition_type == "Совпадают":
                values = set(df1_vals) & set(df2_vals)
            elif condition_type == "Не совпадают":
                values = set(df1_vals) ^ set(df2_vals)
            else:
                raise ValueError(f"Неизвестный тип условия: {condition_type}")

            condition_mask = combined_vals.isin(values)

            if i == 0:
                result_mask = condition_mask
            else:
                if logic == "И":
                    result_mask &= condition_mask
                elif logic == "ИЛИ":
                    result_mask |= condition_mask
                else:
                    raise ValueError(f"Неизвестная логика: {logic}")

        result_df = combined[result_mask].copy()

        # Приводим все строки к нижнему регистру для сравнения и удаления дубликатов
        normalized_df = result_df.apply(lambda col: col.astype(str).str.lower())

        # Удаляем дубликаты по приведённым к нижнему регистру данным
        mask = ~normalized_df.duplicated()

        # Возвращаем оригинальные строки, но только те, что уникальны в нижнем регистре
        return result_df[mask].reset_index(drop=True)


if __name__ == "__main__":
    root = tk.Tk()
    app = FileComparator(root)
    root.mainloop()
