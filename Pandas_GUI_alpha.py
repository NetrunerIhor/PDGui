import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from fpdf import FPDF
import os

class DataProcessor:
    def __init__(self, data):
        self.data = data

    def filter_data(self, column, condition):
        """
        Фільтрація даних за вказаним стовпцем і умовою.
        :param column: Назва стовпця
        :param condition: Функція умови для фільтрації
        :return: Відфільтровані дані
        """
        if column not in self.data.columns:
            raise ValueError(f"Стовпець {column} відсутній у даних.")
        return self.data[self.data[column].apply(condition)]

    def clean_data(self):
        """
        Очищення даних:
        - Заповнення відсутніх значень середнім.
        - Видалення дублювань.
        :return: Очищені дані
        """
        self.data.fillna(self.data.mean(numeric_only=True), inplace=True)
        self.data.drop_duplicates(inplace=True)

    def calculate_statistics(self):
        """
        Обчислення базової статистики для числових стовпців.
        :return: DataFrame зі статистикою
        """
        stats = self.data.describe(include=[np.number])
        return stats

    
    def generate_report(self, output_path, selected_columns=None, include_graphics=True):
        """
        Генерація звіту у форматі PDF.
        :param output_path: Шлях до файлу звіту.
        :param selected_columns: Вибрані стовпці для звіту (за замовчуванням - всі).
        :param include_graphics: Чи включати графіки у звіт.
        """
        from tkinter import filedialog

        pdf = FPDF()
        pdf.add_page()

        # Підключення шрифту DejaVuSans для підтримки Unicode
        font_path = "fonts/DejaVuSans-Bold.ttf"
        if not os.path.exists(font_path):
            print("Шрифт не знайдений!")
            return

        pdf.add_font("DejaVu", "", font_path, uni=True)
        pdf.set_font("DejaVu", size=12)

        # Заголовок
        pdf.set_font("DejaVu", size=16)
        pdf.cell(200, 10, txt="Звіт про дані", ln=True, align="C")
        pdf.ln(10)

        # Вибір стовпців (якщо користувач нічого не вибрав, використовуються всі)
        if selected_columns is None:
            selected_columns = self.data.columns

        # Фільтрування даних за вибраними стовпцями
        filtered_data = self.data[selected_columns]

        # Базова статистика
        pdf.set_font("DejaVu", size=12)
        pdf.cell(200, 10, txt="Базова статистика:", ln=True)
        pdf.ln(5)

        stats = filtered_data.describe(include="all")  # Генерація статистики
        col_width = 45  # Ширина стовпців у таблиці
        row_height = 8  # Висота рядків у таблиці
        page_width = pdf.w - 35  # Ширина сторінки (з урахуванням відступів)
        max_cols_per_page = int(page_width // col_width)  # Максимальна кількість стовпців на сторінці
        max_rows_per_page = int((pdf.h - 40) // row_height)
        # Розбиття таблиці на частини, якщо стовпці не вміщуються
        pdf.set_font("DejaVu", size=10)
        for start_col in range(0, len(stats.columns), max_cols_per_page):
            end_col = start_col + max_cols_per_page
            subset_columns = stats.columns[start_col:end_col]

            # Заголовок таблиці
            pdf.cell(col_width, row_height, txt="Параметр", border=1, align="C")
            for col in subset_columns:
                pdf.cell(col_width, row_height, txt=str(col), border=1, align="C")
            pdf.ln(row_height)

            # Дані таблиці
            for index, row in stats.iterrows():
                if pdf.get_y() + row_height > pdf.h - 20 :
                    pdf.add_page()
                    pdf.cell(col_width, row_height, txt="Параметр", border=1, align="C")
                    for col in subset_columns:
                        pdf.cell(col_width, row_height, txt=str(col), border=1, align="C")
                    pdf.ln(row_height)

                pdf.cell(col_width, row_height, txt=str(index), border=1)
                for col in subset_columns:
                    value = row[col]
                    cell_value = f"{value:.2f}" if pd.api.types.is_numeric_dtype(value) else str(value)
                    pdf.cell(col_width, row_height, txt=cell_value, border=1)
                pdf.ln(row_height)

            # Перехід на нову сторінку для наступної частини таблиці
            pdf.ln(5)
            #pdf.add_page()

        # Додавання графіків
        if include_graphics:
            pdf.add_page()
            pdf.set_font("DejaVu", size=14)
            pdf.cell(200, 10, txt="Графіки:", ln=True)
            pdf.ln(10)

            # Запит у користувача на вибір файлів із графіками
            image_files = filedialog.askopenfilenames(
                title="Оберіть файли графіків",
                filetypes=[("Зображення", "*.png;*.jpg;*.jpeg;*.bmp")],
            )

            for image_path in image_files:
                try:
                    pdf.image(image_path, x=10, y=pdf.get_y(), w=170)
                    pdf.ln(130)
                except Exception as e:
                    print(f"Не вдалося додати графік {image_path}: {e}")

        # Збереження звіту
        pdf.output(output_path, "F")
        print(f"Звіт збережено у файл: {output_path}")

class DataLoaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Завантаження та перегляд даних")
        self.root.geometry("800x600")
        self.root.minsize(800,600)

        # Змінна для збереження даних
        self.data = None
        self.original_data = None
        self.processor = None
        self.selected_columns = set()
        # Інтерфейс
        self.is_dark_mode = False
        self.create_widgets(self.root)
        self.setup_edit_frame()
        self.setup_processing_widgets()
        #self.open_processing_window()

        self.apply_widget_styles()
        self.enable_light_mode()
        self.root.columnconfigure(0, weight=2)
        self.root.rowconfigure(1, weight=1)
    
        self.tree.bind("<Control-Button-1>", self.on_column_select)
        

    def create_widgets(self, root):
        # Заголовок таблиці
        self.label = tk.Label(self.root, text="Попередній перегляд даних:", font=("Arial", 12))
        self.label.grid(row=0, column=0)

        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.grid(row=1, column=0, sticky="nsew")
        
    
        # Скролбар для таблиці
        scrollbar_y = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscroll=scrollbar_y.set)
        scrollbar_y.grid(row=1, column=0, sticky="ens",)

        scrollbar_x = ttk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscroll=scrollbar_x.set)
        scrollbar_x.grid(row=1, column=0, sticky="wes",)

        self.tree.bind('<<TreeviewSelect>>', self.edit_selected_item)

        self.mainmenu = tk.Menu(self.root)
        self.root.config(menu=self.mainmenu)
        filemenu = tk.Menu(self.mainmenu, tearoff = 0)
        filemenu.add_command(label="Відкрити", command=self.load_data)
        filemenu.add_command(label="Зберегти", command=self.save_data)
        self.mainmenu.add_cascade(label="Файл", menu=filemenu)

        self.mainmenu.add_command(label="Темний режим", command=self.toggle_theme)

        self.mainmenu.add_command(label="Довідка", command=self.open_help_window)
        
        #self.report_button = tk.Button(root, text="Створити звіт", command=self.create_report, state="disabled")
        #self.report_button.grid(row=5, column=1,)

        #self.plot_button = tk.Button(self.root, text="Побудувати графік", command=self.plot_selected_columns)
        #self.plot_button.grid(row=4, column=1,)

        #self.apply_widget_styles()
    
    def setup_processing_widgets(self):
        """
        Створює віджети для обробки даних у головному вікні.
        """
        # Рамка для обробки даних
        self.processing_frame = tk.LabelFrame(
            self.root, text="Обробка даних", font=("Arial", 10, "bold")
        )
        self.processing_frame.grid(row=2, column=0, columnspan=2, sticky="nsew", padx=10, pady=10)

        # Очищення даних
        clean_button = tk.Button(
            self.processing_frame,
            text="Очистити дані",
            command=self.clean_data,
        )
        clean_button.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Фільтрація даних
        tk.Label(
            self.processing_frame,
            text="Фільтрація даних:",
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        filter_button = tk.Button(
            self.processing_frame,
            text="Фільтрувати",
            command=self.apply_filter,
        )
        filter_button.grid(row=3, column=1, padx=10, pady=5, sticky="w")

        reset_button = tk.Button(self.processing_frame, text="Скинути фільтри", command=self.reset_filters)
        reset_button.grid(row=3, column=0, pady=10, sticky="w")

        self.report_button = tk.Button(self.processing_frame, text="Створити звіт", command=self.create_report, state="disabled")
        self.report_button.grid(row=4, column=0,sticky="w")

        self.plot_button = tk.Button(self.processing_frame, text="Побудувати графік", command=self.plot_selected_columns)
        self.plot_button.grid(row=4, column=1,sticky="w")
        
        #self.root.grid_rowconfigure(2, weight=1)
        #self.processing_frame.grid_columnconfigure(1, weight=1)

    def setup_processing_widgets_data_loadet(self):

        processing_frame = tk.Label(self.processing_frame, text="Обробка даних", font=("Arial", 10, "bold"))
        processing_frame.grid(row=2, column= 0 ,  sticky="ns", padx=10, pady=10)

        self.column_combo = ttk.Combobox(
            self.processing_frame, values=list(self.data.columns), state="readonly"
        )
        self.column_combo.grid(row=2, column=1, padx=10, pady=5, sticky="w")

        self.condition_entry = tk.Entry(self.processing_frame)
        self.condition_entry.grid(row=2, column=2, padx=10, pady=5, sticky="w")



    def reset_filters(self):
        """
        Скидає всі фільтри та повертає таблицю до початкового стану.
        """
        if self.original_data is None:
            #messagebox.showwarning("Увага", "Оригінальні дані не завантажені або порожні!")
            return
        # Повернення даних до початкового стану
        self.data = self.original_data.copy()
        self.processor = DataProcessor(self.data)
        # Оновлення таблиці
        self.update_tree(self.data)
    
    def apply_filter(self):
        """
        Застосовує фільтрацію даних до таблиці self.tree.
        """
        if self.data is None or self.data.empty:
            messagebox.showwarning("Увага", "Дані не завантажено або порожні!")
            return

        column = self.column_combo.get()
        condition_str = self.condition_entry.get()

        if not column or not condition_str:
            messagebox.showwarning("Увага", "Будь ласка, виберіть стовпець та умову для фільтрації!")
            return

        try:
            # Створення умови фільтрації
            condition = eval(f"lambda x: {condition_str}")

            # Застосування фільтрації
            filtered_data = self.processor.filter_data(column, condition)

            # Оновлення даних і таблиці
            self.update_tree(filtered_data)
            self.data = filtered_data  # Оновлюємо внутрішні дані
            self.processor = DataProcessor(self.data)  # Оновлюємо обробник даних

            messagebox.showinfo("Успіх", "Фільтрацію застосовано!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося застосувати фільтрацію: {e}")
   
    def clean_data(self):
        """
        Виклик очищення даних через DataProcessor.
        """
        if self.processor:
            self.processor.clean_data()
            self.data = self.processor.data
            messagebox.showinfo("Успіх", "Дані очищено!")
            for row in self.tree.get_children():
                self.tree.delete(row)
            
            for _, row in self.data.iterrows():
                self.tree.insert("", "end", values=row.tolist())

        # Вставка очищених даних у Treeview
        for _, row in self.data.iterrows():
            self.tree.insert("", "end", values=row.tolist())


    def update_tree(self, new_data):
        """
        Оновлює таблицю self.tree з новими даними.
        :param new_data: DataFrame із даними для відображення.
        """
        # Очистка таблиці
        for row in self.tree.get_children():
            self.tree.delete(row)

        # Оновлення колонок
        self.tree["columns"] = list(new_data.columns)
        for col in new_data.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, anchor="center")

        # Додавання нових даних
        for index, row in new_data.iterrows():
            self.tree.insert("", "end", values=row.tolist())

    def on_column_select(self, event):
        # Отримуємо обраний стовпець 
        region = self.tree.identify_region(event.x, event.y)
        if region == "heading":
            column_id = self.tree.identify_column(event.x)

            # Індекс стовпця
            col_index = int(column_id.replace("#", "")) - 1
            col_name = self.tree["columns"][col_index]

            # Додаємо або видаляємо зі списку вибраних
            if col_name in self.selected_columns:
                self.selected_columns.remove(col_name)
                self.tree.heading(col_name, text=col_name, anchor="center")  # Скидаємо стиль
            else:
                self.selected_columns.add(col_name)
                self.tree.heading(col_name, text=f"✔ {col_name}", anchor="center")  # Додаємо індикатор вибору

    def load_data(self):
        """
        Завантаження даних із файлу.
        """
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return

        try:
            if file_path.endswith(".csv"):
                self.data = pd.read_csv(file_path)
            elif file_path.endswith(".xlsx"):
                self.data = pd.read_excel(file_path)
            else:
                messagebox.showerror("Помилка", "Непідтримуваний формат файлу!")
                return
            # Зберігаємо оригінальні дані
            self.original_data = self.data.copy()

            self.processor = DataProcessor(self.data)
            self.update_table()
            self.report_button.config(state="normal")
            self.setup_processing_widgets_data_loadet()
            #self.open_processing_window()
            #self.process_button.config(state="normal")

        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося завантажити файл: {e}")
    

        # Розтягнення рамки обробки даних
        #self.root.grid_rowconfigure(2, weight=1)
        #processing_frame.grid_columnconfigure(1, weight=1)

    def filter_data(self, column, condition_str):
        """
        Виклик фільтрації даних через DataProcessor.
        :param column: Назва стовпця для фільтрації
        :param condition_str: Строкова умова для фільтрації
        """
        if not column or not condition_str:
            messagebox.showwarning("Увага", "Будь ласка, заповніть усі поля для фільтрації!")
            return

        try:
            condition = eval(f"lambda x: {condition_str}")
            filtered_data = self.processor.filter_data(column, condition)
            self.data = filtered_data
            self.processor = DataProcessor(self.data)
            messagebox.showinfo("Успіх", "Дані відфільтровано!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося застосувати фільтрацію: {e}")

    def create_report(self):
        if not self.selected_columns:
            self.selected_columns = self.data.columns

        file_path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF Files", "*.pdf")])
        if not file_path:
            return

        include_graphics = messagebox.askyesno("Графіки", "Включити графіки у звіт?")

        try:
            self.processor.generate_report(file_path, list(self.selected_columns), include_graphics)
            messagebox.showinfo("Успіх", "Звіт успішно створено та збережено!")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося створити звіт: {e}")

    def load_file(self):
        # Відкриття діалогового вікна для вибору файлу
        file_path = filedialog.askopenfilename(
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if not file_path:
            return  # Якщо файл не обрано, нічого не робимо

        try:
            # Завантаження даних
            if file_path.endswith(".csv"):
                self.data = pd.read_csv(file_path)
            elif file_path.endswith(".xlsx"):
                self.data = pd.read_excel(file_path)
            else:
                messagebox.showerror("Помилка", "Непідтримуваний формат файлу!")
                return

            # Оновлення таблиці
            self.update_table()
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося завантажити файл: {e}")
    
    def save_data(self):
        """
        Збереження таблиці у файл CSV або XLSX.
        """
        if self.data is None or self.data.empty:
            messagebox.showwarning("Увага", "Дані відсутні для збереження!")
            return

        # Вибір файлу для збереження
        file_path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("Excel files", "*.xlsx")],
            title="Зберегти файл"
        )
        
        if not file_path:
            return  # Якщо користувач скасував дію

        try:
            # Збереження в залежності від формату
            if file_path.endswith(".csv"):
                self.data.to_csv(file_path, index=False)
            elif file_path.endswith(".xlsx"):
                self.data.to_excel(file_path, index=False, engine="openpyxl")
            else:
                messagebox.showwarning("Увага", "Непідтримуваний формат файлу!")
                return
            
            messagebox.showinfo("Успіх", f"Дані збережено у файл: {file_path}")
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося зберегти файл: {e}")

    def update_table(self):
        self.tree.delete(*self.tree.get_children())

        columns = list(self.data.columns)
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100, anchor="center")

        for _, row in self.data.iterrows():
            self.tree.insert("", "end", values=list(row))

        #messagebox.showinfo("Успіх", "Файл успішно завантажено!")

    def setup_edit_frame(self):
        """
        Створення фрейму для редагування елемента з окремим скролінгом.
        """
        # Основний фрейм для редагування
        self.edit_frame = tk.LabelFrame(self.root, text="Редагувати елемент", padx=10, pady=10)
        self.edit_frame.grid(row=0, column=1, rowspan=2, padx=10, pady=10, sticky="nsew")

        # Canvas для скролінгу
        self.edit_canvas = tk.Canvas(self.edit_frame, highlightthickness=0)
        self.edit_canvas.grid(row=0, column=0, sticky="nsew")

        # Scrollbar для Canvas
        self.edit_scrollbar = ttk.Scrollbar(self.edit_frame, orient="vertical", command=self.edit_canvas.yview)
        self.edit_scrollbar.grid(row=0, column=1, sticky="nsew")
        self.edit_canvas.configure(yscrollcommand=self.edit_scrollbar.set)

        # Внутрішній фрейм у Canvas
        self.scrollable_frame = ttk.Frame(self.edit_canvas)
        self.scrollable_frame_id = self.edit_canvas.create_window((0, 0), window=self.scrollable_frame, anchor="center")

        # Відстеження змін розмірів
        self.scrollable_frame.bind("<Configure>", lambda e: self.edit_canvas.configure(scrollregion=self.edit_canvas.bbox("all")))
        self.edit_canvas.bind("<Configure>", self.adjust_canvas_width)

        # Додавання кнопки поза Canvas
        self.save_button = tk.Button(self.edit_frame, text="Зберегти зміни", command=lambda: self.save_edited_item(self.selected_item))
        self.save_button.grid(row=1, column=0, columnspan=2, pady=5)

        # Налаштування адаптивності
        self.edit_frame.rowconfigure(0, weight=1)  # Canvas займає весь простір по вертикалі
        self.edit_frame.columnconfigure(0, weight=1)  # Canvas розтягується по горизонталі

    def adjust_canvas_width(self, event):
        """
        Налаштовує ширину внутрішнього фрейму відповідно до ширини Canvas.
        """
        canvas_width = event.width
        self.edit_canvas.itemconfig(self.scrollable_frame_id, width=canvas_width)

    def edit_selected_item(self, event):
        """
        Редагування обраного елемента в таблиці.
        """
        self.selected_item = self.tree.selection()  # Отримуємо ID вибраного рядка
        if not self.selected_item:
            return

        # Отримуємо дані про вибраний рядок
        values = self.tree.item(self.selected_item, "values")
        if not values:
            messagebox.showwarning("Увага", "Рядок порожній або не знайдено даних.")
            return

        # Очищення попереднього вмісту
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        # Налаштування колонок для розтягування
        self.scrollable_frame.grid_columnconfigure(0, weight=0)  # Для міток
        self.scrollable_frame.grid_columnconfigure(1, weight=1)  # Для полів вводу

        # Створення полів вводу для редагування
        self.edit_entries = []
        for i, (col_name, value) in enumerate(zip(self.tree["columns"], values)):
            ttk.Label(self.scrollable_frame, text=f"{col_name}:", style="TLabel").grid(row=i, column=0, sticky="w", pady=5, padx=5)
            entry = ttk.Entry(self.scrollable_frame, style="TEntry")
            entry.insert(0, value)
            entry.grid(row=i, column=1, pady=5, padx=5, sticky="ew")
            self.edit_entries.append(entry)

    def save_edited_item(self, item_id):
        """
        Збереження відредагованого елемента.
        """
        new_values = [entry.get() for entry in self.edit_entries]
        self.tree.item(item_id, values=new_values)
        messagebox.showinfo("Успіх", "Дані успішно оновлено!")

    def plot_selected_columns(self):
        """
        Функція для побудови графіків з підтримкою різних типів,
        автоматичним вибором графіка та попереднім переглядом.
        """
        if self.data is None or self.data.empty:
            messagebox.showwarning("Увага", "Спочатку завантажте дані!")
            return

        # Створення нового вікна для вибору параметрів графіка
        plot_window = tk.Toplevel(self.root)
        plot_window.title("Параметри графіка")
        plot_window.geometry("600x600")

        # Список стовпців
        column_names = list(self.data.columns)

        # Вибір стовпців для осі X та Y
        tk.Label(plot_window, text="Стовпець для осі X:").grid(row=0, column=0, pady=5, padx=5, sticky="w")
        x_combo = ttk.Combobox(plot_window, values=column_names, state="readonly")
        x_combo.grid(row=0, column=1, pady=5, padx=5, sticky="ew")

        tk.Label(plot_window, text="Стовпець для осі Y:").grid(row=1, column=0, pady=5, padx=5, sticky="w")
        y_combo = ttk.Combobox(plot_window, values=column_names, state="readonly")
        y_combo.grid(row=1, column=1, pady=5, padx=5, sticky="ew")

        # Вибір типу графіка
        tk.Label(plot_window, text="Тип графіка (або автоматичний вибір):").grid(row=2, column=0, pady=5, padx=5, sticky="w")
        plot_types = ["Автоматичний", "Лінійний", "Стовпчастий", "Точковий", "Гістограма", "Кругова діаграма"]
        plot_type_combo = ttk.Combobox(plot_window, values=plot_types, state="readonly")
        plot_type_combo.current(0)
        plot_type_combo.grid(row=2, column=1, pady=5, padx=5, sticky="ew")

        # Обмеження кількості даних
        tk.Label(plot_window, text="Кількість елементів для побудови (максимум):").grid(row=3, column=0, pady=5, padx=5, sticky="w")
        limit_entry = tk.Entry(plot_window)
        limit_entry.insert(0, str(len(self.data)))  # За замовчуванням — всі дані
        limit_entry.grid(row=3, column=1, pady=5, padx=5, sticky="ew")

        # Поле для попереднього перегляду
        preview_canvas = tk.Canvas(plot_window, width=500, height=400, bg="white")
        preview_canvas.grid(row=4, column=0, columnspan=2, pady=10, padx=5, sticky="nsew")

        self.plot_preview_widget = None

        preview_button = tk.Button(plot_window, text="Попередній перегляд", command=lambda: create_plot( preview_only=True))
        preview_button.grid(row=5, column=0, pady=10, padx=5, sticky="ew")

        # Кнопка для побудови графіка
        plot_button = tk.Button(plot_window, text="Побудувати графік", command=lambda: create_plot())
        plot_button.grid(row=5, column=1, pady=10, padx=5, sticky="ew")

        self.update_widget_style()

        def create_plot(preview_only=False):
            """
            Створює графік із можливістю попереднього перегляду.
            """
            # Отримання параметрів
            x_column = x_combo.get()
            y_column = y_combo.get()
            plot_type = plot_type_combo.get()
            try:
                limit = int(limit_entry.get())
            except ValueError:
                messagebox.showwarning("Увага", "Будь ласка, введіть правильне число для кількості елементів!")
                return

            if not x_column or not y_column:
                messagebox.showwarning("Увага", "Будь ласка, оберіть всі параметри!")
                return

            # Вибір даних для побудови
            data_limited = self.data.iloc[:limit]

            # Перевірка на числовий тип стовпців
            x_is_numeric = pd.api.types.is_numeric_dtype(data_limited[x_column])
            y_is_numeric = pd.api.types.is_numeric_dtype(data_limited[y_column])

            # Якщо стовпець не числовий, рахуємо кількість повторень
            if not x_is_numeric:
                # Створення словника для підрахунку повторів
                count_dict = {}
                for value in data_limited[x_column]:
                    if value not in count_dict:
                        # Якщо елемент ще не був зустрінутим, рахуємо його повторення
                        count_dict[value] = (data_limited[x_column] == value).sum()
                
                # Створюємо DataFrame з результатами
                data_limited = pd.DataFrame(list(count_dict.items()), columns=[x_column, "Кількість"])
                x_column = x_column
                y_column = "Кількість"

            if not y_is_numeric:
                count_dict = {}
                for value in data_limited[y_column]:
                    if value not in count_dict:
                        count_dict[value] = (data_limited[y_column] == value).sum()

                data_limited = pd.DataFrame(list(count_dict.items()), columns=[y_column, "Кількість"])
                x_column = y_column
                y_column = "Кількість"

            # Автоматичний вибір типу графіка
            if plot_type == "Автоматичний":
                if data_limited[x_column].nunique() < 10:  # Кругова діаграма для категорій
                    plot_type = "Кругова діаграма"
                elif pd.api.types.is_numeric_dtype(data_limited[y_column]):
                    plot_type = "Лінійний"
                else:
                    plot_type = "Стовпчастий"

            # Побудова графіка
            try:
                fig, ax = plt.subplots(figsize=(6, 4))
                if plot_type == "Лінійний":
                    ax.plot(data_limited[x_column], data_limited[y_column], marker="o")
                    ax.set_xticks(range(len(data_limited[x_column])))
                    ax.set_xticklabels(data_limited[x_column], rotation=45, ha="right")
                elif plot_type == "Стовпчастий":
                    ax.bar(data_limited[x_column], data_limited[y_column])
                    ax.set_xticks(range(len(data_limited[x_column])))
                    ax.set_xticklabels(data_limited[x_column], rotation=45, ha="right")
                elif plot_type == "Точковий":
                    ax.scatter(data_limited[x_column], data_limited[y_column])
                    ax.set_xticks(range(len(data_limited[x_column])))
                    ax.set_xticklabels(data_limited[x_column], rotation=45, ha="right")
                elif plot_type == "Гістограма":
                    ax.hist(data_limited[y_column], bins=10)
                elif plot_type == "Кругова діаграма":
                    data_grouped = data_limited.groupby(x_column)[y_column].sum()
                    ax.pie(data_grouped, labels=data_grouped.index, autopct='%1.1f%%')

                ax.set_title(f"{plot_type} графік: {y_column} vs {x_column}")
                ax.set_xlabel(x_column)
                ax.set_ylabel(y_column)

                # Попередній перегляд у Canvas
                if preview_only:
                    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
                    if self.plot_preview_widget:
                        self.plot_preview_widget.get_tk_widget().destroy()
                    self.plot_preview_widget = FigureCanvasTkAgg(fig, master=preview_canvas)
                    self.plot_preview_widget.draw()
                    self.plot_preview_widget.get_tk_widget().grid(row=0, column=0, sticky="nsew")
                    plt.close(fig)
                else:
                    plt.show()


            except Exception as e:
                messagebox.showerror("Помилка", f"Не вдалося побудувати графік: {e}")

    def apply_widget_styles(self):
        """
        Застосовує стилі до різних віджетів програми.
        """
        style = ttk.Style()

        # Загальний стиль для кнопок
        style.configure(
            "RoundedButton.TButton",
            font=("Arial", 12),
            padding=5,
            relief="flat",
            borderwidth=0,
            focuscolor="",
        )
        style.map(
            "RoundedButton.TButton",
            relief=[("pressed", "flat"), ("active", "solid")],
        )

        # Стиль для Treeview (таблиці)
        style.configure(
            "Custom.Treeview",
            rowheight=25,
            borderwidth=0,
            relief="flat",
            highlightthickness=0,
            gridlines="both",
        )
        style.configure("Custom.Treeview.Heading", font=("Arial", 12, "bold"), anchor="center")

        # Додавання сітки до таблиці
        style.map(
            "Custom.Treeview",
            background=[("selected", "#e0f0ff")],
            foreground=[("selected", "black")],
        )

        # Стиль для Scrollbar
        style.configure(
            "Vertical.TScrollbar",
            gripcount=0,
            borderwidth=1,
            troughcolor="#e0e0e0",
            background="#cccccc",
            arrowcolor="#333333",
        )
        style.configure(
            "Horizontal.TScrollbar",
            gripcount=0,
            borderwidth=1,
            troughcolor="#e0e0e0",
            background="#cccccc",
            arrowcolor="#333333",
        )

        
        def apply_styles_recursively(widget):
            """
            Рекурсивно проходить через усі дочірні елементи і застосовує стилі.
            """
            if isinstance(widget, ttk.Button):
                widget.configure(style="RoundedButton.TButton")
            elif isinstance(widget, ttk.Treeview):
                widget.configure(style="Custom.Treeview")
            elif isinstance(widget, ttk.Scrollbar):
                if widget.cget("orient") == "vertical":
                    widget.configure(style="Vertical.TScrollbar")
                elif widget.cget("orient") == "horizontal":
                    widget.configure(style="Horizontal.TScrollbar")
            elif isinstance(widget, tk.Canvas):
                widget.configure(highlightthickness=0)
            elif isinstance(widget, tk.LabelFrame):
                widget.configure(font=("Arial", 12, "bold"))
            elif isinstance(widget, tk.Toplevel):
                # Налаштування стилів для дочірніх вікон
                widget.configure(bg="#F0F0F0")  # Наприклад, світлий фон
                for child in widget.winfo_children():
                    apply_styles_recursively(child)

            # Проходимо через дочірні елементи
            for child in widget.winfo_children():
                apply_styles_recursively(child)
           
    
    def toggle_theme(self):
        # Перемикає між світлою та темною темами
        if self.is_dark_mode == False:
            self.enable_light_mode()
            self.mainmenu.entryconfig(2, label="Темний режим")
            self.is_dark_mode = True
        else:
            self.enable_dark_mode()
            self.mainmenu.entryconfig(2, label="Світлий режим")
            self.is_dark_mode = False
        return self.is_dark_mode
        #self.is_dark_mode = not self.is_dark_mode

    def enable_dark_mode(self):
        # Темні кольори
        dark_bg = "#2E2E2E"  # Темно-сірий фон
        dark_fg = "#FFFFFF"  # Білий текст
        dark_highlight = "#444444"  # Сірий для виділень

        # Оновлення стилю root і всіх дочірніх віджетів
        self.root.configure(bg=dark_bg)
        self.update_widget_style(self.root, dark_bg, dark_fg, dark_highlight)
        self.mainmenu.configure(bg=dark_bg, fg=dark_fg)
        for index in range(self.mainmenu.index("end") + 1):
            self.mainmenu.entryconfig(index, background=dark_highlight, )
        self.edit_frame.configure(bg=dark_bg, fg=dark_fg)
        for window in self.root.winfo_children():
            if isinstance(window, tk.Toplevel):
                self.update_widget_style(window, dark_bg, dark_fg, dark_highlight)

    def enable_light_mode(self):
        # Світлі кольори
        light_bg = "#F0F0F0"  # Світло-сірий фон
        light_fg = "#000000"  # Чорний текст
        light_highlight = "#E0E0E0"  # Світло-сірий для виділень

        # Оновлення стилю root і всіх дочірніх віджетів
        self.root.configure(bg=light_bg)
        self.update_widget_style(self.root, light_bg, light_fg, light_highlight)
        self.mainmenu.configure(bg=light_bg, fg=light_fg)
        for index in range(self.mainmenu.index("end") + 1):
            self.mainmenu.entryconfig(index, background=light_highlight, )
        for widget in self.root.winfo_children():
            if isinstance(widget, tk.Toplevel):
                widget.configure(bg=light_bg)  # Світлий фон
                self.update_widget_style(self.root, light_bg, light_fg, light_highlight)
        
    def update_widget_style(self, widget, bg_color, fg_color, highlight_color):
        """
        Оновлює стиль віджетів на основі заданих кольорів.
        """
        if isinstance(widget, tk.LabelFrame):
            widget.configure(bg=bg_color, fg=fg_color)
        elif isinstance(widget, tk.Frame):
            widget.configure(bg=bg_color)
        elif isinstance(widget, tk.Label):
            widget.configure(bg=bg_color, fg=fg_color)
        elif isinstance(widget, tk.Button):
            widget.configure(bg=highlight_color,
                            fg=fg_color,
                            relief="flat",
                            highlightbackground=bg_color)
        elif isinstance(widget, ttk.Scrollbar):
            style = ttk.Style()
            style.configure("TScrollbar",
                            background=highlight_color,
                            troughcolor=bg_color,
                            arrowcolor=fg_color)
        elif isinstance(widget, tk.Canvas):
            widget.configure(bg=bg_color, highlightbackground=highlight_color)
        elif isinstance(widget, ttk.Treeview):
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("Treeview",
                            background=bg_color,
                            fieldbackground=bg_color,
                            bordercolor=fg_color,
                            foreground=fg_color,
                            rowheight=25)
            style.configure("Treeview.Heading",
                            background=highlight_color,
                            fieldbackground=bg_color,
                            foreground=fg_color)
        elif isinstance(widget, ttk.Frame):
            style = ttk.Style()
            style.theme_use("clam")
            style.configure("TFrame", background=bg_color)
            widget.configure(style="TFrame")
        elif isinstance(widget, tk.Menu):
            # Застосовуємо стиль до меню
            widget.configure(bg=bg_color,
                            fg=fg_color,
                            activebackground=highlight_color,
                            activeforeground=fg_color)
        elif isinstance(widget, tk.Toplevel):
            widget.configure(bg=bg_color)
            for child in widget.winfo_children():
                self.update_widget_style(child, bg_color, fg_color, highlight_color)
        elif isinstance(widget, ttk.Entry):
            style = ttk.Style()
            style.configure("TEntry", fieldbackground=bg_color, foreground=fg_color)
        elif isinstance(widget, ttk.Label):
            style = ttk.Style()
            style.configure("TLabel", background=bg_color, foreground=fg_color)
        else:
            for child in widget.winfo_children():
                self.update_widget_style(child, bg_color, fg_color, highlight_color)

    def open_help_window(self):
        """
        Відкриває додаткове вікно з текстом із файлу help.txt.
        """
        try:
            with open("help.txt", "r", encoding="utf-8") as file:
                help_text = file.read()
        except FileNotFoundError:
            messagebox.showerror("Помилка", "Файл help.txt не знайдено!")
            return
        except Exception as e:
            messagebox.showerror("Помилка", f"Не вдалося відкрити файл: {e}")
            return

        # Створення нового вікна
        help_window = tk.Toplevel(self.root)
        help_window.title("Довідка")
        help_window.geometry("600x400")

        # Текстовий віджет для відображення вмісту
        text_widget = tk.Text(help_window, wrap="word", font=("Arial", 12))
        text_widget.insert("1.0", help_text)
        text_widget.config(state="disabled")  # Заборонити редагування тексту
        text_widget.pack(expand=True, fill="both")

        # Скролбар для тексту
        scrollbar = ttk.Scrollbar(help_window, orient="vertical", command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
                

if __name__ == "__main__":
    root = tk.Tk()
    app = DataLoaderApp(root)
    root.mainloop()
