import tkinter as tk
import os
import sqlite3
import threading
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from dateutil.parser import parse
import re
from tkinter import font
import pandas as pd
from threading import Thread

class DatabaseHandler:
    def __init__(self, db_name='registers.db'):
        self.conn = sqlite3.connect(db_name, check_same_thread=False)
        self.create_tables()
    
    def get_all_records(self, file_type):
        table = self.get_table_name(file_type)
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'SELECT * FROM {table}')
            columns = [desc[0] for desc in cursor.description]
            return columns, cursor.fetchall()
        
    def create_tables(self):
        tables = {
            'contracts': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                '"Номер договора" INTEGER',
                '"Дата заключения договора" TEXT',
                '"Покупатель, ИНН" TEXT',
                '"Кадастровый номер ЗУ, адрес ЗУ" TEXT',
                '"Площадь ЗУ, кв.м" REAL', 
                '"Разрешенное использование ЗУ" TEXT',
                '"Основание предоставления" TEXT',
                '"Цена ЗУ по договору, руб." REAL',
                '"Срок оплаты по договору" TEXT',
                '"Фактическая дата оплаты" TEXT',
                '"Контроль по дате (""-"" - просрочка)" TEXT',
                '"№ выписки учета поступлений, № ПП" TEXT',
                '"Оплачено" REAL',
                '"Контроль по оплате цены (""-"" - переплата; ""+"" - недоплата)" TEXT',
                '"примечание" TEXT',
                '"начисленные ПЕНИ" REAL',
                '"оплачено пеней" REAL',
                '"неоплаченные ПЕНИ (""+"" - недоплата; ""-"" - переплата)" TEXT',
                '"Дата выписки учета поступлений, № ПП" TEXT',
                '"Возврат имеющейся переплаты" TEXT'
            ],
        }
        
        with self.conn:
            cursor = self.conn.cursor()
            for table_name, columns in tables.items():
                columns_str = ', '.join(columns)
                cursor.execute(f'CREATE TABLE IF NOT EXISTS {table_name} ({columns_str})')
    
    def get_table_name(self, file_type):
        return {
            1: 'contracts',
            2: 'agreements',
            3: 'permits'
        }[file_type]
    
    def update_record(self, file_type, record_id, column, value):
        table = self.get_table_name(file_type)
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'''
                UPDATE {table}
                SET "{column}" = ?
                WHERE id = ?
            ''', (value, record_id))
    
    def import_from_dataframe(self, file_type, df):
        table = self.get_table_name(file_type)
        
        # Нормализация названий колонок
        df.columns = [self.normalize_column_name(col) for col in df.columns]
        
        # Приведение колонок к правильным названиям
        column_mapping = {
            'Контроль по дате (- - просрочка)': 'Контроль по дате ("-" - просрочка)',
            'Контроль по оплате цены (- переплата; + - недоплата)': 
                'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
            'неоплаченные ПЕНИ (+ - недоплата; - переплата)': 
                'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'
        }
        df.rename(columns=column_mapping, inplace=True)
        
        # Преобразование числовых полей
        money_columns = ['Цена ЗУ по договору, руб.', 'Оплачено', 
                        'начисленные ПЕНИ', 'оплачено пеней']
        for col in money_columns:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.replace(r'[^\d,.]', '', regex=True)
                    .str.replace(',', '.')
                    .astype(float)
                )
        
        # Обработка дат
        date_columns = [col for col in df.columns if 'дата' in col.lower()]
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')
        
        # Добавление недостающих колонок в таблицу
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f"PRAGMA table_info({table})")
            existing_columns = [col[1] for col in cursor.fetchall()]
            
            for col in df.columns:
                clean_col = col.split('(')[0].strip().replace('"', '')
                if clean_col not in existing_columns:
                    try:
                        cursor.execute(
                            f'ALTER TABLE {table} ADD COLUMN "{clean_col}" TEXT'
                        )
                    except sqlite3.OperationalError:
                        pass
        
        # Вставка данных
        columns = [f'"{col}"' for col in df.columns]
        placeholders = ', '.join(['?'] * len(columns))
        
        try:
            cursor.executemany(
                f'''INSERT INTO {table} ({', '.join(columns)})
                    VALUES ({placeholders})''',
                df.where(pd.notnull(df), None).values.tolist()
            )
        except sqlite3.Error as e:
            print("SQL error:", e)
            raise
    
    def export_to_dataframe(self, file_type):
        table = self.get_table_name(file_type)
        with self.conn:
            return pd.read_sql(f'SELECT * FROM {table}', self.conn)
    
    def get_paginated_records(self, file_type, offset, limit):
        table = self.get_table_name(file_type)
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'SELECT * FROM {table} LIMIT ? OFFSET ?', (limit, offset))
            return cursor.fetchall()

class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Реестры комитета")
        self.geometry("600x400")
        self.db = DatabaseHandler()
        
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Arial', 12), padding=10)
        
        self.create_widgets()

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.pack(expand=True, fill='both', padx=50, pady=50)
        
        ttk.Label(frame, text="Выберите тип реестра:", font=('Arial', 14)).pack(pady=20)
        
        btn1 = ttk.Button(frame, 
                         text="РЕЕСТР договоров купли-продажи", 
                         command=lambda: FileWindow(self, 1))
        btn1.pack(pady=10, fill='x')
        
        btn2 = ttk.Button(frame, 
                         text="РЕЕСТР соглашений о перераспр.", 
                         command=lambda: FileWindow(self, 2))
        btn2.pack(pady=10, fill='x')
        
        btn3 = ttk.Button(frame, 
                         text="РЕЕСТР разрешений на использование ЗУ", 
                         command=lambda: FileWindow(self, 3))
        btn3.pack(pady=10, fill='x')

class FileWindow(tk.Toplevel):
    expected_columns = {
        1: [
            'Номер договора', 
            'Дата заключения договора', 
            'Покупатель, ИНН', 
            'Кадастровый номер ЗУ, адрес ЗУ',
            'Площадь ЗУ, кв.м',
            'Разрешенное использование ЗУ',
            'Основание предоставления',
            'Цена ЗУ по договору, руб.',
            'Срок оплаты по договору',
            'Фактическая дата оплаты',
            'Контроль по дате ("-" - просрочка)',
            '№ выписки учета поступлений, № ПП',
            'Оплачено',
            'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
            'примечание',
            'начисленные ПЕНИ',
            'оплачено пеней',
            'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)',
            'Дата выписки учета поступлений, № ПП',
            'Возврат имеющейся переплаты'
        ],
        2: ['Номер', 'Территория', 'Стороны', 'Срок'],
        3: ['Номер', 'ЗУ', 'Заявитель', 'Период']
    }  
    
    def __init__(self, parent, file_type): 
        super().__init__(parent)
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self._is_alive = True
        self.loading_label = None
        self.column_widths = {}
        self.parent = parent
        self.file_type = file_type
        self.configure_ui()
        self.create_widgets()
        self.create_toolbar()
        self.update_treeview()
        self.setup_tags()
    
    def process_dataframe(self, df):
        # Явное преобразование числовых колонок
        money_columns = ['Цена ЗУ по договору, руб.', 'Оплачено', 
                        'начисленные ПЕНИ', 'оплачено пеней']
        for col in money_columns:
            if col in df.columns:
                df[col] = (
                    df[col].astype(str)
                    .str.replace(r'[^\d,.]', '', regex=True)
                    .str.replace(',', '.')
                    .astype(float)
                )
        
        # Преобразование дат
        date_columns = [col for col in df.columns if 'дата' in col.lower()]
        for col in date_columns:
            df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')
        
        return df

    def on_close(self):
        self._is_alive = False
        self.destroy()
    
    def configure_ui(self):
        self.title(f"Реестр типа {self.file_type}")
        self.state('zoomed')
        self.style = ttk.Style()
        
        # Настройка стилей
        self.style.configure("Treeview",
                        font=('Arial', 10),
                        rowheight=40,
                        background="#ffffff",
                        fieldbackground="#ffffff",
                        bordercolor="#d3d3d3",
                        borderwidth=1,
                        relief="solid")
        
        self.style.configure("Treeview.Heading", 
                            font=('Arial', 10, 'bold'),
                            background="#e0e0e0",
                            relief="raised")
        
        # Границы между строками
        self.style.map("Treeview",
                  background=[('selected', '#347083')],
                  foreground=[('selected', 'white')])

    def setup_tags(self):
        self.tree.tag_configure('overdue', background='#ffcccc')
        self.tree.tag_configure('warning', background='#ffffcc')

    def create_widgets(self):
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True)
        
        self.tree = ttk.Treeview(container, show='headings')
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.tree.bind('<Double-1>', self.on_double_click)
        self.tree["columns"] = self.expected_columns[self.file_type]
        for col in self.expected_columns[self.file_type]:
            self.tree.heading(col, text=col, anchor='w')
            self.tree.column(col, 
                            width=200, 
                            minwidth=100,
                            stretch=tk.YES,
                            anchor='w')
    
    # Добавляем возможность изменения размера колонок
        self.tree.bind('<ButtonRelease-1>', self.resize_columns)

    def resize_columns(self, event):
        for col in self.tree["columns"]:
            if col not in self.column_widths:
                f = font.Font()
                header_width = f.measure(col) + 20
                max_width = header_width
                
                # Ограничиваем количество проверяемых строк для производительности
                for item in self.tree.get_children()[:100]:  # Проверяем первые 100 строк
                    value = str(self.tree.set(item, col))
                    max_width = max(max_width, f.measure(value) + 20)
                
                self.column_widths[col] = max_width
                
            self.tree.column(col, width=self.column_widths[col])
            
    def validate_date(self, date_str):
        try:
            parse(date_str, dayfirst=True)
            return True
        except:
            return False

    def filter_data(self, event=None):
        query = self.search_var.get().lower()
        col = self.column_var.get()
        
        for item in self.tree.get_children():
            values = self.tree.item(item, 'values')
            col_index = self.tree["columns"].index(col)
            if query in str(values[col_index]).lower():
                self.tree.item(item, tags=('match',))
            else:
                self.tree.item(item, tags=('nomatch',))
        
        self.tree.tag_configure('match', background='')
        self.tree.tag_configure('nomatch', background='gray90')
    
    def create_toolbar(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(toolbar, text="Загрузить файл", command=self.load_file).pack(side='left', padx=2)
        ttk.Button(toolbar, text="Создать новый", command=self.create_new).pack(side='left', padx=2)
        ttk.Button(toolbar, text="Сохранить", command=self.save_file).pack(side='right', padx=2)
        
        # Поле поиска
        search_frame = ttk.Frame(toolbar)
        search_frame.pack(side='right', padx=10)
        
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=20)
        search_entry.pack(side='left')
        search_entry.bind('<KeyRelease>', self.filter_data)
        
        self.column_var = tk.StringVar()
        columns = self.expected_columns[self.file_type]
        search_combo = ttk.Combobox(search_frame, textvariable=self.column_var, 
                                   values=columns, state='readonly')
        search_combo.pack(side='left')
        search_combo.current(0)

    def load_file(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        
        def import_task():
            try:
                df = pd.read_excel(file_path, engine='openpyxl' if '.xlsx' in file_path else 'xlrd')
                
                # Применяем все преобразования из старой версии
                df = self.process_dataframe(df)
                
                self.parent.db.import_from_dataframe(self.file_type, df)
                self.after(0, self.update_treeview)
                
            except Exception as e:
                self.after(0, lambda e=e: messagebox.showerror("Ошибка", str(e)))
        
        Thread(target=import_task, daemon=True).start()
        
        def import_data():
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
                
                # Нормализация названий колонок
                df.columns = [self.normalize_column_name(col) for col in df.columns]
                
                # Обработка данных
                date_columns = [col for col in df.columns if 'дата' in col.lower()]
                for col in date_columns:
                    df[col] = pd.to_datetime(
                        df[col], 
                        dayfirst=True, 
                        errors='coerce'
                    ).dt.strftime('%d.%m.%Y')
                
                # Импорт через новый экземпляр обработчика БД
                db = DatabaseHandler()
                db.import_from_dataframe(self.file_type, df)
                
                self.after(0, self.update_treeview)
                
            except Exception as e:
                self.after(0, lambda: messagebox.showerror("Ошибка", str(e)))
        
        Thread(target=import_data, daemon=True).start()

    @staticmethod
    def normalize_column_name(name):
        return re.sub(r'\s+', ' ', name).strip()

    def create_new(self):
        self.update_treeview()
    
    def show_loading_indicator(self):
        if not self.loading_label:
            self.loading_label = ttk.Label(
                self,
                text="Загрузка данных...",
                font=('Arial', 14),
                background='#ffffff',
                relief='solid'
            )
            self.loading_label.place(relx=0.5, rely=0.5, anchor='center')
        self.loading_label.lift()

    def hide_loading_indicator(self):
        if self.loading_label:
            self.loading_label.destroy()
            self.loading_label = None

    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        records = self.parent.db.get_all_records(self.file_type)
        
        # Получаем реальные названия колонок из базы данных
        with self.parent.db.conn:
            cursor = self.parent.db.conn.cursor()
            cursor.execute(f"PRAGMA table_info({self.parent.db.get_table_name(self.file_type)})")
            db_columns = [col[1] for col in cursor.fetchall()][1:]  # исключаем id
        
        # Обновляем columns в treeview
        self.tree["columns"] = db_columns
        for col in db_columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=200, anchor='w')
        
        # Заполняем данными
        for record in records:
            values = list(record[1:])  # пропускаем первый элемент (id)
            tags = []
            
            try:
                # Проверка просрочки
                due_date = datetime.strptime(values[8], "%d.%m.%Y")
                actual_date = datetime.strptime(values[9], "%d.%m.%Y")
                if (actual_date - due_date).days < 0:
                    tags.append('overdue')
                
                # Проверка оплаты
                price = float(values[7])
                paid = float(values[12])
                if price != paid:
                    tags.append('warning')
                    
            except (ValueError, IndexError, TypeError):
                pass
            
            self.tree.insert('', 'end', values=values, tags=tags)
        
        # Автоподбор ширины колонок
        for col in db_columns:
            max_width = max(
                font.Font().measure(str(col)) + 20,  # ширина заголовка
                *[font.Font().measure(str(self.tree.set(item, col))) + 20 
                for item in self.tree.get_children()]
            )
            self.tree.column(col, width=min(max_width, 400))
    
    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        col_index = int(column[1:]) - 1
        current_value = self.tree.item(item, 'values')[col_index]
        col_name = self.expected_columns[self.file_type][col_index]
        
        edit_win = tk.Toplevel(self)
        edit_win.title("Редактирование")
        
        if col_name in ['Дата заключения', 'Срок оплаты', 'Фактическая оплата']:
            entry = ttk.Entry(edit_win, font=('Arial', 12))
            entry.pack(padx=10, pady=10)
            entry.insert(0, current_value)
            entry.bind('<FocusIn>', lambda e: self.show_calendar_dialog(entry, col_name))
        else:
            entry = ttk.Entry(edit_win, font=('Arial', 12))
            entry.pack(padx=10, pady=10)
            entry.insert(0, current_value)
        
        # Валидация для числовых полей
        if col_name in ['Цена ЗУ', 'Оплачено', 'Начисленные пени', 'Оплачено пеней']:
            validate_cmd = (edit_win.register(self.validate_number), '%P')
            entry.configure(validate='key', validatecommand=validate_cmd)
        
        def save_edit():
            new_value = entry.get()
            record_id = self.tree.item(item, 'values')[0]
            try:
                self.parent.db.update_record(
                    file_type=self.file_type,
                    record_id=record_id,
                    column=col_name,
                    value=new_value
                )
                self.update_treeview()
                edit_win.destroy()
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
        
        ttk.Button(edit_win, text="Сохранить", command=save_edit).pack(pady=5)
        
    def validate_number(self, value):
        return value == "" or value.isdigit()

    def show_calendar_dialog(self, parent_entry, field_name):
        cal = simpledialog.askstring("Ввод даты", 
                                   f"Введите дату для {field_name} (ДД.ММ.ГГГГ):")
        if cal and self.validate_date(cal):
            parent_entry.delete(0, tk.END)
            parent_entry.insert(0, cal)
    
    def save_file(self):
        try:
            df = self.parent.db.export_to_dataframe(self.file_type)
            df = df.drop(columns=['id'])
            
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            
            if not file_path:
                return
            
            if os.path.exists(file_path):
                backup_name = f"{file_path}_backup_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
                os.rename(file_path, backup_name)
            
            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Успех", "Файл успешно сохранен!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()
