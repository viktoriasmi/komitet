import tkinter as tk
import traceback
import sys
import math
import os
import sqlite3
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from dateutil.parser import parse
import re
import pandas as pd
from tkcalendar import Calendar
from datetime import timedelta  

class DatabaseHandler:
    def __init__(self, db_name='registers.db'):
        self.conn = sqlite3.connect(db_name)
        self.calculated_columns = {
        1: [
            'Контроль по дате ("-" - просрочка)',
            'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
            'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'
        ]
        }
        self.create_tables()
        
    def create_tables(self):
        tables = {
            'contracts': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                '"Номер договора" INTEGER',
                '"Дата заключения договора" TEXT',
                '"Покупатель, ИНН" TEXT',
                '"Кадастровый номер ЗУ, адрес ЗУ" TEXT',
                '"Площадь ЗУ, кв. м" REAL',
                '"Разрешенное использование ЗУ" TEXT',
                '"Основание предоставления" TEXT',
                '"Цена ЗУ по договору, руб." REAL',
                '"Срок оплаты по договору" TEXT',
                '"Фактическая дата оплаты" TEXT',
                '"№ выписки учета поступлений, № ПП" TEXT',
                '"Оплачено" REAL',
                '"примечание" TEXT',
                '"начисленные ПЕНИ" REAL',
                '"оплачено пеней" REAL',
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
    
    def get_all_records(self, file_type, columns=None):
        table = self.get_table_name(file_type)
        if columns:
            columns_str = ', '.join([f'"{col}"' for col in columns])
        else:
            columns_str = '*'
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'SELECT {columns_str} FROM {table}')
            return cursor.fetchall()
    
    def update_record(self, file_type, record_id, column, value):
        table = self.get_table_name(file_type)
        # Экранируем двойные кавычки в названии колонки
        column_escaped = column.replace('"', '""')
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'''
                UPDATE {table}
                SET "{column_escaped}" = ?
                WHERE id = ?
            ''', (value, record_id))
    
    def import_from_dataframe(self, file_type, df):
        table = self.get_table_name(file_type)
        if file_type == 1:
            calculated = self.calculated_columns[file_type]
            df = df.drop(columns=[c for c in calculated if c in df.columns])
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d.%m.%Y')
        # Экранируем двойные кавычки внутри названий колонок
        columns = [f'"{col.replace("\"", "\"\"")}"' for col in df.columns.tolist()]
        placeholders = ', '.join(['?'] * len(columns))
        df = df.where(pd.notnull(df), None) 
        with self.conn:
            cursor = self.conn.cursor()
            try:
                cursor.executemany(
                    f'''INSERT INTO {table} ({', '.join(columns)})
                        VALUES ({placeholders})''',
                    df.values.tolist()
                )
            except sqlite3.Error as e:
                print("SQL error:", e)
                raise
    
    def export_to_dataframe(self, file_type):
        table = self.get_table_name(file_type)
        with self.conn:
            return pd.read_sql(f'SELECT * FROM {table}', self.conn)

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
            'Площадь ЗУ, кв. м',
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
    date_columns = {
        1: [
            'Дата заключения договора',
            'Срок оплаты по договору',
            'Фактическая дата оплаты',
            'Дата выписки учета поступлений, № ПП'
        ],
        2: [],
        3: []
    }
    
    def __init__(self, parent, file_type): 
        super().__init__(parent)
        self.parent = parent
        self.transient(parent)  
        self.grab_set()  
        self.file_type = file_type
        self.configure_ui()
        self.create_widgets()
        self.create_toolbar()
        self.update_treeview()
        self.setup_tags()
        self.state('zoomed')  
    
    def calculate_days_diff(self, row):
        try:
            due_date = datetime.strptime(row['Срок оплаты по договору'], "%d.%m.%Y")
            actual_date = datetime.strptime(row['Фактическая дата оплаты'], "%d.%m.%Y")
            return (actual_date - due_date).days
        except:
            return 0

    def calculate_payment_diff(self, row):
        try:
            return float(row['Цена ЗУ по договору, руб.']) - float(row['Оплачено'])
        except:
            return 0

    def calculate_peni_diff(self, row):
        try:
            return float(row['начисленные ПЕНИ']) - float(row['оплачено пеней'])
        except:
            return 0

    def configure_ui(self):
        self.title(f"Реестр типа {self.file_type}")
        self.geometry("1400x800")
        self.style = ttk.Style()
        self.style.configure("Red.Treeview", background="#ffcccc")
        self.style.configure("Yellow.Treeview", background="#ffffcc")

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
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor='center')

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
            parent=self,
            title="Выберите файл",
            filetypes=[
                ("Excel files", "*.xls;*.xlsx;*.xlsm"),  
                ("All files", "*.*")
            ]
        )
        if not file_path:
            return
        
        try:
            # Указываем явные типы для текстовых колонок
            dtype_spec = {
                '№ выписки учета поступлений, № ПП': str,
                'примечание': str,
                'Кадастровый номер ЗУ, адрес ЗУ': str,
                'Покупатель, ИНН': str,
                'Основание предоставления': str,
                'Разрешенное использование ЗУ': str,
                'Номер договора': str
            }

            # Чтение с указанием формата дат и типов
            if file_path.endswith(('.xls', '.xlsx')):
                df = pd.read_excel(
                    file_path, 
                    engine='openpyxl',
                    dtype=dtype_spec,
                    date_parser=lambda x: pd.to_datetime(x, dayfirst=True, errors='coerce')
                )
            else:
                df = pd.read_excel(
                    file_path, 
                    engine='xlrd',
                    dtype=dtype_spec,
                    date_parser=lambda x: pd.to_datetime(x, dayfirst=True, errors='coerce')
                )

            # Нормализация названий колонок
            df.columns = (
                df.columns.str.strip()
                .str.normalize('NFKC')
                .str.replace(r'\s+', ' ', regex=True)
                .str.replace(r'["\']', '', regex=True)  # Удаляем кавычки
            )

            # Точное сопоставление колонок
            column_mapping = {
                'кадастровый номер зу адрес зу': 'Кадастровый номер ЗУ, адрес ЗУ',
                'площадь зу кв м': 'Площадь ЗУ, кв. м',
                'пп6 п2 ст 393 ст 3917 ст 3920 зк рф': 'Основание предоставления',
                'контроль по дате - просрочка': 'Контроль по дате ("-" - просрочка)',
                'оплачено руб': 'Оплачено',
                'начисленные пени': 'начисленные ПЕНИ',
                'оплачено пеней': 'оплачено пеней'
            }
            df.rename(columns=column_mapping, inplace=True)

            # Добавляем недостающие колонки
            for col in self.expected_columns[self.file_type]:
                if col not in df.columns:
                    df[col] = None

            # Упорядочиваем колонки
            df = df[self.expected_columns[self.file_type]]

            # Обработка числовых колонок
            num_cols = ['Площадь ЗУ, кв. м', 'Цена ЗУ по договору, руб.', 
                    'Оплачено', 'начисленные ПЕНИ', 'оплачено пеней']
            for col in num_cols:
                if col in df.columns:
                    # Удаляем пробелы и заменяем запятые
                    df[col] = df[col].astype(str).str.replace('\s+', '', regex=True)
                    df[col] = df[col].str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce')

            # Обработка дат
            date_cols = ['Дата заключения договора', 'Срок оплаты по договору',
                        'Фактическая дата оплаты', 'Дата выписки учета поступлений, № ПП']
            for col in date_cols:
                if col in df.columns:
                    # Конвертируем уже распарсенные даты в строковый формат
                    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')

            # Текстовые колонки - очистка и преобразование
            text_cols = ['примечание', '№ выписки учета поступлений, № ПП']
            for col in text_cols:
                if col in df.columns:
                    # Преобразуем числа в строки без .0
                    df[col] = df[col].astype(str).str.replace(r'\.0$', '', regex=True)
                    df[col] = df[col].replace('nan', '')

            # Замена NaN на None
            df = df.where(pd.notnull(df), None)

            # Импорт в базу
            self.parent.db.import_from_dataframe(self.file_type, df)
            self.update_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка загрузки файла: {str(e)}\n\nТрассировка:\n{traceback.format_exc()}")
    
    def create_new(self):
        try:
            expected = self.expected_columns[self.file_type]
            default_data = {}
            
            # Создаем данные по умолчанию в соответствии с типом реестра
            if self.file_type == 1:
                default_data = {
                    'Номер договора': '',
                    'Дата заключения договора': datetime.now().strftime('%d.%m.%Y'),
                    'Покупатель, ИНН': 'ООО "Пример", ИНН 0000000000',
                    'Кадастровый номер ЗУ, адрес ЗУ': '00:00:0000000:00, г. Пример',
                    'Площадь ЗУ, кв. м': '',
                    'Разрешенное использование ЗУ': '',
                    'Основание предоставления': 'Номер ЗК РФ',
                    'Цена ЗУ по договору, руб.': '0.00',
                    'Срок оплаты по договору': datetime.now().strftime('%d.%m.%Y'),
                    'Фактическая дата оплаты': '',
                    '№ выписки учета поступлений, № ПП': '',
                    'Оплачено': '0,00',
                    'примечание': '',
                    'начисленные ПЕНИ': '0,00',
                    'оплачено пеней': '0,00',
                    'Дата выписки учета поступлений, № ПП': '',
                    'Возврат имеющейся переплаты': ''
                }
            else:
                default_data = {col: '' for col in expected}
            
            # Создаем DataFrame с одной строкой
            df = pd.DataFrame([default_data], columns=expected)
            
            # Применяем преобразования как при импорте из Excel
            # Обработка числовых полей
            numeric_columns = {
                1: ['Цена ЗУ по договору, руб.', 'Оплачено', 
                    'начисленные ПЕНИ', 'оплачено пеней'],
                2: [],
                3: []
            }.get(self.file_type, [])
            
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(r'[^\d,.]', '', regex=True)
                    df[col] = df[col].str.replace(',', '.').astype(float)
            
            # Обработка дат
            date_columns = {
                1: ['Дата заключения договора', 'Срок оплаты по договору', 
                    'Фактическая дата оплаты', 'Дата выписки учета поступлений, № ПП'],
                2: [],
                3: []
            }.get(self.file_type, [])
            
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')
            
            # Добавляем через существующий импортный метод
            self.parent.db.import_from_dataframe(self.file_type, df)
            self.update_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при создании новой записи: {str(e)}")
    
    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        records = self.parent.db.get_all_records(self.file_type)
        
        for record in records:
            record_id = record[0]
            record_dict = dict(zip(self.expected_columns[self.file_type], record[1:]))
            tags = []
            calculated_values = {}

            try:
                if self.file_type == 1:
                    # Получаем значения из словаря
                    due_date_str = record_dict.get('Срок оплаты по договору', '')
                    actual_date_str = record_dict.get('Фактическая дата оплаты', '')
                    
                    # Парсим даты
                    days_diff = 0
                    if due_date_str and actual_date_str:
                        due_date = datetime.strptime(due_date_str, "%d.%m.%Y")
                        actual_date = datetime.strptime(actual_date_str, "%d.%m.%Y")
                        days_diff = (actual_date - due_date).days

                    # Получаем и преобразуем числовые значения
                    price = float(record_dict.get('Цена ЗУ по договору, руб.', 0))
                    paid = float(record_dict.get('Оплачено', 0))
                    accrued = float(record_dict.get('начисленные ПЕНИ', 0))
                    paid_pen = float(record_dict.get('оплачено пеней', 0))

                    # Расчет значений
                    payment_diff = price - paid
                    unpaid_pen = accrued - paid_pen

                    calculated_values = {
                        'Контроль по дате ("-" - просрочка)': days_diff,
                        'Контроль по оплате цены ("-" - переплата; "+" - недоплата)': payment_diff,
                        'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)': unpaid_pen
                    }

                    # Проверка условий подсветки
                    if days_diff < 0:
                        tags.append('overdue')
                    if payment_diff != 0:
                        tags.append('warning')

            except Exception as e:
                print(f"Ошибка расчетов: {e}")

            # Формируем финальные значения
            final_values = []
            for col in self.expected_columns[self.file_type]:
                if col in calculated_values:
                    # Форматируем числовые значения
                    value = calculated_values[col]
                    if isinstance(value, float):
                        final_values.append(f"{value:.2f}".replace('.', ','))
                    else:
                        final_values.append(str(value))
                else:
                    # Берем значение из записи и обрабатываем пустые значения
                    value = record_dict.get(col, '')
                    if value is None:
                        final_values.append('')
                    else:
                        # Для числовых полей добавляем форматирование
                        if col in ['Цена ЗУ по договору, руб.', 'Оплачено', 
                                'начисленные ПЕНИ', 'оплачено пеней']:
                            try:
                                final_values.append(f"{float(value):,.2f}".replace(',', ' ').replace('.', ','))
                            except:
                                final_values.append(str(value))
                        else:
                            final_values.append(str(value))
            
            self.tree.insert('', 'end', values=final_values, tags=tags)
    

    def create_calendar(self, parent, entry, col_name):
        cal_win = tk.Toplevel(parent)
        cal_win.title("Выбор даты")
        
        cal = Calendar(cal_win, 
                    selectmode='day', 
                    date_pattern='dd.mm.yyyy',
                    locale='ru_RU')
        cal.pack(padx=10, pady=10)

        def set_date():
            selected_date = cal.get_date()
            if self.validate_date(selected_date):
                entry.delete(0, tk.END)
                entry.insert(0, selected_date)
                cal_win.destroy()
            else:
                messagebox.showerror("Ошибка", "Некорректный формат даты")

        ttk.Button(cal_win, 
                text="Сохранить", 
                command=set_date
                ).pack(pady=5)


    def save_edit(self, new_value, record_id, col_name, edit_win):
        try:
            self.parent.db.update_record(self.file_type, record_id, col_name, new_value)
            
            if self.file_type == 1:
                self.update_calculations(record_id, col_name, new_value)
            
            self.update_treeview()
            edit_win.destroy()
            messagebox.showinfo("Успех", "Изменения сохранены!")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def update_calculations(self, record_id, edited_col, new_value):
        cursor = self.parent.db.conn.cursor()
        cursor.execute('SELECT * FROM contracts WHERE id = ?', (record_id,))
        record = cursor.fetchone()

        try:
            # Обновление срока оплаты
            if edited_col == 'Дата заключения договора':
                contract_date = datetime.strptime(new_value, '%d.%m.%Y')
                due_date = (contract_date + timedelta(days=7)).strftime('%d.%m.%Y')
                self.parent.db.update_record(1, record_id, 'Срок оплаты по договору', due_date)
                edited_col = 'Срок оплаты по договору'
                new_value = due_date

            # Расчет контроля по дате
            if edited_col in ['Срок оплаты по договору', 'Фактическая дата оплаты']:
                due_date_str = record[8] if edited_col != 'Срок оплаты по договору' else new_value
                actual_date_str = record[9] if edited_col != 'Фактическая дата оплаты' else new_value
                
                if due_date_str and actual_date_str:
                    due_date = datetime.strptime(due_date_str, '%d.%m.%Y')
                    actual_date = datetime.strptime(actual_date_str, '%d.%m.%Y')
                    days_diff = (actual_date - due_date).days  # Правильный порядок вычитания
                    self.parent.db.update_record(1, record_id, 
                                                'Контроль по дате ("-" - просрочка)', 
                                                str(days_diff))  # Убираем инверсию знака
            # Расчет контроля по оплате
            if edited_col in ['Цена ЗУ по договору, руб.', 'Оплачено']:
                price = float(record[7]) if edited_col != 'Цена ЗУ по договору, руб.' else float(new_value)
                paid = float(record[12]) if edited_col != 'Оплачено' else float(new_value)
                payment_diff = price - paid
                self.parent.db.update_record(1, record_id, 
                    'Контроль по оплате цены ("-" - переплата; "+" - недоплата)', f"{payment_diff:.2f}")

            # Расчет неоплаченных пени
            if edited_col in ['начисленные ПЕНИ', 'оплачено пеней']:
                accrued = float(record[15]) if edited_col != 'начисленные ПЕНИ' else float(new_value)
                paid_pen = float(record[16]) if edited_col != 'оплачено пеней' else float(new_value)
                unpaid_pen = math.floor(accrued - paid_pen)
                self.parent.db.update_record(1, record_id, 
                    'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)', str(unpaid_pen))

        except Exception as e:
            print(f"Ошибка в расчетах: {e}")

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        item = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        
        col_index = int(column[1:]) - 1
        current_value = self.tree.item(item, 'values')[col_index]
        col_name = self.expected_columns[self.file_type][col_index]

        if col_name in self.parent.db.calculated_columns.get(self.file_type, []):
            messagebox.showinfo("Информация", "Это поле рассчитывается автоматически и не может быть изменено вручную.")
            return

        edit_win = tk.Toplevel(self)
        edit_win.title("Редактирование")
        edit_win.geometry("400x150")
        
        record_id = self.tree.item(item, 'tags')[0] if self.tree.item(item, 'tags') else self.tree.item(item, 'iid')
        entry = ttk.Entry(edit_win, font=('Arial', 12), width=30)
        entry.pack(padx=20, pady=20, fill='x', expand=True)
        entry.insert(0, current_value)

        if col_name in self.date_columns.get(self.file_type, []):
            self.create_calendar(edit_win, entry, col_name)

        btn_frame = ttk.Frame(edit_win)
        btn_frame.pack(fill='x', padx=20, pady=10)

        ttk.Button(btn_frame, 
                 text="Сохранить", 
                 command=lambda: self.save_edit(entry.get(), record_id, col_name, edit_win)
                 ).pack(side='right')
                 
        ttk.Button(btn_frame, 
                 text="Отмена", 
                 command=edit_win.destroy
                 ).pack(side='right', padx=5)

        # Добавляем валидацию для числовых полей
        if col_name in ['Цена ЗУ по договору, руб.', 'Оплачено', 'начисленные ПЕНИ', 'оплачено пеней']:
            validate_cmd = (edit_win.register(self.validate_number), '%P')
            entry.configure(validate='key', validatecommand=validate_cmd)
        
    def validate_number(self, value):
        try:
            if value.strip() == "": return True
            float(value.replace(',', '.'))
            return True
        except:
            return False

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

            if self.file_type == 1:
                # Форматирование числовых колонок
                num_cols = ['Цена ЗУ по договору, руб.', 'Оплачено', 
                        'начисленные ПЕНИ', 'оплачено пеней']
                for col in num_cols:
                    df[col] = df[col].apply(lambda x: f"{x:,.2f}".replace(',', ' ').replace('.', ',') if pd.notnull(x) else '')
                
                # Форматирование дат
                date_cols = ['Дата заключения договора', 'Срок оплаты по договору',
                            'Фактическая дата оплаты', 'Дата выписки учета поступлений, № ПП']
                for col in date_cols:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%d.%m.%Y')    
                # Применяем расчеты
                df['Контроль по дате ("-" - просрочка)'] = df.apply(self.calculate_days_diff, axis=1)
                df['Контроль по оплате цены ("-" - переплата; "+" - недоплата)'] = df.apply(self.calculate_payment_diff, axis=1)
                df['неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'] = df.apply(self.calculate_peni_diff, axis=1)

            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Успех", "Файл успешно сохранен!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()