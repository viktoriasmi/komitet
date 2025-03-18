import tkinter as tk
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
    
    def get_all_records(self, file_type):
        table = self.get_table_name(file_type)
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'SELECT * FROM {table}')
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
        
        # Экранирование кавычек в названиях колонок
        columns = [f'"{col.replace("\"", "\"\"")}"' for col in df.columns.tolist()]
        placeholders = ', '.join(['?'] * len(columns))
        
        with self.conn:
            cursor = self.conn.cursor()
            try:
                # Удаляем старые данные
                cursor.execute(f'DELETE FROM {table}')
                
                # Вставляем новые данные
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
    calculated_columns = {
        1: [
            'Контроль по дате ("-" - просрочка)',
            'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
            'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'
        ]
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
            filetypes=[("Excel files", "*.xls;*.xlsx;*.xlsm")]
        )
        if not file_path:
            return
        
        try:
            # Чтение файла
            engine = 'xlrd' if file_path.endswith(('.xls', '.xlm')) else 'openpyxl'
            df = pd.read_excel(file_path, engine=engine)

            # Приведение названий колонок к стандартному формату
            df.columns = (
                df.columns.str.strip()
                .str.replace(r'\s+', ' ', regex=True)
                .str.replace(r'[“”„"]', '', regex=True)
            )

            # Явное указание порядка колонок
            expected = self.expected_columns[self.file_type]
            df = df.reindex(columns=expected)

            # Обработка числовых полей
            numeric_columns = {
                1: [
                    'Площадь ЗУ, кв. м',
                    'Цена ЗУ по договору, руб.',
                    'Оплачено',
                    'начисленные ПЕНИ',
                    'оплачено пеней'
                ]
            }.get(self.file_type, [])
            
            for col in numeric_columns:
                if col in df.columns:
                    df[col] = (
                        df[col].astype(str)
                        .str.replace(r'\s+', '', regex=True)
                        .str.replace(',', '.')
                        .apply(pd.to_numeric, errors='coerce')
                    )

            # Обработка дат
            date_columns = {
                1: [
                    'Дата заключения договора',
                    'Срок оплаты по договору',
                    'Фактическая дата оплаты',
                    'Дата выписки учета поступлений, № ПП'
                ]
            }.get(self.file_type, [])
            
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(
                        df[col], 
                        dayfirst=True, 
                        errors='coerce'
                    ).dt.strftime('%d.%m.%Y')

            # Импорт в БД
            self.parent.db.import_from_dataframe(self.file_type, df)
            self.update_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки данных: {str(e)}")
    
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
            values = ['' if v is None else str(v) for v in record[1:]]
            tags = []
            
            # Инициализируем переменные значениями по умолчанию
            days_diff = 0
            payment_diff = 0
            unpaid_pen = 0
            
            try:
                # Обработка дат
                if values[8] and values[9]:  # Проверяем наличие дат
                    due_date = datetime.strptime(values[8], "%d.%m.%Y")
                    actual_date = datetime.strptime(values[9], "%d.%m.%Y")    
                    days_diff = (actual_date - due_date).days
                
                # Обработка числовых значений
                price = float(values[7].replace(' ', '').replace(',', '.')) if values[7] else 0.0
                paid = float(values[12].replace(' ', '').replace(',', '.')) if values[12] else 0.0
                payment_diff = price - paid
                
                penalties = float(values[15].replace(' ', '').replace(',', '.')) if values[15] else 0.0
                paid_pen = float(values[16].replace(' ', '').replace(',', '.')) if values[16] else 0.0
                unpaid_pen = penalties - paid_pen
                
                # Вставляем вычисляемые поля
                if len(values) >= 10:
                    values.insert(10, days_diff)
                if len(values) >= 14:
                    values.insert(13, payment_diff)
                if len(values) >= 18:
                    values.insert(17, unpaid_pen)
                    
            except Exception as e:
                # Добавляем недостающие значения
                while len(values) < 20:
                    values.append('')
            
            if self.file_type == 1:
                # Индексы контрольных колонок
                control_date_idx = 10
                control_payment_idx = 13
                control_peni_idx = 17
                
                try:
                    for idx in [control_date_idx, control_payment_idx, control_peni_idx]:
                        value = values[idx]
                        if isinstance(value, str):
                            value = value.strip().replace(' ', '')
                            if value.startswith('-') and value != '-':
                                tags.append('overdue')
                        elif isinstance(value, (int, float)):
                            # Изменяем условие для отрицательных значений
                            if float(value) < 0:
                                tags.append('overdue')
                except Exception as e:
                    print(f"Ошибка при проверке подсветки: {e}")
            
            self.tree.insert('', 'end', values=values, tags=(record_id, *tags), iid=str(record_id))
    
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
                    # Инвертируем разницу и добавляем минус для просрочки
                    days_diff = (due_date - actual_date).days
                    self.parent.db.update_record(1, record_id, 
                                            'Контроль по дате ("-" - просрочка)', 
                                            str(-days_diff))  # Добавляем минус перед значением

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

        if col_name in self.calculated_columns.get(self.file_type, []):
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