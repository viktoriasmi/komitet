import tkinter as tk
import os
import sqlite3
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from dateutil.parser import parse
import re
import pandas as pd

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
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute(f'''
                UPDATE {table}
                SET "{column}" = ?
                WHERE id = ?
            ''', (value, record_id))
    
    def import_from_dataframe(self, file_type, df):
        table = self.get_table_name(file_type)
        for col in df.columns:
            if pd.api.types.is_datetime64_any_dtype(df[col]):
                df[col] = df[col].dt.strftime('%d.%m.%Y')
        # Экранируем двойные кавычки внутри названий колонок
        columns = [f'"{col.replace("\"", "\"\"")}"' for col in df.columns.tolist()]
        placeholders = ', '.join(['?'] * len(columns))
        
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
    
    def __init__(self, parent, file_type): 
        super().__init__(parent)
        self.parent = parent
        self.file_type = file_type
        self.configure_ui()
        self.create_widgets()
        self.create_toolbar()
        self.update_treeview()
        self.setup_tags()
    
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
                ("Excel files", "*.xlsx"),
                ("Excel 97-2003 files", "*.xls"),
                ("All files", "*.*")
            ]
        )
        if not file_path:
            return
        
        try:
            # Убрали дублирующийся код загрузки Excel
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            elif file_path.endswith('.xlsx'):
                df = pd.read_excel(file_path, engine='openpyxl')

            # Преобразование дат из строк в datetime и обратно в строку
            date_columns = ['Дата заключения договора', 'Срок оплаты по договору', 
                            'Фактическая дата оплаты', 'Дата выписки учета поступлений, № ПП']
            for col in date_columns:
                if col in df.columns:
                    # Преобразуем в datetime с учетом формата день.месяц.год
                    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
                    # Конвертируем обратно в строку с нужным форматом
                    df[col] = df[col].dt.strftime('%d.%m.%Y')
                    
            # Дальнейшая обработка колонок (остается без изменений)
            df.columns = (
                df.columns.str.strip()
                .str.replace(r'\s+', ' ', regex=True)
                .str.replace(r'[“”„"]', '', regex=True)
                .str.replace('("-" -', '("-" -')
            )

            column_mapping = {
                'Контроль по дате (- - просрочка)': 'Контроль по дате ("-" - просрочка)',
                'Контроль по оплате цены (- переплата; + - недоплата)': 
                    'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
                'неоплаченные ПЕНИ (+ - недоплата; - переплата)': 
                    'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'
            }
            df.rename(columns=column_mapping, inplace=True)

            expected = self.expected_columns[self.file_type]
            df = df.loc[:, ~df.columns.duplicated()]
            df = df.reindex(columns=expected)

            money_columns = ['Цена ЗУ по договору, руб.', 'Оплачено', 'начисленные ПЕНИ', 'оплачено пеней']
            for col in money_columns:
                if col in df.columns:
                    df[col] = (
                        df[col].astype(str)
                        .str.replace(r'[^\d,.]', '', regex=True)
                        .str.replace(',', '.')
                    )
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                    df[col] = df[col].fillna(0.0)

            if list(df.columns) != expected:
                mismatch = list(set(expected) - set(df.columns)) + list(set(df.columns) - set(expected))
                messagebox.showerror(
                    "Ошибка", 
                    f"Несовпадение колонок:\n{', '.join(mismatch)}"
                )
                return

            self.parent.db.import_from_dataframe(self.file_type, df)
            self.update_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка обработки данных: {str(e)}")
    
    def create_new(self):
        self.update_treeview()
    
    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        records = self.parent.db.get_all_records(self.file_type)
        
        for record in records:
            values = list(record[1:])
            
            # Инициализируем переменные значениями по умолчанию
            days_diff = 0
            payment_diff = 0
            unpaid_pen = 0
            
            try:
                due_date = datetime.strptime(values[8], "%d.%m.%Y")  
                actual_date = datetime.strptime(values[9], "%d.%m.%Y")    
                days_diff = (actual_date - due_date).days
                values.insert(10, days_diff)
                
                price = float(values[7].replace(' ', '').replace(',', '.'))  
                paid = float(values[12].replace(' ', '').replace(',', '.')) 
                payment_diff = price - paid
                values.insert(13, payment_diff)
                
                penalties = float(values[15].replace(' ', '').replace(',', '.'))  
                paid_pen = float(values[16].replace(' ', '').replace(',', '.'))  
                unpaid_pen = penalties - paid_pen
                values.insert(17, unpaid_pen)
                
            except Exception as e:
                # Добавляем недостающие значения в случае ошибки
                values += [0, 0, 0] 
            
            tags = []
            if days_diff < 0:
                tags.append('overdue')
            if payment_diff != 0:
                tags.append('warning')
            
            self.tree.insert('', 'end', values=values, tags=tags)
    
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