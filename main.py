import tkinter as tk
import os
import sqlite3
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import pandas as pd

class DatabaseHandler:
    def __init__(self, db_name='registers.db'):
        self.conn = sqlite3.connect(db_name)
        self.create_tables()
        
    def create_tables(self):
        tables = {
            'contracts': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                'Номер INTEGER',
                'Участок TEXT',
                'Собственник TEXT',
                'Дата TEXT'
            ],
            'agreements': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                'Номер INTEGER',
                'Территория TEXT',
                'Стороны TEXT',
                'Срок TEXT'
            ],
            'permits': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                'Номер INTEGER',
                'ЗУ TEXT',
                'Заявитель TEXT',
                'Период TEXT'
            ]
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
                SET {column} = ?
                WHERE id = ?
            ''', (value, record_id))
    
    def import_from_dataframe(self, file_type, df):
        table = self.get_table_name(file_type)
        columns = df.columns.tolist()
        placeholders = ', '.join(['?'] * len(columns))
        
        with self.conn:
            cursor = self.conn.cursor()
            cursor.executemany(f'''
                INSERT INTO {table} ({', '.join(columns)})
                VALUES ({placeholders})
            ''', df.values.tolist())
    
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
    expected_columns = {  # Исправлены названия столбцов на русские
        1: ['Номер', 'Участок', 'Собственник', 'Дата'],
        2: ['Номер', 'Территория', 'Стороны', 'Срок'],
        3: ['Номер', 'ЗУ', 'Заявитель', 'Период']
    }
    
    def __init__(self, parent, file_type):
        super().__init__(parent)
        self.parent = parent
        self.file_type = file_type
        self.title(f"Реестр типа {file_type}")
        self.geometry("1000x700")
        self.tree = None  # Важно инициализировать атрибут
        
        # Сначала создаем виджеты, потом обновляем данные
        self.create_widgets()
        self.create_toolbar()
        self.update_treeview()

    # Добавить метод create_widgets из старой версии
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

    # Добавить метод filter_data из старой версии
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
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        
        try:
            df = pd.read_excel(file_path)
            if list(df.columns) != self.expected_columns[self.file_type]:
                messagebox.showerror("Ошибка", "Неверная структура файла!")
                return
            
            # Импортируем данные в БД
            self.parent.db.import_from_dataframe(self.file_type, df)
            self.update_treeview()
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
    
    def create_new(self):
        # Создаем пустую таблицу в Treeview
        self.update_treeview()
    
    def update_treeview(self):
        self.tree.delete(*self.tree.get_children())
        
        # Получаем данные из БД
        records = self.parent.db.get_all_records(self.file_type)
        columns = self.expected_columns[self.file_type]
        
        self.tree["columns"] = columns
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=100)

        for record in records:
            # Пропускаем ID в первой позиции
            self.tree.insert('', 'end', values=record[1:])
    
    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        col_index = int(column[1:]) - 1
        current_value = self.tree.item(item, 'values')[col_index]
        
        edit_win = tk.Toplevel(self)
        edit_win.title("Редактирование")
        
        entry = ttk.Entry(edit_win, font=('Arial', 12))
        entry.pack(padx=10, pady=10)
        entry.insert(0, current_value)
        entry.focus_set()
        
        def save_edit():
            new_value = entry.get()
            col_name = self.expected_columns[self.file_type][col_index]
            
            # Получаем ID записи из первого столбца
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
    
    def save_file(self):
        try:
            # Экспортируем данные из БД в DataFrame
            df = self.parent.db.export_to_dataframe(self.file_type)
            # Удаляем столбец id
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
            
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Успех", "Файл успешно сохранен!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()