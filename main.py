import tkinter as tk
import traceback
import sys
import math
import os
import sqlite3
import hashlib
import shutil
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from dateutil.parser import parse
import re
import pandas as pd
from tkcalendar import Calendar
from datetime import timedelta  
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape
from datetime import datetime

class CustomMessageBox(tk.Toplevel):
    def __init__(self, parent, title, message, icon_path='icon.ico'):
        super().__init__(parent)
        self.title(title)
        try:
            self.iconbitmap(icon_path)
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        
        self.geometry("300x100")
        self.resizable(False, False)
        
        ttk.Label(self, text=message).pack(pady=10)
        ttk.Button(self, text="OK", command=self.destroy).pack(pady=5)
        self.grab_set()

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
    
    def get_user_fio(self, user_id):
        with self.conn:
            cursor = self.conn.cursor()
            cursor.execute("SELECT fio FROM users WHERE id = ?", (user_id,))
            result = cursor.fetchone()
            return result[0] if result else "Неизвестный"

    def create_tables(self):
        tables = {
                'users': [
                    'id INTEGER PRIMARY KEY AUTOINCREMENT',
                    'fio TEXT NOT NULL',
                    'login TEXT UNIQUE NOT NULL',
                    'password TEXT NOT NULL',
                    'is_admin BOOLEAN NOT NULL DEFAULT 0',
                    'can_edit_1 BOOLEAN NOT NULL DEFAULT 0',
                    'can_edit_2 BOOLEAN NOT NULL DEFAULT 0',
                    'login_attempts INTEGER NOT NULL DEFAULT 0',  
                    'is_locked BOOLEAN NOT NULL DEFAULT 0'
                ],
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
                '"Возврат имеющейся переплаты" TEXT',
                '"Контроль по дате (""-"" - просрочка)" REAL',
                '"Контроль по оплате цены (""-"" - переплата; ""+"" - недоплата)" REAL',
                '"неоплаченные ПЕНИ (""+"" - недоплата; ""-"" - переплата)" REAL',
                '"Редактор" INTEGER REFERENCES users(id)'
            ],
            'agreements': [
                'id INTEGER PRIMARY KEY AUTOINCREMENT',
                '"№ соглашения" TEXT',
                '"Дата заключения" TEXT',
                '"Собственник, ИНН" TEXT',
                '"Кадастровый номер образуемого ЗУ, адрес ЗУ" TEXT',
                '"Площадь образуемого ЗУ, кв. м" REAL',
                '"реквизиты приказа ГК ПО по им. Отнош." TEXT',
                '"Размер платы за увеличение площади ЗУ, руб." REAL',
                '"Срок оплаты" TEXT',
                '"Фактическая дата оплаты" TEXT',
                '"Контроль по дате (""-"" - просрочка)" REAL',
                '"№ выписки учета поступлений, № ПП" TEXT',  # Один экземпляр
                '"Оплачено" REAL',
                '"Контроль по оплате цены (""-"" - переплата; ""+"" - недоплата)" REAL',
                '"примечание" TEXT',
                '"начисленные ПЕНИ" REAL',
                '"оплачено пеней" REAL',
                '"неоплаченные ПЕНИ (""+"" - недоплата; ""-"" - переплата)" REAL',
                '"Возврат имеющейся переплаты" TEXT',
                '"Редактор" INTEGER REFERENCES users(id)'
            ]
        }
        triggers = [
            '''CREATE TRIGGER IF NOT EXISTS update_due_date AFTER UPDATE OF "Дата заключения договора" ON contracts
            BEGIN
                UPDATE contracts SET
                    "Срок оплаты по договору" = 
                        strftime('%d.%m.%Y', 
                            date(
                                substr(NEW."Дата заключения договора", 7, 4) || '-' ||
                                substr(NEW."Дата заключения договора", 4, 2) || '-' ||
                                substr(NEW."Дата заключения договора", 1, 2),
                                '+7 days'
                            )
                        )
                WHERE id = NEW.id;
            END;''',
            
            '''CREATE TRIGGER IF NOT EXISTS update_control_date AFTER UPDATE ON contracts
            BEGIN
                UPDATE contracts SET 
                    "Контроль по дате (""-"" - просрочка)" = 
                        julianday(
                            substr(NEW."Срок оплаты по договору", 7, 4) || '-' ||
                            substr(NEW."Срок оплаты по договору", 4, 2) || '-' ||
                            substr(NEW."Срок оплаты по договору", 1, 2)
                        ) - julianday(
                            substr(NEW."Фактическая дата оплаты", 7, 4) || '-' ||
                            substr(NEW."Фактическая дата оплаты", 4, 2) || '-' ||
                            substr(NEW."Фактическая дата оплаты", 1, 2)
                        ),
                    "Контроль по оплате цены (""-"" - переплата; ""+"" - недоплата)" = 
                        NEW."Цена ЗУ по договору, руб." - NEW."Оплачено",
                    "неоплаченные ПЕНИ (""+"" - недоплата; ""-"" - переплата)" = 
                        NEW."начисленные ПЕНИ" - NEW."оплачено пеней"
                WHERE id = NEW.id;
            END;''',
            
            '''CREATE TRIGGER IF NOT EXISTS update_agreement_due_date AFTER UPDATE OF "Дата заключения" ON agreements
            BEGIN
                UPDATE agreements SET
                    "Срок оплаты" = 
                        strftime('%d.%m.%Y', 
                            date(
                                substr(NEW."Дата заключения", 7, 4) || '-' ||
                                substr(NEW."Дата заключения", 4, 2) || '-' ||
                                substr(NEW."Дата заключения", 1, 2),
                                '+7 days'
                            )
                        )
                WHERE id = NEW.id;
            END;''',

            '''CREATE TRIGGER IF NOT EXISTS update_agreement_control AFTER UPDATE ON agreements
            BEGIN
                UPDATE agreements SET 
                    "Контроль по дате (""-"" - просрочка)" = 
                        julianday(
                            substr(NEW."Срок оплаты", 7, 4) || '-' ||
                            substr(NEW."Срок оплаты", 4, 2) || '-' ||
                            substr(NEW."Срок оплаты", 1, 2)
                        ) - julianday(
                            substr(NEW."Фактическая дата оплаты", 7, 4) || '-' ||
                            substr(NEW."Фактическая дата оплаты", 4, 2) || '-' ||
                            substr(NEW."Фактическая дата оплаты", 1, 2)
                        ),
                    "Контроль по оплате цены (""-"" - переплата; ""+"" - недоплата)" = 
                        NEW."Размер платы за увеличение площади ЗУ, руб." - NEW."Оплачено",
                    "неоплаченные ПЕНИ (""+"" - недоплата; ""-"" - переплата)" = 
                        NEW."начисленные ПЕНИ" - NEW."оплачено пеней"
                WHERE id = NEW.id;
            END;''',

            '''CREATE TRIGGER IF NOT EXISTS update_editor AFTER UPDATE ON contracts
            BEGIN
                UPDATE contracts SET "Редактор" = NEW."Редактор" WHERE id = NEW.id;
            END;'''
        ]
        
        with self.conn:
            cursor = self.conn.cursor()
            for table_name, columns in tables.items():
                columns_str = ', '.join(columns)
                cursor.execute(f'CREATE TABLE IF NOT EXISTS {table_name} ({columns_str})')
            for trigger in triggers:
                cursor.execute(trigger)
            cursor.execute("PRAGMA table_info(users)")
            columns = [column[1] for column in cursor.fetchall()]
            if 'login_attempts' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN login_attempts INTEGER NOT NULL DEFAULT 0")
            if 'is_locked' not in columns:
                cursor.execute("ALTER TABLE users ADD COLUMN is_locked BOOLEAN NOT NULL DEFAULT 0")

    def get_table_name(self, file_type):
        return {
            1: 'contracts',
            2: 'agreements'
        }[file_type]
    
    def get_all_records(self, file_type, columns=None):
        table = self.get_table_name(file_type)
        # Добавляем столбец "Редактор" в выборку
        if columns is None:
            columns_str = '*'
        else:
            columns_str = ', '.join([f'"{col}"' for col in columns]) + ', "Редактор"'
        
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

class AuthWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Аутентификация")
        self.geometry("300x200")
        try:
            self.iconbitmap('icon.ico')  # Добавлено
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.create_widgets()
        self.grab_set()

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.pack(expand=True, padx=20, pady=20)

        ttk.Button(frame, text="Вход", command=self.open_login).pack(pady=10, fill='x')
        ttk.Button(frame, text="Регистрация", command=self.open_register).pack(pady=10, fill='x')

    def open_login(self):
        LoginDialog(self)

    def open_register(self):
        RegisterDialog(self)

class RegisterDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Регистрация")
        self.geometry("300x250")
        try:
            self.iconbitmap('icon.ico')  
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.create_widgets()
        self.grab_set()

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill='both', expand=True)

        ttk.Label(frame, text="ФИО:").grid(row=0, column=0, sticky='w', pady=2)
        self.fio_entry = ttk.Entry(frame)
        self.fio_entry.grid(row=0, column=1, pady=2)

        ttk.Label(frame, text="Логин:").grid(row=1, column=0, sticky='w', pady=2)
        self.login_entry = ttk.Entry(frame)
        self.login_entry.grid(row=1, column=1, pady=2)

        ttk.Label(frame, text="Пароль:").grid(row=2, column=0, sticky='w', pady=2)
        self.password_entry = ttk.Entry(frame, show="*")
        self.password_entry.grid(row=2, column=1, pady=2)

        ttk.Label(frame, text="Подтвердите пароль:").grid(row=3, column=0, sticky='w', pady=2)
        self.confirm_entry = ttk.Entry(frame, show="*")
        self.confirm_entry.grid(row=3, column=1, pady=2)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=4, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Зарегистрироваться", command=self.register).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Отмена", command=self.destroy).pack(side='left', padx=5)

    def register(self):
        fio = self.fio_entry.get()
        login = self.login_entry.get().strip()
        password = self.password_entry.get().strip()
        confirm = self.confirm_entry.get().strip()

        if not all([fio, login, password, confirm]):
            messagebox.showerror("Ошибка", "Все поля обязательны для заполнения")
            return

        if password != confirm:
            CustomMessageBox(self, "Ошибка", "Пароли не совпадают").wait_window()
            return

        # Проверка сложности пароля
        if len(password) < 8:
            CustomMessageBox(self, "Ошибка", "Пароль должен содержать не менее 8 символов").wait_window()
            return
        if not re.search(r'[A-Za-z]', password) or not re.search(r'\d', password):
            CustomMessageBox(self,"Ошибка", "Пароль должен содержать буквы и цифры").wait_window()
            return

        db = DatabaseHandler()
        cursor = db.conn.cursor()
        try:
            cursor.execute("SELECT login FROM users WHERE login = ?", (login,))
            if cursor.fetchone():
                CustomMessageBox(self, "Ошибка", "Логин уже занят").wait_window()
                return

            hashed_password = hashlib.sha256(f"salt{password}".encode()).hexdigest()
            cursor.execute(
                "INSERT INTO users (fio, login, password) VALUES (?, ?, ?)",
                (fio, login, hashed_password)
            )
            db.conn.commit()
            messagebox.showinfo("Успех", "Регистрация прошла успешно")
            self.destroy()
        except sqlite3.Error as e:
            messagebox.showerror("Ошибка БД", str(e))
        finally:
            cursor.close()

class LoginDialog(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Вход")
        self.geometry("300x150")
        try:
            self.iconbitmap('icon.ico')  
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.create_widgets()
        self.grab_set()

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.pack(padx=10, pady=10, fill='both', expand=True)

        ttk.Label(frame, text="Логин:").grid(row=0, column=0, sticky='w', pady=2)
        self.login_entry = ttk.Entry(frame)
        self.login_entry.grid(row=0, column=1, pady=2)

        ttk.Label(frame, text="Пароль:").grid(row=1, column=0, sticky='w', pady=2)
        self.password_entry = ttk.Entry(frame, show="*")
        self.password_entry.grid(row=1, column=1, pady=2)

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, columnspan=2, pady=10)

        ttk.Button(btn_frame, text="Войти", command=self.login).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Отмена", command=self.destroy).pack(side='left', padx=5)

    def login(self):
        login = self.login_entry.get().strip()
        password = self.password_entry.get().strip()

        if not login or not password:
            CustomMessageBox(self, "Ошибка", "Введите логин и пароль").wait_window()
            return

        db = DatabaseHandler()
        cursor = db.conn.cursor()
        try:
            cursor.execute(
                "SELECT id, fio, login, password, is_admin, can_edit_1, can_edit_2, login_attempts, is_locked "
                "FROM users WHERE login = ?",
                (login,)
            )
            result = cursor.fetchone()
            if not result:
                messagebox.showerror("Ошибка", "Неверный логин или пароль")
                return
            user_id, fio, login, stored_hash, is_admin, can_edit_1, can_edit_2, login_attempts, is_locked = result
            
            if is_locked:
                CustomMessageBox(self, "Ошибка", "Учетная запись заблокирована").wait_window()
                return

            hashed_password = hashlib.sha256(f"salt{password}".encode()).hexdigest()
            if stored_hash != hashed_password:
                new_attempts = login_attempts + 1
                is_locked_new = new_attempts >= 5
                cursor.execute(
                    "UPDATE users SET login_attempts = ?, is_locked = ? WHERE id = ?",
                    (new_attempts, is_locked_new, user_id)
                )
                db.conn.commit()
                remaining = 5 - new_attempts
                if is_locked_new:
                    messagebox.showerror("Ошибка", "Учетная запись заблокирована после 5 неудачных попыток.")
                else:
                    messagebox.showerror("Ошибка", f"Неверный пароль. Осталось попыток: {remaining}")
                return

            # Сброс попыток при успешном входе
            cursor.execute(
                "UPDATE users SET login_attempts = 0, is_locked = 0 WHERE id = ?",
                (user_id,)
            )
            db.conn.commit()

            messagebox.showinfo("Успех", "Вход выполнен")
            self.parent.destroy()
            self.destroy()
            app = MainApp({
                'id': user_id,
                'is_admin': bool(is_admin),
                'can_edit_1': bool(can_edit_1),
                'can_edit_2': bool(can_edit_2)
            })
            app.mainloop()

        except sqlite3.Error as e:
            messagebox.showerror("Ошибка БД", str(e))
        finally:
            cursor.close()

class MainApp(tk.Tk):
    def __init__(self, user_info):  # Убираем параметр parent
        super().__init__()  # Вызываем конструктор Tk без параметров
        self.title("Реестры комитета")
        self.geometry("600x400")
        try:
            self.iconbitmap('icon.ico')  
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.db = DatabaseHandler()
        self.user_info = user_info
        
        self.style = ttk.Style()
        self.style.configure('TButton', font=('Arial', 12), padding=10)
        self.create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def open_admin_panel(self):
        AdminWindow(self)

    def on_close(self):
        self.destroy()
        sys.exit()

    def create_widgets(self):
        frame = ttk.Frame(self)
        frame.pack(expand=True, fill='both', padx=50, pady=50)
        
        ttk.Label(frame, text="Выберите тип реестра:", font=('Arial', 14)).pack(pady=20)
        
        # Кнопки реестров с проверкой прав
        btn1_state = 'normal' if self.user_info['can_edit_1'] or self.user_info['is_admin'] else 'disabled'
        btn1 = ttk.Button(frame, 
                        text="РЕЕСТР договоров купли-продажи", 
                        command=lambda: FileWindow(self, 1),
                        state=btn1_state)
        btn1.pack(pady=10, fill='x')
        
        btn2_state = 'normal' if self.user_info['can_edit_2'] or self.user_info['is_admin'] else 'disabled'
        btn2 = ttk.Button(frame, 
                        text="РЕЕСТР соглашений о перераспр.", 
                        command=lambda: FileWindow(self, 2),
                        state=btn2_state)
        btn2.pack(pady=10, fill='x')

        # Кнопка админ-панели
        if self.user_info['is_admin']:
            ttk.Button(frame, 
                     text="Управление пользователями", 
                     command=self.open_admin_panel
                     ).pack(pady=20)
            
        ttk.Button(frame, 
                 text="Выход", 
                 command=self.logout
                 ).pack(pady=20)
        
    def logout(self):
        self.destroy()  
        root = tk.Tk()  
        root.withdraw()
        auth_window = AuthWindow(root)
        auth_window.mainloop()

class AdminWindow(tk.Toplevel):
    def __init__(self, parent):
        super().__init__(parent)
        self.parent = parent
        self.title("Управление пользователями")
        self.geometry("1200x600")
        try:
            self.iconbitmap('icon.ico')  
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.db = DatabaseHandler()
        self.filter_status = tk.StringVar(value='Все')
        self.create_widgets()
        self.load_users()

    def create_widgets(self):
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True, padx=10, pady=10)
        
        filter_frame = ttk.Frame(container)
        filter_frame.grid(row=0, column=0, sticky='ew', pady=5)
    
        ttk.Label(filter_frame, text="Фильтр по статусу:").pack(side='left', padx=5)
        
        filter_combo = ttk.Combobox(
            filter_frame, 
            textvariable=self.filter_status,
            values=['Все', 'Активные', 'Заблокированные'],
            state='readonly',
            width=15
        )
        filter_combo.pack(side='left', padx=5)
        filter_combo.bind('<<ComboboxSelected>>', self.apply_filter)

        # Treeview с прокрутками
        self.tree = ttk.Treeview(
            container,
            columns=('fio', 'login', 'admin', 'edit1', 'edit2', 'attempts', 'locked'),
            show='headings'
        )
        
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.bind('<ButtonRelease-1>', self.on_edit)

        # Настройка колонок
        columns_config = [
            ('fio', 'ФИО', 200),
            ('login', 'Логин', 150),
            ('admin', 'Админ', 80),
            ('edit1', 'Реестр 1', 80),
            ('edit2', 'Реестр 2', 80),
            ('attempts', 'Попытки', 80),
            ('locked', 'Блокировка', 90)
        ]
        
        for col, text, width in columns_config:
            self.tree.heading(col, text=text)
            anchor = 'center' if col in ['admin', 'edit1', 'edit2', 'attempts', 'locked'] else 'w'
            self.tree.column(col, width=width, anchor=anchor)

        # Размещение элементов
        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        hsb.grid(row=1, column=0, sticky='ew')

        # Кнопки управления
        btn_frame = ttk.Frame(container)
        btn_frame.grid(row=2, column=0, columnspan=2, pady=10, sticky='ew')
        
        ttk.Button(btn_frame, text="Разблокировать", command=self.unlock_user).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Сбросить попытки", command=self.reset_attempts).pack(side='left', padx=5)
        ttk.Button(btn_frame, text="Сохранить изменения", command=self.save_changes).pack(side='right')

        # Настройка расширения областей
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def apply_filter(self, event=None):
        """Обработчик применения фильтра"""
        self.load_users()

    def load_users(self):
        status = self.filter_status.get()
        self.tree.delete(*self.tree.get_children())
        cursor = self.db.conn.cursor()
        query = """
            SELECT id, fio, login, is_admin, can_edit_1, can_edit_2, login_attempts, is_locked 
            FROM users
        """
        if status == 'Активные':
            query += " WHERE is_locked = 0"
        elif status == 'Заблокированные':
            query += " WHERE is_locked = 1"
        
        cursor.execute(query)
        for row in cursor.fetchall():
            self.tree.insert('', 'end', 
                values=(
                    row[1],  # fio
                    row[2],  # login
                    '✓' if row[3] else '✗',  # admin
                    '✓' if row[4] else '✗',  # edit1
                    '✓' if row[5] else '✗',  # edit2
                    row[6],  # attempts
                    '✓' if row[7] else '✗'  # locked
                ),
                tags=(row[0],)
            )

    def unlock_user(self):
        selected = self.tree.selection()
        if not selected:
            CustomMessageBox(self, "Ошибка", "Выберите пользователя!").wait_window()
            return
        
        user_id = self.tree.item(selected[0], 'tags')[0]
        with self.db.conn:
            self.db.conn.execute("UPDATE users SET is_locked=0 WHERE id=?", (user_id,))
        self.load_users()
        CustomMessageBox(self, "Успех", "Пользователь разблокирован").wait_window()

    def reset_attempts(self):
        selected = self.tree.selection()
        if not selected:
            CustomMessageBox(self, "Ошибка", "Выберите пользователя!").wait_window()
            return
        
        user_id = self.tree.item(selected[0], 'tags')[0]
        with self.db.conn:
            self.db.conn.execute("UPDATE users SET login_attempts=0 WHERE id=?", (user_id,))
        self.load_users()
        CustomMessageBox(self, "Успех", "Счетчик попыток сброшен").wait_window()
    
    def on_edit(self, event):
        selected = self.tree.selection()
        if not selected:  # Проверяем, есть ли выбранный элемент
            return
        item = selected[0]
        column = self.tree.identify_column(event.x)

        if column in ['#3', '#4', '#5']:  # Колонки Admin, Edit1, Edit2
            user_id = self.tree.item(item, 'tags')[0]
            current_value = self.tree.set(item, column)
            new_value = '✓' if current_value == '✗' else '✗'
            self.tree.set(item, column, new_value)
    
    def save_changes(self):
        cursor = self.db.conn.cursor()
        for item in self.tree.get_children():
            user_id = self.tree.item(item, 'tags')[0]
            values = self.tree.item(item, 'values')
            
            updates = {
                'is_admin': 1 if values[2] == '✓' else 0,
                'can_edit_1': 1 if values[3] == '✓' else 0,
                'can_edit_2': 1 if values[4] == '✓' else 0
            }
            
            cursor.execute('''UPDATE users SET
                is_admin = ?,
                can_edit_1 = ?,
                can_edit_2 = ?
                WHERE id = ?''',
                (updates['is_admin'], updates['can_edit_1'], updates['can_edit_2'], user_id))
        
        self.db.conn.commit()
        messagebox.showinfo("Сохранено", "Изменения успешно сохранены!")

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
        '№ выписки учета поступлений, № ПП',
        'Оплачено',
        'примечание',
        'начисленные ПЕНИ',
        'оплачено пеней',
        'Дата выписки учета поступлений, № ПП',
        'Возврат имеющейся переплаты',
        # Расчетные колонки в конце
        'Контроль по дате ("-" - просрочка)',
        'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
        'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)',
        'Редактор'
        ],
        2: [
            '№ соглашения',
            'Дата заключения',
            'Собственник, ИНН',
            'Кадастровый номер образуемого ЗУ, адрес ЗУ',
            'Площадь образуемого ЗУ, кв. м',
            'реквизиты приказа ГК ПО по им. Отнош.',
            'Размер платы за увеличение площади ЗУ, руб.',
            'Срок оплаты',
            'Фактическая дата оплаты',
            'Контроль по дате ("-" - просрочка)',
            '№ выписки учета поступлений, № ПП',
            'Оплачено',
            'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
            'примечание',
            'начисленные ПЕНИ',
            'оплачено пеней',
            'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)',
            'Возврат имеющейся переплаты',
            'Редактор'
        ]
    }  
    date_columns = {
        1: [
            'Дата заключения договора',
            'Срок оплаты по договору',
            'Фактическая дата оплаты',
            'Дата выписки учета поступлений, № ПП'
        ],
        2: [
            'Дата заключения',
            'Срок оплаты',
            'Фактическая дата оплаты'
        ]
    }
    
    def show_tooltip(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            if self.tooltip_timer:
                self.after_cancel(self.tooltip_timer)
            
            self.tooltip_timer = self.after(500, self._create_tooltip, event)
        else:
            self.hide_tooltip()

    def _create_tooltip(self, event):
        item = self.tree.identify_row(event.y)
        col = self.tree.identify_column(event.x)
        
        # Получаем индекс данных (игнорируя скрытый столбец id)
        col_index = int(col[1:]) - 2  # -2 так как первый столбец id (индекс 0)
        
        if 0 <= col_index < len(self.expected_columns[self.file_type]):
            value = self.tree.item(item, 'values')[col_index + 1]  # +1 из-за id
            
            # Получаем координаты ячейки
            x, y, width, height = self.tree.bbox(item, column=col)
            
            # Преобразуем координаты в абсолютные
            x_root = self.tree.winfo_rootx() + x + width//2
            y_root = self.tree.winfo_rooty() + y + height
            
            if self.tooltip:
                self.tooltip.destroy()
            
            self.tooltip = tk.Toplevel(self)
            self.tooltip.wm_overrideredirect(True)
            self.tooltip.wm_geometry(f"+{x_root}+{y_root}")
            
            label = ttk.Label(self.tooltip, 
                            text=value, 
                            background="#ffffe0",
                            relief='solid',
                            borderwidth=1,
                            wraplength=300)
            label.pack()

    def hide_tooltip(self, event=None):
        if self.tooltip_timer:
            self.after_cancel(self.tooltip_timer)
            self.tooltip_timer = None
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None

    def __init__(self, parent, file_type): 
        super().__init__(parent)
        self.parent = parent
        self.file_type = file_type
        # Инициализация user_info и show_editor до создания виджетов
        self.user_info = parent.user_info
        self.show_editor = parent.user_info['is_admin']
        
        # Настройка стилей
        self.style = ttk.Style(self)
        self.style.configure("Treeview", 
                            font=('Arial', 10), 
                            rowheight=30,
                            background="#ffffff",
                            fieldbackground="#ffffff")
        self.style.configure("Treeview.Heading", 
                            font=('Arial', 10, 'bold'),
                            background="#e1e1e1",
                            relief="flat")
        self.style.map("Treeview.Heading", 
                      background=[('active', '#d0d0d0')])
        
        self.configure_ui()
        try:
            self.iconbitmap('icon.ico')  
        except Exception as e:
            print("Ошибка загрузки иконки:", e)
        self.create_widgets() 
        self.create_toolbar()
        self.setup_tags()
        self.update_treeview()

        self.state('zoomed')
        self.tooltip = None
        self.tooltip_timer = None
    
    def has_permission(self):
        return (self.user_info['is_admin'] or 
               (self.file_type == 1 and self.user_info['can_edit_1']) or 
               (self.file_type == 2 and self.user_info['can_edit_2']))

    def calculate_days_diff(self, row):
        try:
            due_date = datetime.strptime(row['Срок оплаты по договору'], "%d.%m.%Y")
            actual_date = datetime.strptime(row['Фактическая дата оплаты'], "%d.%m.%Y")
            return (due_date - actual_date).days  # Правильный порядок вычитания
        except Exception as e:
            print(f"Ошибка расчета дней: {str(e)}")
            return 0

    # В методах расчета:
    def calculate_payment_diff(self, row):
        try:
            price = float(row.get('Цена ЗУ по договору, руб.') or 0)
            paid = float(row.get('Оплачено') or 0)
            return price - paid
        except:
            return 0

    def calculate_peni_diff(self, row):
        try:
            accrued = float(row.get('начисленные ПЕНИ') or 0)
            paid_pen = float(row.get('оплачено пеней') or 0)
            return accrued - paid_pen
        except:
            return 0

    def configure_ui(self):
        titles = {
        1: "Реестр договоров купли-продажи земельных участков, гос. собств-ть на кот. не разграничена",
        2: "Реестр соглашений о перераспр. земель, гос. собств-ть на кот. не разграничена"
        }
        self.title(titles[self.file_type])
        self.geometry("1400x800")
        self.style.configure("Red.Treeview", background="#ffcccc")
        self.style.configure("Yellow.Treeview", background="#ffffcc")

    def setup_tags(self):
        self.tree.tag_configure('overdue', background='#ffcccc')
        self.tree.tag_configure('warning', background='#ffffcc')
        self.tree.tag_configure('evenrow', background='#f8f8f8')
        self.tree.tag_configure('oddrow', background='#ffffff')

    def sort_treeview(self, col, reverse):
        data = [(self.tree.set(child, col), child) 
              for child in self.tree.get_children('')]
        
        # Пытаемся преобразовать к числам для правильной сортировки
        try:
            data.sort(key=lambda t: float(t[0].replace(',', '.')), reverse=reverse)
        except:
            data.sort(reverse=reverse)
        
        for index, (val, child) in enumerate(data):
            self.tree.move(child, '', index)
        
        self.tree.heading(col, 
                        command=lambda: self.sort_treeview(col, not reverse))
        
        # Обновляем цвета строк после сортировки
        self.update_row_colors()
       
        #self.filter_data()

    def update_row_colors(self):
        for i, child in enumerate(self.tree.get_children('')):
            tags = list(self.tree.item(child, 'tags'))
            # Удаляем старые теги строк
            tags = [t for t in tags if t not in ('evenrow', 'oddrow')]
            # Добавляем новые теги
            tags.append('evenrow' if i % 2 == 0 else 'oddrow')
            self.tree.item(child, tags=tags)

    def show_tooltip(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region == "cell":
            item = self.tree.identify_row(event.y)
            col = self.tree.identify_column(event.x)
            col_index = int(col[1:]) - 1
            if 0 <= col_index < len(self.expected_columns[self.file_type]):
                value = self.tree.item(item, 'values')[col_index + 1]
                
                # Создаем подсказку
                x, y, _, _ = self.tree.bbox(item, col)
                x += self.tree.winfo_rootx() + 20
                y += self.tree.winfo_rooty() + 20
                
                self.tooltip = tk.Toplevel(self)
                self.tooltip.wm_overrideredirect(True)
                self.tooltip.wm_geometry(f"+{x}+{y}")
                
                label = ttk.Label(self.tooltip, 
                                text=value, 
                                background="#ffffe0",
                                relief='solid', 
                                borderwidth=1,
                                wraplength=300)  # Перенос текста
                label.pack()
                self.tooltip_label = label

    def hide_tooltip(self, event):
        if hasattr(self, 'tooltip'):
            self.tooltip.destroy()
            del self.tooltip

    def create_widgets(self):
        container = ttk.Frame(self)
        container.pack(fill='both', expand=True)
        
        self.tree = ttk.Treeview(container, show='headings')
        vsb = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        # Добавляем скрытый столбец ID
        self.tree["columns"] = ['id'] + self.expected_columns[self.file_type]
        for col in self.tree["columns"]:
            if col == 'id':
                self.tree.column(col, width=0, stretch=False)
                self.tree.heading(col, text='', anchor='center')  # Добавьте это
            elif col == 'Редактор' and not self.show_editor:
                self.tree.column(col, width=0, stretch=False)
                self.tree.heading(col, text='')  # Скрываем заголовок
            else:
                self.tree.heading(col, text=col, anchor='center')
                self.tree.column(col, width=150, minwidth=100, stretch=True)
        
        for col in self.tree["columns"][1:]:  # Пропускаем столбец 'id'
            self.tree.heading(col, 
                            text=col, 
                            anchor='center',
                            command=lambda c=col: self.sort_treeview(c, False))
        
        # Привязка событий для всплывающих подсказок
        # self.tree.bind('<Motion>', self.show_tooltip)
        # self.tree.bind('<Leave>', self.hide_tooltip)
        # self.tree.bind('<ButtonPress>', self.hide_tooltip)

        # Добавляем горизонтальную прокрутку
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(xscrollcommand=hsb.set)
        hsb.grid(row=1, column=0, sticky='ew')

        self.tree.grid(row=0, column=0, sticky='nsew')
        vsb.grid(row=0, column=1, sticky='ns')
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        self.tree.bind('<Double-1>', self.on_double_click)

    def validate_date(self, date_str):
        try:
            datetime.strptime(date_str, "%d.%m.%Y")
            return True
        except ValueError:
            return False

    def filter_data(self, event=None):
        query = self.search_var.get().lower().strip()
        col = self.column_var.get()
        
        if col not in self.tree["columns"]:
            return
    
        col_index = self.tree["columns"].index(col)

        # Определяем числовые колонки для текущего типа файла
        numeric_columns = {
            1: ['Цена ЗУ по договору, руб.', 'Оплачено', 
                'начисленные ПЕНИ', 'оплачено пеней',
                'Площадь ЗУ, кв. м', 'Контроль по дате ("-" - просрочка)',
                'Контроль по оплате цены ("-" - переплата; "+" - недоплата)',
                'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)'],
            2: [],
            3: []
        }.get(self.file_type, [])
        
        is_numeric_col = col in numeric_columns

        for item in self.tree.get_children(''):
            values = self.tree.item(item, 'values')

            if col not in self.tree["columns"]:
                return 
            col_index = self.tree["columns"].index(col)
            raw_value = str(values[col_index])

            match = False
            
            try:
                if is_numeric_col and query:
                    # Нормализуем числовые значения
                    cell_num = float(raw_value.replace(',', '.').replace(' ', ''))
                    query_num = float(query.replace(',', '.'))
                    match = math.isclose(cell_num, query_num, rel_tol=1e-9)
                else:
                    # Текстовый поиск
                    cell_value = raw_value.lower()
                    match = query in cell_value
            except:
                # В случае ошибки преобразования - используем текстовый поиск
                cell_value = raw_value.lower()
                match = query in cell_value
            
            # Обновляем теги
            current_tags = list(self.tree.item(item, 'tags'))
            current_tags = [t for t in current_tags if t not in ('match', 'nomatch')]
            
            if match:
                current_tags.append('match')
            else:
                current_tags.append('nomatch')
            
            self.tree.item(item, tags=current_tags)
        
        # Настройка внешнего вида тегов
        self.tree.tag_configure('match', background='#90EE90')
        self.tree.tag_configure('nomatch', background='gray90')
    
    def create_toolbar(self):
        toolbar = ttk.Frame(self)
        toolbar.pack(fill='x', padx=5, pady=5)
        
        btn_load = ttk.Button(toolbar, text="Загрузить файл", 
                            command=self.load_file,
                            state='normal' if self.has_permission() else 'disabled')
        btn_load.pack(side='left', padx=2)
        
        btn_new = ttk.Button(toolbar, text="Создать новый", 
                           command=self.create_new,
                           state='normal' if self.has_permission() else 'disabled')
        btn_new.pack(side='left', padx=2)
        
        btn_save = ttk.Button(toolbar, text="Сохранить", 
                            command=self.save_file,
                            state='normal' if self.has_permission() else 'disabled')
        btn_save.pack(side='right', padx=2)

        btn_xml = ttk.Button(toolbar, text="Экспорт в XML", 
                           command=self.generate_xml_report)
        btn_xml.pack(side='left', padx=2)

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

    def generate_xml_report(self):
        try:
            # Получаем данные из БД
            df = self.parent.db.export_to_dataframe(self.file_type)
            
            # Создаем XML структуру
            root = ET.Element('Реестр')
            root.set('ВерсияФормата', '1.0')
            root.set('ДатаФормирования', datetime.now().strftime('%Y-%m-%d'))  # Исправленный формат даты

            for _, row in df.iterrows():
                record = ET.SubElement(root, 'Запись')
                
                for col in self.expected_columns[self.file_type]:
                    if col in ['Редактор', 'id']:
                        continue
                    
                    # Генерируем валидное имя тега
                    tag_name = re.sub(r'[^a-zA-Zа-яА-Я0-9_]', '_', col)
                    tag_name = tag_name.strip('_').rstrip('_')
                    
                    # Заменяем множественные подчеркивания на одно
                    tag_name = re.sub(r'_+', '_', tag_name)
                    
                    field = ET.SubElement(record, tag_name)
                    
                    # Обработка значений
                    value = str(row[col]) if not pd.isna(row[col]) else ''
                    field.text = escape(value)

            # Сохраняем файл с правильной кодировкой
            file_path = filedialog.asksaveasfilename(
                defaultextension=".xml",
                filetypes=[("XML files", "*.xml")]
            )
            
            if file_path:
                # Создаем дерево с правильным объявлением кодировки
                tree = ET.ElementTree(root)
                
                # Принудительно записываем XML декларацию с кодировкой
                with open(file_path, 'wb') as f:
                    f.write('<?xml version="1.0" encoding="UTF-8"?>\n'.encode('utf-8'))
                    tree.write(f, 
                            encoding='utf-8',
                            xml_declaration=False,  # Уже записали декларацию вручную
                            method='xml')
                
                messagebox.showinfo("Успех", "XML-отчет успешно сформирован!")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка генерации XML: {str(e)}")

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
        
        elif self.file_type == 2:
            numeric_columns = [
                'Площадь образуемого ЗУ, кв. м',
                'Размер платы за увеличение площади ЗУ, руб.',
                'Оплачено',
                'начисленные ПЕНИ',
                'оплачено пеней'
            ]
            date_columns = [
                'Дата заключения',
                'Срок оплаты',
                'Фактическая дата оплаты'
            ]
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
                    df[col] = df[col].astype(str).str.replace(r'\s+', '', regex=True)
                    df[col] = df[col].str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)

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

            for col in numeric_columns:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.replace(r'\s+', '', regex=True)
                    df[col] = df[col].str.replace(',', '.', regex=False)
                    df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0)
            
            # Обработка дат
            for col in date_columns:
                if col in df.columns:
                    df[col] = pd.to_datetime(df[col], dayfirst=True, errors='coerce').dt.strftime('%d.%m.%Y')
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
                contract_date = datetime.now()
                due_date = (contract_date + timedelta(days=7)).strftime('%d.%m.%Y')  # Добавляем 7 дней
                default_data = {
                    'Номер договора': '',
                    'Дата заключения договора': contract_date.strftime('%d.%m.%Y'),
                    'Покупатель, ИНН': 'ООО "Пример", ИНН 0000000000',
                    'Кадастровый номер ЗУ, адрес ЗУ': '00:00:0000000:00, г. Пример',
                    'Площадь ЗУ, кв. м': '',
                    'Разрешенное использование ЗУ': '',
                    'Основание предоставления': 'Номер ЗК РФ',
                    'Цена ЗУ по договору, руб.': '0.00',
                    'Срок оплаты по договору': '',
                    'Фактическая дата оплаты': '',
                    '№ выписки учета поступлений, № ПП': '',
                    'Оплачено': '0.00',
                    'примечание': '',
                    'начисленные ПЕНИ': '0.00',
                    'оплачено пеней': '0.00',
                    'Дата выписки учета поступлений, № ПП': '',
                    'Возврат имеющейся переплаты': ''
                }
            elif self.file_type == 2:
                default_data = {
                    '№ соглашения': 'б/н',
                    'Дата заключения': datetime.now().strftime('%d.%m.%Y'),
                    'Собственник, ИНН': 'Иванов Иван Иванович, ИНН 0000000000',
                    'Кадастровый номер образуемого ЗУ, адрес ЗУ': '60:00:0000000:000, г. Пример',
                    'Площадь образуемого ЗУ, кв. м': '1000',
                    'реквизиты приказа ГК ПО по им. Отнош.': 'п. 13 ст. 39.29 ЗК РФ',
                    'Размер платы за увеличение площади ЗУ, руб.': '10000.00',
                    'Срок оплаты': '',
                    'Фактическая дата оплаты': '',
                    '№ выписки учета поступлений, № ПП': '',
                    'Оплачено': '0.00',
                    'примечание': '',
                    'начисленные ПЕНИ': '0.00',
                    'оплачено пеней': '0.00'
                }
            else:
                default_data = {col: '' for col in expected}
            
            # Создаем DataFrame с одной строкой
            df = pd.DataFrame([default_data], columns=expected)
            df['Редактор'] = self.parent.user_info['id']
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
                2: []
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
            # Преобразуем запись в список для модификации
            record_list = list(record)
            
            # Заменяем ID редактора на ФИО
            editor_id = record_list[-1]  # Последний элемент - ID редактора
            editor_name = self.parent.db.get_user_fio(editor_id) if editor_id else ""
            record_list[-1] = editor_name  # Заменяем ID на ФИО
            
            tags = []
            try:
                # Рассчитываем условия подсветки
                days_diff = float(record_list[self.expected_columns[self.file_type].index('Контроль по дате ("-" - просрочка)') + 1])
                payment_diff = float(record_list[self.expected_columns[self.file_type].index('Контроль по оплате цены ("-" - переплата; "+" - недоплата)') + 1])
                
                if days_diff < 0:
                    tags.append('overdue')
                if payment_diff > 0:
                    tags.append('warning')
            except Exception as e:
                print(f"Ошибка расчетов: {e}")

            # Формируем значения для вставки (пропускаем скрытый ID записи)
            values = [str(x) if x is not None else "" for x in record_list[1:]]
            
            # Вставляем одну запись
            self.tree.insert('', 'end', values=[record_id] + values, tags=tags)
            
        # Обновляем ширину колонок
        for col in self.tree["columns"][1:]:  # Пропускаем 'id'
            max_len = max(
                [len(str(self.tree.set(child, col))) 
                for child in self.tree.get_children('')] + [len(col)]
            )
            self.tree.column(col, width=int(max_len * 8.5))

        # Обновляем цвета строк
        self.update_row_colors()
    

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
            if not self.has_permission():
                messagebox.showwarning("Доступ запрещен")
                return
            processed_value = None
            # Обработка числовых полей
            if col_name in ['Цена ЗУ по договору, руб.', 'Оплачено', 
                          'начисленные ПЕНИ', 'оплачено пеней', 'Площадь ЗУ, кв. м']:
                cleaned_value = new_value.replace(' ', '').replace(',', '.').strip()
                processed_value = float(cleaned_value) if cleaned_value else 0.0
            
            # Обработка дат
            elif col_name in self.date_columns.get(self.file_type, []):
                if not self.validate_date(new_value):
                    raise ValueError("Некорректный формат даты (требуется ДД.ММ.ГГГГ)")
                processed_value = new_value
            
            # Обработка текстовых полей
            else:
                processed_value = new_value.strip()

            # Обновление записи
            self.parent.db.update_record(
            self.file_type, 
            record_id, 
            col_name, 
            processed_value
            )

            cursor = self.parent.db.conn.cursor()
            cursor.execute(
                f"UPDATE {self.parent.db.get_table_name(self.file_type)} "
                "SET \"Редактор\" = ? WHERE id = ?",
                (self.parent.user_info['id'], record_id)
            )
            self.parent.db.conn.commit()
            # Обновление интерфейса
            self.update_treeview()
            edit_win.destroy()
            messagebox.showinfo("Успех", "Изменения успешно сохранены!")

        except ValueError as ve:
            messagebox.showerror("Ошибка", f"Некорректный ввод: {str(ve)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка сохранения: {str(e)}")

    # def update_calculations(self, record_id, edited_col, new_value):
    #     cursor = self.parent.db.conn.cursor()
    #     cursor.execute('SELECT * FROM contracts WHERE id = ?', (record_id,))
    #     record = cursor.fetchone()
        
    #     if not record:
    #         print(f"Запись с id {record_id} не найдена")
    #         return

    #     try:
    #         # Преобразуем запись в словарь
    #         columns = [col[0] for col in cursor.description]
    #         record_dict = dict(zip(columns, record))
            
    #         # Обновление срока оплаты
    #         if edited_col in ['Срок оплаты по договору', 'Фактическая дата оплаты']:
    #             cursor.execute('''
    #                 SELECT 
    #                     "Срок оплаты по договору",
    #                     "Фактическая дата оплаты",
    #                     "Контроль по дате (""-"" - просрочка)"
    #                 FROM contracts WHERE id = ?
    #             ''', (record_id,))
    #             due_date_str, actual_date_str, _ = cursor.fetchone()
                
    #             if due_date_str and actual_date_str:
    #                 try:
    #                     due_date = datetime.strptime(due_date_str, '%d.%m.%Y')
    #                     actual_date = datetime.strptime(actual_date_str, '%d.%m.%Y')
    #                     days_diff = (actual_date - due_date).days  # Правильный порядок вычисления
    #                     self.parent.db.update_record(
    #                         1, 
    #                         record_id, 
    #                         'Контроль по дате (""-"" - просрочка)', 
    #                         days_diff
    #                     )
    #                 except Exception as e:
    #                     print(f"Ошибка расчета дней: {str(e)}")
    #         # Расчет контроля по оплате
    #         if edited_col in ['Цена ЗУ по договору, руб.', 'Оплачено']:
    #             price = float(record[7]) if edited_col != 'Цена ЗУ по договору, руб.' else float(new_value)
    #             paid = float(record[12]) if edited_col != 'Оплачено' else float(new_value)
    #             payment_diff = price - paid
    #             self.parent.db.update_record(1, record_id, 
    #                 'Контроль по оплате цены ("-" - переплата; "+" - недоплата)', f"{payment_diff:.2f}")

    #         # Расчет неоплаченных пени
    #         if edited_col in ['начисленные ПЕНИ', 'оплачено пеней']:
    #             accrued = float(record[15]) if edited_col != 'начисленные ПЕНИ' else float(new_value)
    #             paid_pen = float(record[16]) if edited_col != 'оплачено пеней' else float(new_value)
    #             unpaid_pen = math.floor(accrued - paid_pen)
    #             self.parent.db.update_record(1, record_id, 
    #                 'неоплаченные ПЕНИ ("+" - недоплата; "-" - переплата)', str(unpaid_pen))

    #     except Exception as e:
    #         print(f"Ошибка в расчетах: {e}")
    #         # Добавляем запись в лог
    #         with open('error.log', 'a') as f:
    #             f.write(f"{datetime.now()} - Ошибка: {str(e)}\n")

    def clear_placeholder(self, entry, original):
        if entry.get() in ["Только целые числа", "Число с точкой/запятой"]:
            entry.delete(0, tk.END)
            entry.config(foreground='black')
            if original:
                entry.insert(0, original)

    def on_double_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        item = self.tree.identify_row(event.y)
        if not item:  # Добавленная проверка
            return

        if not self.has_permission():
            messagebox.showwarning("Доступ запрещен")
            return

        item = self.tree.identify_row(event.y)
        if not item:  # Проверка на существование элемента
            return  
        
        column = self.tree.identify_column(event.x)
        col_index = int(column[1:]) - 2  # Изменено с -1 на -2
        if col_index < 0 or col_index >= len(self.expected_columns[self.file_type]):
            return
        col_name = self.expected_columns[self.file_type][col_index]
        
        # Проверяем является ли колонка датой для текущего типа реестра
        date_columns = self.date_columns.get(self.file_type, [])
        
        if col_name in self.parent.db.calculated_columns.get(self.file_type, []):
            messagebox.showinfo("Информация", "Это поле рассчитывается автоматически и не может быть изменено вручную.")
            return
        
        current_value = self.tree.item(item, 'values')[col_index + 1]
        edit_win = tk.Toplevel(self)
        edit_win.title("Редактирование")
        edit_win.geometry("400x200")  # Увеличим высоту для календаря
        
        record_id = int(self.tree.item(item, 'values')[0])
        entry = ttk.Entry(edit_win, font=('Arial', 12), width=30)
        entry.pack(padx=20, pady=20, fill='x', expand=True)
        entry.insert(0, current_value)

        # Добавляем календарь для дат
        if col_name in date_columns:
            self.create_calendar(edit_win, entry, col_name)
        
        # Остальной код без изменений...
        if col_name == 'Номер договора':
            entry.insert(0, "Только целые числа")
            entry.config(foreground='grey')
        elif col_name == 'Площадь ЗУ, кв. м':
            entry.insert(0, "Число с точкой/запятой")
            entry.config(foreground='grey')
        
        entry.bind('<FocusIn>', lambda e: self.clear_placeholder(entry, current_value))

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

        # Валидация для числовых полей
        if col_name in ['Цена ЗУ по договору, руб.', 'Оплачено', 'начисленные ПЕНИ', 'оплачено пеней']:
            validate_cmd = (edit_win.register(self.validate_number), '%P')
            entry.configure(validate='key', validatecommand=validate_cmd)
        
    def validate_number(self, value):
        try:
            value = value.strip()
            if not value:
                return True
            value = value.replace(',', '.')
            
            # Проверка на несколько точек
            if value.count('.') > 1:
                return False
                
            # Разделение на целую и дробную части
            parts = value.split('.')
            
            # Проверка, что все части - числа
            if not all(part.isdigit() for part in parts if part):
                return False
                
            # Проверка количества знаков после запятой
            if len(parts) > 1 and len(parts[1]) > 2:
                return False
                
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
            elif self.file_type == 2:
                # Форматирование числовых колонок
                num_cols = [
                    'Размер платы за увеличение площади ЗУ, руб.',
                    'Оплачено',
                    'начисленные ПЕНИ',
                    'оплачено пеней'
                ]
                for col in num_cols:
                    df[col] = df[col].apply(lambda x: f"{x:,.2f}".replace(',', ' ').replace('.', ',') if pd.notnull(x) else '')
                
                # Форматирование дат
                date_cols = [
                    'Дата заключения',
                    'Срок оплаты',
                    'Фактическая дата оплаты'
                ]
                for col in date_cols:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%d.%m.%Y')

            if not self.parent.user_info['is_admin'] and 'Редактор' in df.columns:
                df = df.drop(columns=['Редактор'])
            
            df.to_excel(file_path, index=False, engine='openpyxl')
            messagebox.showinfo("Успех", "Файл успешно сохранен!")
            
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))
        
def backup_database():
    DB_NAME = 'registers.db'
    BACKUP_DIR = 'backups'
    
    try:
        # Создаем директорию для бэкапов
        if not os.path.exists(BACKUP_DIR):
            os.makedirs(BACKUP_DIR)
        
        # Генерируем имя файла с timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_name = f"{BACKUP_DIR}/registers_{timestamp}.bak"
        
        # Создаем соединение с базой для гарантии целостности
        conn = sqlite3.connect(DB_NAME)
        conn.backup(sqlite3.connect(backup_name))
        conn.close()
        
        print(f"Backup created: {backup_name}")
        return True
    
    except Exception as e:
        print(f"Backup error: {str(e)}")
        return False

if __name__ == "__main__":
    backup_database()
    root = tk.Tk()
    root.withdraw()
    auth_window = AuthWindow(root)
    auth_window.mainloop()