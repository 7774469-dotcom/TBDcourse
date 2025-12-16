import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pyodbc
import csv
from datetime import datetime

SERVER_NAME = 'HOME-PC\SQLEXPRESS'  
DATABASE_NAME = 'FinalAttestationDB'

CONNECTION_STRING = f'DRIVER={{SQL Server}};SERVER={SERVER_NAME};DATABASE={DATABASE_NAME};Trusted_Connection=yes;'

# ==========================================
# 🗄️ БЭКЕНД (РАБОТА С БАЗОЙ)
# ==========================================
class DBManager:
    def __init__(self):
        self.conn = None

    def connect(self):
        try:
            self.conn = pyodbc.connect(CONNECTION_STRING)
            return True
        except Exception as e:
            messagebox.showerror("Ошибка подключения", f"Проверь имя сервера!\nОшибка: {e}")
            return False

    # Получить список всех студентов для выпадающего меню (Combobox)
    def get_student_list_for_login(self):
        query = "SELECT StudentID, LastName + ' ' + FirstName + ' (' + RecordBookNumber + ')' as FullName FROM Students ORDER BY LastName"
        cursor = self.conn.cursor()
        cursor.execute(query)
        # Возвращаем список кортежей: [(1, 'Ахметов Серик (ZK-..)'), ...]
        return cursor.fetchall()

    # АДМИН: Получить полную таблицу + ID результата (чтобы можно было менять)
    def get_all_results_admin(self, search_text=""):
        query = """
        SELECT 
            ar.ResultID,
            s.LastName + ' ' + s.FirstName, 
            g.GroupName, 
            at.TypeName, 
            ar.Grade, 
            cm.FullName, 
            ar.ExamDate
        FROM AttestationResults ar
        JOIN Students s ON ar.StudentID = s.StudentID
        JOIN StudentGroups g ON s.GroupID = g.GroupID
        JOIN AttestationTypes at ON ar.TypeID = at.TypeID
        JOIN CommissionMembers cm ON ar.MemberID = cm.MemberID
        WHERE s.LastName LIKE ? OR s.FirstName LIKE ? OR g.GroupName LIKE ?
        ORDER BY g.GroupName, s.LastName
        """
        params = (f'%{search_text}%', f'%{search_text}%', f'%{search_text}%')
        cursor = self.conn.cursor()
        cursor.execute(query, params)
        return cursor.fetchall()

    # АДМИН: Обновить оценку (UPDATE)
    def update_grade(self, result_id, new_grade):
        try:
            cursor = self.conn.cursor()
            cursor.execute("UPDATE AttestationResults SET Grade = ? WHERE ResultID = ?", (new_grade, result_id))
            self.conn.commit()
            return True
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось обновить: {e}")
            return False

    # АДМИН: Статистика
    def get_group_stats(self):
        query = """
        SELECT g.GroupName, COUNT(s.StudentID), AVG(CAST(ar.Grade AS FLOAT))
        FROM AttestationResults ar
        JOIN Students s ON ar.StudentID = s.StudentID
        JOIN StudentGroups g ON s.GroupID = g.GroupID
        GROUP BY g.GroupName
        ORDER BY AVG(CAST(ar.Grade AS FLOAT)) DESC
        """
        cursor = self.conn.cursor()
        cursor.execute(query)
        return cursor.fetchall()

    # СТУДЕНТ: Данные + Средний балл
    def get_student_results(self, student_id):
        cursor = self.conn.cursor()
        # Оценки
        cursor.execute("""
            SELECT at.TypeName, ar.Grade, ar.Topic, cm.FullName, ar.ExamDate
            FROM AttestationResults ar
            JOIN AttestationTypes at ON ar.TypeID = at.TypeID
            JOIN CommissionMembers cm ON ar.MemberID = cm.MemberID
            WHERE ar.StudentID = ?
        """, (student_id,))
        results = cursor.fetchall()
        
        # Средний балл
        cursor.execute("SELECT AVG(CAST(Grade AS FLOAT)) FROM AttestationResults WHERE StudentID = ?", (student_id,))
        avg_grade = cursor.fetchone()[0]
        
        return results, avg_grade

# ==========================================
# 🖥️ ИНТЕРФЕЙС (GUI)
# ==========================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Система Учета Аттестации (Курсовая)")
        self.geometry("1000x650")
        self.db = DBManager()
        
        # Красивая тема
        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("Treeview", font=('Segoe UI', 10), rowheight=28)
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'), background="#2c3e50", foreground="white")
        style.map("Treeview", background=[('selected', '#3498db')])
        
        if not self.db.connect():
            self.destroy()
            return

        # Загружаем список студентов один раз при запуске
        self.student_map = {} # Словарь "Имя Фамилия" -> ID
        raw_students = self.db.get_student_list_for_login()
        self.student_names = []
        for s_id, s_name in raw_students:
            self.student_map[s_name] = s_id
            self.student_names.append(s_name)

        self.show_login_screen()

    def clear_screen(self):
        for widget in self.winfo_children():
            widget.destroy()

    # -----------------------------------------------------------
    # ЭКРАН 1: ВХОД (LOGIN)
    # -----------------------------------------------------------
    def show_login_screen(self):
        self.clear_screen()
        bg_color = "#ecf0f1"
        frame = tk.Frame(self, bg=bg_color)
        frame.pack(fill="both", expand=True)

        # Логотип / Заголовок
        tk.Label(frame, text="🎓 Итоговая Аттестация", font=("Segoe UI", 26, "bold"), bg=bg_color, fg="#2c3e50").pack(pady=(80, 10))
        tk.Label(frame, text="Выберите роль для входа в систему", font=("Segoe UI", 12), bg=bg_color, fg="#7f8c8d").pack(pady=(0, 40))

        # Блок входа
        login_frame = tk.Frame(frame, bg="white", padx=40, pady=40, relief="raised", bd=1)
        login_frame.pack()

        # Выбор студента (Combobox)
        tk.Label(login_frame, text="Войти как Студент:", font=("Segoe UI", 10, "bold"), bg="white").pack(anchor="w")
        self.combo_students = ttk.Combobox(login_frame, values=self.student_names, width=40, state="readonly")
        self.combo_students.set("Выберите свое имя...")
        self.combo_students.pack(pady=5)
        
        tk.Button(login_frame, text="Войти (Студент)", command=self.login_as_student, bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=5).pack(pady=10)

        tk.Label(login_frame, text="___________________________", bg="white", fg="#bdc3c7").pack(pady=10)

        # Вход админа
        tk.Label(login_frame, text="Пароль Администратора:", font=("Segoe UI", 10, "bold"), bg="white").pack(anchor="w")
        self.entry_admin_pass = tk.Entry(login_frame, show="•", width=43, bg="#ecf0f1", relief="flat")
        self.entry_admin_pass.pack(pady=5)

        tk.Button(login_frame, text="Войти (Админ)", command=self.login_as_admin, bg="#e74c3c", fg="white", font=("Segoe UI", 10, "bold"), relief="flat", padx=20, pady=5).pack(pady=10)

        tk.Label(frame, text="Подсказка для защиты: Пароль админа - admin", bg=bg_color, fg="gray").pack(side="bottom", pady=20)

    def login_as_student(self):
        selection = self.combo_students.get()
        if selection in self.student_map:
            student_id = self.student_map[selection]
            self.show_student_dashboard(student_id, selection)
        else:
            messagebox.showwarning("Ошибка", "Пожалуйста, выберите студента из списка!")

    def login_as_admin(self):
        password = self.entry_admin_pass.get()
        if password == "admin":
            self.show_admin_dashboard()
        else:
            messagebox.showerror("Ошибка", "Неверный пароль администратора")

    # -----------------------------------------------------------
    # ЭКРАН 2: АДМИНКА
    # -----------------------------------------------------------
    def show_admin_dashboard(self):
        self.clear_screen()
        
        # Верхняя панель
        top = tk.Frame(self, bg="#2c3e50", height=60, padx=20)
        top.pack(fill="x")
        tk.Label(top, text="🔧 Панель управления", font=("Segoe UI", 16, "bold"), bg="#2c3e50", fg="white").pack(side="left", pady=10)
        tk.Button(top, text="Выйти", command=self.show_login_screen, bg="#c0392b", fg="white", relief="flat").pack(side="right")

        # Панель инструментов
        toolbar = tk.Frame(self, bg="#ecf0f1", padx=10, pady=10)
        toolbar.pack(fill="x")

        tk.Label(toolbar, text="Поиск:", bg="#ecf0f1", font=("Segoe UI", 11)).pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace("w", self.filter_admin_table) # Живой поиск
        tk.Entry(toolbar, textvariable=self.search_var, width=30).pack(side="left", padx=10)

        tk.Button(toolbar, text="💾 Экспорт в Excel", command=self.export_to_csv, bg="#27ae60", fg="white").pack(side="right")
        tk.Label(toolbar, text="ℹ️ Дважды кликните по строке, чтобы изменить оценку", bg="#ecf0f1", fg="gray").pack(side="right", padx=20)

        # Таблица
        cols = ("ID", "Student", "Group", "Type", "Grade", "Comm", "Date")
        self.tree_admin = ttk.Treeview(self, columns=cols, show="headings")
        
        # Настраиваем колонки (ID скрываем визуально, но он нужен для логики)
        self.tree_admin.heading("ID", text="ID")
        self.tree_admin.column("ID", width=0, stretch=False) # Скрытая колонка
        
        self.tree_admin.heading("Student", text="Студент")
        self.tree_admin.column("Student", width=200)
        
        self.tree_admin.heading("Group", text="Группа")
        self.tree_admin.column("Group", width=80, anchor="center")
        
        self.tree_admin.heading("Type", text="Тип аттестации")
        self.tree_admin.column("Type", width=200)
        
        self.tree_admin.heading("Grade", text="Оценка")
        self.tree_admin.column("Grade", width=60, anchor="center")
        
        self.tree_admin.heading("Comm", text="Преподаватель")
        self.tree_admin.heading("Date", text="Дата")

        self.tree_admin.pack(fill="both", expand=True, padx=10, pady=5)
        
        # Привязываем двойной клик
        self.tree_admin.bind("<Double-1>", self.on_double_click_admin)

        self.filter_admin_table() # Первичная загрузка

    def filter_admin_table(self, *args):
        # Очистка
        for row in self.tree_admin.get_children():
            self.tree_admin.delete(row)
        
        search = self.search_var.get()
        data = self.db.get_all_results_admin(search)
        
        for row in data:
            # row[0] это ID, row[4] это оценка. Раскрасим двойки красным
            tags = ('bad_mark',) if row[4] == 2 else ()
            self.tree_admin.insert("", "end", values=list(row), tags=tags)
        
        self.tree_admin.tag_configure('bad_mark', foreground='red')

    def on_double_click_admin(self, event):
        item = self.tree_admin.selection()
        if not item: return
        
        # Получаем данные строки
        values = self.tree_admin.item(item, "values")
        result_id = values[0]
        student_name = values[1]
        current_grade = values[4]

        # Открываем всплывающее окно
        self.open_edit_window(result_id, student_name, current_grade)

    def open_edit_window(self, result_id, name, grade):
        win = tk.Toplevel(self)
        win.title("Редактирование оценки")
        win.geometry("300x250")
        
        tk.Label(win, text=f"Студент: {name}", font=("bold"), wraplength=280).pack(pady=10)
        tk.Label(win, text="Новая оценка:").pack()
        
        # Шкала выбора оценки
        scale = tk.Scale(win, from_=2, to=5, orient="horizontal", length=200, tickinterval=1)
        scale.set(grade)
        scale.pack(pady=10)

        def save():
            if self.db.update_grade(result_id, scale.get()):
                messagebox.showinfo("Успех", "Оценка изменена!")
                self.filter_admin_table() # Обновляем таблицу
                win.destroy()

        tk.Button(win, text="Сохранить", command=save, bg="#3498db", fg="white", width=15).pack(pady=20)

    def export_to_csv(self):
        try:
            filename = f"Report_{datetime.now().strftime('%Y-%m-%d_%H-%M')}.csv"
            with open(filename, mode="w", newline="", encoding="utf-8-sig") as file:
                writer = csv.writer(file, delimiter=";")
                writer.writerow(["ID", "Студент", "Группа", "Тип", "Оценка", "Преподаватель", "Дата"])
                
                # Берем данные прямо из таблицы
                for row_id in self.tree_admin.get_children():
                    row = self.tree_admin.item(row_id)['values']
                    writer.writerow(row)
            
            messagebox.showinfo("Экспорт", f"Отчет успешно сохранен:\n{filename}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить: {e}")

    # -----------------------------------------------------------
    # ЭКРАН 3: СТУДЕНТ
    # -----------------------------------------------------------
    def show_student_dashboard(self, student_id, full_name):
        self.clear_screen()
        
        # Шапка
        top = tk.Frame(self, bg="#2980b9", height=80, padx=20)
        top.pack(fill="x")
        
        tk.Label(top, text=full_name, font=("Segoe UI", 18, "bold"), bg="#2980b9", fg="white").pack(side="left", pady=20)
        tk.Button(top, text="Выйти", command=self.show_login_screen, bg="white", fg="#2980b9").pack(side="right")

        # Контент
        content = tk.Frame(self, padx=20, pady=20)
        content.pack(fill="both", expand=True)

        # Получаем данные
        data, avg_grade = self.db.get_student_results(student_id)

        # Карточка GPA
        gpa_color = "green" if avg_grade and avg_grade >= 4.5 else "orange" if avg_grade and avg_grade >= 3.5 else "red"
        
        tk.Label(content, text=f"Ваш средний балл (GPA): {avg_grade:.2f}" if avg_grade else "Нет оценок", 
                 font=("Segoe UI", 16), fg=gpa_color).pack(anchor="w", pady=(0, 20))

        # Таблица
        cols = ("Type", "Grade", "Topic", "Comm", "Date")
        tree = ttk.Treeview(content, columns=cols, show="headings", height=10)
        
        tree.heading("Type", text="Дисциплина / Вид")
        tree.heading("Grade", text="Оценка")
        tree.heading("Topic", text="Тема / Билет")
        tree.heading("Comm", text="Принимал")
        tree.heading("Date", text="Дата")

        tree.column("Grade", width=50, anchor="center")
        tree.column("Topic", width=250)

        tree.pack(fill="both", expand=True)

        for row in data:
            tree.insert("", "end", values=list(row))

if __name__ == "__main__":
    app = App()
    app.mainloop()