import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sqlite3
from datetime import datetime
from tkcalendar import DateEntry
import openpyxl
import logging

# Basic logging configuration
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

class PatientManagementApp:
    def __init__(self, root):
        self.root = root
        self.root.title("سامانه مدیریت بیماران")
        self.root.geometry("1400x800")
        self.root.minsize(1350, 700)
        self.root.config(bg="#f0f0f0")
        
        # Style configuration
        style = ttk.Style(self.root)
        style.configure("TLabel", font=('Tahoma', 10), background="#f0f0f0", anchor="e")
        style.configure("TButton", font=('Tahoma', 10, 'bold'))
        style.configure("TEntry", font=('Tahoma', 10))
        style.configure("TCombobox", font=('Tahoma', 10))
        style.configure("Treeview.Heading", font=('Tahoma', 11, 'bold'))
        style.configure("Treeview", rowheight=25, font=('Tahoma', 10))
        style.configure("TLabelframe.Label", font=('Tahoma', 11, 'bold'))

        self.db_name = 'hospital_patients.db'
        self.conn = None
        self.cursor = None
        self.connect_db()

        self.create_widgets()
        self.display_patients()

        self.selected_patient_db_id = None

    def connect_db(self):
        try:
            self.conn = sqlite3.connect(self.db_name)
            self.cursor = self.conn.cursor()
            logging.info("Database connection successful.")
            self.create_table()
            self.create_specialists_table()
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در اتصال به پایگاه داده: {e}")
            self.root.destroy()

    def create_table(self):
        try:
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS patients (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    patient_name TEXT NOT NULL,
                    last_name TEXT NOT NULL,
                    age INTEGER NOT NULL,
                    ward TEXT NOT NULL,
                    patient_code TEXT NOT NULL,
                    specialist TEXT NOT NULL,
                    submission_date TEXT NOT NULL,
                    submission_time TEXT NOT NULL
                )
            ''')
            self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("خطای ساخت جدول", f"خطا در ایجاد جدول پایگاه داده: {e}")

    def create_specialists_table(self):
        try:
            self.cursor.execute('''
                CREATE TABLE IF NOT EXISTS specialists (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    specialist_name TEXT NOT NULL,
                    is_active INTEGER NOT NULL DEFAULT 1
                )
            ''')
            self.conn.commit()
            # Insert initial specialists if table is empty
            self.cursor.execute("SELECT COUNT(*) FROM specialists")
            if self.cursor.fetchone()[0] == 0:
                initial_specialists = [
                    "قلب و عروق - دکتر کریمی", "داخلی - دکتر رضایی", "اطفال - دکتر محمدی",
                    "پوست - دکتر قاسمی", "چشم - دکتر احمدی", "ارتوپدی - دکتر حسینی",
                    "گوش و حلق و بینی - دکتر نوری", "مغز و اعصاب - دکتر مرادی",
                    "جراحی - دکتر یوسفی", "اورولوژی - دکتر بهرامی", "زنان و زایمان - دکتر علوی",
                    "ریه - دکتر پارسا", "غدد - دکتر اکبری", "گوارش - دکتر شجاعی",
                    "روانپزشکی - دکتر جمشیدی"
                ]
                for spec in initial_specialists:
                    self.cursor.execute("INSERT INTO specialists (specialist_name, is_active) VALUES (?, 1)", (spec,))
                self.conn.commit()
        except sqlite3.Error as e:
            messagebox.showerror("خطای ساخت جدول", f"خطا در ایجاد جدول متخصصین: {e}")

    def get_active_specialists(self):
        try:
            self.cursor.execute("SELECT specialist_name FROM specialists WHERE is_active=1")
            return [row[0] for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در دریافت متخصصین: {e}")
            return []

    def get_all_specialists(self):
        try:
            self.cursor.execute("SELECT specialist_name FROM specialists")
            return [row[0] for row in self.cursor.fetchall()]
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در دریافت همه متخصصین: {e}")
            return []

    def create_widgets(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill="both", expand=True)

        # --- Input Frame ---
        input_frame = ttk.LabelFrame(main_frame, text="ثبت و ویرایش اطلاعات بیمار", padding="10", labelanchor="ne")
        input_frame.pack(fill="x", pady=5)
        for i in range(6): input_frame.grid_columnconfigure(i, weight=1)

        labels_texts = [":نام بیمار", ":نام خانوادگی", ":سن", ":بخش", ":کد بیمار", ":پزشک متخصص"]
        self.entries = {}
        
        ttk.Label(input_frame, text=labels_texts[0], anchor="e").grid(row=0, column=5, sticky="e", padx=5, pady=5)
        self.entries["نام بیمار"] = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.entries["نام بیمار"], justify='right').grid(row=0, column=4, sticky="ew", padx=5, pady=5)
        
        ttk.Label(input_frame, text=labels_texts[1], anchor="e").grid(row=0, column=3, sticky="e", padx=5, pady=5)
        self.entries["نام خانوادگی"] = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.entries["نام خانوادگی"], justify='right').grid(row=0, column=2, sticky="ew", padx=5, pady=5)
        
        ttk.Label(input_frame, text=labels_texts[2], anchor="e").grid(row=0, column=1, sticky="e", padx=5, pady=5)
        self.entries["سن"] = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.entries["سن"], justify='right', width=10).grid(row=0, column=0, sticky="ew", padx=5, pady=5)

        ttk.Label(input_frame, text=labels_texts[3], anchor="e").grid(row=1, column=5, sticky="e", padx=5, pady=5)
        self.entries["بخش"] = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.entries["بخش"], justify='right').grid(row=1, column=4, sticky="ew", padx=5, pady=5)

        ttk.Label(input_frame, text=labels_texts[4], anchor="e").grid(row=1, column=3, sticky="e", padx=5, pady=5)
        self.entries["کد بیمار"] = tk.StringVar()
        ttk.Entry(input_frame, textvariable=self.entries["کد بیمار"], justify='right').grid(row=1, column=2, sticky="ew", padx=5, pady=5)

        ttk.Label(input_frame, text=labels_texts[5], anchor="e").grid(row=2, column=5, sticky="e", padx=5, pady=5)
        self.specialists = self.get_active_specialists()
        self.specialist_var = tk.StringVar()
        self.specialist_combo = ttk.Combobox(input_frame, textvariable=self.specialist_var, values=self.specialists,
                                            state="readonly", justify='right', height=len(self.specialists))
        self.specialist_combo.grid(row=2, column=2, columnspan=3, sticky="ew", padx=5, pady=5)
        if self.specialists:
            self.specialist_var.set(self.specialists[0])
        
        self.submit_button = ttk.Button(input_frame, text="ثبت اطلاعات", command=self.add_patient)
        self.submit_button.grid(row=2, column=0, columnspan=2, sticky="ew", padx=5, pady=10)

        # --- Specialist Management Frame ---
        specialist_frame = ttk.LabelFrame(main_frame, text="مدیریت پزشکان ویزیت‌کننده", padding="10", labelanchor="ne")
        specialist_frame.pack(fill="x", pady=5)
        for i in range(6): specialist_frame.grid_columnconfigure(i, weight=1)

        ttk.Label(specialist_frame, text=":نام پزشک", anchor="e").grid(row=0, column=5, sticky="e", padx=5, pady=5)
        self.new_specialist_var = tk.StringVar()
        ttk.Entry(specialist_frame, textvariable=self.new_specialist_var, justify='right').grid(row=0, column=4, sticky="ew", padx=5, pady=5)

        ttk.Button(specialist_frame, text="اضافه کردن پزشک", command=self.add_specialist).grid(row=0, column=2, columnspan=2, sticky="ew", padx=5, pady=5)
        ttk.Button(specialist_frame, text="حذف پزشک", command=self.delete_specialist).grid(row=0, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # --- Filter Frame ---
        filter_frame = ttk.LabelFrame(main_frame, text="جستجو و فیلتر", padding="10", labelanchor="ne")
        filter_frame.pack(fill="x", pady=10)
        for i in range(8): filter_frame.grid_columnconfigure(i, weight=1)

        ttk.Label(filter_frame, text=":فیلتر تخصص", font=('Tahoma', 10), anchor="e").grid(row=0, column=7, sticky="e", padx=5)
        self.filter_specialist_var = tk.StringVar(value="همه متخصصین")
        self.filter_specialist_combo = ttk.Combobox(filter_frame, textvariable=self.filter_specialist_var,
                                                  values=["همه متخصصین"] + self.get_all_specialists(), state="readonly", justify='right',
                                                  height=len(self.get_all_specialists()) + 1)
        self.filter_specialist_combo.grid(row=0, column=6, sticky="ew", padx=5)
        self.filter_specialist_combo.bind("<<ComboboxSelected>>", self.filter_patients_by_specialist)

        ttk.Label(filter_frame, text=":از تاریخ", font=('Tahoma', 10), anchor="e").grid(row=0, column=5, sticky="e", padx=5)
        self.date1_entry = DateEntry(filter_frame, date_pattern='yyyy-mm-dd', locale='fa_IR', font=('Tahoma', 9))
        self.date1_entry.grid(row=0, column=4, sticky="ew", padx=5)
        
        ttk.Label(filter_frame, text=":تا تاریخ", font=('Tahoma', 10), anchor="e").grid(row=0, column=3, sticky="e", padx=5)
        self.date2_entry = DateEntry(filter_frame, date_pattern='yyyy-mm-dd', locale='fa_IR', font=('Tahoma', 9))
        self.date2_entry.grid(row=0, column=2, sticky="ew", padx=5)

        ttk.Button(filter_frame, text="اعمال فیلتر تاریخ", command=self.filter_patients_by_date_range).grid(row=0, column=0, columnspan=2, sticky="ew", padx=5)

        ttk.Label(filter_frame, text=":جستجوی کد بیمار", font=('Tahoma', 10), anchor="e").grid(row=1, column=7, sticky="e", padx=5, pady=5)
        self.search_code_var = tk.StringVar()
        search_entry = ttk.Entry(filter_frame, textvariable=self.search_code_var, justify='right', font=('Tahoma', 10))
        search_entry.grid(row=1, column=5, columnspan=2, sticky="ew", padx=5, pady=5)
        
        ttk.Button(filter_frame, text="جستجو", command=self.search_patient_by_code).grid(row=1, column=4, sticky="ew", padx=5, pady=5)

        ttk.Button(filter_frame, text="نمایش همه و بازنشانی", command=self.reset_filters_and_display_all).grid(row=1, column=0, columnspan=2, sticky="ew", padx=5, pady=5)

        # --- Treeview Frame ---
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)

        self.display_columns_order = ("نام بیمار", "نام خانوادگی", "سن", "بخش", "کد بیمار", "پزشک متخصص", "تاریخ ثبت", "زمان ثبت", "id")
        
        self.patient_tree = ttk.Treeview(tree_frame, columns=self.display_columns_order, displaycolumns=self.display_columns_order, show="headings")
        
        self.patient_tree.heading("نام بیمار", text="نام بیمار", anchor="e")
        self.patient_tree.column("نام بیمار", width=150, anchor="e", stretch=tk.YES)

        self.patient_tree.heading("نام خانوادگی", text="نام خانوادگی", anchor="e")
        self.patient_tree.column("نام خانوادگی", width=180, anchor="e", stretch=tk.YES)
        
        self.patient_tree.heading("سن", text="سن", anchor="center")
        self.patient_tree.column("سن", width=60, anchor="center", stretch=tk.NO)

        self.patient_tree.heading("بخش", text="بخش", anchor="e")
        self.patient_tree.column("بخش", width=120, anchor="e", stretch=tk.YES)

        self.patient_tree.heading("کد بیمار", text="کد بیمار", anchor="e")
        self.patient_tree.column("کد بیمار", width=120, anchor="e", stretch=tk.NO)

        self.patient_tree.heading("پزشک متخصص", text="پزشک متخصص", anchor="e")
        self.patient_tree.column("پزشک متخصص", width=250, anchor="e", stretch=tk.YES)
        
        self.patient_tree.heading("تاریخ ثبت", text="تاریخ ثبت", anchor="center")
        self.patient_tree.column("تاریخ ثبت", width=120, anchor="center", stretch=tk.NO)
        
        self.patient_tree.heading("زمان ثبت", text="زمان ثبت", anchor="center")
        self.patient_tree.column("زمان ثبت", width=120, anchor="center", stretch=tk.NO)

        self.patient_tree.heading("id", text="ردیف", anchor="center")
        self.patient_tree.column("id", width=60, anchor="center", stretch=tk.NO)
        
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.patient_tree.yview)
        self.patient_tree.configure(yscrollcommand=vsb.set)
        
        vsb.pack(side="right", fill="y")
        self.patient_tree.pack(side="left", fill="both", expand=True)

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(pady=10, fill="x")
        for i in range(3): action_frame.grid_columnconfigure(i, weight=1)

        ttk.Button(action_frame, text="ویرایش بیمار منتخب", command=self.edit_patient).grid(row=0, column=2, padx=5, sticky="ew")
        ttk.Button(action_frame, text="حذف بیمار(ان) منتخب", command=self.delete_selected_patients).grid(row=0, column=1, padx=5, sticky="ew")
        ttk.Button(action_frame, text="خروجی اکسل", command=self.export_to_excel).grid(row=0, column=0, padx=5, sticky="ew")

    def add_specialist(self):
        specialist_name = self.new_specialist_var.get().strip()
        if not specialist_name:
            messagebox.showwarning("ورودی ناقص", "لطفا نام پزشک را وارد کنید.")
            return
        try:
            self.cursor.execute("SELECT specialist_name FROM specialists WHERE specialist_name=?", (specialist_name,))
            if self.cursor.fetchone():
                messagebox.showwarning("تکراری", "این پزشک قبلا ثبت شده است.")
                return
            self.cursor.execute("INSERT INTO specialists (specialist_name, is_active) VALUES (?, 1)", (specialist_name,))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "پزشک با موفقیت اضافه شد.")
            self.new_specialist_var.set("")
            self.specialists = self.get_active_specialists()
            self.specialist_combo['values'] = self.specialists
            if self.specialists:
                self.specialist_var.set(self.specialists[0])
            self.filter_specialist_combo['values'] = ["همه متخصصین"] + self.get_all_specialists()
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در اضافه کردن پزشک: {e}")

    def delete_specialist(self):
        specialist_name = self.new_specialist_var.get().strip()
        if not specialist_name:
            messagebox.showwarning("انتخاب کنید", "لطفا نام پزشک را در کادر وارد کنید.")
            return
        try:
            self.cursor.execute("SELECT COUNT(*) FROM specialists WHERE specialist_name=? AND is_active=1", (specialist_name,))
            if self.cursor.fetchone()[0] == 0:
                messagebox.showwarning("خطا", "پزشک مورد نظر یافت نشد یا غیرفعال است.")
                return
            self.cursor.execute("SELECT COUNT(*) FROM patients WHERE specialist=?", (specialist_name,))
            if self.cursor.fetchone()[0] > 0:
                messagebox.showwarning("خطا", "نمی‌توان پزشک را حذف کرد زیرا در سوابق بیماران استفاده شده است.")
                return
            self.cursor.execute("UPDATE specialists SET is_active=0 WHERE specialist_name=?", (specialist_name,))
            self.conn.commit()
            messagebox.showinfo("موفقیت", "پزشک با موفقیت غیرفعال شد.")
            self.new_specialist_var.set("")
            self.specialists = self.get_active_specialists()
            self.specialist_combo['values'] = self.specialists
            if self.specialists:
                self.specialist_var.set(self.specialists[0])
            else:
                self.specialist_var.set("")
            self.filter_specialist_combo['values'] = ["همه متخصصین"] + self.get_all_specialists()
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در غیرفعال کردن پزشک: {e}")

    def add_patient(self):
        name = self.entries["نام بیمار"].get().strip()
        last_name = self.entries["نام خانوادگی"].get().strip()
        age_str = self.entries["سن"].get().strip()
        ward = self.entries["بخش"].get().strip()
        code = self.entries["کد بیمار"].get().strip()
        specialist = self.specialist_var.get()

        if not all([name, last_name, age_str, ward, code, specialist]):
            messagebox.showwarning("ورودی ناقص", ".لطفا تمام فیلدها را پر کنید")
            return

        try:
            age = int(age_str)
            if not (0 < age < 150): raise ValueError
        except ValueError:
            messagebox.showwarning("ورودی نامعتبر", ".سن باید یک عدد صحیح معتبر باشد")
            return

        if self.selected_patient_db_id:
            self.update_patient_data(name, last_name, age, ward, code, specialist)
        else:
            self.insert_new_patient(name, last_name, age, ward, code, specialist)

    def insert_new_patient(self, name, last_name, age, ward, code, specialist):
        try:
            submission_date = datetime.now().strftime("%Y-%m-%d")
            submission_time = datetime.now().strftime("%H:%M:%S")
            self.cursor.execute("""
                INSERT INTO patients (patient_name, last_name, age, ward, patient_code, specialist, submission_date, submission_time)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """, (name, last_name, age, ward, code, specialist, submission_date, submission_time))
            self.conn.commit()
            messagebox.showinfo("موفقیت", ".اطلاعات بیمار با موفقیت ثبت شد")
            self.clear_entries()
            self.display_patients()
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در ثبت اطلاعات: {e}")

    def update_patient_data(self, name, last_name, age, ward, code, specialist):
        try:
            self.cursor.execute("""
                UPDATE patients
                SET patient_name=?, last_name=?, age=?, ward=?, patient_code=?, specialist=?
                WHERE id=?
            """, (name, last_name, age, ward, code, specialist, self.selected_patient_db_id))
            self.conn.commit()
            messagebox.showinfo("موفقیت", ".اطلاعات بیمار با موفقیت به‌روزرسانی شد")
            self.clear_entries()
            self.display_patients()
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در به‌روزرسانی اطلاعات: {e}")

    def clear_entries(self):
        for var in self.entries.values():
            var.set("")
        if self.specialists:
            self.specialist_var.set(self.specialists[0])
        if self.selected_patient_db_id:
            self.selected_patient_db_id = None
            self.submit_button.config(text="ثبت اطلاعات")
            self.root.title("سامانه مدیریت بیماران")

    def reset_filters_and_display_all(self):
        self.search_code_var.set("")
        self.filter_specialist_var.set("همه متخصصین")
        today = datetime.now()
        self.date1_entry.set_date(today)
        self.date2_entry.set_date(today)
        self.clear_entries()
        self.display_patients()

    def display_patients(self, data=None):
        for item in self.patient_tree.get_children():
            self.patient_tree.delete(item)

        rows = data
        if data is None:
            try:
                self.cursor.execute("SELECT id, patient_name, last_name, age, ward, patient_code, specialist, submission_date, submission_time FROM patients ORDER BY id DESC")
                rows = self.cursor.fetchall()
            except sqlite3.Error as e:
                messagebox.showerror("خطای پایگاه داده", f"خطا در بازیابی اطلاعات: {e}")
                return
        
        row_counter = 1
        for row in rows:
            db_id = row[0]
            display_values = row[1:] + (row_counter,)
            self.patient_tree.insert("", "end", iid=db_id, values=display_values)
            row_counter += 1

    def edit_patient(self):
        selected_items = self.patient_tree.selection()
        if not selected_items:
            messagebox.showwarning("انتخاب کنید", ".لطفا یک بیمار را برای ویرایش انتخاب کنید")
            return
        if len(selected_items) > 1:
            messagebox.showwarning("انتخاب چندگانه", ".فقط یک بیمار را می‌توان در هر لحظه ویرایش کرد")
            return

        item_db_id = selected_items[0]
        self.selected_patient_db_id = item_db_id

        try:
            self.cursor.execute("SELECT patient_name, last_name, age, ward, patient_code, specialist FROM patients WHERE id=?", (item_db_id,))
            db_data = self.cursor.fetchone()
            if not db_data:
                messagebox.showerror("خطا", ".بیمار مورد نظر در پایگاه داده یافت نشد")
                return

            self.entries["نام بیمار"].set(db_data[0])
            self.entries["نام خانوادگی"].set(db_data[1])
            self.entries["سن"].set(str(db_data[2]))
            self.entries["بخش"].set(db_data[3])
            self.entries["کد بیمار"].set(db_data[4])
            self.specialist_var.set(db_data[5])
            
            self.submit_button.config(text="به‌روزرسانی اطلاعات")
            self.root.title(f"در حال ویرایش بیمار: {db_data[0]} {db_data[1]}")
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در خواندن اطلاعات برای ویرایش: {e}")

    def delete_selected_patients(self):
        selected_items = self.patient_tree.selection()
        if not selected_items:
            messagebox.showwarning("انتخاب کنید", ".لطفا یک یا چند بیمار را برای حذف انتخاب کنید")
            return

        confirm = messagebox.askyesno("تایید حذف", f"آیا از حذف {len(selected_items)} بیمار منتخب اطمینان دارید؟ این عمل قابل بازگشت نیست.")
        if confirm:
            try:
                ids_to_delete = [(item_id,) for item_id in selected_items]
                self.cursor.executemany("DELETE FROM patients WHERE id=?", ids_to_delete)
                self.conn.commit()
                messagebox.showinfo("موفقیت", ".بیمار(ان) با موفقیت حذف شدند")
                self.clear_entries()
                self.display_patients()
            except sqlite3.Error as e:
                messagebox.showerror("خطای پایگاه داده", f"خطا در حذف اطلاعات: {e}")

    def filter_patients_by_specialist(self, event=None):
        selected_specialist = self.filter_specialist_var.get()
        query = "SELECT * FROM patients"
        params = []
        if selected_specialist != "همه متخصصین":
            query += " WHERE specialist=?"
            params.append(selected_specialist)
        query += " ORDER BY id DESC"
        
        try:
            self.cursor.execute(query, params)
            self.display_patients(self.cursor.fetchall())
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در فیلتر کردن: {e}")

    def filter_patients_by_date_range(self):
        try:
            start_date = self.date1_entry.get_date().strftime('%Y-%m-%d')
            end_date = self.date2_entry.get_date().strftime('%Y-%m-%d')
            self.cursor.execute("SELECT * FROM patients WHERE submission_date BETWEEN ? AND ? ORDER BY id DESC", (start_date, end_date))
            self.display_patients(self.cursor.fetchall())
        except Exception as e:
            messagebox.showerror("خطای تاریخ", f"خطا در فیلتر تاریخ: {e}")

    def search_patient_by_code(self, event=None):
        search_term = self.search_code_var.get().strip()
        if not search_term:
            self.filter_patients_by_specialist()
            return
        try:
            query = "SELECT * FROM patients WHERE patient_code LIKE ?"
            params = [f'%{search_term}%']
            
            selected_specialist = self.filter_specialist_var.get()
            if selected_specialist != "همه متخصصین":
                query += " AND specialist=?"
                params.append(selected_specialist)

            query += " ORDER BY id DESC"
            self.cursor.execute(query, params)
            self.display_patients(self.cursor.fetchall())
        except sqlite3.Error as e:
            messagebox.showerror("خطای پایگاه داده", f"خطا در جستجو: {e}")

    def export_to_excel(self):
        if not self.patient_tree.get_children():
            messagebox.showwarning("داده‌ای وجود ندارد", ".جدول خالی است. داده‌ای برای خروجی گرفتن وجود ندارد")
            return

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
            title="ذخیره فایل اکسل"
        )
        if not file_path:
            return

        try:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "گزارش بیماران"
            sheet.sheet_view.rightToLeft = True

            headers = [self.patient_tree.heading(col)["text"] for col in self.display_columns_order]
            sheet.append(headers)

            items_in_order = self.patient_tree.get_children('')
            for item_id in items_in_order:
                row_values = self.patient_tree.item(item_id, 'values')
                sheet.append(list(row_values))

            workbook.save(file_path)
            messagebox.showinfo("موفقیت", f"اطلاعات با موفقیت در فایل زیر ذخیره شد:{file_path}")
        except Exception as e:
            messagebox.showerror("خطا در خروجی", f"خطا در تولید فایل اکسل: {e}")

    def on_closing(self):
        if self.conn:
            self.conn.close()
            logging.info("Database connection closed.")
        self.root.destroy()

if __name__ == "__main__":
    root = tk.Tk()
    app = PatientManagementApp(root)
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    root.mainloop()