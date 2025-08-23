# Patient Management System

This is a Python-based desktop application built with Tkinter for managing patient records in a hospital setting. It uses SQLite for data storage and provides a user-friendly interface for adding, editing, deleting, and filtering patient records, as well as exporting data to Excel.

## Features
- Add, edit, and delete patient records with details like name, last name, age, ward, patient code, specialist, and submission date/time.
- Manage a list of medical specialists with options to add or deactivate specialists.
- Filter patient records by specialist, date range, or patient code.
- Display patient records in a sortable table with vertical scrolling.
- Export patient data to Excel (.xlsx) files.
- Persian-centric interface with right-to-left text support and Persian calendar integration.
- Error handling for database operations, invalid inputs, and file exports.
- Logging for database connections and key actions.

## Requirements
- Python 3.7+
- `tkinter` (included with Python standard library)
- `tkcalendar` library (`pip install tkcalendar`)
- `openpyxl` library (`pip install openpyxl`)
- `sqlite3` (included with Python standard library)
- Optional: Tahoma font for optimal Persian text rendering

## Setup
1. Install dependencies using `pip install -r requirements.txt` (create a `requirements.txt` with `tkcalendar` and `openpyxl`).
2. Optionally, ensure the Tahoma font is installed on your system for proper Persian text display.
3. Run the application with `python app.py`.

## Usage
- Launch the application to open the main window.
- **Add Patient**: Enter patient details (name, last name, age, ward, patient code, specialist) and click "ثبت اطلاعات" to save.
- **Edit Patient**: Select a patient from the table, click "ویرایش بیمار منتخب", modify details, and click "به‌روزرسانی اطلاعات".
- **Delete Patient(s)**: Select one or more patients from the table and click "حذف بیمار(ان) منتخب" to remove them.
- **Manage Specialists**: Add new specialists or deactivate existing ones in the "مدیریت پزشکان ویزیت‌کننده" section.
- **Filter Records**: Use the specialist dropdown, date range picker (Persian calendar), or patient code search to filter the table.
- **Export to Excel**: Click "خروجی اکسل" to save the current table data as an Excel file.
- **Reset Filters**: Click "نمایش همه و بازنشانی" to clear filters and show all records.

## Database Structure
- `patients`: Stores patient records (id, patient_name, last_name, age, ward, patient_code, specialist, submission_date, submission_time).
- `specialists`: Stores specialist details (id, specialist_name, is_active).

## Code Structure
- `PatientManagementApp`: Main application class handling the UI, database operations, and logic.
  - `connect_db`, `create_table`, `create_specialists_table`: Initialize SQLite database and tables.
  - `add_patient`, `update_patient_data`, `delete_selected_patients`: Manage patient records.
  - `add_specialist`, `delete_specialist`: Manage specialist list.
  - `filter_patients_by_specialist`, `filter_patients_by_date_range`, `search_patient_by_code`: Filter and search patient records.
  - `export_to_excel`: Export table data to Excel.
  - `display_patients`: Populate the table with patient records.
  - `on_closing`: Ensure proper database cleanup on exit.

## Notes
- The application uses a Persian calendar (`tkcalendar` with `locale='fa_IR'`) for date selection.
- Specialists cannot be deleted if associated with patient records to maintain data integrity.
- The interface is optimized for right-to-left text with Persian labels and supports the Tahoma font for better readability.
- Logging is configured to track database connections and key actions in the console.

## License
MIT License

---

# سامانه مدیریت بیماران

این یک برنامه دسکتاپ مبتنی بر پایتون است که با استفاده از Tkinter برای مدیریت سوابق بیماران در محیط بیمارستانی طراحی شده است. این برنامه از SQLite برای ذخیره داده‌ها استفاده می‌کند و رابط کاربری ساده‌ای برای افزودن، ویرایش، حذف و فیلتر کردن سوابق بیماران و همچنین خروجی گرفتن به فرمت اکسل ارائه می‌دهد.

## ویژگی‌ها
- افزودن، ویرایش و حذف سوابق بیماران با جزئیاتی مانند نام، نام خانوادگی، سن، بخش، کد بیمار، پزشک متخصص و تاریخ/زمان ثبت.
- مدیریت لیست پزشکان متخصص با امکان افزودن یا غیرفعال کردن متخصصین.
- فیلتر کردن سوابق بیماران بر اساس تخصص، بازه زمانی یا کد بیمار.
- نمایش سوابق بیماران در جدولی قابل مرتب‌سازی با قابلیت اسکرول عمودی.
- خروجی گرفتن داده‌های بیماران به فایل‌های اکسل (.xlsx).
- رابط کاربری متمرکز بر پارسی با پشتیبانی از متن راست‌به‌چپ و ادغام تقویم پارسی.
- مدیریت خطاها برای عملیات پایگاه داده، ورودی‌های نامعتبر و خروجی فایل.
- ثبت لاگ برای اتصال به پایگاه داده و اقدامات کلیدی.

## پیش‌نیازها
- پایتون نسخه 3.7 یا بالاتر
- `tkinter` (موجود در کتابخانه استاندارد پایتون)
- کتابخانه `tkcalendar` (نصب با `pip install tkcalendar`)
- کتابخانه `openpyxl` (نصب با `pip install openpyxl`)
- `sqlite3` (موجود در کتابخانه استاندارد پایتون)
- اختیاری: فونت Tahoma برای نمایش بهینه متن پارسی

## راه‌اندازی
1. وابستگی‌ها را با استفاده از `pip install -r requirements.txt` نصب کنید (فایل `requirements.txt` را با درج `tkcalendar` و `openpyxl` ایجاد کنید).
2. در صورت تمایل، فونت Tahoma را روی سیستم خود نصب کنید تا نمایش متن پارسی بهینه باشد.
3. برنامه را با اجرای `python app.py` راه‌اندازی کنید.

## استفاده
- برنامه را اجرا کنید تا پنجره اصلی باز شود.
- **افزودن بیمار**: جزئیات بیمار (نام، نام خانوادگی، سن، بخش، کد بیمار، متخصص) را وارد کرده و روی "ثبت اطلاعات" کلیک کنید.
- **ویرایش بیمار**: بیمار را از جدول انتخاب کنید، روی "ویرایش بیمار منتخب" کلیک کنید، جزئیات را تغییر دهید و روی "به‌روزرسانی اطلاعات" کلیک کنید.
- **حذف بیمار(ان)**: یک یا چند بیمار را از جدول انتخاب کرده و روی "حذف بیمار(ان) منتخب" کلیک کنید.
- **مدیریت متخصصین**: در بخش "مدیریت پزشکان ویزیت‌کننده" متخصص جدید اضافه کنید یا متخصص موجود را غیرفعال کنید.
- **فیلتر سوابق**: از منوی کشویی تخصص، انتخابگر بازه زمانی (تقویم پارسی) یا جستجوی کد بیمار برای فیلتر کردن جدول استفاده کنید.
- **خروجی به اکسل**: روی "خروجی اکسل" کلیک کنید تا داده‌های جدول به صورت فایل اکسل ذخیره شوند.
- **بازنشانی فیلترها**: روی "نمایش همه و بازنشانی" کلیک کنید تا فیلترها پاک شده و همه سوابق نمایش داده شوند.

## ساختار پایگاه داده
- `patients`: ذخیره سوابق بیماران (شناسه، نام بیمار، نام خانوادگی، سن، بخش، کد بیمار، متخصص، تاریخ ثبت، زمان ثبت).
- `specialists`: ذخیره جزئیات متخصصین (شناسه، نام متخصص، وضعیت فعال).

## ساختار کد
- `PatientManagementApp`: کلاس اصلی برنامه که رابط کاربری، عملیات پایگاه داده و منطق را مدیریت می‌کند.
  - `connect_db`، `create_table`، `create_specialists_table`: راه‌اندازی پایگاه داده SQLite و جداول.
  - `add_patient`، `update_patient_data`، `delete_selected_patients`: مدیریت سوابق بیماران.
  - `add_specialist`، `delete_specialist`: مدیریت لیست متخصصین.
  - `filter_patients_by_specialist`، `filter_patients_by_date_range`، `search_patient_by_code`: فیلتر و جستجوی سوابق بیماران.
  - `export_to_excel`: خروجی گرفتن داده‌های جدول به اکسل.
  - `display_patients`: پر کردن جدول با سوابق بیماران.
  - `on_closing`: اطمینان از تمیز کردن پایگاه داده هنگام خروج.

## نکات
- برنامه از تقویم پارسی (`tkcalendar` با `locale='fa_IR'`) برای انتخاب تاریخ استفاده می‌کند.
- متخصصین در صورتی که در سوابق بیماران استفاده شده باشند، قابل حذف نیستند تا یکپارچگی داده‌ها حفظ شود.
- رابط کاربری برای متن راست‌به‌چپ با برچسب‌های پارسی بهینه شده و از فونت Tahoma برای خوانایی بهتر پشتیبانی می‌کند.
- ثبت لاگ برای رصد اتصال به پایگاه داده و اقدامات کلیدی در کنسول تنظیم شده است.

## مجوز
مجوز MIT

---

# 患者管理系统

这是一个基于Python的桌面应用程序，使用Tkinter构建，用于在医院环境中管理患者记录。它使用SQLite进行数据存储，提供用户友好的界面，用于添加、编辑、删除和过滤患者记录，并支持将数据导出到Excel。

## 功能
- 添加、编辑和删除患者记录，包含姓名、姓氏、年龄、病房、患者代码、专科医生和提交日期/时间等详细信息。
- 管理医疗专家列表，支持添加或停用专家。
- 按专科、日期范围或患者代码过滤患者记录。
- 在可排序的表格中显示患者记录，支持垂直滚动。
- 将患者数据导出到Excel (.xlsx) 文件。
- 以波斯语为中心，支持从右到左的文本和波斯日历集成。
- 处理数据库操作、无效输入和文件导出的错误。
- 记录数据库连接和关键操作的日志。

## 要求
- Python 3.7或更高版本
- `tkinter`（Python标准库中包含）
- `tkcalendar`库（使用`pip install tkcalendar`安装）
- `openpyxl`库（使用`pip install openpyxl`安装）
- `sqlite3`（Python标准库中包含）
- 可选：Tahoma字体，用于优化波斯文本显示

## 设置
1. 使用`pip install -r requirements.txt`安装依赖项（创建一个包含`tkcalendar`和`openpyxl`的`requirements.txt`文件）。
2. 可选：确保系统上安装了Tahoma字体以优化波斯文本显示。
3. 使用`python app.py`运行应用程序。

## 使用
- 启动应用程序以打开主窗口。
- **添加患者**：输入患者详细信息（姓名、姓氏、年龄、病房、患者代码、专科医生），点击“ثبت اطلاعات”保存。
- **编辑患者**：从表格中选择患者，点击“ویرایش بیمار منتخب”，修改详细信息，然后点击“به‌روزرسانی اطلاعات”。
- **删除患者**：从表格中选择一个或多个患者，点击“حذف بیمار(ان) منتخب”删除。
- **管理专家**：在“مدیریت پزشکان ویزیت‌کننده”部分添加新专家或停用现有专家。
- **过滤记录**：使用专科下拉菜单、日期范围选择器（波斯日历）或患者代码搜索来过滤表格。
- **导出到Excel**：点击“خروجی اکسل”将当前表格数据保存为Excel文件。
- **重置过滤器**：点击“نمایش همه و بازنشانی”清除过滤器并显示所有记录。

## 数据库结构
- `patients`：存储患者记录（ID、患者姓名、姓氏、年龄、病房、患者代码、专科医生、提交日期、提交时间）。
- `specialists`：存储专家详细信息（ID、专家姓名、活跃状态）。

## 代码结构
- `PatientManagementApp`：主应用程序类，处理用户界面、数据库操作和逻辑。
  - `connect_db`、`create_table`、`create_specialists_table`：初始化SQLite数据库和表。
  - `add_patient`、`update_patient_data`、`delete_selected_patients`：管理患者记录。
  - `add_specialist`、`delete_specialist`：管理专家列表。
  - `filter_patients_by_specialist`、`filter_patients_by_date_range`、`search_patient_by_code`：过滤和搜索患者记录。
  - `export_to_excel`：将表格数据导出到Excel。
  - `display_patients`：用患者记录填充表格。
  - `on_closing`：在退出时确保正确清理数据库。

## 注意事项
- 应用程序使用波斯日历（`tkcalendar`，设置`locale='fa_IR'`）进行日期选择。
- 如果专家与患者记录相关联，则无法删除，以保持数据完整性。
- 界面针对从右到左的文本进行了优化，带有波斯标签，并支持Tahoma字体以提高可读性。
- 日志配置为在控制台中跟踪数据库连接和关键操作。

## 许可证
MIT许可证