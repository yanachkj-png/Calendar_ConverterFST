import sys
import os
import tempfile
import re
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

USE_FIXED_ADDRESS = True
FIXED_ADDRESS = "Факультет Социальных Технологий ул. Черняховского, 6/10, Санкт-Петербург, Россия, 191119"
TARGET_GROUP = "ЖУР-3-24-03"

def normalize_group_string(s):
    if pd.isna(s):
        return ""
    s = str(s)
    s = re.sub(r'\s+', '', s)
    return s

def group_contains_target(group_str, target):
    if pd.isna(group_str):
        return False
    cleaned = normalize_group_string(group_str)
    parts = re.split(r'[,]+', cleaned)
    for p in parts:
        if p == normalize_group_string(target):
            return True
    return False

def parse_time_range(time_str, event_type):
    if pd.isna(time_str) or time_str == "-" or str(time_str).strip() == "":
        return None, None
    s = str(time_str).strip()
    if "-" in s and "_" not in s:
        parts = s.split("-")
        start_raw, end_raw = parts[0], parts[1]
        start = start_raw.replace(".", ":").strip()
        end = end_raw.replace(".", ":").strip()
        start_time = datetime.strptime(start, "%H:%M").time()
        end_time = datetime.strptime(end, "%H:%M").time()
        return start_time, end_time
    else:
        single = s.replace("_", ":").replace(".", ":").strip()
        start_time = datetime.strptime(single, "%H:%M").time()
        if event_type == "Э":
            delta = timedelta(hours=2)
        else:
            delta = timedelta(hours=1)
        end_dt = datetime.combine(date.today(), start_time) + delta
        end_time = end_dt.time()
        return start_time, end_time

def extract_multiday_dates(text):
    pattern = r"\((\d{2})\.(\d{2})\.(\d{4})-(\d{2})\.(\d{2})\.(\d{4})\)"
    match = re.search(pattern, text)
    if match:
        s_day, s_month, s_year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        e_day, e_month, e_year = int(match.group(4)), int(match.group(5)), int(match.group(6))
        return date(s_year, s_month, s_day), date(e_year, e_month, e_day)
    return None, None

def get_event_year(month):
    now = datetime.now()
    current_year = now.year
    current_month = now.month
    if month < current_month:
        return current_year + 1
    else:
        return current_year

def convert_xls_to_xlsx(xls_path, progress_callback=None):
    try:
        import win32com.client as win32
    except ImportError:
        raise Exception("Для конвертации .xls требуется pywin32. Установите: py -m pip install pywin32")
    try:
        excel = win32.gencache.EnsureDispatch('Excel.Application')
    except Exception as e:
        raise Exception(f"Не удалось запустить Excel. Убедитесь, что Microsoft Excel установлен.\nОшибка: {str(e)}")
    excel.Visible = False
    excel.DisplayAlerts = False
    try:
        excel.FileValidation = 2
    except:
        pass
    try:
        if progress_callback:
            progress_callback(5, "Открытие Excel...")
        wb = excel.Workbooks.Open(
            xls_path,
            ReadOnly=False,
            Notify=False,
            IgnoreReadOnlyRecommended=True,
            Local=True
        )
        fd, temp_xlsx = tempfile.mkstemp(suffix='.xlsx')
        os.close(fd)
        if progress_callback:
            progress_callback(10, "Сохранение в .xlsx...")
        wb.SaveAs(temp_xlsx, FileFormat=51)
        wb.Close()
        excel.Quit()
        return temp_xlsx
    except Exception as e:
        excel.Quit()
        raise Exception(f"Ошибка при открытии .xls файла. Возможно, файл в защищённом просмотре. Откройте вручную, разрешите редактирование и повторите.\nОшибка: {str(e)}")

def process_excel(excel_path, output_path, progress_callback=None):
    temp_file = None
    try:
        ext = os.path.splitext(excel_path)[1].lower()
        if ext == '.xls':
            if progress_callback:
                progress_callback(5, "Конвертация .xls -> .xlsx")
            temp_file = convert_xls_to_xlsx(excel_path, progress_callback)
            read_path = temp_file
        elif ext == '.xlsx':
            read_path = excel_path
        else:
            raise Exception("Неподдерживаемый формат. Используйте .xls или .xlsx")

        if progress_callback:
            progress_callback(20, "Чтение файла Excel")
        df = pd.read_excel(read_path, engine='openpyxl')
        if temp_file and os.path.exists(temp_file):
            os.unlink(temp_file)
        df = df.dropna(how='all')

        if progress_callback:
            progress_callback(30, "Фильтрация группы")
        mask = df["Группа"].apply(lambda x: group_contains_target(x, TARGET_GROUP))
        df_filtered = df[mask].copy()
        if df_filtered.empty:
            raise Exception(f"Не найдено строк для группы {TARGET_GROUP}")

        events = []
        total = len(df_filtered)
        for idx, (_, row) in enumerate(df_filtered.iterrows()):
            if progress_callback:
                progress_callback(30 + int(40 * idx / total), f"Обработка строки {idx+1}/{total}")
            day = int(row["День"])
            month = int(row["Месяц"])
            discipline = str(row["Дисциплина"]) if not pd.isna(row["Дисциплина"]) else ""
            teacher = str(row["Преподаватель"]) if not pd.isna(row["Преподаватель"]) else ""
            room = str(row["Ауд"]) if not pd.isna(row["Ауд"]) else ""
            event_type = str(row["Тип"]) if not pd.isna(row["Тип"]) else ""
            if event_type == "-":
                continue

            event_type = event_type.strip()

            if event_type == "Л":
                subject = f"Л: {discipline}"
            elif event_type == "Пз":
                subject = f"ПЗ: {discipline}"
            else:
                subject = discipline

            if room == "СДО":
                location = "СДО"
            elif USE_FIXED_ADDRESS and room != "":
                location = FIXED_ADDRESS
            else:
                location = ""

            description = f"Преподаватель: {teacher}" if teacher else ""

            multiday_start, multiday_end = extract_multiday_dates(discipline)
            if multiday_start and multiday_end and event_type == "П":
                events.append({
                    "Subject": subject,
                    "Start Date": multiday_start.strftime("%Y-%m-%d"),
                    "Start Time": "",
                    "End Date": multiday_end.strftime("%Y-%m-%d"),
                    "End Time": "",
                    "All Day Event": "True",
                    "Description": description,
                    "Location": location
                })
                continue

            year = get_event_year(month)
            try:
                event_date = date(year, month, day)
            except ValueError:
                raise Exception(f"Ошибка даты: {day}.{month}.{year}")

            time_raw = row["Время_xl"]
            start_time, end_time = parse_time_range(time_raw, event_type)
            start_date_str = event_date.strftime("%Y-%m-%d")
            end_date_str = start_date_str

            if start_time is None:
                events.append({
                    "Subject": subject,
                    "Start Date": start_date_str,
                    "Start Time": "",
                    "End Date": start_date_str,
                    "End Time": "",
                    "All Day Event": "True",
                    "Description": description,
                    "Location": location
                })
            else:
                events.append({
                    "Subject": subject,
                    "Start Date": start_date_str,
                    "Start Time": start_time.strftime("%H:%M"),
                    "End Date": end_date_str,
                    "End Time": end_time.strftime("%H:%M"),
                    "All Day Event": "False",
                    "Description": description,
                    "Location": location
                })

        if progress_callback:
            progress_callback(90, "Сохранение CSV")
        output_df = pd.DataFrame(events)
        output_df.to_csv(output_path, index=False, encoding='utf-8-sig')
        if progress_callback:
            progress_callback(100, "Готово")
        return len(events)
    except Exception as e:
        if temp_file and os.path.exists(temp_file):
            try: os.unlink(temp_file)
            except: pass
        raise e

def show_error_window(error_text):
    err_root = tk.Tk()
    err_root.title("Ошибка конвертации")
    err_root.geometry("600x400")
    err_root.resizable(False, False)
    frame = tk.Frame(err_root, padx=10, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)
    tk.Label(frame, text="Произошла ошибка:", fg="red", font=('Segoe UI', 10, 'bold')).pack(anchor='w')
    text_area = tk.Text(frame, wrap=tk.WORD, font=('Consolas', 9))
    text_area.pack(fill=tk.BOTH, expand=True, pady=5)
    text_area.insert(tk.END, error_text)
    text_area.config(state=tk.DISABLED)
    def copy_text():
        err_root.clipboard_clear()
        err_root.clipboard_append(error_text)
        messagebox.showinfo("Скопировано", "Текст ошибки скопирован")
    btn_frame = tk.Frame(frame)
    btn_frame.pack(pady=5)
    tk.Button(btn_frame, text="Копировать ошибку", command=copy_text, bg='#e0e0e0', padx=10).pack(side=tk.LEFT, padx=5)
    tk.Button(btn_frame, text="Закрыть", command=err_root.destroy, padx=10).pack(side=tk.LEFT)
    err_root.mainloop()

def console_mode(file_path):
    print(f"Обработка файла: {file_path}")
    default_csv = os.path.join(os.path.dirname(file_path), "google_calendar_import.csv")
    save_path = input(f"Путь для сохранения CSV (Enter для '{default_csv}'): ").strip()
    if not save_path:
        save_path = default_csv
    try:
        def progress(p, msg):
            print(f"{msg}... {p}%")
        count = process_excel(file_path, save_path, progress_callback=progress)
        print(f"Успешно! Экспортировано {count} событий в {save_path}")
    except Exception as e:
        print(f"Ошибка: {str(e)}")
        input("Нажмите Enter для выхода")

def show_about_window():
    about = tk.Toplevel()
    about.title("О программе")
    about.geometry("300x150")
    about.resizable(False, False)
    about.configure(bg='#f0f0f0')
    about.transient()
    about.grab_set()
    frame = tk.Frame(about, bg='#f0f0f0', padx=20, pady=20)
    frame.pack(fill=tk.BOTH, expand=True)
    text = "Евгений Резчиков, 2026\n\ntelegram @goatunheim\n\n24/7"
    label = tk.Label(frame, text=text, font=('Segoe UI', 10), bg='#f0f0f0', fg='#333', justify='center')
    label.pack(expand=True)

def gui_mode():
    root = tk.Tk()
    root.title("Calendar Converter version 1.0")
    root.geometry("650x380")
    root.resizable(False, False)
    root.configure(bg='#f0f0f0')

    main_frame = tk.Frame(root, bg='#f0f0f0', padx=25, pady=15)
    main_frame.pack(fill=tk.BOTH, expand=True)

    header_frame = tk.Frame(main_frame, bg='#f0f0f0')
    header_frame.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0,15))
    tk.Label(header_frame, text="Загрузи расписание для Google Calendar!",
             font=('Segoe UI', 15, 'bold'), bg='#f0f0f0', fg='#333').pack(side=tk.LEFT)
    about_btn = tk.Button(header_frame, text="ⓘ", font=('Segoe UI', 12), bg='#f0f0f0', fg='#555', relief=tk.FLAT,
                          command=show_about_window)
    about_btn.pack(side=tk.RIGHT)

    tk.Label(main_frame, text="Файл расписания:", font=('Segoe UI', 11), bg='#f0f0f0', anchor='w').grid(row=1, column=0, sticky='w', pady=(0,5))
    file_path_var = tk.StringVar()
    file_entry = tk.Entry(main_frame, textvariable=file_path_var, font=('Segoe UI', 10), width=50, bg='white')
    file_entry.grid(row=2, column=0, padx=(0,10), sticky='ew')
    def on_focus_in_file(event):
        if file_path_var.get() == "Обзор":
            file_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
            file_path_var.set("")
    def on_focus_out_file(event):
        if not file_path_var.get():
            file_entry.config(fg='gray', font=('Segoe UI', 10, 'italic'))
            file_path_var.set("Обзор")
    file_path_var.set("Обзор")
    file_entry.config(fg='gray', font=('Segoe UI', 10, 'italic'))
    file_entry.bind("<FocusIn>", on_focus_in_file)
    file_entry.bind("<FocusOut>", on_focus_out_file)

    def select_file():
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if p:
            file_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
            file_path_var.set(p)
            default = os.path.join(os.path.dirname(p), "google_calendar_import.csv")
            save_path_var.set(default)
            save_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
    tk.Button(main_frame, text="Обзор", command=select_file, bg='#e0e0e0', fg='#333', font=('Segoe UI', 10), width=10).grid(row=2, column=1, sticky='w')

    tk.Label(main_frame, text="Выберите место сохранения:", font=('Segoe UI', 11), bg='#f0f0f0', anchor='w').grid(row=3, column=0, sticky='w', pady=(12,5))
    save_path_var = tk.StringVar()
    save_entry = tk.Entry(main_frame, textvariable=save_path_var, font=('Segoe UI', 10), width=50, bg='white')
    save_entry.grid(row=4, column=0, padx=(0,10), sticky='ew')
    def on_focus_in_save(event):
        if save_path_var.get() == "Обзор":
            save_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
            save_path_var.set("")
    def on_focus_out_save(event):
        if not save_path_var.get():
            save_entry.config(fg='gray', font=('Segoe UI', 10, 'italic'))
            save_path_var.set("Обзор")
    save_path_var.set("Обзор")
    save_entry.config(fg='gray', font=('Segoe UI', 10, 'italic'))
    save_entry.bind("<FocusIn>", on_focus_in_save)
    save_entry.bind("<FocusOut>", on_focus_out_save)

    def set_save_file():
        init = save_path_var.get()
        if init == "Обзор":
            init = ""
        p = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")],
                                         initialfile=os.path.basename(init) if init else "google_calendar_import.csv",
                                         initialdir=os.path.dirname(init) if init else "")
        if p:
            save_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
            save_path_var.set(p)
    tk.Button(main_frame, text="Сохранить как...", command=set_save_file, bg='#e0e0e0', fg='#333', font=('Segoe UI', 10), width=14).grid(row=4, column=1, sticky='w')

    tk.Label(main_frame, text="Прогресс", font=('Segoe UI', 11), bg='#f0f0f0', anchor='w').grid(row=5, column=0, sticky='w', pady=(15,5))
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(main_frame, variable=progress_var, maximum=100, length=470)
    progress_bar.grid(row=6, column=0, columnspan=2, sticky='ew', pady=(0,5))
    status_label = tk.Label(main_frame, text="Готов к работе", font=('Segoe UI', 9, 'italic'), bg='#f0f0f0', fg='#555')
    status_label.grid(row=7, column=0, columnspan=2, sticky='w')

    def run():
        in_path = file_path_var.get()
        if not in_path or in_path == "Обзор":
            messagebox.showwarning("Внимание", "Выберите файл расписания")
            return
        out_path = save_path_var.get()
        if not out_path or out_path == "Обзор":
            default = os.path.join(os.path.dirname(in_path), "google_calendar_import.csv")
            out_path = default
            save_path_var.set(out_path)
            save_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
        progress_var.set(0)
        status_label.config(text="Начинаем...")
        root.update()

        def progress_cb(val, msg):
            progress_var.set(val)
            status_label.config(text=msg)
            root.update()
        try:
            count = process_excel(in_path, out_path, progress_callback=progress_cb)
            progress_var.set(100)
            status_label.config(text="Готово!")
            messagebox.showinfo("Успех", f"Экспортировано {count} событий\nФайл: {out_path}")
        except Exception as e:
            progress_var.set(0)
            status_label.config(text="Ошибка")
            show_error_window(str(e))

    btn = tk.Button(main_frame, text="Конвертировать", command=run, bg='#c0c0c0', fg='#000', font=('Segoe UI', 11), width=22, height=1)
    btn.grid(row=8, column=0, columnspan=2, pady=10)

    copyright_label = tk.Label(main_frame, text="© Евгений Резчиков, 2026 — Abricos Prod.",
                               font=('Segoe UI', 8), bg='#f0f0f0', fg='#888')
    copyright_label.grid(row=9, column=0, columnspan=2, sticky='w', pady=(0,0))

    main_frame.columnconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=0)
    root.mainloop()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        console_mode(sys.argv[1])
    else:
        gui_mode()