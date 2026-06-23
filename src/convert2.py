import sys
import os
import re
from datetime import datetime, date, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd

USE_FIXED_ADDRESS = True
FIXED_ADDRESS = "Факультет Социальных Технологий ул. Черняховского, 6/10, Санкт-Петербург, Россия, 191119"

def normalize_group_name(name):
    if not name:
        return ""
    normalized = re.sub(r'[^A-Za-zА-Яа-я0-9]', '', str(name))
    return normalized.upper()

def expand_group_list(group_str):
    parts = [p.strip() for p in re.split(r'[,;]+', group_str) if p.strip()]
    if not parts:
        return []
    if len(parts) == 1:
        return parts
    first = parts[0]
    match = re.match(r'^(.*?)(\d+)$', first)
    if match:
        prefix = match.group(1)
        result = [first]
        for p in parts[1:]:
            if p.isdigit():
                result.append(prefix + p)
            else:
                result.append(p)
        seen = set()
        unique = []
        for g in result:
            if g not in seen:
                seen.add(g)
                unique.append(g)
        return unique
    else:
        return parts

def is_valid_group(group_str):
    """
    Возвращает True, если строка - реальная группа (не перечисление, не составной номер).
    Критерии:
    - Нет запятых/точек с запятой.
    - При разбиении по дефису получается ровно 4 части (формат БУКВЫ-цифра-цифра-две цифры).
    """
    if re.search(r'[,;]', group_str):
        return False
    parts = group_str.split('-')
    # Пример: "МК-3-25-01" -> 4 части. "МК-3-25-01-03" -> 5 частей.
    return len(parts) == 4

def extract_all_groups(df):
    groups = set()
    for val in df["Группа"].dropna():
        s = str(val).strip()
        if is_valid_group(s):
            groups.add(s)
    return sorted(groups)

def filter_by_group(df, target_group):
    target_norm = normalize_group_name(target_group)
    mask = []
    for val in df["Группа"]:
        if pd.isna(val):
            mask.append(False)
            continue
        groups = expand_group_list(str(val))
        found = any(normalize_group_name(g) == target_norm for g in groups)
        mask.append(found)
    return df[mask].copy()

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

def process_excel(excel_path, target_group, output_path, progress_callback=None):
    try:
        ext = os.path.splitext(excel_path)[1].lower()
        if ext not in ('.xls', '.xlsx'):
            raise Exception("Неподдерживаемый формат. Используйте .xls или .xlsx")

        if progress_callback:
            progress_callback(20, "Чтение файла Excel...")
        try:
            df = pd.read_excel(excel_path, engine='calamine')
        except ImportError:
            if ext == '.xlsx':
                df = pd.read_excel(excel_path, engine='openpyxl')
            else:
                raise Exception("Для чтения .xls требуется установить python-calamine: py -m pip install python-calamine")
        
        df = df.dropna(how='all')

        if progress_callback:
            progress_callback(30, "Фильтрация группы")
        df_filtered = filter_by_group(df, target_group)
        if df_filtered.empty:
            raise Exception(f"Не найдено строк для группы {target_group}")

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
    try:
        ext = os.path.splitext(file_path)[1].lower()
        if ext == '.xls':
            try:
                df = pd.read_excel(file_path, engine='calamine')
            except ImportError:
                print("Для чтения .xls требуется установить python-calamine: py -m pip install python-calamine")
                input("Нажмите Enter для выхода")
                return
        else:
            df = pd.read_excel(file_path, engine='openpyxl')
        groups = extract_all_groups(df)
        if not groups:
            print("В файле не найдено групп.")
            input("Нажмите Enter для выхода")
            return
        print("Доступные группы:")
        for i, g in enumerate(groups, 1):
            print(f"{i}. {g}")
        choice = input("Выберите номер группы: ").strip()
        try:
            idx = int(choice) - 1
            target_group = groups[idx]
        except:
            print("Неверный выбор.")
            input("Нажмите Enter для выхода")
            return
        default_csv = os.path.join(os.path.dirname(file_path), "google_calendar_import.csv")
        save_path = input(f"Путь для сохранения CSV (Enter для '{default_csv}'): ").strip()
        if not save_path:
            save_path = default_csv
        def progress(p, msg):
            print(f"{msg}... {p}%")
        count = process_excel(file_path, target_group, save_path, progress_callback=progress)
        print(f"Успешно! Экспортировано {count} событий в {save_path}")
    except Exception as e:
        print(f"Ошибка: {str(e)}")
    input("Нажмите Enter для выхода")

def show_group_selection(groups, callback):
    win = tk.Toplevel()
    win.title("Выбор группы")
    win.geometry("450x350")
    win.resizable(False, False)
    win.configure(bg='#f0f0f0')
    win.transient()
    win.grab_set()
    
    tk.Label(win, text="Выберите вашу группу из списка:", font=('Segoe UI', 11, 'bold'), bg='#f0f0f0').pack(pady=(15,5))
    
    frame_list = tk.Frame(win, bg='#f0f0f0')
    frame_list.pack(fill=tk.BOTH, expand=True, padx=20, pady=5)
    
    scrollbar = tk.Scrollbar(frame_list)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    listbox = tk.Listbox(frame_list, font=('Segoe UI', 10), yscrollcommand=scrollbar.set, height=12)
    listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
    scrollbar.config(command=listbox.yview)
    
    for g in groups:
        listbox.insert(tk.END, g)
    
    button_frame = tk.Frame(win, bg='#f0f0f0')
    button_frame.pack(pady=15)
    
    def on_ok():
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("Внимание", "Выберите группу из списка")
            return
        target = groups[sel[0]]
        win.destroy()
        callback(target)
    
    def on_cancel():
        win.destroy()
        callback(None)
    
    def on_double_click(event):
        on_ok()
    
    listbox.bind("<Double-Button-1>", on_double_click)
    
    tk.Button(button_frame, text="Выбрать", command=on_ok, bg='#c0c0c0', font=('Segoe UI', 10), width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Отмена", command=on_cancel, bg='#e0e0e0', font=('Segoe UI', 10), width=12).pack(side=tk.LEFT, padx=5)
    
    win.wait_window()
    return

def gui_mode():
    root = tk.Tk()
    root.title("Calendar Converter version 2.9")
    root.geometry("650x400")
    root.resizable(False, False)
    root.configure(bg='#f0f0f0')

    main_frame = tk.Frame(root, bg='#f0f0f0', padx=25, pady=15)
    main_frame.pack(fill=tk.BOTH, expand=True)

    header_frame = tk.Frame(main_frame, bg='#f0f0f0')
    header_frame.grid(row=0, column=0, columnspan=3, sticky='ew', pady=(0,15))
    tk.Label(header_frame, text="Загрузи расписание для Google Calendar!",
             font=('Segoe UI', 15, 'bold'), bg='#f0f0f0', fg='#333').pack(side=tk.LEFT)
    def show_about():
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
    about_btn = tk.Button(header_frame, text="ⓘ", font=('Segoe UI', 12), bg='#f0f0f0', fg='#555', relief=tk.FLAT,
                          command=show_about)
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

    group_label_var = tk.StringVar()
    group_label_var.set("Группа не выбрана")
    group_label = tk.Label(main_frame, textvariable=group_label_var, font=('Segoe UI', 9, 'italic'), bg='#f0f0f0', fg='#555')
    group_label.grid(row=3, column=0, columnspan=2, sticky='w', pady=(5,0))

    def select_file():
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xls *.xlsx")])
        if p:
            file_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
            file_path_var.set(p)
            if hasattr(root, 'target_group'):
                del root.target_group
            group_label_var.set("Группа не выбрана")
            try:
                ext = os.path.splitext(p)[1].lower()
                if ext == '.xls':
                    try:
                        df = pd.read_excel(p, engine='calamine')
                    except ImportError:
                        messagebox.showerror("Ошибка", "Для чтения .xls требуется установить python-calamine. Выполните: py -m pip install python-calamine")
                        return
                else:
                    df = pd.read_excel(p, engine='openpyxl')
                groups = extract_all_groups(df)
                if not groups:
                    messagebox.showerror("Ошибка", "В файле не найдено групп.")
                    return
                def on_group_selected(selected_group):
                    if selected_group is None:
                        group_label_var.set("Группа не выбрана (отменено)")
                        if hasattr(root, 'target_group'):
                            del root.target_group
                    else:
                        root.target_group = selected_group
                        group_label_var.set(f"Выбрана группа: {selected_group}")
                        default_save = os.path.join(os.path.dirname(p), f"{selected_group}_calendar.csv")
                        save_path_var.set(default_save)
                        save_entry.config(fg='black', font=('Segoe UI', 10, 'normal'))
                show_group_selection(groups, on_group_selected)
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось прочитать файл:\n{str(e)}")
    tk.Button(main_frame, text="Обзор", command=select_file, bg='#e0e0e0', fg='#333', font=('Segoe UI', 10), width=10).grid(row=2, column=1, sticky='w')

    tk.Label(main_frame, text="Выберите место сохранения:", font=('Segoe UI', 11), bg='#f0f0f0', anchor='w').grid(row=4, column=0, sticky='w', pady=(12,5))
    save_path_var = tk.StringVar()
    save_entry = tk.Entry(main_frame, textvariable=save_path_var, font=('Segoe UI', 10), width=50, bg='white')
    save_entry.grid(row=5, column=0, padx=(0,10), sticky='ew')
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
    tk.Button(main_frame, text="Сохранить как...", command=set_save_file, bg='#e0e0e0', fg='#333', font=('Segoe UI', 10), width=14).grid(row=5, column=1, sticky='w')

    tk.Label(main_frame, text="Прогресс", font=('Segoe UI', 11), bg='#f0f0f0', anchor='w').grid(row=6, column=0, sticky='w', pady=(15,5))
    progress_var = tk.IntVar()
    progress_bar = ttk.Progressbar(main_frame, variable=progress_var, maximum=100, length=470)
    progress_bar.grid(row=7, column=0, columnspan=2, sticky='ew', pady=(0,5))
    status_label = tk.Label(main_frame, text="Готов к работе", font=('Segoe UI', 9, 'italic'), bg='#f0f0f0', fg='#555')
    status_label.grid(row=8, column=0, columnspan=2, sticky='w')

    def run():
        in_path = file_path_var.get()
        if not in_path or in_path == "Обзор":
            messagebox.showwarning("Внимание", "Выберите файл расписания")
            return
        if not hasattr(root, 'target_group') or not root.target_group:
            messagebox.showwarning("Внимание", "Сначала выберите группу (нажмите Обзор и выберите группу)")
            return
        out_path = save_path_var.get()
        if not out_path or out_path == "Обзор":
            default = os.path.join(os.path.dirname(in_path), f"{root.target_group}_calendar.csv")
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
            count = process_excel(in_path, root.target_group, out_path, progress_callback=progress_cb)
            progress_var.set(100)
            status_label.config(text="Готово!")
            messagebox.showinfo("Успех", f"Экспортировано {count} событий\nФайл: {out_path}")
        except Exception as e:
            progress_var.set(0)
            status_label.config(text="Ошибка")
            show_error_window(str(e))

    btn = tk.Button(main_frame, text="Конвертировать", command=run, bg='#c0c0c0', fg='#000', font=('Segoe UI', 11), width=22, height=1)
    btn.grid(row=9, column=0, columnspan=2, pady=10)

    copyright_label = tk.Label(main_frame, text="© Евгений Резчиков, 2026 — Abricos Prod.",
                               font=('Segoe UI', 8), bg='#f0f0f0', fg='#888')
    copyright_label.grid(row=10, column=0, columnspan=2, sticky='w', pady=(0,0))

    main_frame.columnconfigure(0, weight=1)
    main_frame.columnconfigure(1, weight=0)
    root.mainloop()

if __name__ == "__main__":
    if len(sys.argv) > 1:
        console_mode(sys.argv[1])
    else:
        gui_mode()