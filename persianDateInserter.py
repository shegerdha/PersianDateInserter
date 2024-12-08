import keyboard
import pyautogui
import jdatetime
import time
import tkinter as tk
from tkinter import ttk
import os
from PIL import Image, ImageTk, ImageDraw
import datetime
import pickle
import sys
import threading
import pystray
import ctypes

# اطمینان از اجرای اسکریپت با دسترسی ادمین برای کتابخانه keyboard
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, __file__, None, 1)
    sys.exit()

def save_last_state(state):
    with open("last_state.pkl", "wb") as f:
        pickle.dump(state, f)

def load_last_state():
    try:
        with open("last_state.pkl", "rb") as f:
            state = pickle.load(f)
            if (datetime.datetime.now() - state['timestamp']).total_seconds() < 600:
                return state
    except (FileNotFoundError, EOFError):
        pass
    return None

def insert_persian_date(date):
    date = date.translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')).translate(str.maketrans('٠١٢٣٤٥٦٧٨٩', '0123456789'))  # تبدیل اعداد فارسی به لاتین
    time.sleep(0.2)  # کاهش تأخیر برای بهبود UX
    pyautogui.write(date)

# متغیر برای پیگیری وضعیت باز یا بسته بودن پنجره
is_window_open = False
window_instance = None  # ذخیره نمونه پنجره برای مدیریت

def show_date_menu():
    global is_window_open, window_instance

    if is_window_open:
        # اگر پنجره باز است، آن را ببند
        if window_instance is not None:
            window_instance.destroy()
            is_window_open = False
            update_systray_menu()
        return

    is_window_open = True
    update_systray_menu()  # به‌روزرسانی منوی System Tray

    last_state = load_last_state()

    root = tk.Tk()
    window_instance = root  # ذخیره نمونه پنجره

    def on_closing():
        global is_window_open, window_instance
        is_window_open = False
        window_instance = None
        root.destroy()
        update_systray_menu()  # به‌روزرسانی منوی System Tray

    root.protocol("WM_DELETE_WINDOW", on_closing)

    root.title("انتخاب تاریخ شمسی")
    root.geometry("400x330")
    root.configure(bg="#f0f0f0")
    root.attributes("-topmost", True)  # همیشه در بالاترین سطح
    root.update()  # به‌روزرسانی پنجره برای اطمینان از رندر شدن
    root.focus_force()  # تمرکز روی پنجره
    root.grab_set()  # اطمینان از اینکه همه ورودی‌ها به این پنجره هدایت می‌شوند

    # تنظیم آیکون پنجره
    current_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(current_dir, 'assets')  # پوشه assets
    icons_dir = os.path.join(assets_dir, 'icons')  # پوشه icons داخل assets

    icon_path = os.path.join(icons_dir, 'systray-icon.png')
    if os.path.exists(icon_path):
        window_icon = tk.PhotoImage(file=icon_path)
        root.iconphoto(False, window_icon)
    else:
        # ایجاد آیکون پیش‌فرض در صورت عدم وجود فایل
        window_icon = tk.PhotoImage(width=1, height=1)
        root.iconphoto(False, window_icon)

    today_jalali = last_state['today_jalali'] if last_state else jdatetime.date.today()
    today_gregorian = last_state['today_gregorian'] if last_state else datetime.date.today()
    date_var = tk.StringVar()
    date_var.set(today_jalali.strftime('%Y-%m-%d').translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')))
    day_of_week_var = tk.StringVar()
    day_of_week_var.set(today_jalali.strftime('%A'))
    gregorian_date_var = tk.StringVar()
    gregorian_date_var.set(today_gregorian.strftime('%Y-%m-%d'))
    formatted_gregorian_var = tk.StringVar()
    formatted_gregorian_var.set(today_gregorian.strftime('%d/%b/%y'))

    def update_date(delta_days):
        nonlocal today_jalali, today_gregorian
        today_jalali += jdatetime.timedelta(days=delta_days)
        today_gregorian += datetime.timedelta(days=delta_days)
        date_var.set(today_jalali.strftime('%Y-%m-%d').translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')))
        day_of_week_var.set(today_jalali.strftime('%A'))
        gregorian_date_var.set(today_gregorian.strftime('%Y-%m-%d'))
        formatted_gregorian_var.set(today_gregorian.strftime('%d/%b/%y'))

    def on_confirm():
        root.destroy()
        insert_persian_date(date_var.get())
        save_last_state({
            'today_jalali': today_jalali,
            'today_gregorian': today_gregorian,
            'timestamp': datetime.datetime.now()
        })
        global is_window_open, window_instance
        is_window_open = False
        window_instance = None
        update_systray_menu()  # به‌روزرسانی منوی System Tray

    # استایل‌دهی
    style = ttk.Style()
    style.configure("TButton", font=("IRANYekanX", 12), padding=5)
    style.configure("TLabel", font=("IRANYekanX", 14), background="#f0f0f0")
    style.configure("TCombobox", font=("IRANYekanX", 12))

    # لیست کشویی تاریخ میلادی
    gregorian_date_label = ttk.Label(root, text="تاریخ میلادی:")
    gregorian_date_label.pack(pady=5)
    gregorian_date_dropdown = ttk.Combobox(root, textvariable=gregorian_date_var, state="readonly", width=20)
    gregorian_date_dropdown['values'] = [(today_gregorian + datetime.timedelta(days=i)).strftime('%Y-%m-%d') for i in range(-30, 31)]
    gregorian_date_dropdown.pack(pady=5)

    def on_gregorian_date_change(event):
        nonlocal today_jalali, today_gregorian
        selected_date = gregorian_date_var.get()
        today_gregorian = datetime.datetime.strptime(selected_date, '%Y-%m-%d').date()
        today_jalali = jdatetime.date.fromgregorian(date=today_gregorian)
        date_var.set(today_jalali.strftime('%Y-%m-%d').translate(str.maketrans('۰۱۲۳۴۵۶۷۸۹', '0123456789')))
        day_of_week_var.set(today_jalali.strftime('%A'))
        formatted_gregorian_var.set(today_gregorian.strftime('%d/%b/%y'))

    gregorian_date_dropdown.bind('<<ComboboxSelected>>', on_gregorian_date_change)

    # لیبل تاریخ میلادی قالب‌بندی‌شده
    formatted_gregorian_label = ttk.Label(root, textvariable=formatted_gregorian_var)
    formatted_gregorian_label.pack(pady=5)

    # لیبل تاریخ شمسی
    date_label = ttk.Label(root, textvariable=date_var)
    date_label.pack(pady=5)

    # لیبل روز هفته
    day_of_week_label = ttk.Label(root, textvariable=day_of_week_var)
    day_of_week_label.pack(pady=5)

    frame = tk.Frame(root, bg="#f0f0f0")
    frame.pack()

    # بارگذاری آیکون‌ها برای دکمه‌ها از پوشه assets/icons
    try:
        plus_icon_path = os.path.join(icons_dir, "plus_icon.png")
        minus_icon_path = os.path.join(icons_dir, "minus_icon.png")

        plus_icon = ImageTk.PhotoImage(Image.open(plus_icon_path).resize((30, 30)))
        minus_icon = ImageTk.PhotoImage(Image.open(minus_icon_path).resize((30, 30)))
    except FileNotFoundError:
        plus_icon = None
        minus_icon = None

    btn_back = ttk.Button(frame, image=minus_icon if minus_icon else None, text="-" if not minus_icon else "", command=lambda: update_date(-1))
    btn_back.pack(side="left", padx=10)

    btn_forward = ttk.Button(frame, image=plus_icon if plus_icon else None, text="+" if not plus_icon else "", command=lambda: update_date(1))
    btn_forward.pack(side="left", padx=10)

    btn_confirm = ttk.Button(root, text="درج تاریخ", command=on_confirm)
    btn_confirm.pack(pady=10)

    root.bind('<Return>', lambda event: on_confirm())  # اتصال کلید Enter به عملکرد درج تاریخ

    root.mainloop()

def on_quit_callback(icon, item):
    icon.stop()
    sys.exit()

def update_systray_menu():
    global systray_icon, is_window_open

    # ایجاد منوی جدید با توجه به وضعیت پنجره
    if is_window_open:
        open_close_text = 'بستن'
    else:
        open_close_text = 'باز کردن'

    systray_icon.menu = pystray.Menu(
        pystray.MenuItem(open_close_text, on_open_close_callback),
        pystray.MenuItem('خروج', on_quit_callback)
    )
    systray_icon.update_menu()

def on_open_close_callback(icon, item):
    # عملکرد باز یا بستن پنجره
    show_date_menu()

def main():
    global systray_icon

    # اجرای Listener هات‌کی در Thread جداگانه
    keyboard_thread = threading.Thread(target=lambda: keyboard.add_hotkey('ctrl+alt+d', show_date_menu), daemon=True)
    keyboard_thread.start()

    # اطمینان از وجود آیکون یا ایجاد یک آیکون پیش‌فرض
    current_dir = os.path.dirname(os.path.abspath(__file__))
    assets_dir = os.path.join(current_dir, 'assets')  # پوشه assets
    icons_dir = os.path.join(assets_dir, 'icons')  # پوشه icons داخل assets

    icon_path = os.path.join(icons_dir, 'systray-icon.png')
    if os.path.exists(icon_path):
        image = Image.open(icon_path)
    else:
        # ایجاد یک آیکون ساده در صورت عدم وجود systray-icon.png
        image = Image.new('RGB', (64, 64), color=(255, 0, 0))
        draw = ImageDraw.Draw(image)
        draw.rectangle([16, 16, 48, 48], fill=(255, 255, 255))

    # ایجاد آیکون System Tray
    systray_icon = pystray.Icon('Persian Date Picker', image, 'انتخاب تاریخ شمسی')
    update_systray_menu()  # تنظیم منوی اولیه

    # اجرای آیکون System Tray
    systray_icon.run()

if __name__ == '__main__':
    print("The program is running... Press Ctrl+Alt+D to open the Persian date menu.")
    main()
