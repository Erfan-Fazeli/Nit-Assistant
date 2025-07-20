import tkinter as tk
from tkinter import messagebox, filedialog
import json, os, threading, time, psutil, pyautogui, pystray
from PIL import Image, ImageDraw
from pathlib import Path
from enum import Enum
from typing import Optional, Callable

# ایمپورت کتابخانه جدید و تنظیم تم
import customtkinter

try:
    from win32gui import GetForegroundWindow, GetWindowText
    from win32process import GetWindowThreadProcessId
    import win32con, ctypes
    from ctypes import wintypes
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

class Config:
    """پیکربندی مرکزی برنامه."""
    class App:
        NAME, VERSION = "NIT Group Personal Assistant", "Final v3.2" # نسخه آپدیت شد
        GEOMETRY_COLLAPSED, GEOMETRY_EXPANDED = "380x300", "380x500"
        FOLDER_NAME, SETTINGS_FILE = "PersonalAssistant", "config.json"
        # <--- رنگ سرمه‌ای پس‌زمینه برای استفاده در هدر
        BACKGROUND_COLOR = "#0F172A"

    class UI:
        class Theme:
            PRIMARY = "#27ae60"
            ACCENT_RED = "#e74c3c"
            ACCENT_PURPLE = "#9b59b6"
            WARNING = "#f39c12"
            INFO = "#3498db"

class LogLevel(Enum):
    SUCCESS = (Config.UI.Theme.PRIMARY, "✅"); WARNING = (Config.UI.Theme.WARNING, "⚠️"); ERROR   = (Config.UI.Theme.ACCENT_RED, "❌"); INFO    = (Config.UI.Theme.INFO, "ℹ️"); STARTUP = (Config.UI.Theme.ACCENT_PURPLE, "🚀"); ACTIVE  = ("#e67e22", "🎯"); SAVE    = ("#2ecc71", "💾")

class CustomSpinbox(customtkinter.CTkFrame):
    def __init__(self, parent, from_=1, to=60, textvariable=None, command=None):
        super().__init__(parent, fg_color="transparent")
        
        self.textvariable = textvariable
        self.from_ = from_
        self.to = to
        self.command = command

        self.grid_columnconfigure(0, weight=1)

        self.entry = customtkinter.CTkEntry(self, textvariable=self.textvariable, width=50, justify='center')
        self.entry.grid(row=0, column=0, rowspan=2, padx=(0, 5), sticky="ew")

        # <--- تغییر ۱: کوچک‌تر کردن دکمه‌ها و فونت آنها
        button_font = customtkinter.CTkFont(size=10)
        up_button = customtkinter.CTkButton(self, text="▲", width=18, height=12, font=button_font, command=self._increment)
        up_button.grid(row=0, column=1)
        
        down_button = customtkinter.CTkButton(self, text="▼", width=18, height=12, font=button_font, command=self._decrement)
        down_button.grid(row=1, column=1)

    def _increment(self):
        try:
            current_value = int(self.textvariable.get())
            self.textvariable.set(min(current_value + 1, self.to))
            if self.command: self.command()
        except ValueError:
            self.textvariable.set(self.from_)

    def _decrement(self):
        try:
            current_value = int(self.textvariable.get())
            self.textvariable.set(max(current_value - 1, self.from_))
            if self.command: self.command()
        except ValueError:
            self.textvariable.set(self.from_)
            
    def configure(self, state=None):
        self.entry.configure(state=state)

class AppUI(customtkinter.CTk):
    instance = None
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        AppUI.instance = self
        
        # <--- رنگ پس‌زمینه از کانفیگ خوانده می‌شود
        self.configure(fg_color=Config.App.BACKGROUND_COLOR)

        self.script = AutoSaveScript(self.add_log, self.update_status)
        self.tray_manager = TrayManager(self._schedule_show_from_tray, self._schedule_quit_application)
        self.window_hidden = False
        self.log_expanded = True

        self._configure_root_style()
        self._create_widgets()
        
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.script.start()
        self.update_status("Initializing...", "INITIALIZING")
    
    def _configure_root_style(self):
        self.title(Config.App.NAME)
        self.geometry(Config.App.GEOMETRY_EXPANDED)
        self.resizable(False, False)

    def _create_widgets(self):
        self.container = customtkinter.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=20)

        self._create_header(self.container)
        self._create_control_panel(self.container)
        self._create_log_section(self.container)
        self._create_footer(self.container)
    
    def _create_header(self, parent):
        header = customtkinter.CTkFrame(parent, height=55, corner_radius=10, fg_color=Config.UI.Theme.PRIMARY)
        header.pack(fill="x", pady=(0, 20))
        header.pack_propagate(False)
        
        # <--- تغییر ۲: استایل جدید فونت‌ها و رنگ هدر
        title_font = customtkinter.CTkFont(family="Fixedsys", size=14, weight="bold")
        version_font = customtkinter.CTkFont(family="Fixedsys", size=8)
        header_text_color = Config.App.BACKGROUND_COLOR # رنگ سرمه‌ای

        customtkinter.CTkLabel(header, text=Config.App.NAME, font=title_font, text_color=header_text_color).pack(pady=(7,0))
        customtkinter.CTkLabel(header, text=f"v{Config.App.VERSION}", font=version_font, text_color=header_text_color).pack()

    def _create_control_panel(self, parent):
        card = customtkinter.CTkFrame(parent)
        card.pack(fill="x", pady=(0, 15))
        
        content = customtkinter.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=15, pady=12)
        
        customtkinter.CTkLabel(content, text="📊 System Status", font=customtkinter.CTkFont(size=12, weight="bold")).pack(anchor="w")
        
        status_fr = customtkinter.CTkFrame(content, fg_color="transparent")
        status_fr.pack(fill="x", pady=(10, 5))
        
        customtkinter.CTkLabel(status_fr, text="Status:", font=customtkinter.CTkFont(size=11, weight="bold")).pack(side="left")
        self.status_label = customtkinter.CTkLabel(status_fr, text="Initializing...")
        self.status_label.pack(side="left", padx=5)
        
        customtkinter.CTkButton(status_fr, text="⚙️ Settings", width=100, command=self._open_settings_dialog).pack(side="right")

    def _create_log_section(self, parent):
        self.log_fr = customtkinter.CTkFrame(parent)
        self.log_fr.pack(fill="both", expand=True)
        
        log_header_fr = customtkinter.CTkFrame(self.log_fr, fg_color="transparent")
        log_header_fr.pack(fill="x", padx=15, pady=(10,5))
        
        self.log_toggle_button = customtkinter.CTkButton(log_header_fr, text="▼", width=20, command=self._toggle_log)
        self.log_toggle_button.pack(side="left", anchor='n')
        customtkinter.CTkLabel(log_header_fr, text="📋 Activity Log", font=customtkinter.CTkFont(size=12, weight="bold")).pack(side="left", padx=10)
        
        self.log_text_container = customtkinter.CTkFrame(self.log_fr, fg_color="transparent")
        self.log_text_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))

        self.log_text = customtkinter.CTkTextbox(self.log_text_container, wrap="none", state="disabled", border_width=0)
        self.log_text.pack(fill="both", expand=True)

    def _create_footer(self, parent):
        self.footer = customtkinter.CTkFrame(parent, fg_color="transparent")
        self.footer.pack(fill="x", side="bottom", pady=(10, 0))
        
        footer_text = "Developed With ❤️ By Erf For NIT Group"
        footer_label = customtkinter.CTkLabel(self.footer, text=footer_text, font=customtkinter.CTkFont(size=9), text_color="gray50")
        footer_label.pack()

    def _open_settings_dialog(self):
        dialog = customtkinter.CTkToplevel(self)
        dialog.title("Settings")
        dialog.geometry("420x480")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        save_enabled_var = tk.BooleanVar(value=self.script.settings['auto_save_enabled'])
        save_interval_var = tk.StringVar(value=str(self.script.settings['auto_save_interval']))
        backup_enabled_var = tk.BooleanVar(value=self.script.settings.get('smart_backup_enabled', False))
        backup_interval_var = tk.StringVar(value=str(self.script.settings.get('smart_backup_interval', 60)))
        startup_var = tk.BooleanVar(value=self.script.settings.get('start_with_windows', False))

        main_frame = customtkinter.CTkFrame(dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        automation_card = customtkinter.CTkFrame(main_frame)
        automation_card.pack(fill="x", pady=(0, 15), ipady=10)
        customtkinter.CTkLabel(automation_card, text="⚡️ Automation", font=customtkinter.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15)
        
        automation_content = customtkinter.CTkFrame(automation_card, fg_color="transparent")
        automation_content.pack(fill="x", padx=15, pady=5)
        automation_content.grid_columnconfigure((0, 1), weight=1)

        save_col = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        save_col.grid(row=0, column=0, padx=(0, 10), sticky="ew")
        customtkinter.CTkLabel(save_col, text="Auto-Save").pack(side="left")
        save_switch = customtkinter.CTkSwitch(save_col, text="", variable=save_enabled_var)
        save_switch.pack(side="right")
        
        save_interval_frame = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        save_interval_frame.grid(row=1, column=0, padx=(0, 10), pady=(5,0), sticky="ew")
        customtkinter.CTkLabel(save_interval_frame, text="Interval (sec):").pack(side="left")
        save_spinbox = CustomSpinbox(save_interval_frame, from_=1, to=60, textvariable=save_interval_var)
        save_spinbox.pack(side="right")

        backup_col = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        backup_col.grid(row=0, column=1, padx=(10, 0), sticky="ew")
        customtkinter.CTkLabel(backup_col, text="Smart Backup").pack(side="left")
        backup_switch = customtkinter.CTkSwitch(backup_col, text="", variable=backup_enabled_var)
        backup_switch.pack(side="right")

        backup_interval_frame = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        backup_interval_frame.grid(row=1, column=1, padx=(10, 0), pady=(5,0), sticky="ew")
        customtkinter.CTkLabel(backup_interval_frame, text="Interval (min):").pack(side="left")
        backup_spinbox = CustomSpinbox(backup_interval_frame, from_=1, to=1440, textvariable=backup_interval_var)
        backup_spinbox.pack(side="right")

        general_card = customtkinter.CTkFrame(main_frame)
        general_card.pack(fill="x", pady=(0, 15), ipady=10)
        customtkinter.CTkLabel(general_card, text="⚙️ General", font=customtkinter.CTkFont(size=12, weight="bold")).pack(anchor="w", padx=15)
        
        startup_frame = customtkinter.CTkFrame(general_card, fg_color="transparent")
        startup_frame.pack(fill="x", padx=15, pady=10)
        customtkinter.CTkLabel(startup_frame, text="Start with Windows").pack(side="left")
        startup_switch = customtkinter.CTkSwitch(startup_frame, text="", variable=startup_var)
        startup_switch.pack(side="right")

        customtkinter.CTkButton(main_frame, text="📂 Add Application...", command=self._on_add_app_browse).pack(fill="x", ipady=4, pady=5)

        def toggle_spinners_state():
            save_spinbox.configure(state='normal' if save_enabled_var.get() else 'disabled')
            backup_spinbox.configure(state='normal' if backup_enabled_var.get() else 'disabled')

        save_switch.configure(command=toggle_spinners_state)
        backup_switch.configure(command=toggle_spinners_state)
        toggle_spinners_state()

        def save_and_close():
            try:
                self.script.update_settings({ 'auto_save_enabled': save_enabled_var.get(), 'auto_save_interval': int(save_interval_var.get()), 'smart_backup_enabled': backup_enabled_var.get(), 'smart_backup_interval': int(backup_interval_var.get()), 'start_with_windows': startup_var.get() })
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Invalid Input", "Intervals must be valid numbers.", parent=dialog)
        
        btn_container = customtkinter.CTkFrame(main_frame, fg_color="transparent")
        btn_container.pack(side="bottom", pady=(10, 0))
        customtkinter.CTkButton(btn_container, text="Save Settings", command=save_and_close).pack(side="left", padx=5)
        customtkinter.CTkButton(btn_container, text="Close", command=dialog.destroy, fg_color=Config.UI.Theme.ACCENT_RED, hover_color="#c0392b").pack(side="left", padx=5)

    def _toggle_log(self):
        if self.log_expanded:
            self.log_text_container.pack_forget()
            self.geometry(Config.App.GEOMETRY_COLLAPSED)
            self.log_toggle_button.configure(text="▶")
        else:
            self.footer.pack_forget()
            self.log_text_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
            self.footer.pack(fill="x", side="bottom", pady=(10, 0))
            self.geometry(Config.App.GEOMETRY_EXPANDED)
            self.log_toggle_button.configure(text="▼")
        self.log_expanded = not self.log_expanded

    def add_log(self, message: str, level: LogLevel):
        self.log_text.configure(state="normal")
        log_entry = f"[{time.strftime('%H:%M:%S')}] {message}\n"
        self.log_text.insert(tk.END, log_entry)
        self.log_text.see(tk.END)
        self.log_text.configure(state="disabled")

    def update_status(self, text: str, status_type: str, app_name: str = ""):
        self.status_label.configure(text=text)
        status_colors = {
            "INITIALIZING": "gray", "WAITING": Config.UI.Theme.ACCENT_PURPLE, "ACTIVE": Config.UI.Theme.PRIMARY,
            "PAUSED": Config.UI.Theme.WARNING, "SAVED": Config.UI.Theme.INFO
        }
        self.status_label.configure(text_color=status_colors.get(status_type, "gray"))
        if status_type == "SAVED":
            self.after(2000, lambda: self.update_status("Active", "ACTIVE"))

    def _on_add_app_browse(self):
        filepath = filedialog.askopenfilename(title="Select Application", filetypes=[("Executable files", "*.exe")])
        if filepath:
            self.script.add_monitored_app(filepath)
            messagebox.showinfo("Success", f"{Path(filepath).name} added to the watchlist.")

    def _schedule_show_from_tray(self): self.after(0, self.show_from_tray)
    def _schedule_quit_application(self): self.after(0, self.quit_application)

    def hide_to_tray(self, silent: bool = False):
        if not self.window_hidden:
            self.withdraw()
            self.tray_manager.run_in_thread()
            self.window_hidden = True
            if not silent: self.add_log("Minimized to system tray.", LogLevel.INFO)

    def show_from_tray(self):
        if self.window_hidden:
            self.tray_manager.stop()
            self.deiconify()
            self.lift()
            self.focus_force()
            self.window_hidden = False

    def quit_application(self):
        self.script.cleanup()
        self.tray_manager.stop()
        self.destroy()

# ==============================================================================
# کلاس‌های بدون تغییر
# ==============================================================================
class SettingsManager:
    def __init__(self):
        appdata_dir = Path(os.getenv('APPDATA', '.')) / Config.App.FOLDER_NAME; appdata_dir.mkdir(parents=True, exist_ok=True)
        self._path = appdata_dir / Config.App.SETTINGS_FILE
        self.settings = { 'monitored_apps': ['photoshop.exe', 'afterfx.exe', 'premiere.exe', 'illustrator.exe', 'indesign.exe', 'acrobat.exe', 'animate.exe', 'lightroom.exe', 'audition.exe', 'figma.exe', 'resolve.exe', 'capcut.exe'], 'auto_save_enabled': True, 'auto_save_interval': 3, 'smart_backup_enabled': False, 'smart_backup_interval': 60, 'start_with_windows': False }
    def load(self):
        if self._path.exists():
            try:
                with open(self._path, 'r') as f: loaded_settings = json.load(f)
                for key, value in self.settings.items(): loaded_settings.setdefault(key, value)
                self.settings = loaded_settings
            except (json.JSONDecodeError, IOError): self.save()
        else: self.save()
        return self.settings
    def save(self):
        try:
            with open(self._path, 'w') as f: json.dump(self.settings, f, indent=4)
            return True
        except IOError: return False

class WindowsStartupManager:
    def __init__(self):
        self.startup_folder = Path(os.getenv('APPDATA')) / 'Microsoft/Windows/Start Menu/Programs/Startup'; self.shortcut_path = self.startup_folder / f"{Config.App.NAME}.lnk"
    def set_startup(self, enable: bool):
        if not WINDOWS_API_AVAILABLE: return
        try:
            if enable: self._create_shortcut()
            elif self.shortcut_path.exists(): self.shortcut_path.unlink()
        except Exception as e: print(f"Error managing startup shortcut: {e}")
    def _create_shortcut(self):
        import sys
        try:
            from win32com.client import Dispatch
            shell = Dispatch('WScript.Shell'); shortcut = shell.CreateShortCut(str(self.shortcut_path))
            target = sys.executable
            if target.endswith("python.exe"): target = target.replace("python.exe", "pythonw.exe")
            shortcut.Targetpath = target; shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
            shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__)); shortcut.save()
        except ImportError: print("pywin32 is required to manage startup entries.")
        except Exception as e: print(f"Failed to create shortcut: {e}")

class ProcessMonitor:
    def __init__(self, log_cb: Callable): self._log, self.hook, self.hook_proc = log_cb, None, None
    def get_active_window_info(self) -> Optional[tuple[str, str]]:
        if not WINDOWS_API_AVAILABLE: return None
        try:
            hwnd = GetForegroundWindow()
            if not hwnd: return None
            _, pid = GetWindowThreadProcessId(hwnd); process_name = psutil.Process(pid).name().lower(); window_title = GetWindowText(hwnd)
            return (process_name, window_title)
        except (psutil.Error, AttributeError): return None
    def start_monitoring(self, check_cb: Callable):
        if not WINDOWS_API_AVAILABLE: self._log("Windows API not found. Using fallback polling.", LogLevel.WARNING); return self._start_fallback_monitoring(check_cb)
        try:
            def handler(nCode, wParam, lParam):
                if nCode >= 0 and wParam == win32con.HCBT_ACTIVATE: AppUI.instance.after(100, check_cb)
                return ctypes.windll.user32.CallNextHookEx(self.hook, nCode, wParam, lParam)
            HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM); self.hook_proc = HOOKPROC(handler)
            self.hook = ctypes.windll.user32.SetWindowsHookExW(win32con.WH_CBT, self.hook_proc, ctypes.windll.kernel32.GetModuleHandleW(None), 0)
            if not self.hook: raise RuntimeError("Hook failed.")
            self._log("Real-time monitoring enabled.", LogLevel.SUCCESS)
        except Exception: self._start_fallback_monitoring(check_cb)
    def _start_fallback_monitoring(self, check_cb: Callable):
        def loop():
            last_info = None
            while True:
                info = self.get_active_window_info()
                if info and info[0] != (last_info and last_info[0]): last_info = info; AppUI.instance.after(0, check_cb)
                time.sleep(2)
        threading.Thread(target=loop, daemon=True).start()
    def cleanup(self):
        if self.hook and WINDOWS_API_AVAILABLE: ctypes.windll.user32.UnhookWindowsHookEx(self.hook)

class AutoSaveScript:
    def __init__(self, log_cb: Callable, status_cb: Callable):
        self._log, self._update_status = log_cb, status_cb; self.settings_manager = SettingsManager(); self.process_monitor = ProcessMonitor(self._log)
        self.startup_manager = WindowsStartupManager(); self.settings = self.settings_manager.load()
        self.current_app: Optional[str] = None; self._save_timer: Optional[threading.Timer] = None; self._backup_timer: Optional[threading.Timer] = None
    def start(self): self._log("Personal Assistant started.", LogLevel.STARTUP); self.process_monitor.start_monitoring(self._check_active_window)
    def _is_target(self, name: str) -> bool: return name in self.settings.get('monitored_apps', [])
    def _check_active_window(self):
        info = self.process_monitor.get_active_window_info()
        if info and self._is_target(info[0]):
            app_name = info[0].replace('.exe', '').title()
            if self.current_app != app_name:
                self.current_app = app_name; self.settings = self.settings_manager.load()
                if self.settings.get('auto_save_enabled', True): self._update_status("Active", "ACTIVE", app_name)
                else: self._update_status("Paused", "PAUSED", app_name)
                self._log(f"Active application: {app_name}", LogLevel.ACTIVE); self._start_timers()
        elif self.current_app:
            self._log(f"Stopped monitoring {self.current_app}", LogLevel.INFO); self._update_status("Waiting", "WAITING")
            self.current_app = None; self._stop_timers()
    def _start_timers(self):
        self._stop_timers()
        if self.settings.get('auto_save_enabled', True): self._start_save_timer()
        if self.settings.get('smart_backup_enabled', False): self._start_backup_timer()
    def _stop_timers(self): self._stop_save_timer(); self._stop_backup_timer()
    def _start_save_timer(self):
        self._stop_save_timer()
        def task():
            if self.current_app and self.settings.get('auto_save_enabled', True):
                try: pyautogui.hotkey('ctrl', 's')
                except Exception: self._log("Auto-Save command failed.", LogLevel.ERROR)
                self._log(f"Auto-saved in {self.current_app}", LogLevel.SAVE); self._update_status("Saved!", "SAVED")
                if self.current_app: self._start_save_timer()
        self._save_timer = threading.Timer(self.settings['auto_save_interval'], task); self._save_timer.start()
    def _stop_save_timer(self):
        if self._save_timer: self._save_timer.cancel(); self._save_timer = None
    def _start_backup_timer(self):
        self._stop_backup_timer()
        def task():
            if self.current_app and self.settings.get('smart_backup_enabled', False):
                self._create_backup()
                if self.current_app: self._start_backup_timer()
        self._backup_timer = threading.Timer(self.settings.get('smart_backup_interval', 60) * 60, task); self._backup_timer.start()
    def _stop_backup_timer(self):
        if self._backup_timer: self._backup_timer.cancel(); self._backup_timer = None
    def _create_backup(self):
        info = self.process_monitor.get_active_window_info()
        if not info or not info[1] or '*' not in info[1]: return
        try:
            title = info[1].split(' @')[0].split(' - ')[0].replace('*', '').strip(); original_path = Path(title)
            if not original_path.is_file(): self._log("Backup failed: Cannot determine file path.", LogLevel.WARNING); return
            backup_dir = original_path.parent / "Smart Backup"; backup_dir.mkdir(exist_ok=True)
            timestamp = time.strftime("%Y%m%d-%H%M%S"); backup_filename = f"{original_path.stem}-{timestamp}{original_path.suffix}"
            backup_path = backup_dir / backup_filename; pyautogui.hotkey('ctrl', 'shift', 's'); time.sleep(1)
            pyautogui.write(str(backup_path), interval=0.01); pyautogui.press('enter')
            self._log(f"Smart Backup created: {backup_filename}", LogLevel.SUCCESS); self._update_status("Backup!", "SAVED")
        except Exception: self._log("Backup failed unexpectedly.", LogLevel.ERROR)
    def update_settings(self, new_settings):
        self.settings.update(new_settings)
        if self.settings_manager.save():
            self._log("Settings updated.", LogLevel.SUCCESS); self.startup_manager.set_startup(self.settings.get('start_with_windows', False))
        else: self._log("Failed to save settings.", LogLevel.ERROR)
        if self.current_app:
            self._stop_timers(); self._start_timers()
            if self.settings.get('auto_save_enabled', True): self._update_status("Active", "ACTIVE", self.current_app)
            else: self._update_status("Paused", "PAUSED", self.current_app)
    def add_monitored_app(self, app_path: str):
        app_name = Path(app_path).name.lower()
        if app_name and app_name not in self.settings['monitored_apps']:
            self.settings['monitored_apps'].append(app_name)
            if self.settings_manager.save(): self._log(f"Added to watchlist: {app_name}", LogLevel.SUCCESS)
            else: self._log("Failed to save new app.", LogLevel.ERROR)
    def cleanup(self): self._stop_timers(); self.process_monitor.cleanup(); self._log("Personal Assistant has been shut down.", LogLevel.INFO)

class TrayManager:
    def __init__(self, show_cb: Callable, exit_cb: Callable):
        self._show, self._exit = show_cb, exit_cb
        self._icon: Optional[pystray.Icon] = None

    def run_in_thread(self):
        if self._icon and self._icon.visible:
            return
        def create_image():
            img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.ellipse((4, 4, 60, 60), fill=Config.UI.Theme.PRIMARY)
            return img
        menu = pystray.Menu(pystray.MenuItem("Show", self._show, default=True), pystray.MenuItem("Exit", self._exit))
        self._icon = pystray.Icon(Config.App.NAME, create_image(), Config.App.NAME, menu)
        threading.Thread(target=self._icon.run, daemon=True).start()

    def stop(self):
        if self._icon:
            self._icon.stop()
            self._icon = None


if __name__ == "__main__":
    if WINDOWS_API_AVAILABLE:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception: pass
    
    app = AppUI()
    app.mainloop()