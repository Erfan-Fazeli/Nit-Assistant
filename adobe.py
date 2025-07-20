import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, scrolledtext, filedialog
import json, os, threading, time, psutil, pyautogui, pystray
from PIL import Image, ImageDraw
from pathlib import Path
from enum import Enum
from typing import Optional, Callable

try:
    from win32gui import GetForegroundWindow, GetWindowText
    from win32process import GetWindowThreadProcessId
    import win32con, ctypes
    from ctypes import wintypes
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

class Config:
    """پیکربندی مرکزی برنامه."""
    class App:
        NAME, VERSION = "NIT Group Personal Assistant", "Final v2.4" # نسخه آپدیت شد
        GEOMETRY_COLLAPSED, GEOMETRY_EXPANDED = "350x280", "350x450"
        FOLDER_NAME, SETTINGS_FILE = "PersonalAssistant", "config.json"

    class UI:
        class Theme:
            BACKGROUND, CARD_BG = "#0F172A", "#1A202C"
            LOG_AREA_BG = "#0A0F1A"
            PRIMARY, PRIMARY_HOVER = "#27ae60", "#229954"
            SECONDARY, SECONDARY_HOVER = "#34495e", "#2c3e50"
            ACCENT_PURPLE, ACCENT_RED, ACCENT_BLUE = "#9b59b6", "#e74c3c", "#3498db"
            ACCENT_RED_HOVER = "#c0392b"
            WARNING = "#f39c12"
            TEXT_LIGHT, TEXT_MUTED = "#E2E8F0", "#94A3B8"
        
        class Font:
            FAMILY_PRIMARY, FAMILY_MONO = "Consolas", "Consolas"
            TITLE, HEADER, BODY, FOOTER, MONO = 14, 11, 10, 9, 9

class LogLevel(Enum):
    SUCCESS = ("#27ae60", "✅"); WARNING = ("#f39c12", "⚠️"); ERROR   = ("#e74c3c", "❌"); INFO    = ("#3498db", "ℹ️"); STARTUP = ("#9b59b6", "🚀"); ACTIVE  = ("#e67e22", "🎯"); SAVE    = ("#2ecc71", "💾")

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
        if self._icon and self._icon.visible: return
        def create_image():
            img = Image.new('RGBA', (64, 64), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            draw.ellipse((4, 4, 60, 60), fill=Config.UI.Theme.PRIMARY)
            draw.rectangle((18, 18, 46, 46), fill=Config.UI.Theme.BACKGROUND)
            draw.rectangle((22, 12, 42, 30), fill=Config.UI.Theme.TEXT_LIGHT)
            return img
        menu = pystray.Menu(pystray.MenuItem("Show", self._show, default=True), pystray.MenuItem("Exit", self._exit))
        self._icon = pystray.Icon(Config.App.NAME, create_image(), Config.App.NAME, menu)
        threading.Thread(target=self._icon.run, daemon=True).start()
    def stop(self):
        if self._icon: self._icon.stop()

class StyledSpinbox(tk.Frame):
    def __init__(self, parent, from_=1, to=60, textvariable=None, command=None, **kwargs):
        super().__init__(parent, **kwargs)
        self.configure(bg=Config.UI.Theme.CARD_BG)
        self.textvariable, self.command = textvariable, command
        self._from, self._to = from_, to
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        # <<-- این خط دیگر خطا نمی‌دهد چون متد _font اصلاح شده است -->>
        entry_font, btn_font = self._font(font_cfg.BODY), self._font(7, mono=True)
        self.entry = tk.Entry(self, textvariable=self.textvariable, width=4, relief="flat", justify='center', bg=theme.CARD_BG, fg=theme.TEXT_LIGHT, font=entry_font, insertbackground=theme.TEXT_LIGHT, disabledbackground=theme.CARD_BG, disabledforeground=theme.TEXT_MUTED)
        self.up_button = tk.Button(self, text="▲", command=self._increment, relief="flat", font=btn_font, bg=theme.CARD_BG, fg=theme.TEXT_LIGHT, activebackground=theme.PRIMARY_HOVER, activeforeground=theme.TEXT_LIGHT, width=2)
        self.down_button = tk.Button(self, text="▼", command=self._decrement, relief="flat", font=btn_font, bg=theme.CARD_BG, fg=theme.TEXT_LIGHT, activebackground=theme.PRIMARY_HOVER, activeforeground=theme.TEXT_LIGHT, width=2)
        self.entry.pack(side="left", fill="y", expand=True, padx=(0, 1))
        btn_frame = tk.Frame(self, bg=theme.CARD_BG)
        btn_frame.pack(side="left", fill="y")
        self.up_button.pack(side="top", expand=True, fill="both")
        self.down_button.pack(side="bottom", expand=True, fill="both")
        self.entry.bind('<Return>', self._validate)
        self.entry.bind('<FocusOut>', self._validate)
    
    # <<-- شروع اصلاحیه نهایی: تعریف متد _font در این کلاس کامل شد -->>
    def _font(self, size, bold=False, mono=False):
        return (Config.UI.Font.FAMILY_MONO if mono else Config.UI.Font.FAMILY_PRIMARY,
                size, "bold" if bold else "normal")
    # <<-- پایان اصلاحیه نهایی -->>
    
    def _increment(self):
        try:
            self.textvariable.set(min(int(self.textvariable.get()) + 1, self._to))
            self._validate()
        except ValueError: self.textvariable.set(self._from)
    def _decrement(self):
        try:
            self.textvariable.set(max(int(self.textvariable.get()) - 1, self._from))
            self._validate()
        except ValueError: self.textvariable.set(self._from)
    def _validate(self, event=None):
        if self.command: self.command()
    def configure(self, state=None, **kwargs):
        super().configure(**kwargs)
        if state is not None:
            for widget in [self.entry, self.up_button, self.down_button]:
                widget.configure(state=state)
    config = configure

class ModernSwitch(tk.Canvas):
    def __init__(self, parent, variable, command=None, **kwargs):
        super().__init__(parent, width=50, height=26, bg=Config.UI.Theme.CARD_BG, highlightthickness=0, bd=0, **kwargs)
        self.variable = variable
        self.command = command
        self.is_on = self.variable.get()
        self.is_animating = False
        self.on_color = Config.UI.Theme.PRIMARY
        self.off_color = Config.UI.Theme.SECONDARY
        self.handle_color = Config.UI.Theme.TEXT_LIGHT
        self.handle_x_on = 38
        self.handle_x_off = 12
        self.animation_steps = 15  # کاهش گام‌ها برای سادگی
        self.animation_delay = 30  # افزایش تأخیر برای آهسته‌تر شدن
        self.bind("<Button-1>", self._on_click)
        self._draw_switch()

    def _draw_base(self, color):
        self.delete("base", "shadow")
        # سایه با تنظیم دقیق برای حذف فضای مشکی
        self.create_oval(4, 4, 22, 22, fill="#1A202C", outline="", tags="shadow")
        self.create_oval(28, 4, 46, 22, fill="#1A202C", outline="", tags="shadow")
        # پایه اصلی
        self.create_oval(3, 3, 23, 23, fill=color, outline="", tags="base")
        self.create_oval(27, 3, 47, 23, fill=color, outline="", tags="base")
        self.create_rectangle(13, 3, 37, 23, fill=color, outline="", tags="base")

    def _draw_handle(self, x):
        self.delete("handle")
        self.create_oval(x-9, 4, x+9, 22, fill=self.handle_color, outline="", tags="handle")

    def _draw_switch(self):
        bg_color = self.on_color if self.is_on else self.off_color
        handle_x = self.handle_x_on if self.is_on else self.handle_x_off
        self._draw_base(bg_color)
        self._draw_handle(handle_x)

    def _on_click(self, event):
        if self.is_animating: return
        self.is_on = not self.is_on
        self.variable.set(self.is_on)
        self.is_animating = True
        start_x = self.handle_x_off if self.is_on else self.handle_x_on
        end_x = self.handle_x_on if self.is_on else self.handle_x_off
        self._draw_base(self.on_color if self.is_on else self.off_color)
        self._animate(1, self.animation_steps, start_x, end_x)

    def _animate(self, step, total_steps, start_x, end_x):
        if step > total_steps:
            self.is_animating = False
            if self.command: self.command()
            return
        progress = step / total_steps
        current_x = start_x + (end_x - start_x) * progress
        self.coords("handle", current_x-9, 4, current_x+9, 22)
        self.after(self.animation_delay, self._animate, step + 1, total_steps, start_x, end_x)


class AppUI(tk.Toplevel):
    instance = None
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)
        AppUI.instance = self
        self.master.withdraw()
        self.script = AutoSaveScript(self.add_log, self.update_status)
        self.tray_manager = TrayManager(self._schedule_show_from_tray, self._schedule_quit_application)
        self.window_hidden = False
        self.log_expanded = True
        self._configure_root_style()
        self._create_widgets()
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.script.start()
        self.update_status("Initializing...", "INITIALIZING")
    
    def _font(self, s, b=False, m=False):
        return (Config.UI.Font.FAMILY_PRIMARY, s, "bold" if b else "normal")
    def _configure_root_style(self):
        self.title(Config.App.NAME)
        self.geometry(Config.App.GEOMETRY_EXPANDED)
        self.resizable(False, False)
        self.configure(bg=Config.UI.Theme.BACKGROUND)
    def _create_widgets(self):
        self.container = tk.Frame(self, bg=Config.UI.Theme.BACKGROUND)
        self.container.pack(fill="both", expand=True, padx=20, pady=20)
        self._create_header(self.container)
        self._create_control_panel(self.container)
        self._create_log_section(self.container)
        self._create_footer(self.container)
    def _create_header(self, parent):
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        header = tk.Frame(parent, bg=theme.PRIMARY, height=55)
        header.pack(fill="x", pady=(0, 20))
        header.pack_propagate(False)
        tk.Label(header, text=Config.App.NAME, font=("Fixedsys", font_cfg.TITLE, "bold"), fg="white", bg=theme.PRIMARY).pack(pady=(5, 0))
        tk.Label(header, text=f"v{Config.App.VERSION}", font=("Fixedsys", font_cfg.FOOTER, "bold"), fg="black", bg=theme.PRIMARY).pack()
    def _create_control_panel(self, parent):
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        card = tk.Frame(parent, bg=theme.CARD_BG)
        card.pack(fill="x", pady=(0, 15))
        content = tk.Frame(card, bg=theme.CARD_BG)
        content.pack(fill="both", expand=True, padx=15, pady=12)
        tk.Label(content, text="📊 System Status", font=self._font(font_cfg.HEADER, b=True), fg=theme.TEXT_LIGHT, bg=theme.CARD_BG).pack(anchor="w")
        status_fr = tk.Frame(content, bg=theme.CARD_BG)
        status_fr.pack(fill="x", pady=(10, 5))
        tk.Label(status_fr, text="Status:", font=self._font(font_cfg.BODY, b=True), fg=theme.TEXT_LIGHT, bg=theme.CARD_BG).pack(side="left")
        self.status_label = tk.Label(status_fr, text="Initializing...", font=self._font(font_cfg.BODY), fg=theme.TEXT_MUTED, bg=theme.CARD_BG)
        self.status_label.pack(side="left", padx=5)
        tk.Button(status_fr, text="⚙️ Settings", font=self._font(font_cfg.BODY), bg=theme.SECONDARY, fg=theme.TEXT_LIGHT, activebackground=theme.SECONDARY_HOVER, relief="flat", cursor="hand2", command=self._open_settings_dialog).pack(side="right")
    def _create_log_section(self, parent):
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        self.log_fr = tk.Frame(parent, bg=theme.CARD_BG)
        self.log_fr.pack(fill="both", expand=True)
        log_header_fr = tk.Frame(self.log_fr, bg=theme.CARD_BG)
        log_header_fr.pack(fill="x", padx=15, pady=(10, 5))
        self.log_toggle_button = tk.Button(log_header_fr, text="▼", command=self._toggle_log, relief="flat", bg=theme.CARD_BG, fg=theme.TEXT_LIGHT, font=self._font(8))
        self.log_toggle_button.pack(side="left", anchor='n')
        tk.Label(log_header_fr, text="📋 Activity Log", font=self._font(font_cfg.HEADER, b=True), fg=theme.TEXT_LIGHT, bg=theme.CARD_BG).pack(side="left")
        self.log_text_container = tk.Frame(self.log_fr, bg=theme.CARD_BG)
        self.log_text_container.pack(fill="both", expand=True)
        self.log_text = scrolledtext.ScrolledText(self.log_text_container, height=10, font=self._font(font_cfg.MONO, m=True), bg=theme.LOG_AREA_BG, fg=theme.TEXT_LIGHT, wrap="none", state="disabled", relief="flat", bd=0)
        self.log_text.pack(fill="both", expand=True, padx=15, pady=(0, 15))
    def _create_footer(self, parent):
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        footer = tk.Frame(parent, bg=theme.BACKGROUND)
        footer.pack(fill="x", side="bottom", pady=(10, 0))
        footer_content = tk.Frame(footer, bg=theme.BACKGROUND)
        footer_content.pack()
        segs = [{"text": "Developed With ", "fg": theme.TEXT_LIGHT}, {"text": "❤️", "fg": theme.ACCENT_RED}, {"text": " By ", "fg": theme.TEXT_LIGHT}, {"text": "Erf", "fg": theme.ACCENT_PURPLE}, {"text": " For ", "fg": theme.TEXT_LIGHT}, {"text": "NIT Group", "fg": theme.ACCENT_RED}]
        for seg in segs:
            tk.Label(footer_content, font=self._font(font_cfg.FOOTER, b=True), bg=theme.BACKGROUND, **seg).pack(side="left")

    def _open_settings_dialog(self):
        theme, font_cfg = Config.UI.Theme, Config.UI.Font
        dialog = tk.Toplevel(self)
        dialog.title("Settings")
        dialog.geometry("420x480")
        dialog.resizable(False, False)
        dialog.configure(bg=theme.BACKGROUND)
        dialog.transient(self)
        dialog.grab_set()

        save_enabled_var = tk.BooleanVar(value=self.script.settings['auto_save_enabled'])
        save_interval_var = tk.StringVar(value=str(self.script.settings['auto_save_interval']))
        backup_enabled_var = tk.BooleanVar(value=self.script.settings.get('smart_backup_enabled', False))
        backup_interval_var = tk.StringVar(value=str(self.script.settings.get('smart_backup_interval', 60)))
        startup_var = tk.BooleanVar(value=self.script.settings.get('start_with_windows', False))
        
        main_frame = tk.Frame(dialog, bg=theme.BACKGROUND)
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        automation_card = tk.Frame(main_frame, bg=theme.CARD_BG)
        automation_card.pack(fill="x", pady=(0, 15))
        tk.Label(automation_card, text="⚡️ Automation", font=self._font(font_cfg.HEADER, b=True), bg=theme.CARD_BG, fg=theme.TEXT_LIGHT).pack(anchor="w", padx=15, pady=(10, 5))
        automation_content = tk.Frame(automation_card, bg=theme.CARD_BG)
        automation_content.pack(fill="x", padx=15, pady=5)
        
        save_col = tk.Frame(automation_content, bg=theme.CARD_BG)
        save_col.pack(side="left", expand=True, fill="x", padx=(0, 10))
        save_header = tk.Frame(save_col, bg=theme.CARD_BG)
        save_header.pack(fill="x", pady=(5,0))
        tk.Label(save_header, text="Auto-Save", bg=theme.CARD_BG, fg=theme.TEXT_LIGHT).pack(side="left")
        save_switch = ModernSwitch(save_header, variable=save_enabled_var)
        save_switch.pack(side="right")
        save_interval_frame = tk.Frame(save_col, bg=theme.CARD_BG)
        save_interval_frame.pack(fill="x", pady=(5, 10))
        tk.Label(save_interval_frame, text="Interval (sec):", bg=theme.CARD_BG, fg=theme.TEXT_MUTED).pack(side="left")
        save_spinbox = StyledSpinbox(save_interval_frame, from_=1, to=60, textvariable=save_interval_var)
        save_spinbox.pack(side="right")

        backup_col = tk.Frame(automation_content, bg=theme.CARD_BG)
        backup_col.pack(side="left", expand=True, fill="x", padx=(10, 0))
        backup_header = tk.Frame(backup_col, bg=theme.CARD_BG)
        backup_header.pack(fill="x", pady=(5,0))
        tk.Label(backup_header, text="Smart Backup", bg=theme.CARD_BG, fg=theme.TEXT_LIGHT).pack(side="left")
        backup_switch = ModernSwitch(backup_header, variable=backup_enabled_var)
        backup_switch.pack(side="right")
        backup_interval_frame = tk.Frame(backup_col, bg=theme.CARD_BG)
        backup_interval_frame.pack(fill="x", pady=(5, 10))
        tk.Label(backup_interval_frame, text="Interval (min):", bg=theme.CARD_BG, fg=theme.TEXT_MUTED).pack(side="left")
        backup_spinbox = StyledSpinbox(backup_interval_frame, from_=1, to=1440, textvariable=backup_interval_var)
        backup_spinbox.pack(side="right")
        
        general_card = tk.Frame(main_frame, bg=theme.CARD_BG)
        general_card.pack(fill="x", pady=(0, 15))
        tk.Label(general_card, text="⚙️ General", font=self._font(font_cfg.HEADER, b=True), bg=theme.CARD_BG, fg=theme.TEXT_LIGHT).pack(anchor="w", padx=15, pady=(10, 10))
        startup_frame = tk.Frame(general_card, bg=theme.CARD_BG)
        startup_frame.pack(fill="x", padx=15, pady=(0, 10))
        tk.Label(startup_frame, text="Start with Windows", bg=theme.CARD_BG, fg=theme.TEXT_LIGHT).pack(side="left")
        startup_switch = ModernSwitch(startup_frame, variable=startup_var)
        startup_switch.pack(side="right", anchor="e")
        add_app_frame = tk.Frame(general_card, bg=theme.CARD_BG)
        add_app_frame.pack(fill="x", padx=15, pady=(5, 15))
        tk.Button(add_app_frame, text="📂 Add Application...", command=self._on_add_app_browse, font=self._font(font_cfg.BODY), bg=theme.PRIMARY, fg="white", activebackground=theme.PRIMARY_HOVER, relief="flat").pack(ipady=4)

        def toggle_spinners_state():
            save_spinbox.config(state='normal' if save_enabled_var.get() else 'disabled')
            backup_spinbox.config(state='normal' if backup_enabled_var.get() else 'disabled')
        save_switch.command = toggle_spinners_state
        backup_switch.command = toggle_spinners_state
        toggle_spinners_state()

        btn_container = tk.Frame(main_frame, bg=theme.BACKGROUND)
        btn_container.pack(fill="x", side="bottom", pady=(10, 0))
        btn_frame = tk.Frame(btn_container, bg=theme.BACKGROUND)
        btn_frame.pack()
        
        def save_and_close():
            try:
                self.script.update_settings({ 'auto_save_enabled': save_enabled_var.get(), 'auto_save_interval': int(save_interval_var.get()), 'smart_backup_enabled': backup_enabled_var.get(), 'smart_backup_interval': int(backup_interval_var.get()), 'start_with_windows': startup_var.get() })
                dialog.destroy()
            except ValueError:
                messagebox.showerror("Invalid Input", "Intervals must be valid numbers.", parent=dialog)
        tk.Button(btn_frame, text="Save Settings", command=save_and_close, font=self._font(font_cfg.BODY, b=True), bg=theme.PRIMARY, fg="white", activebackground=theme.PRIMARY_HOVER, relief="flat", width=12, height=1, bd=0).pack(side="left", padx=5, ipady=4)
        tk.Button(btn_frame, text="Close", command=dialog.destroy, font=self._font(font_cfg.BODY, b=True), bg=theme.ACCENT_RED, fg="white", activebackground=theme.ACCENT_RED_HOVER, relief="flat", width=8, height=1, bd=0).pack(side="left", padx=5, ipady=4)

    def _toggle_log(self):
        if self.log_expanded:
            self.log_text_container.pack_forget()
            self.geometry(Config.App.GEOMETRY_COLLAPSED)
            self.log_toggle_button.config(text="▶")
        else:
            self.log_text_container.pack(fill="both", expand=True)
            self.geometry(Config.App.GEOMETRY_EXPANDED)
            self.log_toggle_button.config(text="▼")
        self.log_expanded = not self.log_expanded

    def add_log(self, message: str, level: LogLevel):
        log_map = { "Personal Assistant started.": "🚀", "Real-time monitoring enabled.": "🔎", "Windows API not found.": "⚠️", "Active application:": "🎯", "Stopped monitoring": "💤", "Auto-saved in": "💾", "Auto-Save command failed.": "💥", "Settings updated.": "⚙️", "Failed to save settings.": "🔒", "Added to watchlist:": "➕", "Moving to tray in 3s...": "⏳", "Minimized to system tray.": "🔽", "Personal Assistant has been shut down.": "🛑", "Smart Backup created:": "🛡️", "Backup failed:": "⚠️" }
        color, default_emoji = level.value
        emoji = next((emj for key, emj in log_map.items() if message.startswith(key)), default_emoji)
        self.log_text.config(state="normal")
        self.log_text.tag_config(level.name, foreground=color)
        self.log_text.insert(tk.END, f"[{time.strftime('%H:%M:%S')}] {emoji} {message}\n", level.name)
        self.log_text.see(tk.END)
        self.log_text.config(state="disabled")

    def update_status(self, text: str, status_type: str, app_name: str = ""):
        theme = Config.UI.Theme
        status_map = { "INITIALIZING": (f"{text}", theme.TEXT_MUTED), "WAITING": (f"{text}", theme.ACCENT_PURPLE), "ACTIVE": (f"{text}", theme.PRIMARY), "PAUSED": (f"{text}", theme.WARNING), "SAVED": (f"{text}", theme.ACCENT_BLUE) }
        display_text, color = status_map.get(status_type, (text, theme.TEXT_MUTED))
        self.status_label.config(text=display_text, fg=color)
        if status_type == "SAVED":
            previous_status = "ACTIVE" if self.script.settings.get('auto_save_enabled') else "PAUSED"
            previous_text = "Active" if previous_status == "ACTIVE" else "Paused"
            self.after(2000, lambda: self.update_status(previous_text, previous_status, self.script.current_app))

    def _on_add_app_browse(self):
        filepath = filedialog.askopenfilename(title="Select Application", filetypes=[("Executable files", "*.exe")])
        if filepath:
            self.script.add_monitored_app(filepath)
            messagebox.showinfo("Success", f"{Path(filepath).name} added to the watchlist.", parent=self.focus_get())

    def _schedule_show_from_tray(self):
        self.after(0, self.show_from_tray)

    def _schedule_quit_application(self):
        self.after(0, self.quit_application)

    def hide_to_tray(self, silent: bool = False):
        if not self.window_hidden:
            self.withdraw()
            self.tray_manager.run_in_thread()
            self.window_hidden = True
            if not silent:
                self.add_log("Minimized to system tray.", LogLevel.INFO)

    def show_from_tray(self):
        if self.window_hidden:
            self.tray_manager.stop()
            self.deiconify()
            self.lift()
            self.focus_force()
            self.window_hidden = False
            self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)

    def quit_application(self):
        self.script.cleanup()
        self.tray_manager.stop()
        self.master.destroy()

if __name__ == "__main__":
    if WINDOWS_API_AVAILABLE:
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    root.withdraw()
    app = AppUI(master=root)
    root.mainloop()