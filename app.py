import tkinter as tk
from tkinter import messagebox, filedialog
import json, os, threading, time, psutil, pyautogui, pystray
from PIL import Image, ImageDraw
from pathlib import Path
from enum import Enum
from typing import Optional, Callable

import customtkinter
import tkfontawesome as fa 

try:
    from win32gui import GetForegroundWindow, GetWindowText
    from win32process import GetWindowThreadProcessId
    import win32con, ctypes
    from ctypes import wintypes
    WINDOWS_API_AVAILABLE = True
except ImportError:
    WINDOWS_API_AVAILABLE = False

# --- App Appearance Setup ---
customtkinter.set_appearance_mode("dark")
customtkinter.set_default_color_theme("green")

class Config:
    """Central configuration for the application."""
    class App:
        NAME, VERSION = "NIT Personal Assistant", "Version: v0.8"
        GEOMETRY_COLLAPSED, GEOMETRY_EXPANDED = "350x420", "350x520" 
        FOLDER_NAME, SETTINGS_FILE = "NitPersonalAssistant", "config.json"
        BACKGROUND_COLOR = "#0F172A"

    class UI:
        class Theme:
            PRIMARY = "#27ae60"
            ACCENT_RED = "#e74c3c"
            ACCENT_PURPLE = "#9b59b6"
            WARNING = "#f39c12"
            INFO = "#3498db"
            PRIMARY_DISABLED = "#2a6a43"
            ICON_DISABLED = "#6b7280" 

class LogLevel(Enum):
    SUCCESS = (Config.UI.Theme.PRIMARY, "")
    WARNING = (Config.UI.Theme.WARNING, "")
    ERROR = (Config.UI.Theme.ACCENT_RED, "")
    INFO = (Config.UI.Theme.INFO, "")
    STARTUP = (Config.UI.Theme.ACCENT_PURPLE, "")
    ACTIVE = ("#e67e22", "")
    SAVE = ("#2ecc71", "")

class AppUI(customtkinter.CTk):
    instance = None
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        AppUI.instance = self
        
        self.configure(fg_color=Config.App.BACKGROUND_COLOR)

        self.script = AutoSaveScript(self.add_log, self.update_status)
        self.tray_manager = TrayManager(self._schedule_show_from_tray, self._schedule_quit_application)
        self.window_hidden = False
        self.log_expanded = False

        self._configure_root_style()
        self._create_widgets()
        
        self.protocol("WM_DELETE_WINDOW", self.hide_to_tray)
        self.script.start()
        self.update_status("Initializing...", "INITIALIZING")
        self._animate_gradient()

    def _configure_root_style(self):
        self.title(Config.App.NAME)
        self.geometry(Config.App.GEOMETRY_COLLAPSED)
        self.resizable(False, False)

    def _create_widgets(self):
        self.container = customtkinter.CTkFrame(self, fg_color="transparent")
        self.container.pack(fill="both", expand=True, padx=20, pady=20)
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        main_content_frame = customtkinter.CTkFrame(self.container, fg_color="transparent")
        main_content_frame.grid(row=0, column=0, sticky="nsew")

        self._create_header(main_content_frame)
        self._create_control_panel(main_content_frame)
        self._create_action_buttons(main_content_frame)
        self._create_log_section(main_content_frame)

        self._create_footer(self.container)

    def _create_header(self, parent):
        header = customtkinter.CTkFrame(parent, corner_radius=10, fg_color=Config.UI.Theme.PRIMARY)
        header.pack(fill="x", pady=(0, 20))
        
        text_container = customtkinter.CTkFrame(header, fg_color="transparent")
        text_container.pack(expand=True, pady=8)

        title_font = customtkinter.CTkFont(family="Fixedsys", size=20, weight="bold")
        version_font = customtkinter.CTkFont(family="Fixedsys", size=10, weight="normal")
        header_text_color = Config.App.BACKGROUND_COLOR

        customtkinter.CTkLabel(text_container, text=Config.App.NAME, font=title_font, text_color=header_text_color).pack()
        customtkinter.CTkLabel(text_container, text=f"{Config.App.VERSION}", font=version_font, text_color=header_text_color).pack(pady=0)

    def _create_control_panel(self, parent):
        card = customtkinter.CTkFrame(parent)
        card.pack(fill="x", pady=(0, 15))
        
        content = customtkinter.CTkFrame(card, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=15, pady=12)
        
        header_frame = customtkinter.CTkFrame(content, fg_color="transparent")
        header_frame.pack(anchor="w")
        
        status_icon = fa.icon_to_image(name="chart-line", scale_to_width=18, fill="white") 
        
        customtkinter.CTkLabel(header_frame, text="", image=status_icon).pack(side="left")
        customtkinter.CTkLabel(header_frame, text=" System Status", font=customtkinter.CTkFont(size=12, weight="bold")).pack(side="left", padx=5)
        
        status_fr = customtkinter.CTkFrame(content, fg_color="transparent")
        status_fr.pack(fill="x", pady=(10, 5))
        
        customtkinter.CTkLabel(status_fr, text="Status:", font=customtkinter.CTkFont(size=11, weight="bold")).pack(side="left")
        self.status_label = customtkinter.CTkLabel(status_fr, text="Initializing...")
        self.status_label.pack(side="left", padx=5)

    def _create_action_buttons(self, parent):
        action_card = customtkinter.CTkFrame(parent)
        action_card.pack(fill="x", pady=(0, 15))
        
        content_frame = customtkinter.CTkFrame(action_card, fg_color="transparent")
        content_frame.pack(fill="both", expand=True, padx=15, pady=15)
        
        content_frame.grid_columnconfigure((0, 1), weight=1)
        
        exit_icon = fa.icon_to_image(name="power-off", scale_to_width=16, fill="white")
        tools_icon = fa.icon_to_image(name="sliders-h", scale_to_width=16, fill="white")

        customtkinter.CTkButton(content_frame, text=" Exit", image=exit_icon, compound="left",
                              fg_color=Config.UI.Theme.ACCENT_RED, hover_color="#c0392b",
                              command=self.quit_application, height=35).grid(row=0, column=0, sticky="ew", padx=(0, 5))
        customtkinter.CTkButton(content_frame, text=" Tools", image=tools_icon, compound="left",
                              command=self._open_settings_dialog, height=35).grid(row=0, column=1, sticky="ew", padx=(5, 0))

    def _create_log_section(self, parent):
        self.log_fr = customtkinter.CTkFrame(parent)
        self.log_fr.pack(fill="both", expand=True)
        
        log_header_fr = customtkinter.CTkFrame(self.log_fr, fg_color="transparent")
        log_header_fr.pack(fill="x", padx=15, pady=(10, 5))
        
        self.icon_log_closed = fa.icon_to_image(name="chevron-right", scale_to_width=12, fill="white")
        self.icon_log_open = fa.icon_to_image(name="chevron-down", scale_to_width=12, fill="white")
        
        self.log_toggle_button = customtkinter.CTkButton(
            log_header_fr,
            text="", 
            image=self.icon_log_closed, 
            width=24, height=24,
            fg_color="transparent", 
            hover_color="#334155",
            command=self._toggle_log
        )
        self.log_toggle_button.pack(side="left", padx=(0, 10))
        
        customtkinter.CTkLabel(log_header_fr, text="Activity Log", font=customtkinter.CTkFont(size=12, weight="bold")).pack(side="left")
        
        self.log_text_container = customtkinter.CTkFrame(self.log_fr, fg_color="transparent")
        
        self.log_text = customtkinter.CTkTextbox(self.log_text_container, wrap="none", state="disabled", border_width=0, fg_color="#1E293B")
        self.log_text.pack(fill="both", expand=True)
    
    def _hex_to_rgb(self, hex_color):
        hex_color = hex_color.lstrip('#')
        return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

    def _rgb_to_hex(self, rgb_color):
        return f"#{int(rgb_color[0]):02x}{int(rgb_color[1]):02x}{int(rgb_color[2]):02x}"

    def _create_smooth_gradient(self, key_colors, steps_between):
        palette = []
        key_colors_loop = key_colors + [key_colors[0]]
        for i in range(len(key_colors_loop) - 1):
            start_rgb = self._hex_to_rgb(key_colors_loop[i])
            end_rgb = self._hex_to_rgb(key_colors_loop[i+1])
            for step in range(steps_between):
                interpolated_rgb = [
                    start_rgb[j] + (end_rgb[j] - start_rgb[j]) * step / steps_between
                    for j in range(3)
                ]
                palette.append(self._rgb_to_hex(interpolated_rgb))
        return palette

    def _create_footer(self, parent):
        self.footer = customtkinter.CTkFrame(parent, fg_color="transparent")
        self.footer.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        
        gradient_frame = customtkinter.CTkFrame(self.footer, fg_color="transparent")
        gradient_frame.pack()

        theme_palette = [
            Config.UI.Theme.PRIMARY, Config.UI.Theme.INFO, Config.UI.Theme.ACCENT_PURPLE,
            Config.UI.Theme.WARNING, Config.UI.Theme.ACCENT_RED
        ]
        self.gradient_colors = self._create_smooth_gradient(theme_palette, steps_between=30)
        
        self.gradient_labels = []
        self.gradient_offset = 0
        footer_text = "Developed By Erfan Fazeli For NIT TM"
        footer_font = customtkinter.CTkFont(size=12, weight="bold")

        for char in footer_text:
            label = customtkinter.CTkLabel(gradient_frame, text=char, font=footer_font)
            label.pack(side="left")
            self.gradient_labels.append(label)

    def _animate_gradient(self):
        num_labels = len(self.gradient_labels)
        num_colors = len(self.gradient_colors)
        
        for i, label in enumerate(self.gradient_labels):
            color_index = (i * 2 + self.gradient_offset) % num_colors
            label.configure(text_color=self.gradient_colors[color_index])
            
        self.gradient_offset = (self.gradient_offset - 1) % num_colors
        self.after(120, self._animate_gradient)

    def _open_settings_dialog(self):
        dialog = customtkinter.CTkToplevel(self)
        dialog.title("Tools")
        dialog.geometry("410x480")
        dialog.resizable(False, False)
        dialog.transient(self)
        dialog.grab_set()

        dialog.configure(fg_color=Config.App.BACKGROUND_COLOR)

        save_enabled_var = tk.BooleanVar(value=self.script.settings['auto_save_enabled'])
        backup_enabled_var = tk.BooleanVar(value=self.script.settings.get('smart_backup_enabled', False))
        startup_var = tk.BooleanVar(value=self.script.settings.get('start_with_windows', False))

        main_frame = customtkinter.CTkFrame(dialog, fg_color="transparent")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        general_icon = fa.icon_to_image(name="sliders-h", scale_to_width=18, fill="white")
        automation_icon = fa.icon_to_image(name="microchip", scale_to_width=18, fill="white")
        actions_icon = fa.icon_to_image(name="folder-plus", scale_to_width=18, fill="white")
        
        icon_plus_enabled = fa.icon_to_image(name="plus", scale_to_width=14, fill=Config.UI.Theme.PRIMARY)
        icon_plus_disabled = fa.icon_to_image(name="plus", scale_to_width=14, fill=Config.UI.Theme.ICON_DISABLED)
        icon_minus_enabled = fa.icon_to_image(name="minus", scale_to_width=14, fill=Config.UI.Theme.PRIMARY)
        icon_minus_disabled = fa.icon_to_image(name="minus", scale_to_width=14, fill=Config.UI.Theme.ICON_DISABLED)

        general_card = customtkinter.CTkFrame(main_frame)
        general_card.pack(fill="x", pady=(0, 15), ipady=10)
        general_header = customtkinter.CTkFrame(general_card, fg_color="transparent")
        general_header.pack(anchor="w", padx=15, pady=(5,0))
        customtkinter.CTkLabel(general_header, image=general_icon, text="").pack(side="left")
        customtkinter.CTkLabel(general_header, text=" General", font=customtkinter.CTkFont(size=12, weight="bold")).pack(side="left", padx=5)

        startup_frame = customtkinter.CTkFrame(general_card, fg_color="transparent")
        startup_frame.pack(fill="x", padx=15, pady=10)
        startup_frame.grid_columnconfigure(0, weight=1) 
        startup_frame.grid_columnconfigure(1, weight=0)
        customtkinter.CTkLabel(startup_frame, text="Run on Windows Startup").grid(row=0, column=0, sticky="w")
        startup_switch = customtkinter.CTkSwitch(startup_frame, text="", variable=startup_var, switch_width=50, switch_height=25)
        startup_switch.grid(row=0, column=1, sticky="e")

        automation_card = customtkinter.CTkFrame(main_frame)
        automation_card.pack(fill="x", pady=(0, 15), ipady=10)
        automation_header = customtkinter.CTkFrame(automation_card, fg_color="transparent")
        automation_header.pack(anchor="w", padx=15, pady=(5,0))
        customtkinter.CTkLabel(automation_header, image=automation_icon, text="").pack(side="left")
        customtkinter.CTkLabel(automation_header, text=" Automation", font=customtkinter.CTkFont(size=12, weight="bold")).pack(side="left", padx=5)

        automation_content = customtkinter.CTkFrame(automation_card, fg_color="transparent")
        automation_content.pack(fill="x", padx=15, pady=10)
        
        # --- START: Reverted to the stable Grid structure ---
        # 1. Configure a 4-column grid, with an expanding spacer in the middle
        automation_content.grid_columnconfigure(0, weight=0) # Label
        automation_content.grid_columnconfigure(1, weight=0) # Switch
        automation_content.grid_columnconfigure(2, weight=1) # Expanding Spacer (absorbs jiggle)
        automation_content.grid_columnconfigure(3, weight=0) # Controls
        
        # --- Auto-Save Row ---
        # 2. Place widgets directly into the grid
        customtkinter.CTkLabel(automation_content, text="Auto-Save").grid(row=0, column=0, sticky="w", pady=5)
        save_switch = customtkinter.CTkSwitch(automation_content, text="", variable=save_enabled_var, switch_width=50, switch_height=25)
        save_switch.grid(row=0, column=1, sticky="w", padx=10, pady=5)
        
        save_controls_fr = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        save_controls_fr.grid(row=0, column=3, sticky="e", pady=5)
        
        save_minus_label = customtkinter.CTkLabel(save_controls_fr, text="", image=icon_minus_enabled)
        save_minus_label.pack(side="left", padx=5)
        save_entry = customtkinter.CTkEntry(save_controls_fr, width=50, justify="center")
        save_entry.insert(0, str(self.script.settings['auto_save_interval']))
        save_entry.pack(side="left", padx=5)
        save_plus_label = customtkinter.CTkLabel(save_controls_fr, text="", image=icon_plus_enabled)
        save_plus_label.pack(side="left", padx=5)
        
        save_minus_label.bind("<Button-1>", lambda e: update_save_value(-1))
        save_plus_label.bind("<Button-1>", lambda e: update_save_value(1))
        save_minus_label.bind("<Enter>", lambda e: save_minus_label.configure(cursor="hand2"))
        save_minus_label.bind("<Leave>", lambda e: save_minus_label.configure(cursor=""))
        save_plus_label.bind("<Enter>", lambda e: save_plus_label.configure(cursor="hand2"))
        save_plus_label.bind("<Leave>", lambda e: save_plus_label.configure(cursor=""))

        def update_save_value(delta):
            if not save_enabled_var.get(): return
            try: value = int(save_entry.get()) + delta
            except ValueError: return
            if 1 <= value <= 60: save_entry.delete(0, tk.END); save_entry.insert(0, str(value))

        # --- Smart Backup Row (same structure) ---
        customtkinter.CTkLabel(automation_content, text="Smart Backup").grid(row=1, column=0, sticky="w", pady=5)
        backup_switch = customtkinter.CTkSwitch(automation_content, text="", variable=backup_enabled_var, switch_width=50, switch_height=25)
        backup_switch.grid(row=1, column=1, sticky="w", padx=10, pady=5)
        
        backup_controls_fr = customtkinter.CTkFrame(automation_content, fg_color="transparent")
        backup_controls_fr.grid(row=1, column=3, sticky="e", pady=5)

        backup_minus_label = customtkinter.CTkLabel(backup_controls_fr, text="", image=icon_minus_enabled)
        backup_minus_label.pack(side="left", padx=5)
        backup_entry = customtkinter.CTkEntry(backup_controls_fr, width=50, justify="center")
        backup_entry.insert(0, str(self.script.settings.get('smart_backup_interval', 60)))
        backup_entry.pack(side="left", padx=5)
        backup_plus_label = customtkinter.CTkLabel(backup_controls_fr, text="", image=icon_plus_enabled)
        backup_plus_label.pack(side="left", padx=5)

        backup_minus_label.bind("<Button-1>", lambda e: update_backup_value(-1))
        backup_plus_label.bind("<Button-1>", lambda e: update_backup_value(1))
        backup_minus_label.bind("<Enter>", lambda e: backup_minus_label.configure(cursor="hand2"))
        backup_minus_label.bind("<Leave>", lambda e: backup_minus_label.configure(cursor=""))
        backup_plus_label.bind("<Enter>", lambda e: backup_plus_label.configure(cursor="hand2"))
        backup_plus_label.bind("<Leave>", lambda e: backup_plus_label.configure(cursor=""))
        # --- END: Reverted to the stable Grid structure ---

        def update_backup_value(delta):
            if not backup_enabled_var.get(): return
            try: value = int(backup_entry.get()) + delta
            except ValueError: return
            if 1 <= value <= 120: backup_entry.delete(0, tk.END); backup_entry.insert(0, str(value))

        actions_card = customtkinter.CTkFrame(main_frame)
        actions_card.pack(fill="x", pady=(15, 0))
        actions_content = customtkinter.CTkFrame(actions_card, fg_color="transparent")
        actions_content.pack(fill="x", expand=True, padx=15, pady=15)
        customtkinter.CTkButton(actions_content, text=" Add Application...", image=actions_icon, compound="left", command=self._on_add_app_browse).pack(fill="x", ipady=4, pady=(0, 15))

        btn_container = customtkinter.CTkFrame(actions_content, fg_color="transparent")
        btn_container.pack(side="bottom")
        customtkinter.CTkButton(btn_container, text="Close", command=dialog.destroy, fg_color=Config.UI.Theme.ACCENT_RED, hover_color="#c0392b").pack(side="left", padx=5)
        customtkinter.CTkButton(btn_container, text="Save Settings", command=lambda: save_and_close()).pack(side="left", padx=5)

        def toggle_controls_state():
            if save_enabled_var.get():
                save_entry.configure(state='normal')
                save_minus_label.configure(image=icon_minus_enabled, cursor="hand2")
                save_plus_label.configure(image=icon_plus_enabled, cursor="hand2")
                save_minus_label.bind("<Button-1>", lambda e: update_save_value(-1))
                save_plus_label.bind("<Button-1>", lambda e: update_save_value(1))
            else:
                save_entry.configure(state='disabled')
                save_minus_label.configure(image=icon_minus_disabled, cursor="")
                save_plus_label.configure(image=icon_plus_disabled, cursor="")
                save_minus_label.unbind("<Button-1>")
                save_plus_label.unbind("<Button-1>")

            if backup_enabled_var.get():
                backup_entry.configure(state='normal')
                backup_minus_label.configure(image=icon_minus_enabled, cursor="hand2")
                backup_plus_label.configure(image=icon_plus_enabled, cursor="hand2")
                backup_minus_label.bind("<Button-1>", lambda e: update_backup_value(-1))
                backup_plus_label.bind("<Button-1>", lambda e: update_backup_value(1))
            else:
                backup_entry.configure(state='disabled')
                backup_minus_label.configure(image=icon_minus_disabled, cursor="")
                backup_plus_label.configure(image=icon_plus_disabled, cursor="")
                backup_minus_label.unbind("<Button-1>")
                backup_plus_label.unbind("<Button-1>")

        save_switch.configure(command=toggle_controls_state)
        backup_switch.configure(command=toggle_controls_state)
        toggle_controls_state()

        def save_and_close():
            try:
                auto_save_interval = int(save_entry.get())
                smart_backup_interval = int(backup_entry.get())
                if not (1 <= auto_save_interval <= 60 and 1 <= smart_backup_interval <= 120):
                    messagebox.showerror("Error", "Invalid intervals.")
                    return
                self.script.update_settings({
                    'auto_save_enabled': save_enabled_var.get(), 'auto_save_interval': auto_save_interval,
                    'smart_backup_enabled': backup_enabled_var.get(), 'smart_backup_interval': smart_backup_interval,
                    'start_with_windows': startup_var.get()
                })
                dialog.destroy()
            except ValueError: messagebox.showerror("Error", "Please enter valid numbers.")

    def _toggle_log(self):
        if self.log_expanded:
            self.log_text_container.pack_forget()
            self.geometry(Config.App.GEOMETRY_COLLAPSED)
            self.log_toggle_button.configure(image=self.icon_log_closed)
        else:
            self.log_text_container.pack(fill="both", expand=True, padx=15, pady=(0, 15))
            self.geometry(Config.App.GEOMETRY_EXPANDED)
            self.log_toggle_button.configure(image=self.icon_log_open)
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

# The rest of the classes (SettingsManager, WindowsStartupManager, etc.) remain unchanged.
# ... (paste the rest of your unchanged classes here)
class SettingsManager:
    def __init__(self):
        appdata_dir = Path(os.getenv('APPDATA', '.')) / Config.App.FOLDER_NAME
        appdata_dir.mkdir(parents=True, exist_ok=True)
        self._path = appdata_dir / Config.App.SETTINGS_FILE
        self.settings = {
            'monitored_apps': [
                'photoshop.exe', 'afterfx.exe', 'premiere.exe', 'illustrator.exe', 
                'indesign.exe', 'acrobat.exe', 'animate.exe', 'lightroom.exe', 
                'audition.exe', 'figma.exe', 'resolve.exe', 'capcut.exe'
            ],
            'auto_save_enabled': True,
            'auto_save_interval': 3,
            'smart_backup_enabled': False,
            'smart_backup_interval': 60,
            'start_with_windows': False
        }
    def load(self):
        if self._path.exists():
            try:
                with open(self._path, 'r') as f: 
                    loaded_settings = json.load(f)
                for key, value in self.settings.items():
                    loaded_settings.setdefault(key, value)
                self.settings = loaded_settings
            except (json.JSONDecodeError, IOError): 
                self.save()
        else:
            self.save()
        return self.settings
    def save(self):
        try:
            with open(self._path, 'w') as f: 
                json.dump(self.settings, f, indent=4)
            return True
        except IOError: 
            return False

class WindowsStartupManager:
    def __init__(self):
        self.startup_folder = Path(os.getenv('APPDATA')) / 'Microsoft/Windows/Start Menu/Programs/Startup'
        self.shortcut_path = self.startup_folder / f"{Config.App.NAME}.lnk"
    def set_startup(self, enable: bool):
        if not WINDOWS_API_AVAILABLE: 
            return
        try:
            if enable: 
                self._create_shortcut()
            elif self.shortcut_path.exists(): 
                self.shortcut_path.unlink()
        except Exception as e: 
            print(f"Error managing startup shortcut: {e}")
    def _create_shortcut(self):
        import sys
        try:
            from win32com.client import Dispatch
            shell = Dispatch('WScript.Shell')
            shortcut = shell.CreateShortCut(str(self.shortcut_path))
            target = sys.executable
            if target.endswith("python.exe"): 
                target = target.replace("python.exe", "pythonw.exe")
            shortcut.Targetpath = target
            shortcut.Arguments = f'"{os.path.abspath(__file__)}"'
            shortcut.WorkingDirectory = os.path.dirname(os.path.abspath(__file__))
            shortcut.save()
        except ImportError: 
            print("pywin32 is required to manage startup entries.")
        except Exception as e: 
            print(f"Failed to create shortcut: {e}")

class ProcessMonitor:
    def __init__(self, log_cb: Callable): 
        self._log, self.hook, self.hook_proc = log_cb, None, None
    def get_active_window_info(self) -> Optional[tuple[str, str]]:
        if not WINDOWS_API_AVAILABLE: 
            return None
        try:
            hwnd = GetForegroundWindow()
            if not hwnd: 
                return None
            _, pid = GetWindowThreadProcessId(hwnd)
            process_name = psutil.Process(pid).name().lower()
            window_title = GetWindowText(hwnd)
            return (process_name, window_title)
        except (psutil.Error, AttributeError): 
            return None
    def start_monitoring(self, check_cb: Callable):
        if not WINDOWS_API_AVAILABLE: 
            self._log("Windows API not found. Using fallback polling.", LogLevel.WARNING)
            return self._start_fallback_monitoring(check_cb)
        try:
            def handler(nCode, wParam, lParam):
                if nCode >= 0 and wParam == win32con.HCBT_ACTIVATE: 
                    AppUI.instance.after(100, check_cb)
                return ctypes.windll.user32.CallNextHookEx(self.hook, nCode, wParam, lParam)
            HOOKPROC = ctypes.WINFUNCTYPE(ctypes.c_int, ctypes.c_int, wintypes.WPARAM, wintypes.LPARAM)
            self.hook_proc = HOOKPROC(handler)
            self.hook = ctypes.windll.user32.SetWindowsHookExW(win32con.WH_CBT, self.hook_proc, ctypes.windll.kernel32.GetModuleHandleW(None), 0)
            if not self.hook: 
                raise RuntimeError("Hook failed.")
            self._log("Real-time monitoring enabled.", LogLevel.SUCCESS)
        except Exception: 
            self._start_fallback_monitoring(check_cb)
    def _start_fallback_monitoring(self, check_cb: Callable):
        def loop():
            last_info = None
            while True:
                info = self.get_active_window_info()
                if info and info[0] != (last_info and last_info[0]): 
                    last_info = info
                    AppUI.instance.after(0, check_cb)
                time.sleep(2)
        threading.Thread(target=loop, daemon=True).start()
    def cleanup(self):
        if self.hook and WINDOWS_API_AVAILABLE: 
            ctypes.windll.user32.UnhookWindowsHookEx(self.hook)

class AutoSaveScript:
    def __init__(self, log_cb: Callable, status_cb: Callable):
        self._log, self._update_status = log_cb, status_cb
        self.settings_manager = SettingsManager()
        self.process_monitor = ProcessMonitor(self._log)
        self.startup_manager = WindowsStartupManager()
        self.settings = self.settings_manager.load()
        self.current_app: Optional[str] = None
        self._save_timer: Optional[threading.Timer] = None
        self._backup_timer: Optional[threading.Timer] = None
    def start(self): 
        self._log("Personal Assistant started.", LogLevel.STARTUP)
        self.process_monitor.start_monitoring(self._check_active_window)
    def _is_target(self, name: str) -> bool: 
        return name in self.settings.get('monitored_apps', [])
    def _check_active_window(self):
        info = self.process_monitor.get_active_window_info()
        if info and self._is_target(info[0]):
            app_name = info[0].replace('.exe', '').title()
            if self.current_app != app_name:
                self.current_app = app_name
                self.settings = self.settings_manager.load()
                if self.settings.get('auto_save_enabled', True): 
                    self._update_status("Active", "ACTIVE", app_name)
                else: 
                    self._update_status("Paused", "PAUSED", app_name)
                self._log(f"Active application: {app_name}", LogLevel.ACTIVE)
                self._start_timers()
        elif self.current_app:
            self._log(f"Stopped monitoring {self.current_app}", LogLevel.INFO)
            self._update_status("Waiting", "WAITING")
            self.current_app = None
            self._stop_timers()
    def _start_timers(self):
        self._stop_timers()
        if self.settings.get('auto_save_enabled', True): 
            self._start_save_timer()
        if self.settings.get('smart_backup_enabled', False): 
            self._start_backup_timer()
    def _stop_timers(self): 
        self._stop_save_timer()
        self._stop_backup_timer()
    def _start_save_timer(self):
        self._stop_save_timer()
        def task():
            if self.current_app and self.settings.get('auto_save_enabled', True):
                try: 
                    pyautogui.hotkey('ctrl', 's')
                except Exception: 
                    self._log("Auto-Save command failed.", LogLevel.ERROR)
                self._log(f"Auto-saved in {self.current_app}", LogLevel.SAVE)
                self._update_status("Saved!", "SAVED")
                if self.current_app: 
                    self._start_save_timer()
        self._save_timer = threading.Timer(self.settings['auto_save_interval'], task)
        self._save_timer.start()
    def _stop_save_timer(self):
        if self._save_timer: 
            self._save_timer.cancel()
            self._save_timer = None
    def _start_backup_timer(self):
        self._stop_backup_timer()
        def task():
            if self.current_app and self.settings.get('smart_backup_enabled', False):
                self._create_backup()
                if self.current_app: 
                    self._start_backup_timer()
        self._backup_timer = threading.Timer(self.settings.get('smart_backup_interval', 60) * 60, task)
        self._backup_timer.start()
    def _stop_backup_timer(self):
        if self._backup_timer: 
            self._backup_timer.cancel()
            self._backup_timer = None
    def _create_backup(self):
        info = self.process_monitor.get_active_window_info()
        if not info or not info[1] or '*' not in info[1]: 
            return
        try:
            title = info[1].split(' @')[0].split(' - ')[0].replace('*', '').strip()
            original_path = Path(title)
            if not original_path.is_file(): 
                self._log("Backup failed: Cannot determine file path.", LogLevel.WARNING)
                return
            backup_dir = original_path.parent / "Smart Backup"
            backup_dir.mkdir(exist_ok=True)
            timestamp = time.strftime("%Y%m%d-%H%M%S")
            backup_filename = f"{original_path.stem}-{timestamp}{original_path.suffix}"
            backup_path = backup_dir / backup_filename
            pyautogui.hotkey('ctrl', 'shift', 's')
            time.sleep(1)
            pyautogui.write(str(backup_path), interval=0.01)
            pyautogui.press('enter')
            self._log(f"Smart Backup created: {backup_filename}", LogLevel.SUCCESS)
            self._update_status("Backup!", "SAVED")
        except Exception: 
            self._log("Backup failed unexpectedly.", LogLevel.ERROR)
    def update_settings(self, new_settings):
        self.settings.update(new_settings)
        if self.settings_manager.save():
            self._log("Settings updated.", LogLevel.SUCCESS)
            self.startup_manager.set_startup(self.settings.get('start_with_windows', False))
        else: 
            self._log("Failed to save settings.", LogLevel.ERROR)
        if self.current_app:
            self._stop_timers()
            self._start_timers()
            if self.settings.get('auto_save_enabled', True): 
                self._update_status("Active", "ACTIVE", self.current_app)
            else: 
                self._update_status("Paused", "PAUSED", self.current_app)
    def add_monitored_app(self, app_path: str):
        app_name = Path(app_path).name.lower()
        if app_name and app_name not in self.settings['monitored_apps']:
            self.settings['monitored_apps'].append(app_name)
            if self.settings_manager.save(): 
                self._log(f"Added to watchlist: {app_name}", LogLevel.SUCCESS)
            else: 
                self._log("Failed to save new app.", LogLevel.ERROR)
    def cleanup(self): 
        self._stop_timers()
        self.process_monitor.cleanup()
        self._log("Personal Assistant has been shut down.", LogLevel.INFO)

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
        except Exception: 
            pass
    
    app = AppUI()
    app.mainloop()