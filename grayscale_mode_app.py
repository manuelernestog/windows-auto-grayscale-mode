import winreg
import tkinter as tk
from tkinter import ttk, messagebox, BooleanVar
import threading
import time
import sys
import os
import pystray
from PIL import Image, ImageDraw
from datetime import datetime, time as dt_time
import json
import pyautogui
import pathlib

REG_PATH = r"Software\Microsoft\ColorFiltering"
REG_VALUE_NAME = "Active"

APP_CONFIG_DIR_NAME = "Windows Auto Grayscale Mode"
CONFIG_FILE_NAME = "schedule_config.json"

def get_config_path():
    """Gets the full path to the schedule configuration file in AppData."""
    appdata_path = pathlib.Path(os.environ['APPDATA'])
    config_dir = appdata_path / APP_CONFIG_DIR_NAME
    config_file_path = config_dir / CONFIG_FILE_NAME
    return config_file_path

def set_grayscale(active):
    """Toggles the grayscale filter by simulating the hotkey."""
    try:
        time.sleep(0.1)

        pyautogui.hotkey('win', 'ctrl', 'c')
        print("Simulated Windows + Control + C via pyautogui.")

    except Exception as e:
        print(f"Error simulating hotkey: {e}")

def get_grayscale_status():
    """Gets the current status of the grayscale filter from the registry."""
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, REG_PATH, 0, winreg.KEY_READ)
        value, reg_type = winreg.QueryValueEx(key, REG_VALUE_NAME)
        winreg.CloseKey(key)
        return value == 1
    except FileNotFoundError:
        return False
    except Exception as e:
        print(f"Error getting grayscale status: {e}")
        return False 

def load_schedule_config():
    """Loads schedule configuration from the JSON file in AppData."""
    config_path = get_config_path()
    config = {"start_time": "22:00", "end_time": "06:00", "enabled": False}

    if config_path.exists():
        try:
            with open(config_path, 'r') as f:
                loaded_config = json.load(f)
                config.update(loaded_config)
        except Exception as e:
            print(f"Error loading config file from {config_path}: {e}")
    else:
        old_config_path = pathlib.Path(CONFIG_FILE_NAME)
        if old_config_path.exists():
             print(f"Migrating config from old location: {old_config_path}")
             try:
                 with open(old_config_path, 'r') as f:
                     loaded_config = json.load(f)
                     config.update(loaded_config)
                 save_schedule_config(config)
                 old_config_path.unlink() 
             except Exception as e:
                 print(f"Error migrating config file: {e}")


    return config

def save_schedule_config(config):
    """Saves schedule configuration to the JSON file in AppData."""
    config_path = get_config_path()
    config_dir = config_path.parent

    try:
        config_dir.mkdir(parents=True, exist_ok=True)
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=4)
    except Exception as e:
        print(f"Error saving config file to {config_path}: {e}")


class RestModeApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Windows Auto Grayscale Mode")
        self.root.geometry("300x350")
        self.root.resizable(False, False) 

        try:
            icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
            if getattr(sys, 'frozen', False):
                 icon_path = os.path.join(sys._MEIPASS, "icon.png")
            self.root.iconphoto(True, tk.PhotoImage(file=icon_path))
        except Exception as e:
            print(f"Error setting window icon: {e}")

        self.schedule_config = load_schedule_config()
        self.is_scheduled = self.schedule_config.get("enabled", False)
        self.schedule_thread = None
        self.stop_schedule_event = threading.Event()

        self.icon = None
        self.setup_tray_icon()

        manual_frame = ttk.LabelFrame(root, text="Manual Control")
        manual_frame.pack(pady=10, padx=10, fill="x")

        self.status_label = ttk.Label(manual_frame, text="Status: Unknown")
        self.status_label.pack(pady=5)

        button_frame = ttk.Frame(manual_frame)
        button_frame.pack(pady=5)

        self.toggle_button = ttk.Button(button_frame, text="Toggle Grayscale", command=self.activate_manual) 
        self.toggle_button.pack(pady=5)

        schedule_frame = ttk.LabelFrame(root, text="Automatic Scheduling")
        schedule_frame.pack(pady=10, padx=10, fill="x")

        self.schedule_enabled_var = BooleanVar(value=self.is_scheduled)
        self.schedule_checkbox = ttk.Checkbutton(schedule_frame, text="Enable Scheduling", variable=self.schedule_enabled_var, command=self.toggle_schedule_enabled)
        self.schedule_checkbox.grid(row=0, column=0, columnspan=2, padx=5, pady=5, sticky="w")


        ttk.Label(schedule_frame, text="Start Time (HH:MM):").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.start_time_entry = ttk.Entry(schedule_frame, width=5)
        self.start_time_entry.insert(0, self.schedule_config.get("start_time", "22:00"))
        self.start_time_entry.grid(row=1, column=1, padx=5, pady=5, sticky="w")

        ttk.Label(schedule_frame, text="End Time (HH:MM):").grid(row=2, column=0, padx=5, pady=5, sticky="w")
        self.end_time_entry = ttk.Entry(schedule_frame, width=5)
        self.end_time_entry.insert(0, self.schedule_config.get("end_time", "06:00"))
        self.end_time_entry.grid(row=2, column=1, padx=5, pady=5, sticky="w")

        self.start_time_entry.bind("<KeyRelease>", self.save_schedule_config_callback)
        self.end_time_entry.bind("<KeyRelease>", self.save_schedule_config_callback)

        info_about_frame = ttk.Frame(root)
        info_about_frame.pack(pady=10)

        info_button = ttk.Button(info_about_frame, text="Info", command=self.show_info)
        info_button.pack(side="left", padx=5)

        about_button = ttk.Button(info_about_frame, text="About", command=self.show_about)
        about_button.pack(side="left", padx=5)

        self.update_status()

        self.root.protocol("WM_DELETE_WINDOW", self.hide_window)

        if self.is_scheduled:
            self.start_scheduling_thread()


    def setup_tray_icon(self):
        icon_path = os.path.join(os.path.dirname(__file__), "icon.png")
        if getattr(sys, 'frozen', False):
                icon_path = os.path.join(sys._MEIPASS, "icon.png")
        image = Image.open(icon_path)

        menu = (pystray.MenuItem('Show', self.show_window),
                pystray.MenuItem('Toggle Grayscale', self.activate_manual), 
                pystray.MenuItem('About', self.show_about),
                pystray.MenuItem(text='Start with Windows', action=self.add_to_startup, visible=not self.is_running_on_startup()),
                pystray.MenuItem(text='Do not start with Windows', action=self.remove_from_startup, visible=self.is_running_on_startup()), 
                pystray.MenuItem('Exit', self.quit_app))

        self.icon = pystray.Icon("Windows Auto Grayscale Mode", image, "Windows Auto Grayscale Mode", menu, on_activate=self.show_window)

        threading.Thread(target=self.icon.run, daemon=True).start()

    def show_window(self):
        """Shows the main application window."""
        self.root.deiconify()
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after_idle(self.root.attributes, '-topmost', False)

    def hide_window(self):
        """Hides the main application window to the tray."""
        self.root.withdraw()

    def quit_app(self):
        """Quits the application."""
        self.stop_scheduling_thread() 
        if self.icon:
            self.icon.stop()
        self.root.quit() 

    def update_status(self):
        """Updates the status label based on the current grayscale state."""
        status = "Active" if get_grayscale_status() else "Inactive"
        self.status_label.config(text=f"Status: {status}")
        self.root.after(5000, self.update_status)

    def activate_manual(self):
        """Triggers the grayscale filter toggle via hotkey simulation."""
        set_grayscale(True)

    def toggle_schedule_enabled(self):
        """Handles the checkbox state change for scheduling."""
        self.is_scheduled = self.schedule_enabled_var.get()
        self.schedule_config["enabled"] = self.is_scheduled
        save_schedule_config(self.schedule_config)

        if self.is_scheduled:
            self.start_scheduling_thread()
        else:
            self.stop_scheduling_thread()

    def start_scheduling_thread(self):
        """Starts the background thread for scheduling."""
        if self.schedule_thread is not None and self.schedule_thread.is_alive():
            print("Scheduling thread is already running.")
            return

        start_time_str = self.start_time_entry.get()
        end_time_str = self.end_time_entry.get()

        try:
            time.strptime(start_time_str, '%H:%M')
            time.strptime(end_time_str, '%H:%M')
        except ValueError:
            messagebox.showerror("Format Error", "Please enter the time in HH:MM format (e.g. 08:00)")
            self.schedule_enabled_var.set(False) 
            self.is_scheduled = False
            self.schedule_config["enabled"] = False
            save_schedule_config(self.schedule_config)
            return

        self.schedule_config["start_time"] = start_time_str
        self.schedule_config["end_time"] = end_time_str
        save_schedule_config(self.schedule_config)

        self.stop_schedule_event.clear()
        self.schedule_thread = threading.Thread(target=self.run_schedule, args=(start_time_str, end_time_str))
        self.schedule_thread.daemon = True
        self.schedule_thread.start()
        print("Scheduling thread started.")

    def stop_scheduling_thread(self):
        """Stops the background thread for scheduling."""
        if self.schedule_thread is not None and self.schedule_thread.is_alive():
            self.stop_schedule_event.set()
            print("Scheduling thread stop requested.")
        self.schedule_thread = None 
        print("Scheduling thread stopped.")

    def save_schedule_config_callback(self, event=None):
        """Reads time entries and saves the schedule configuration."""
        self.schedule_config["start_time"] = self.start_time_entry.get()
        self.schedule_config["end_time"] = self.end_time_entry.get()
        save_schedule_config(self.schedule_config)
        print("Schedule times saved via callback.")


    def run_schedule(self, start_time_str, end_time_str):
        """Background thread function to check and apply schedule."""
        try:
            start_hour, start_minute = map(int, start_time_str.split(':'))
            end_hour, end_minute = map(int, end_time_str.split(':'))
            start_time_obj = dt_time(start_hour, start_minute)
            end_time_obj = dt_time(end_hour, end_minute)
        except ValueError:
            print("Error parsing schedule times in thread.")
            return

        print(f"Schedule thread running. Start: {start_time_str}, End: {end_time_str}")

        while not self.stop_schedule_event.is_set():
            now = datetime.now().time()

            if start_time_obj <= end_time_obj:
                should_be_active = start_time_obj <= now < end_time_obj
            else:
                should_be_active = now >= start_time_obj or now < end_time_obj

            current_status = get_grayscale_status()


            if should_be_active and not current_status:
                print("Scheduled activation triggered.")
                set_grayscale(True) 
            elif not should_be_active and current_status:
                print("Scheduled deactivation triggered.")
                set_grayscale(False) 


            # Sleep until the next minute starts or for a shorter interval
            # Check more frequently than once a minute to react faster to schedule changes
            time.sleep(30) # Check every 30 seconds

        print("Schedule thread finished.")


    def show_info(self):
        """Shows the Info dialog with instructions for enabling the hotkey."""
        info_message = (
            "Windows Auto Grayscale Mode requires the Windows Color Filters keyboard shortcut to be enabled.\n\n"
            "If the app is not working, please go to:\n"
            "System Settings > Accessibility > Color Filters\n"
            "and ensure 'Keyboard shortcut for color filters' is turned ON."
        )
        messagebox.showinfo("Windows Auto Grayscale Mode Info", info_message)

    def show_about(self):
        """Shows the About dialog."""
        messagebox.showinfo("About Windows Auto Grayscale Mode", "Windows Auto Grayscale Mode\nVersion: 1.0\nDeveloper: Manuel Ernesto Garcia")

    def on_closing(self):
        """Handles the window closing event."""
        if messagebox.askokcancel("Exit", "Do you want to close the application?"):
            self.stop_scheduling_thread() 
            self.root.destroy()

    def add_to_startup(self):
        """Adds the application to Windows startup."""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE)
            app_path = os.path.abspath(sys.argv[0])

            command = f'"{app_path}"'
            winreg.SetValueEx(key, "Windows Auto Grayscale Mode", 0, winreg.REG_SZ, command)
            winreg.CloseKey(key)
            print("Added to startup.")
            messagebox.showinfo("Automatic Startup", "Windows Auto Grayscale Mode has been added to Windows startup.")
            self.update_tray_menu() 
        except Exception as e:
            print(f"Error adding to startup: {e}")
            messagebox.showerror("Error", f"Could not add Windows Auto Grayscale Mode to Windows startup:\n{e}")

    def remove_from_startup(self):
        """Removes the application from Windows startup."""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_WRITE)
            winreg.DeleteValue(key, "Windows Auto Grayscale Mode")
            winreg.CloseKey(key)
            print("Removed from startup.")
            messagebox.showinfo("Automatic Startup", "Windows Auto Grayscale Mode has been removed from Windows startup.")
            self.update_tray_menu()
        except FileNotFoundError:
            print("Not found in startup.")
            messagebox.showinfo("Automatic Startup", "Windows Auto Grayscale Mode was not configured to start with Windows.")
        except Exception as e:
            print(f"Error removing from startup: {e}")
            messagebox.showerror("Error", f"Could not remove Windows Auto Grayscale Mode from Windows startup:\n{e}")

    def is_running_on_startup(self):
        """Checks if the application is configured to run on startup."""
        try:
            key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ)
            winreg.QueryValueEx(key, "Windows Auto Grayscale Mode")
            winreg.CloseKey(key)
            return True
        except FileNotFoundError:
            return False
        except Exception as e:
            print(f"Error checking startup status: {e}")
            return False

    def update_tray_menu(self):
        """Updates the tray icon menu based on current status."""
        if self.icon:
            menu = (pystray.MenuItem('Show', self.show_window),
                    pystray.MenuItem('Toggle Grayscale', self.activate_manual),
                    pystray.MenuItem('About', self.show_about),
                    pystray.MenuItem(text='Start with Windows', action=self.add_to_startup, visible=not self.is_running_on_startup()),
                    pystray.MenuItem(text='Do not start with Windows', action=self.remove_from_startup, visible=self.is_running_on_startup()), 
                    pystray.MenuItem('Exit', self.quit_app))

            self.icon.menu = menu


if __name__ == "__main__":
    root = tk.Tk()
    app = RestModeApp(root)
    root.withdraw()
    root.mainloop()
