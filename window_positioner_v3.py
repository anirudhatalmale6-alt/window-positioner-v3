"""
Window Positioner v4 - With Resize, Zoom Control + Go to Link
Auto-arrange browser profile windows in a grid + Apply zoom to all profiles
+ Open URL in new tabs for all profiles + Resize windows to fixed size
"""

import ctypes
from ctypes import wintypes
import math
import sys
import os
import time
import re
import json
import threading

# Check and install dependencies
def install_deps():
    import subprocess
    try:
        import pystray
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pystray", "pillow"])
    try:
        import keyboard
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "keyboard"])
install_deps()

import tkinter as tk
from tkinter import ttk, messagebox
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item
import keyboard

# Windows API
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
EnumWindows = user32.EnumWindows
EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
GetWindowTextW = user32.GetWindowTextW
GetWindowTextLengthW = user32.GetWindowTextLengthW
IsWindowVisible = user32.IsWindowVisible
SetWindowPos = user32.SetWindowPos
ShowWindow = user32.ShowWindow
GetClassName = user32.GetClassNameW
GetWindowThreadProcessId = user32.GetWindowThreadProcessId
SetForegroundWindow = user32.SetForegroundWindow
SW_RESTORE = 9

# Settings file
SETTINGS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "settings.json")

DEFAULT_SETTINGS = {
    "grid_cols": 0,
    "grid_rows": 0,
    "h_gap": 10,
    "v_gap": 10,
    "hotkey": "ctrl+shift+p",
    "zoom_level": 35,
    "window_width": 550,
    "window_height": 600
}

PROFILE_INDICATORS = ['whoerip', 'whoer', 'mimic']
EXCLUDE_INDICATORS = ['multilogin x app', 'multilogin app']
PROFILE_PATTERNS = [r'\b[A-Z]{2}\d+\b']


def load_settings():
    try:
        if os.path.exists(SETTINGS_FILE):
            with open(SETTINGS_FILE, 'r') as f:
                saved = json.load(f)
                settings = DEFAULT_SETTINGS.copy()
                settings.update(saved)
                return settings
    except:
        pass
    return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    try:
        with open(SETTINGS_FILE, 'w') as f:
            json.dump(settings, f, indent=2)
    except:
        pass


class WindowPositioner:
    def __init__(self, settings):
        self.settings = settings

    def get_work_area(self):
        """Get the work area of the PRIMARY monitor only (not all monitors)"""
        # Get primary monitor info
        SM_CXSCREEN = 0  # Primary screen width
        SM_CYSCREEN = 1  # Primary screen height

        # Get primary monitor dimensions
        screen_w = user32.GetSystemMetrics(SM_CXSCREEN)
        screen_h = user32.GetSystemMetrics(SM_CYSCREEN)

        # Get work area (screen minus taskbar) for primary monitor
        class RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                        ('right', ctypes.c_long), ('bottom', ctypes.c_long)]
        rect = RECT()
        user32.SystemParametersInfoW(0x0030, 0, ctypes.byref(rect), 0)

        # Use the smaller of screen size or work area to stay on primary monitor
        work_w = min(rect.right - rect.left, screen_w)
        work_h = min(rect.bottom - rect.top, screen_h)

        return work_w, work_h

    def get_window_title(self, hwnd):
        length = GetWindowTextLengthW(hwnd)
        if length == 0:
            return ""
        buffer = ctypes.create_unicode_buffer(length + 1)
        GetWindowTextW(hwnd, buffer, length + 1)
        return buffer.value

    def get_window_class(self, hwnd):
        buffer = ctypes.create_unicode_buffer(256)
        GetClassName(hwnd, buffer, 256)
        return buffer.value

    def get_process_creation_time(self, hwnd):
        pid = wintypes.DWORD()
        GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not handle:
            return 0
        class FILETIME(ctypes.Structure):
            _fields_ = [('dwLowDateTime', wintypes.DWORD), ('dwHighDateTime', wintypes.DWORD)]
        creation_time = FILETIME()
        exit_time = FILETIME()
        kernel_time = FILETIME()
        user_time = FILETIME()
        result = kernel32.GetProcessTimes(handle, ctypes.byref(creation_time), ctypes.byref(exit_time),
                                          ctypes.byref(kernel_time), ctypes.byref(user_time))
        kernel32.CloseHandle(handle)
        if result:
            return (creation_time.dwHighDateTime << 32) | creation_time.dwLowDateTime
        return 0

    def is_profile_window(self, hwnd):
        if not IsWindowVisible(hwnd):
            return False
        title = self.get_window_title(hwnd)
        title_lower = title.lower()
        class_name = self.get_window_class(hwnd).lower()
        if not title:
            return False
        for exclude in EXCLUDE_INDICATORS:
            if exclude in title_lower:
                return False
        for indicator in PROFILE_INDICATORS:
            if indicator in title_lower:
                return True
        for pattern in PROFILE_PATTERNS:
            if re.search(pattern, title):
                if 'chrome_widgetwin' in class_name:
                    return True
        return False

    def get_profile_windows(self):
        windows = []
        def callback(hwnd, lParam):
            if self.is_profile_window(hwnd):
                creation_time = self.get_process_creation_time(hwnd)
                windows.append((hwnd, self.get_window_title(hwnd), creation_time))
            return True
        EnumWindows(EnumWindowsProc(callback), 0)
        windows.sort(key=lambda x: x[2])
        return [(hwnd, title) for hwnd, title, _ in windows]

    def get_window_size(self, hwnd):
        class RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                        ('right', ctypes.c_long), ('bottom', ctypes.c_long)]
        rect = RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        return rect.right - rect.left, rect.bottom - rect.top

    def position_windows(self, resize=False):
        """Position windows in a grid. If resize=True, also resize to saved size."""
        windows = self.get_profile_windows()
        if not windows:
            return 0

        work_w, work_h = self.get_work_area()
        h_gap = self.settings["h_gap"]
        v_gap = self.settings["v_gap"]
        num_windows = len(windows)

        # Get grid settings
        cols = self.settings["grid_cols"]
        rows = self.settings["grid_rows"]

        # Use saved window size from settings
        win_w = self.settings.get("window_width", 550)
        win_h = self.settings.get("window_height", 600)

        # Auto calculate grid if set to 0
        if cols == 0:
            cols = max(1, work_w // (win_w + h_gap))
        if rows == 0:
            rows = max(1, work_h // (win_h + v_gap))

        # Make sure we have enough cells
        while cols * rows < num_windows:
            cols += 1

        SWP_NOZORDER = 0x0004

        for i, (hwnd, title) in enumerate(windows):
            col = i % cols
            row = i // cols
            x = col * (win_w + h_gap)
            y = row * (win_h + v_gap)

            ShowWindow(hwnd, SW_RESTORE)
            time.sleep(0.03)
            # Position AND resize window
            SetWindowPos(hwnd, None, x, y, win_w, win_h, SWP_NOZORDER)

        return num_windows

    def apply_zoom_to_all(self, zoom_percent):
        """Apply zoom to all profile windows using Ctrl+0 (reset) then Ctrl+- to adjust"""
        windows = self.get_profile_windows()
        if not windows:
            return 0

        # Zoom levels roughly: 100% -> 90% -> 80% -> 75% -> 67% -> 50% -> 33% -> 25%
        zoom_steps = {
            100: 0,   # Ctrl+0 only
            90: 1,    # Ctrl+0 then 1x Ctrl+-
            80: 2,    # Ctrl+0 then 2x Ctrl+-
            75: 3,    # Ctrl+0 then 3x Ctrl+-
            67: 4,    # Ctrl+0 then 4x Ctrl+-
            50: 5,    # Ctrl+0 then 5x Ctrl+-
            33: 6,    # Ctrl+0 then 6x Ctrl+-
            25: 7,    # Ctrl+0 then 7x Ctrl+-
            35: 6,    # Approximate
        }

        # Find closest zoom level
        closest = min(zoom_steps.keys(), key=lambda x: abs(x - zoom_percent))
        steps = zoom_steps[closest]

        for hwnd, title in windows:
            try:
                # Bring window to front
                ShowWindow(hwnd, SW_RESTORE)
                SetForegroundWindow(hwnd)
                time.sleep(0.15)

                # Reset zoom to 100% (Ctrl+0) - using keyboard module (doesn't move mouse)
                keyboard.press_and_release('ctrl+0')
                time.sleep(0.08)

                # Apply zoom out steps
                for _ in range(steps):
                    keyboard.press_and_release('ctrl+-')
                    time.sleep(0.08)

            except Exception as e:
                print(f"Error zooming {title}: {e}")

        return len(windows)

    def resize_all_windows(self, width, height):
        """Resize all profile windows to fixed size - keeps position"""
        windows = self.get_profile_windows()
        if not windows:
            return 0

        SWP_NOZORDER = 0x0004

        class RECT(ctypes.Structure):
            _fields_ = [('left', ctypes.c_long), ('top', ctypes.c_long),
                        ('right', ctypes.c_long), ('bottom', ctypes.c_long)]

        for hwnd, title in windows:
            try:
                ShowWindow(hwnd, SW_RESTORE)
                time.sleep(0.05)
                # Get current position
                rect = RECT()
                user32.GetWindowRect(hwnd, ctypes.byref(rect))
                x = rect.left
                y = rect.top
                # Set new size while keeping the same position
                SetWindowPos(hwnd, None, x, y, width, height, SWP_NOZORDER)
            except Exception as e:
                print(f"Error resizing {title}: {e}")

        return len(windows)

    def open_url_in_all(self, url, apply_zoom_after=False, zoom_percent=50):
        """Open URL in new tab for all profile windows"""
        windows = self.get_profile_windows()
        if not windows:
            return 0

        if not url or not url.strip():
            return 0

        url = url.strip()
        # Add https:// if no protocol specified
        if not url.startswith('http://') and not url.startswith('https://'):
            url = 'https://' + url

        # First pass: Open URL in all windows
        for hwnd, title in windows:
            try:
                # Bring window to front
                ShowWindow(hwnd, SW_RESTORE)
                SetForegroundWindow(hwnd)
                time.sleep(0.15)

                # Open new tab with Ctrl+T
                keyboard.press_and_release('ctrl+t')
                time.sleep(0.3)

                # Type the URL
                keyboard.write(url, delay=0.01)
                time.sleep(0.15)

                # Press Enter to navigate
                keyboard.press_and_release('enter')
                time.sleep(0.2)

            except Exception as e:
                print(f"Error opening URL in {title}: {e}")

        # Second pass: Apply zoom after pages have started loading
        if apply_zoom_after:
            time.sleep(2)  # Wait for pages to start loading
            self.apply_zoom_to_all(zoom_percent)

        return len(windows)


def create_tray_icon():
    img = Image.new('RGB', (64, 64), color=(70, 130, 180))
    draw = ImageDraw.Draw(img)
    draw.rectangle([8, 8, 28, 28], fill='white')
    draw.rectangle([36, 8, 56, 28], fill='white')
    draw.rectangle([8, 36, 28, 56], fill='white')
    draw.rectangle([36, 36, 56, 56], fill='white')
    return img


class App:
    def __init__(self):
        self.settings = load_settings()
        self.positioner = WindowPositioner(self.settings)
        self.tray_icon = None
        self.root = None
        self.hidden = False
        self.create_window()
        self.setup_hotkey()

    def create_window(self):
        self.root = tk.Tk()
        self.root.title("Window Positioner v4")
        self.root.geometry("320x560")
        self.root.resizable(False, False)
        self.root.protocol("WM_DELETE_WINDOW", self.minimize_to_tray)

        main = ttk.Frame(self.root, padding="15")
        main.pack(fill=tk.BOTH, expand=True)

        # Title
        ttk.Label(main, text="Window Positioner", font=('Helvetica', 12, 'bold')).pack(pady=(0, 10))

        # Position Settings
        settings = ttk.LabelFrame(main, text="Position Settings", padding="10")
        settings.pack(fill=tk.X, pady=(0, 10))

        # Grid layout
        row1 = ttk.Frame(settings)
        row1.pack(fill=tk.X, pady=3)
        ttk.Label(row1, text="Grid:").pack(side=tk.LEFT)
        self.cols_var = tk.StringVar(value=str(self.settings["grid_cols"]))
        ttk.Entry(row1, textvariable=self.cols_var, width=4).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="x").pack(side=tk.LEFT)
        self.rows_var = tk.StringVar(value=str(self.settings["grid_rows"]))
        ttk.Entry(row1, textvariable=self.rows_var, width=4).pack(side=tk.LEFT, padx=5)
        ttk.Label(row1, text="(0=auto)").pack(side=tk.LEFT, padx=5)

        # Gap
        row2 = ttk.Frame(settings)
        row2.pack(fill=tk.X, pady=3)
        ttk.Label(row2, text="Gap:").pack(side=tk.LEFT)
        self.hgap_var = tk.StringVar(value=str(self.settings["h_gap"]))
        ttk.Entry(row2, textvariable=self.hgap_var, width=4).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="x").pack(side=tk.LEFT)
        self.vgap_var = tk.StringVar(value=str(self.settings["v_gap"]))
        ttk.Entry(row2, textvariable=self.vgap_var, width=4).pack(side=tk.LEFT, padx=5)
        ttk.Label(row2, text="px").pack(side=tk.LEFT, padx=5)

        # Position Button inside settings frame
        ttk.Button(settings, text="POSITION WINDOWS",
                   command=self.position_windows).pack(fill=tk.X, pady=5)

        # Zoom Settings
        zoom_frame = ttk.LabelFrame(main, text="Zoom Control", padding="10")
        zoom_frame.pack(fill=tk.X, pady=(0, 10))

        # Zoom slider
        zoom_row = ttk.Frame(zoom_frame)
        zoom_row.pack(fill=tk.X, pady=3)
        ttk.Label(zoom_row, text="Zoom:").pack(side=tk.LEFT)
        self.zoom_var = tk.IntVar(value=self.settings.get("zoom_level", 35))
        self.zoom_slider = ttk.Scale(zoom_row, from_=25, to=100, variable=self.zoom_var,
                                     orient=tk.HORIZONTAL, length=150, command=self.update_zoom_label)
        self.zoom_slider.pack(side=tk.LEFT, padx=5)
        self.zoom_label = ttk.Label(zoom_row, text=f"{self.zoom_var.get()}%", width=5)
        self.zoom_label.pack(side=tk.LEFT)

        # Preset zoom buttons
        preset_row = ttk.Frame(zoom_frame)
        preset_row.pack(fill=tk.X, pady=5)
        for zoom in [25, 35, 50, 75, 100]:
            btn = ttk.Button(preset_row, text=f"{zoom}%", width=5,
                           command=lambda z=zoom: self.set_zoom_preset(z))
            btn.pack(side=tk.LEFT, padx=2)

        # Apply Zoom Button
        ttk.Button(zoom_frame, text="APPLY ZOOM TO ALL",
                   command=self.apply_zoom_all).pack(fill=tk.X, pady=5)

        # Window Resize Settings
        resize_frame = ttk.LabelFrame(main, text="Window Size", padding="10")
        resize_frame.pack(fill=tk.X, pady=(0, 10))

        # Size inputs
        size_row = ttk.Frame(resize_frame)
        size_row.pack(fill=tk.X, pady=3)
        ttk.Label(size_row, text="Size:").pack(side=tk.LEFT)
        self.width_var = tk.StringVar(value=str(self.settings.get("window_width", 550)))
        ttk.Entry(size_row, textvariable=self.width_var, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(size_row, text="x").pack(side=tk.LEFT)
        self.height_var = tk.StringVar(value=str(self.settings.get("window_height", 600)))
        ttk.Entry(size_row, textvariable=self.height_var, width=5).pack(side=tk.LEFT, padx=5)
        ttk.Label(size_row, text="px").pack(side=tk.LEFT, padx=5)

        # Resize Button
        ttk.Button(resize_frame, text="RESIZE ALL WINDOWS",
                   command=self.resize_all).pack(fill=tk.X, pady=5)

        # Go to Link Settings
        link_frame = ttk.LabelFrame(main, text="Go to Link", padding="10")
        link_frame.pack(fill=tk.X, pady=(0, 10))

        # URL input
        url_row = ttk.Frame(link_frame)
        url_row.pack(fill=tk.X, pady=3)
        ttk.Label(url_row, text="URL:").pack(side=tk.LEFT)
        self.url_var = tk.StringVar()
        self.url_entry = ttk.Entry(url_row, textvariable=self.url_var, width=30)
        self.url_entry.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)

        # Auto-apply zoom checkbox
        self.auto_zoom_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(link_frame, text="Apply zoom after page loads",
                        variable=self.auto_zoom_var).pack(anchor=tk.W, pady=3)

        # Go to Link Button
        ttk.Button(link_frame, text="OPEN URL IN ALL PROFILES",
                   command=self.open_url_all).pack(fill=tk.X, pady=5)

        # Minimize button
        ttk.Button(main, text="Minimize to Tray", command=self.minimize_to_tray).pack(fill=tk.X, pady=5)

        # Status
        self.status_var = tk.StringVar(value="Hotkey: CTRL+SHIFT+P")
        ttk.Label(main, textvariable=self.status_var, font=('Helvetica', 8)).pack(pady=(10, 0))

    def update_zoom_label(self, val):
        self.zoom_label.config(text=f"{int(float(val))}%")

    def set_zoom_preset(self, zoom):
        self.zoom_var.set(zoom)
        self.zoom_label.config(text=f"{zoom}%")

    def apply_zoom_all(self):
        self.apply_settings()
        zoom = self.zoom_var.get()
        self.status_var.set(f"Applying {zoom}% zoom...")
        self.root.update()

        # Run in thread to not block UI
        def zoom_thread():
            count = self.positioner.apply_zoom_to_all(zoom)
            self.root.after(0, lambda: self.status_var.set(f"Done! Zoom applied to {count} windows"))

        threading.Thread(target=zoom_thread, daemon=True).start()

    def resize_all(self):
        try:
            width = int(self.width_var.get())
            height = int(self.height_var.get())
        except ValueError:
            self.status_var.set("Invalid size values")
            return

        self.settings["window_width"] = width
        self.settings["window_height"] = height
        save_settings(self.settings)

        self.status_var.set(f"Resizing to {width}x{height}...")
        self.root.update()

        # Run in thread to not block UI
        def resize_thread():
            count = self.positioner.resize_all_windows(width, height)
            self.root.after(0, lambda: self.status_var.set(f"Done! {count} windows resized to {width}x{height}"))

        threading.Thread(target=resize_thread, daemon=True).start()

    def open_url_all(self):
        url = self.url_var.get().strip()
        if not url:
            self.status_var.set("Please enter a URL")
            return

        apply_zoom = self.auto_zoom_var.get()
        zoom = self.zoom_var.get()

        if apply_zoom:
            self.status_var.set(f"Opening URL + applying zoom...")
        else:
            self.status_var.set(f"Opening URL in all profiles...")
        self.root.update()

        # Run in thread to not block UI
        def url_thread():
            count = self.positioner.open_url_in_all(url, apply_zoom_after=apply_zoom, zoom_percent=zoom)
            if apply_zoom:
                self.root.after(0, lambda: self.status_var.set(f"Done! URL opened + zoom applied to {count} windows"))
            else:
                self.root.after(0, lambda: self.status_var.set(f"Done! URL opened in {count} windows"))

        threading.Thread(target=url_thread, daemon=True).start()

    def setup_hotkey(self):
        try:
            keyboard.add_hotkey(self.settings["hotkey"], self.hotkey_triggered)
        except:
            pass

    def hotkey_triggered(self):
        self.root.after(0, self.position_windows)

    def apply_settings(self):
        try:
            self.settings["grid_cols"] = int(self.cols_var.get())
            self.settings["grid_rows"] = int(self.rows_var.get())
            self.settings["h_gap"] = int(self.hgap_var.get())
            self.settings["v_gap"] = int(self.vgap_var.get())
            self.settings["zoom_level"] = self.zoom_var.get()
            save_settings(self.settings)
            self.positioner.settings = self.settings
            return True
        except:
            return False

    def position_windows(self):
        self.apply_settings()
        self.status_var.set("Positioning...")
        self.root.update()
        count = self.positioner.position_windows()
        if count == 0:
            self.status_var.set("No windows found")
        else:
            self.status_var.set(f"Done! {count} windows positioned")

    def minimize_to_tray(self):
        self.hidden = True
        self.root.withdraw()
        if self.tray_icon is None:
            menu = pystray.Menu(
                item('Position Windows', self.tray_position),
                item('Resize Windows', self.tray_resize),
                item('Apply Zoom', self.tray_zoom),
                item('Open URL', self.tray_open_url),
                item('Show', self.show_window),
                item('Exit', self.quit_app)
            )
            self.tray_icon = pystray.Icon("WindowPositioner", create_tray_icon(), "Window Positioner", menu)
            threading.Thread(target=self.tray_icon.run, daemon=True).start()

    def tray_position(self):
        self.position_windows()

    def tray_resize(self):
        width = self.settings.get("window_width", 550)
        height = self.settings.get("window_height", 600)
        self.positioner.resize_all_windows(width, height)

    def tray_zoom(self):
        zoom = self.settings.get("zoom_level", 35)
        self.positioner.apply_zoom_to_all(zoom)

    def tray_open_url(self):
        url = self.url_var.get().strip()
        if url:
            self.positioner.open_url_in_all(url)

    def show_window(self):
        self.hidden = False
        self.root.deiconify()
        self.root.lift()

    def quit_app(self):
        if self.tray_icon:
            self.tray_icon.stop()
        self.root.destroy()

    def run(self):
        self.root.mainloop()


if __name__ == "__main__":
    App().run()