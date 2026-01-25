"""
Controller Macro Runner (Camera + Serial + Script Engine + Editor)

Install:
  pip install numpy pillow pyserial pyaudio

FFmpeg:
  Ensure ffmpeg is on PATH:
    ffmpeg -version
"""
import os
import sys
import json
import re
import math
import time
import queue
import threading
import subprocess
import tkinter as tk
import copy
from tkinter import ttk, messagebox, filedialog, simpledialog
import numpy as np
from PIL import Image, ImageTk

from typing import Optional

# Import from local modules
import InputRedirection
import SerialController
import ScriptEngine
import ScriptToPy
from utils import (
    ffmpeg_path,
    safe_script_filename,
    list_script_files,
    list_com_ports,
    load_settings,
    save_settings,
    get_default_keybindings,
    normalize_theme_setting,
    resolve_theme_mode,
    is_python_available,
    is_ffmpeg_available,
    is_tesseract_available,
    download_ffmpeg,
    download_tesseract,
    get_ffmpeg_dir,
    get_tesseract_dir,
    FFMPEG_VERSION,
    TESSERACT_VERSION,
)
from camera import (
    list_dshow_video_devices,
    scale_image_to_fit,
    CameraPopoutWindow,
    RegionSelectorWindow,
    ColorPickerWindow,
)
from audio import (
    PYAUDIO_AVAILABLE,
    pyaudio,
    list_audio_devices,
)
from dialogs import CommandEditorDialog, SettingsDialog, PythonDownloadDialog, DependencyDownloadDialog

# Windows-specific flag to hide console window for subprocesses
_SUBPROCESS_FLAGS = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

THEME_COLORS = {
    "light": {
        "bg": "#f4f5f7",
        "panel": "#ffffff",
        "text": "#1f2328",
        "muted": "#5f6368",
        "border": "#c9cdd4",
        "accent": "#1f6feb",
        "entry_bg": "#ffffff",
        "button_bg": "#f6f7f9",
        "button_fg": "#1f2328",
        "select_bg": "#dbeafe",
        "select_fg": "#111827",
        "tree_bg": "#ffffff",
        "tree_fg": "#1f2328",
        "text_bg": "#fbfbfc",
        "text_fg": "#1f2328",
        "insert_fg": "#1f2328",
        "text_sel_bg": "#dbeafe",
        "text_sel_fg": "#111827",
        "pane_bg": "#e9eaee",
        "ip_bg": "#dbeafe",
        "comment_fg": "#228B22",
        "variable_fg": "#0066CC",
        "math_fg": "#b58900",
        "filepath_fg": "#CC0000",
        "selected_bg": "#e0e0e0",
    },
    "dark": {
        "bg": "#171c22",
        "panel": "#1f252d",
        "text": "#d8dde6",
        "muted": "#9aa3b0",
        "border": "#2d3541",
        "accent": "#6fa8ff",
        "entry_bg": "#232a33",
        "button_bg": "#27303a",
        "button_fg": "#d8dde6",
        "select_bg": "#2f3b4a",
        "select_fg": "#f5f7fa",
        "tree_bg": "#1a2028",
        "tree_fg": "#d8dde6",
        "text_bg": "#161c23",
        "text_fg": "#d8dde6",
        "insert_fg": "#f5f7fa",
        "text_sel_bg": "#2f3b4a",
        "text_sel_fg": "#f5f7fa",
        "pane_bg": "#14181e",
        "ip_bg": "#2a3646",
        "comment_fg": "#7bd88f",
        "variable_fg": "#7fb2ff",
        "math_fg": "#f2c374",
        "filepath_fg": "#ff6b6b",
        "selected_bg": "#24303d",
    },
}


# ----------------------------
# Tkinter App
# ----------------------------

class App:
    def __init__(self, root):
        self.root = root
        self._ui_thread = threading.current_thread()
        self.script_path = None
        self.dirty = False
        self._status_queue = queue.Queue()
        self._status_poll_ms = 50
        self._status_poll_id = None
        self._key_debug = os.environ.get("CMR_KEY_DEBUG", "").strip().lower() in (
            "1",
            "true",
            "yes",
            "on",
        )

        self.root.title("Controller Macro Runner")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.geometry("1280x700")
        try:
            self.root.iconbitmap("bin/icon.ico")
        except:
            pass

        # --- Load settings from file
        self._settings = load_settings()
        self._theme_setting = normalize_theme_setting(self._settings.get("theme", "auto"))
        self._resolved_theme = None
        self._theme_colors = None
        self._theme_poll_id = None
        self.style = ttk.Style(self.root)
        self.video_ratio_options = [
            "3:2 (GBA)",
            "16:9 (Standard)",
            "4:3 (DS Single Screen)",
            "2:3 (DS Dual Screen)",
            "5:3 (3DS Top Screen)",
            "5:6 (3DS Dual Screen)",
        ]
        self._initial_camera_ratio = str(self._settings.get("camera_ratio", "3:2 (GBA)")).strip()
        if self._initial_camera_ratio not in self.video_ratio_options:
            self._initial_camera_ratio = "3:2 (GBA)"
        self._settings["camera_ratio"] = self._initial_camera_ratio

        # --- keyboard controller mode (manual control)
        self.kb_enabled = tk.BooleanVar(value=False)
        self.kb_camera_focused = False  # True when a camera frame widget has focus for keyboard control
        self.kb_bindings = self._settings.get("keybindings", get_default_keybindings())
        self.kb_down = set()         # set of pressed Tk keysyms (normalized)
        self.kb_buttons_held = set() # controller buttons currently held due to keyboard
        self.kb_left_stick_dirs = set()
        self.kb_right_stick_dirs = set()

        # Global key events (manual controller)
        self.root.bind_all("<KeyPress>", self._on_key_press)
        self.root.bind_all("<KeyRelease>", self._on_key_release)

        # Disable keyboard focus on all buttons/checkbuttons to prevent
        # Space/Return from activating them while using keyboard control
        # Bind FocusIn to immediately move focus away from buttons
        # Only applies to main window buttons, not dialog buttons
        def _skip_button_focus(event):
            # Only skip focus for buttons in the main window, not dialogs
            try:
                widget = event.widget
                toplevel = widget.winfo_toplevel()
                # Only redirect focus if button is in the main root window
                if toplevel == self.root:
                    self.root.focus_set()
                    return "break"
            except tk.TclError:
                pass
            return None  # Allow normal focus handling for dialogs

        for widget_class in ("TButton", "Button", "TCheckbutton", "Checkbutton", "TRadiobutton", "Radiobutton"):
            self.root.bind_class(widget_class, "<FocusIn>", _skip_button_focus)

        # Global click event to detect focus loss (clicking outside camera frame)
        self.root.bind_all("<Button-1>", self._check_focus_loss, add="+")

        # camera state
        self.cam_width = 640
        self.cam_height = 426
        self.cam_fps = 30
        self.cam_proc = None
        self.cam_thread = None
        self.cam_running = False
        self.frame_lock = threading.Lock()
        self.latest_frame_bgr = None
        self.video_mouse_xy_var = tk.StringVar(value="x: -, y: -")
        self._last_video_xy = None  # (x,y) in frame coords or None
        self._disp_img_w = 0
        self._disp_img_h = 0
        self.camera_panel_hidden = True
        self._saved_sash_x = None
        self.base_video_width = 640  # adjust if you want
        self.popout_window = None  # Camera popout window

        # audio state
        self.audio_input_var = tk.StringVar()
        self.audio_output_var = tk.StringVar()
        self.audio_pyaudio = None
        self.audio_stream = None  # Legacy single stream
        self.audio_input_stream = None  # New separate input stream
        self.audio_output_stream = None  # New separate output stream
        self.audio_queue = None  # Queue for passing data between streams
        self.audio_running = False
        self.audio_thread = None
        self.audio_input_devices = []
        self.audio_output_devices = []




        # serial + engine
        self.serial = SerialController.SerialController(status_cb=self.set_status,app = self)
        self.engine = ScriptEngine.ScriptEngine(
            self.serial,
            get_frame_fn=self.get_latest_frame,
            status_cb=self.set_status,
            on_ip_update=self.on_ip_update,
            on_tick=self.on_engine_tick,
            on_python_needed=self.on_python_needed,
            on_error=self.on_script_error,
            on_prompt_input=self.on_prompt_input,
            on_prompt_choice=self.on_prompt_choice,
        )

        # Output backend selection
        self.backend_var = tk.StringVar(value="USB Serial")  # "USB Serial" or "3DS Input Redirection"
        threeds_settings = self._settings.get("threeds", {})
        self.threeds_ip_var = tk.StringVar(value=threeds_settings.get("ip", "192.168.1.1"))
        self.threeds_port_var = tk.StringVar(value=str(threeds_settings.get("port", 4950)))
        self.input_redirection_backend: Optional[InputRedirection.InputRedirectionBackend] = None

        # Active backend points to either self.serial or self.input_redirection_backend
        self.active_backend = self.serial

        self.engine.set_backend_getter(lambda: self.active_backend)
        self.engine.set_settings_getter(lambda: self._settings)

        self._build_ui()
        self._schedule_status_drain()
        self.apply_theme_setting(self._theme_setting)
        self._build_context_menu()
        self.apply_video_ratio(persist=False)
        self._schedule_frame_update()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.refresh_cameras()
        self.refresh_ports()
        self.refresh_scripts()
        self.refresh_audio_devices()
        self._update_title()

        # Check for missing dependencies after UI is built
        self.root.after(100, self._check_dependencies_startup)

    # ---- title/dirty
    def mark_dirty(self, dirty=True):
        self.dirty = dirty
        self._update_title()

    def _update_title(self):
        name = os.path.basename(self.script_path) if self.script_path else "(unsaved script)"
        star = " *" if self.dirty else ""
        self.root.title(f"Controller Macro Runner - {name}{star}")

    def _persist_setting_value(self, key: str, value: str):
        value = (value or "").strip()
        if not value or self._settings.get(key) == value:
            return
        self._settings[key] = value
        if not save_settings(self._settings):
            self.set_status("Settings save failed.")

    def _check_dependencies_startup(self):
        """Check for missing dependencies on first startup and offer to download them.

        Only shows the prompt once. After dismissing (yes or no), the prompt
        won't appear again. Users can download dependencies from Settings.
        """
        # Only show this prompt once ever
        if self._settings.get("dependency_check_shown"):
            return

        missing = []
        if not is_ffmpeg_available():
            missing.append("FFmpeg")
        if not is_tesseract_available():
            missing.append("Tesseract OCR")

        if not missing:
            # No dependencies missing, mark as shown so we don't check again
            self._settings["dependency_check_shown"] = True
            save_settings(self._settings)
            return

        # Build message
        if len(missing) == 1:
            msg = f"{missing[0]} is not installed.\n\n"
            if missing[0] == "FFmpeg":
                msg += "FFmpeg is required for camera capture functionality."
            else:
                msg += "Tesseract is required for text recognition (read_text command)."
        else:
            msg = "The following dependencies are not installed:\n\n"
            msg += "  - FFmpeg (required for camera capture)\n"
            msg += "  - Tesseract OCR (required for text recognition)\n"

        msg += "\n\nWould you like to download them now?\n\n"
        msg += "(You can also download later from Settings > Dependencies.)"

        result = messagebox.askyesno(
            "Missing Dependencies",
            msg,
            parent=self.root
        )

        # Mark as shown regardless of choice - only prompt once
        self._settings["dependency_check_shown"] = True
        save_settings(self._settings)

        if not result:
            return

        # Download missing dependencies sequentially
        if "FFmpeg" in missing:
            dialog = DependencyDownloadDialog(
                self.root,
                dependency_name="FFmpeg",
                download_fn=download_ffmpeg,
                size_hint="~140 MB",
                version=f"FFmpeg {FFMPEG_VERSION}",
                location=get_ffmpeg_dir(),
            )
            self.root.wait_window(dialog)
            if dialog.result:
                self.set_status("FFmpeg installed successfully.")

        if "Tesseract OCR" in missing:
            dialog = DependencyDownloadDialog(
                self.root,
                dependency_name="Tesseract OCR",
                download_fn=download_tesseract,
                size_hint="~48 MB",
                version=f"Tesseract {TESSERACT_VERSION}",
                location=get_tesseract_dir(),
            )
            self.root.wait_window(dialog)
            if dialog.result:
                self.set_status("Tesseract installed successfully.")

    def apply_theme_setting(self, theme_setting: str):
        theme_setting = normalize_theme_setting(theme_setting)
        self._theme_setting = theme_setting
        self._apply_theme(resolve_theme_mode(theme_setting))
        self._restart_theme_poll()

    def _restart_theme_poll(self):
        if self._theme_poll_id is not None:
            try:
                self.root.after_cancel(self._theme_poll_id)
            except Exception:
                pass
            self._theme_poll_id = None
        if self._theme_setting == "auto":
            self._theme_poll_id = self.root.after(2000, self._poll_system_theme)

    def _poll_system_theme(self):
        if not self.root.winfo_exists() or self._theme_setting != "auto":
            self._theme_poll_id = None
            return
        mode = resolve_theme_mode("auto")
        if mode != self._resolved_theme:
            self._apply_theme(mode)
        self._theme_poll_id = self.root.after(2000, self._poll_system_theme)

    def _apply_theme(self, mode: str):
        if mode not in ("dark", "light", "custom"):
            mode = "light"
        if mode == "custom":
            colors = self._get_custom_theme_colors()
        else:
            colors = THEME_COLORS[mode]
        if mode == self._resolved_theme and colors == self._theme_colors:
            return
        outline = colors["border"]
        self._resolved_theme = mode
        self._theme_colors = colors

        try:
            self.style.theme_use("clam")
        except tk.TclError:
            pass

        self.root.configure(bg=colors["bg"])
        self.style.configure(
            ".",
            background=colors["bg"],
            foreground=colors["text"],
            fieldbackground=colors["entry_bg"],
        )
        self.style.configure("TFrame", background=colors["bg"])
        self.style.configure("TLabel", background=colors["bg"], foreground=colors["text"])
        self.style.configure(
            "TLabelframe",
            background=colors["bg"],
            foreground=colors["text"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.configure("TLabelframe.Label", background=colors["bg"], foreground=colors["text"])
        self.style.configure(
            "TButton",
            background=colors["button_bg"],
            foreground=colors["button_fg"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
            focuscolor=outline,
        )
        self.style.map(
            "TButton",
            background=[("active", colors["select_bg"])],
            foreground=[("active", colors["button_fg"])],
        )
        self.style.configure(
            "TCheckbutton",
            background=colors["bg"],
            foreground=colors["text"],
        )
        self.style.map(
            "TCheckbutton",
            background=[("active", colors["panel"])],
            foreground=[("active", colors["text"])],
        )
        self.style.configure(
            "TRadiobutton",
            background=colors["bg"],
            foreground=colors["text"],
        )
        self.style.map(
            "TRadiobutton",
            background=[("active", colors["panel"])],
            foreground=[("active", colors["text"])],
        )
        self.style.configure(
            "TEntry",
            fieldbackground=colors["entry_bg"],
            foreground=colors["text"],
            insertcolor=colors["insert_fg"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.map(
            "TEntry",
            fieldbackground=[("disabled", colors["bg"]), ("!disabled", colors["entry_bg"])],
            foreground=[("disabled", colors["muted"]), ("!disabled", colors["text"])],
        )
        self.style.configure(
            "TCombobox",
            fieldbackground=colors["entry_bg"],
            foreground=colors["text"],
            background=colors["panel"],
            insertcolor=colors["insert_fg"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.map(
            "TCombobox",
            fieldbackground=[("readonly", colors["entry_bg"])],
            foreground=[("readonly", colors["text"])],
        )
        self.style.configure(
            "TScrollbar",
            background=colors["button_bg"],
            troughcolor=colors["pane_bg"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.map(
            "TScrollbar",
            background=[("active", colors["select_bg"])],
        )
        self.style.configure(
            "Volume.Horizontal.TScale",
            background=colors["button_bg"],
            troughcolor=colors["pane_bg"],
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.map(
            "Volume.Horizontal.TScale",
            background=[("active", colors["select_bg"])],
        )
        self.root.option_add("*TCombobox*Listbox.background", colors["tree_bg"])
        self.root.option_add("*TCombobox*Listbox.foreground", colors["tree_fg"])
        self.root.option_add("*TCombobox*Listbox.selectBackground", colors["select_bg"])
        self.root.option_add("*TCombobox*Listbox.selectForeground", colors["select_fg"])
        self.root.option_add("*Listbox.background", colors["tree_bg"])
        self.root.option_add("*Listbox.foreground", colors["tree_fg"])
        self.root.option_add("*Listbox.selectBackground", colors["select_bg"])
        self.root.option_add("*Listbox.selectForeground", colors["select_fg"])
        self._update_combobox_listboxes(colors)
        self.style.configure(
            "Treeview",
            background=colors["tree_bg"],
            fieldbackground=colors["tree_bg"],
            foreground=colors["tree_fg"],
            relief="flat",
            bordercolor=outline,
            lightcolor=outline,
            darkcolor=outline,
        )
        self.style.map(
            "Treeview",
            background=[("selected", colors["select_bg"])],
            foreground=[("selected", colors["select_fg"])],
        )
        self.style.configure(
            "Treeview.Heading",
            background=colors["panel"],
            foreground=colors["text"],
            relief="flat",
        )
        self.style.map(
            "Treeview.Heading",
            background=[("active", colors["select_bg"])],
        )
        self.style.configure(
            "TNotebook",
            background=colors["bg"],
        )
        self.style.configure(
            "TNotebook.Tab",
            background=colors["panel"],
            foreground=colors["text"],
            padding=(8, 4),
        )
        self.style.map(
            "TNotebook.Tab",
            background=[("selected", colors["bg"]), ("active", colors["select_bg"])],
            foreground=[("selected", colors["text"])],
        )

        if hasattr(self, "main_pane"):
            self.main_pane.configure(bg=colors["pane_bg"])
        if hasattr(self, "right_split"):
            self.right_split.configure(bg=colors["pane_bg"])
        if hasattr(self, "script_text"):
            self._apply_script_text_theme(colors)

    def _update_combobox_listboxes(self, colors: dict):
        def _walk(widget):
            for child in widget.winfo_children():
                yield child
                yield from _walk(child)

        for widget in _walk(self.root):
            if isinstance(widget, ttk.Combobox):
                try:
                    popdown = widget.tk.call("ttk::combobox::PopdownWindow", str(widget))
                    listbox = f"{popdown}.f.l"
                    widget.tk.call(
                        listbox, "configure",
                        "-background", colors["tree_bg"],
                        "-foreground", colors["tree_fg"],
                        "-selectbackground", colors["select_bg"],
                        "-selectforeground", colors["select_fg"],
                    )
                except tk.TclError:
                    pass

    def _get_custom_theme_colors(self):
        base = THEME_COLORS["dark"]
        custom = self._settings.get("custom_theme") or {}
        colors = {}
        for key, value in base.items():
            custom_value = custom.get(key, value)
            colors[key] = custom_value if isinstance(custom_value, str) else value
        return colors

    def _apply_script_text_theme(self, colors: dict):
        self.script_text.configure(
            background=colors["text_bg"],
            foreground=colors["text_fg"],
            insertbackground=colors["insert_fg"],
            selectbackground=colors["text_sel_bg"],
            selectforeground=colors["text_sel_fg"],
            highlightbackground=colors["border"],
            highlightcolor=colors["border"],
        )
        self.script_text.tag_configure("ip", background=colors["ip_bg"])
        self.script_text.tag_configure("comment", foreground=colors["comment_fg"])
        self.script_text.tag_configure("variable", foreground=colors["variable_fg"])
        self.script_text.tag_configure("math", foreground=colors["math_fg"])
        self.script_text.tag_configure("filepath", foreground=colors["filepath_fg"])
        self.script_text.tag_configure("selected", background=colors["selected_bg"])

    def _stop_theme_poll(self):
        if self._theme_poll_id is not None:
            try:
                self.root.after_cancel(self._theme_poll_id)
            except Exception:
                pass
            self._theme_poll_id = None

    # ---- status
    def _schedule_status_drain(self):
        if not self.root.winfo_exists():
            return
        if self._status_poll_id is not None:
            try:
                self.root.after_cancel(self._status_poll_id)
            except Exception:
                pass
        self._status_poll_id = self.root.after(self._status_poll_ms, self._drain_status_queue)

    def _drain_status_queue(self):
        if not self.root.winfo_exists():
            return
        last_msg = None
        while True:
            try:
                msg = self._status_queue.get_nowait()
            except queue.Empty:
                break
            if not msg:
                continue
            last_msg = msg
        if last_msg is not None:
            try:
                self.status_var.set(last_msg)
            except tk.TclError:
                pass
        self._status_poll_id = self.root.after(self._status_poll_ms, self._drain_status_queue)

    def set_status(self, msg):
        if not msg:
            return
        try:
            self._status_queue.put_nowait(msg)
        except Exception:
            pass

    def on_prompt_input(self, title, message, default_display, confirm):
        if not self.root.winfo_exists():
            return None
        if threading.current_thread() == self._ui_thread:
            return self._prompt_input_with_confirm(title, message, default_display, confirm)

        result = {"value": None}
        done = threading.Event()

        def show_dialog():
            try:
                result["value"] = self._prompt_input_with_confirm(title, message, default_display, confirm)
            finally:
                done.set()

        self.root.after(0, show_dialog)
        done.wait()
        return result["value"]

    def on_prompt_choice(self, title, message, choices, default_index, confirm, display_mode):
        if not self.root.winfo_exists():
            return None
        if threading.current_thread() == self._ui_thread:
            return self._prompt_choice_with_confirm(title, message, choices, default_index, confirm, display_mode)

        result = {"value": None}
        done = threading.Event()

        def show_dialog():
            try:
                result["value"] = self._prompt_choice_with_confirm(title, message, choices, default_index, confirm, display_mode)
            finally:
                done.set()

        self.root.after(0, show_dialog)
        done.wait()
        return result["value"]

    def _prompt_input_with_confirm(self, title, message, default_display, confirm):
        title = "" if title is None else str(title)
        message = "" if message is None else str(message)
        current = "" if default_display is None else str(default_display)
        while True:
            result = self._show_themed_input_dialog(title, message, current)
            if result is None:
                return None
            if not confirm:
                return result
            confirm_msg = f"Use this value?\n\n{result}"
            if self._show_themed_confirm_dialog("Confirm Input", confirm_msg):
                return result
            current = result

    def _prompt_choice_with_confirm(self, title, message, choices, default_index, confirm, display_mode):
        title = "" if title is None else str(title)
        message = "" if message is None else str(message)
        choices_list = [] if choices is None else list(choices)
        if not choices_list:
            return None

        display = "" if display_mode is None else str(display_mode)
        display = display.strip().lower()
        if display not in ("dropdown", "buttons"):
            display = "dropdown"

        current_index = default_index if isinstance(default_index, int) else 0
        if not (0 <= current_index < len(choices_list)):
            current_index = 0

        while True:
            result_index = self._show_themed_choice_dialog(
                title, message, choices_list, current_index, display
            )
            if result_index is None:
                return None
            if not confirm:
                return result_index
            chosen = choices_list[result_index]
            confirm_msg = f"Use this value?\n\n{chosen}"
            if self._show_themed_confirm_dialog("Confirm Choice", confirm_msg):
                return result_index
            current_index = result_index

    def _show_themed_input_dialog(self, title, message, initial_value):
        colors = self._theme_colors or THEME_COLORS.get("light", {})
        bg = colors.get("bg", self.root.cget("bg"))

        dlg = tk.Toplevel(self.root)
        dlg.title(title or "Input")
        dlg.configure(bg=bg)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.columnconfigure(0, weight=1)

        frame = ttk.Frame(dlg, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        msg = message if message else "Enter value:"
        msg = str(msg)
        ttk.Label(frame, text=msg, wraplength=420, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )

        var = tk.StringVar(value=initial_value or "")
        entry = ttk.Entry(frame, textvariable=var, width=40)
        entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 10))

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=2, column=0, columnspan=2, sticky="e")

        result = {"value": None}

        def on_ok():
            result["value"] = var.get()
            dlg.destroy()

        def on_cancel():
            result["value"] = None
            dlg.destroy()

        ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", on_cancel)
        dlg.bind("<Return>", lambda _e: on_ok())
        dlg.bind("<Escape>", lambda _e: on_cancel())

        entry.focus_set()
        entry.selection_range(0, tk.END)
        dlg.update_idletasks()
        try:
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            dlg_w = dlg.winfo_width()
            dlg_h = dlg.winfo_height()
            x = max(0, int(root_x + (root_w - dlg_w) / 2))
            y = max(0, int(root_y + (root_h - dlg_h) / 2))
            dlg.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass
        self.root.wait_window(dlg)
        return result["value"]

    def _show_themed_choice_dialog(self, title, message, choices, default_index, display_mode):
        colors = self._theme_colors or THEME_COLORS.get("light", {})
        bg = colors.get("bg", self.root.cget("bg"))

        dlg = tk.Toplevel(self.root)
        dlg.title(title or "Choose")
        dlg.configure(bg=bg)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.columnconfigure(0, weight=1)

        frame = ttk.Frame(dlg, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        msg = message if message else "Select a value:"
        msg = str(msg)
        ttk.Label(frame, text=msg, wraplength=420, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )

        values = [str(v) for v in choices]
        result = {"index": None}
        display = (display_mode or "dropdown").strip().lower()
        if display not in ("dropdown", "buttons"):
            display = "dropdown"

        if display == "buttons":
            buttons_frame = ttk.Frame(frame)
            buttons_frame.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 10))

            def choose_columns(count, max_cols=3):
                max_cols = min(max_cols, count)
                best_cols = 1
                best_empty = None
                best_rows = None
                for cols in range(1, max_cols + 1):
                    rows = math.ceil(count / cols)
                    last = count - cols * (rows - 1)
                    empty = cols - last
                    if (
                        best_empty is None
                        or empty < best_empty
                        or (empty == best_empty and rows < best_rows)
                    ):
                        best_cols = cols
                        best_empty = empty
                        best_rows = rows
                return best_cols

            columns = choose_columns(len(values), 3)

            for col in range(columns):
                buttons_frame.columnconfigure(col, weight=1)

            def choose(idx):
                result["index"] = idx
                dlg.destroy()

            total = len(values)
            rows = math.ceil(total / columns) if columns else 0
            for row in range(rows):
                start = row * columns
                end = min(total, start + columns)
                items_in_row = end - start
                offset = (columns - items_in_row) // 2 if items_in_row < columns else 0
                for i in range(start, end):
                    col = offset + (i - start)
                    ttk.Button(
                        buttons_frame,
                        text=values[i],
                        command=lambda idx=i: choose(idx)
                    ).grid(row=row, column=col, sticky="ew", padx=4, pady=4)
            def on_cancel():
                result["index"] = None
                dlg.destroy()

            dlg.protocol("WM_DELETE_WINDOW", on_cancel)
            dlg.bind("<Escape>", lambda _e: on_cancel())
        else:
            var = tk.StringVar()
            combo = ttk.Combobox(
                frame,
                textvariable=var,
                values=values,
                state="readonly",
                width=40
            )
            combo.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 10))

            if 0 <= default_index < len(values):
                combo.current(default_index)
            elif values:
                combo.current(0)

            btn_frame = ttk.Frame(frame)
            btn_frame.grid(row=2, column=0, columnspan=2, sticky="e")

            def on_ok():
                idx = combo.current()
                result["index"] = idx if idx >= 0 else None
                dlg.destroy()

            def on_cancel():
                result["index"] = None
                dlg.destroy()

            ttk.Button(btn_frame, text="Cancel", command=on_cancel).pack(side="right", padx=(6, 0))
            ttk.Button(btn_frame, text="OK", command=on_ok).pack(side="right")

            dlg.protocol("WM_DELETE_WINDOW", on_cancel)
            dlg.bind("<Return>", lambda _e: on_ok())
            dlg.bind("<Escape>", lambda _e: on_cancel())

            combo.focus_set()
        dlg.update_idletasks()
        try:
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            dlg_w = dlg.winfo_width()
            dlg_h = dlg.winfo_height()
            x = max(0, int(root_x + (root_w - dlg_w) / 2))
            y = max(0, int(root_y + (root_h - dlg_h) / 2))
            dlg.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass

        self.root.wait_window(dlg)
        return result["index"]

    def _show_themed_confirm_dialog(self, title, message):
        colors = self._theme_colors or THEME_COLORS.get("light", {})
        bg = colors.get("bg", self.root.cget("bg"))

        dlg = tk.Toplevel(self.root)
        dlg.title(title or "Confirm")
        dlg.configure(bg=bg)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)
        dlg.columnconfigure(0, weight=1)

        frame = ttk.Frame(dlg, padding=12)
        frame.grid(row=0, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)

        msg = "" if message is None else str(message)
        ttk.Label(frame, text=msg, wraplength=420, justify="left").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )

        btn_frame = ttk.Frame(frame)
        btn_frame.grid(row=1, column=0, columnspan=2, sticky="e", pady=(10, 0))

        result = {"value": False}

        def on_yes():
            result["value"] = True
            dlg.destroy()

        def on_no():
            result["value"] = False
            dlg.destroy()

        ttk.Button(btn_frame, text="No", command=on_no).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Yes", command=on_yes).pack(side="right")

        dlg.protocol("WM_DELETE_WINDOW", on_no)
        dlg.bind("<Return>", lambda _e: on_yes())
        dlg.bind("<Escape>", lambda _e: on_no())

        dlg.update_idletasks()
        try:
            root_x = self.root.winfo_rootx()
            root_y = self.root.winfo_rooty()
            root_w = self.root.winfo_width()
            root_h = self.root.winfo_height()
            dlg_w = dlg.winfo_width()
            dlg_h = dlg.winfo_height()
            x = max(0, int(root_x + (root_w - dlg_w) / 2))
            y = max(0, int(root_y + (root_h - dlg_h) / 2))
            dlg.geometry(f"+{x}+{y}")
        except tk.TclError:
            pass
        self.root.wait_window(dlg)
        return result["value"]

    # ---- engine tick (live vars)
    def on_engine_tick(self):
        self.root.after(0, self.refresh_vars_view)

    # ---- python download prompt
    def on_python_needed(self):
        """Called when run_python command needs Python but it's not available."""
        def show_dialog():
            result = messagebox.askyesno(
                "Python Required",
                "The run_python command requires Python to execute scripts.\n\n"
                "Python is not currently installed. Would you like to download it now?\n\n"
                "(This is a one-time ~11 MB download)",
                parent=self.root
            )
            if result:
                dialog = PythonDownloadDialog(self.root)
                self.root.wait_window(dialog)
                if dialog.result:
                    self.set_status("Python installed. You can now re-run the script.")

        # Schedule on main thread since this may be called from engine thread
        self.root.after(0, show_dialog)

    # ---- script error handler
    def on_script_error(self, title, message):
        """Called when script execution encounters an error. Shows error dialog from main thread."""
        def show_dialog():
            messagebox.showerror(title, message, parent=self.root)
        # Schedule on main thread since this is called from engine thread
        self.root.after(0, show_dialog)

    # ---- frame access
    def get_latest_frame(self):
        with self.frame_lock:
            return None if self.latest_frame_bgr is None else self.latest_frame_bgr.copy()

    # ---- UI build
    def _build_ui(self):
        outer = ttk.Frame(self.root, padding=10)
        outer.grid(row=0, column=0, sticky="nsew")
        outer.columnconfigure(0, weight=1)
        outer.rowconfigure(1, weight=1)

        # Top bar
        top = ttk.Frame(outer)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(13, weight=1)

        # Backend Selectors
        ttk.Label(top, text="Output:").grid(row=1, column=4, sticky="w")
        self.backend_combo = ttk.Combobox(
            top, textvariable=self.backend_var, state="readonly",
            values=["USB Serial", "3DS Input Redirection"], width=12
        )
        self.backend_combo.grid(row=1, column=5, sticky="ew", padx=(6, 6),pady=(4,0))
        self.backend_combo.bind("<<ComboboxSelected>>", lambda e: self.on_backend_changed())

        # initialize backend selection
        self.on_backend_changed()

        ttk.Button(top, text="Settings...", command=self.open_settings_dialog).grid(row=1, column=6, padx=(0, 0))

        # Camera controls
        ttk.Label(top, text="Camera:").grid(row=0, column=0, sticky="w")
        self.cam_var = tk.StringVar()
        self.cam_combo = ttk.Combobox(top, textvariable=self.cam_var, state="readonly", width=16)
        self.cam_combo.grid(row=0, column=1, sticky="w", padx=(6, 6))
        self.cam_combo.bind("<<ComboboxSelected>>", self._on_camera_selected)
        ttk.Button(top, text="Refresh", command=self.refresh_cameras).grid(row=0, column=2, padx=(0, 0))
        self.cam_toggle_btn = ttk.Button(top, text="Start Cam", command=self.toggle_camera)
        self.cam_toggle_btn.grid(row=0, column=3, padx=(0, 6))

        self.cam_display_btn = ttk.Button(top, text="Show Cam", command=self.toggle_camera_panel)
        self.cam_display_btn.grid(row=1, column=3, padx=(0,6))

        ttk.Label(top, text="Cam Ratio:").grid(row=1, column=0, sticky="w")

        self.ratio_var = tk.StringVar(value=self._initial_camera_ratio)
        self.ratio_combo = ttk.Combobox(
            top, textvariable=self.ratio_var, state="readonly",
            values=self.video_ratio_options, width=16
        )
        self.ratio_combo.grid(row=1, column=1, sticky="w", padx=(6, 6))

        ttk.Button(top, text="Apply", command=self.apply_video_ratio).grid(row=1, column=2, padx=(0, 0))



        # Serial controls
        ttk.Label(top, text="COM:").grid(row=0, column=4, sticky="w")
        self.com_var = tk.StringVar()
        self.com_combo = ttk.Combobox(top, textvariable=self.com_var, state="readonly", width=12)
        self.com_combo.grid(row=0, column=5, sticky="w", padx=(6, 6))
        self.com_combo.bind("<<ComboboxSelected>>", self._on_com_selected)
        ttk.Button(top, text="Refresh", command=self.refresh_ports).grid(row=0, column=6, padx=(0, 0))
        self.ser_btn = ttk.Button(top, text="Connect", command=self.toggle_serial)
        self.ser_btn.grid(row=0, column=7, sticky="w", padx=(0, 4))

        # Channel controls
        ttk.Label(top, text="Channel:").grid(row=0, column=8, sticky="w")  # adjust column if needed

        self.chan_var = tk.StringVar(value="1")
        self.chan_combo = ttk.Combobox(
            top,
            textvariable=self.chan_var,
            state="readonly",
            width=4,
            values=[str(i) for i in range(1, 17)]
        )
        self.chan_combo.grid(row=0, column=9, sticky="w", padx=(6, 6))

        ttk.Button(top, text="Set Channel", command=self.set_channel).grid(row=0, column=10, padx=(0, 4))

        

        ttk.Checkbutton(top, text="Camera-less \nKeyboard Control", variable=self.kb_enabled,
                        command=self._on_keyboard_toggle).grid(row=2, column=0, columnspan=2, padx=(10, 6), sticky="w")

        



        # Script file controls
        ttk.Label(top, text="Script:").grid(row=0, column=12, sticky="w")
        self.script_var = tk.StringVar()
        self.script_combo = ttk.Combobox(top, textvariable=self.script_var, state="readonly", width=24)
        self.script_combo.grid(row=0, column=13, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_scripts).grid(row=0, column=14, padx=(0, 0))
        ttk.Button(top, text="Load", command=self.load_script_from_dropdown).grid(row=0, column=15, padx=(0, 0))
        ttk.Button(top, text="New", command=self.new_script).grid(row=0, column=16, padx=(0, 6))

        ttk.Button(top, text="Export .py", command= lambda: ScriptToPy.export_script_to_python(self)).grid(row=1, column=14, padx=(0, 0))
        ttk.Button(top, text="Save", command=self.save_script).grid(row=1, column=15, padx=(0, 0))
        ttk.Button(top, text="Save As", command=self.save_script_as).grid(row=1, column=16, padx=(0, 6))

        # Run controls
        runbar = ttk.Frame(outer)
        runbar.grid(row=2, column=0, sticky="ew", pady=(8, 0))
        self.run_btn = ttk.Button(runbar, text="Run", command=self.run_script)
        self.run_btn.grid(row=0, column=0, padx=(0, 6))
        self.stop_btn = ttk.Button(runbar, text="Stop", command=self.stop_script)
        self.stop_btn.grid(row=0, column=1, padx=(0, 18))

        self.status_var = tk.StringVar(value="Idle")
        ttk.Label(outer, textvariable=self.status_var).grid(row=3, column=0, sticky="ew", pady=(8, 0))

        # Main split (tk.PanedWindow so the sash handle is visible)
        main = tk.PanedWindow(
            outer,
            orient=tk.HORIZONTAL,
            sashrelief=tk.RAISED,
            sashwidth=8,
            showhandle=True,
            bd=0
        )
        main.grid(row=1, column=0, sticky="nsew", pady=(10, 0))

        # Left: video
        left = ttk.Frame(main)
        left.rowconfigure(0, weight=1)
        left.columnconfigure(0, weight=1)

        self.video_label = ttk.Label(left,anchor="nw")
        self.video_label.grid(row=0, column=0, sticky="nsew")

        self.video_label.bind("<Motion>", self._on_video_mouse_move)
        self.video_label.bind("<Leave>", self._on_video_mouse_leave)
        self.video_label.bind("<Button-1>", self._on_video_click)
        self.video_label.bind("<Double-Button-1>", self._on_video_double_click)


        # Coordinate readout
        coord_bar = ttk.Frame(left)
        coord_bar.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        coord_bar.columnconfigure(0, weight=1)
        ttk.Label(coord_bar, textvariable=self.video_mouse_xy_var).grid(row=0, column=0, sticky="w")

        # Audio controls (input)
        self.audio_input_frame = ttk.Frame(left)
        self.audio_input_frame.grid(row=2, column=0, sticky="ew", pady=(6, 0))
        self.audio_input_frame.columnconfigure(1, weight=1)

        ttk.Label(self.audio_input_frame, text="Audio Input:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.audio_input_combo = ttk.Combobox(self.audio_input_frame, textvariable=self.audio_input_var, state="readonly", width=30)
        self.audio_input_combo.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        self.audio_input_combo.bind("<<ComboboxSelected>>", self._on_audio_input_selected)
        ttk.Button(self.audio_input_frame, text="Refresh", command=self.refresh_audio_devices).grid(row=0, column=2, padx=(0, 6))
        self.audio_toggle_btn = ttk.Button(self.audio_input_frame, text="Start Audio", command=self.toggle_audio)
        self.audio_toggle_btn.grid(row=0, column=3)

        # Audio controls (output)
        self.audio_output_frame = ttk.Frame(left)
        self.audio_output_frame.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self.audio_output_frame.columnconfigure(1, weight=1)

        ttk.Label(self.audio_output_frame, text="Audio Output:").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self.audio_output_combo = ttk.Combobox(self.audio_output_frame, textvariable=self.audio_output_var, state="readonly", width=30)
        self.audio_output_combo.grid(row=0, column=1, sticky="ew")
        self.audio_output_combo.bind("<<ComboboxSelected>>", self._on_audio_output_selected)

        # Initially hide audio controls (camera panel starts hidden)
        self.audio_input_frame.grid_remove()
        self.audio_output_frame.grid_remove()

        # Right: script viewer + vars
        right = ttk.Frame(main)
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        main.add(left, minsize=0)
        main.add(right, minsize=320)

        self.main_pane = main

        # Set default split once the window is actually on-screen
        self.root.after_idle(self._set_default_split_retry)

        # Also re-apply once on first resize/map (helps on some setups)
        self._did_initial_split = False
        self.root.bind("<Configure>", self._on_first_configure)



        # Right pane: vertically split Script/Vars
        right_split = tk.PanedWindow(right, orient=tk.VERTICAL, sashrelief=tk.RAISED, sashwidth=6, showhandle=False, bd=0)
        right_split.grid(row=0, column=0, sticky="nsew")
        right.rowconfigure(0, weight=1)
        self.right_split = right_split

        # --- Script viewer
        script_box = ttk.LabelFrame(right_split, text="Script Commands (right-click menu, double-click edit)")
        script_box.rowconfigure(0, weight=1)
        script_box.columnconfigure(0, weight=1)

        # Use Text widget instead of Treeview for per-character coloring
        self.script_text = tk.Text(
            script_box,
            wrap="none",
            height=12,
            font=("Consolas", 9),
            state="disabled",  # Read-only by default
            cursor="arrow"
        )
        self.script_text.grid(row=0, column=0, sticky="nsew")

        scr_y = ttk.Scrollbar(script_box, orient="vertical", command=self.script_text.yview)
        self.script_text.configure(yscrollcommand=scr_y.set)
        scr_y.grid(row=0, column=1, sticky="ns")

        scr_x = ttk.Scrollbar(script_box, orient="horizontal", command=self.script_text.xview)
        self.script_text.configure(xscrollcommand=scr_x.set)
        scr_x.grid(row=1, column=0, sticky="ew")

        btnrow = ttk.Frame(script_box)
        btnrow.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        ttk.Button(btnrow, text="Add", command=self.add_command).pack(side="left", padx=2,pady=(0,4))
        ttk.Button(btnrow, text="Edit", command=self.edit_command).pack(side="left", padx=2,pady=(0,4))
        ttk.Button(btnrow, text="Delete", command=self.delete_command).pack(side="left", padx=2,pady=(0,4))
        ttk.Button(btnrow, text="Up", command=lambda: self.move_command(-1)).pack(side="left", padx=2,pady=(0,4))
        ttk.Button(btnrow, text="Down", command=lambda: self.move_command(1)).pack(side="left", padx=2,pady=(0,4))
        ttk.Button(btnrow, text="Comment", command=self.add_comment).pack(side="left", padx=2,pady=(0,4))

        # Indent view toggle (if you already have it, keep yours)
        self.indent_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            btnrow,
            text="Indent view",
            variable=self.indent_var,
            command=lambda: self.populate_script_view(preserve_view=True)
        ).pack(side="right", padx=6)

        # --- Vars
        vars_box = ttk.LabelFrame(right_split, text="Variables")
        vars_box.rowconfigure(0, weight=1)
        vars_box.columnconfigure(0, weight=1)

        self.vars_tree = ttk.Treeview(vars_box, columns=("name", "value"), show="headings")
        self.vars_tree.heading("name", text="name")
        self.vars_tree.heading("value", text="value")
        self.vars_tree.column("name", width=140, anchor="w")
        self.vars_tree.column("value", width=360, anchor="w")
        self.vars_tree.grid(row=0, column=0, sticky="nsew")

        vsc = ttk.Scrollbar(vars_box, orient="vertical", command=self.vars_tree.yview)
        self.vars_tree.configure(yscrollcommand=vsc.set)
        vsc.grid(row=0, column=1, sticky="ns")

        # Add panes
        right_split.add(script_box, minsize=220)
        right_split.add(vars_box, minsize=140)

        # Tags / bindings for Text widget
        self.script_text.tag_configure("ip", background="#dbeafe")
        self.script_text.tag_configure("comment", foreground="#228B22")  # Forest green
        self.script_text.tag_configure("variable", foreground="#0066CC")  # Blue
        self.script_text.tag_configure("math", foreground="#b58900")  # Yellow
        self.script_text.tag_configure("filepath", foreground="#CC0000")  # Red
        self.script_text.tag_configure("selected", background="#e0e0e0")  # Selected line
        self.script_text.tag_raise("variable", "math")
        self.script_text.tag_raise("filepath", "variable")
        self.script_text.bind("<Button-3>", self._on_script_right_click)
        self.script_text.bind("<Double-1>", self._on_script_double_click)
        self.script_text.bind("<Button-1>", self._on_script_click)
        self.script_text.bind("<Delete>", self._on_script_delete_key)
        self.script_text.bind("<Control-c>", self._on_script_copy_key)
        self.script_text.bind("<Control-v>", self._on_script_paste_key)

        # Track selected line for editing
        self.selected_script_line = None


    def _normalize_keysym(self, event):
        """
        Normalize Tk keysym into something consistent for binding lookup.
        Examples:
        'w' -> 'w'
        'W' -> 'w'
        'Return' -> 'enter'
        'space' -> 'space'
        """
        ks = event.keysym
        if not ks:
            return None
        ks = ks.lower()
        if ks in ("shift", "shift_l", "shift_r"):
            ks = self._normalize_shift_keysym(ks, event)
        if ks == "return":
            ks = "enter"
        return ks

    def _debug_key_event(self, label, event, normalized=None):
        if not self._key_debug:
            return
        try:
            keysym = getattr(event, "keysym", None)
            keycode = getattr(event, "keycode", None)
            state = getattr(event, "state", None)
            char = getattr(event, "char", None)
            print(
                "[KEYDBG] "
                f"{label} keysym={keysym} keycode={keycode} state={state} "
                f"char={char!r} normalized={normalized} "
                f"kb_down={sorted(self.kb_down)} held={sorted(self.kb_buttons_held)}"
            )
        except Exception:
            pass

    def _normalize_shift_keysym(self, ks, event):
        keycode = getattr(event, "keycode", None)
        if keycode in (160, 50):  # Windows VK_LSHIFT / X11
            return "shift_l"
        if keycode in (161, 62):  # Windows VK_RSHIFT / X11
            return "shift_r"
        if ks == "shift":
            if "shift_l" in self.kb_down and "shift_r" not in self.kb_down:
                return "shift_l"
            if "shift_r" in self.kb_down and "shift_l" not in self.kb_down:
                return "shift_r"
            if "shift_r" in self.kb_bindings and "shift_l" not in self.kb_bindings:
                return "shift_r"
            if "shift_l" in self.kb_bindings and "shift_r" not in self.kb_bindings:
                return "shift_l"
        return ks

    def _release_keyboard_binding(self, ks):
        if ks.startswith("shift"):
            self._debug_key_event("release_binding_start", None, ks)
        target = self.kb_bindings.get(ks)

        if ks in self.kb_down:
            self.kb_down.remove(ks)

        if not target:
            return False

        if target in SerialController.LEFT_STICK_BINDINGS:
            if target in self.kb_left_stick_dirs:
                self.kb_left_stick_dirs.remove(target)
                self._update_keyboard_sticks()
            return True

        if target in SerialController.RIGHT_STICK_BINDINGS:
            if target in self.kb_right_stick_dirs:
                self.kb_right_stick_dirs.remove(target)
                self._update_keyboard_sticks()
            return True

        if target in self.kb_buttons_held:
            self.kb_buttons_held.remove(target)

        self._select_active_backend()
        if self.active_backend and getattr(self.active_backend, "connected", False):
            self.active_backend.set_buttons(sorted(self.kb_buttons_held))

        if ks.startswith("shift"):
            self._debug_key_event("release_binding_done", None, ks)
        return True

    def _stick_dirs_to_xy(self, dirs, prefix):
        x = 0.0
        y = 0.0
        if f"{prefix} Left" in dirs:
            x -= 1.0
        if f"{prefix} Right" in dirs:
            x += 1.0
        if f"{prefix} Up" in dirs:
            y += 1.0
        if f"{prefix} Down" in dirs:
            y -= 1.0
        if x and y:
            scale = 1.0 / math.sqrt(2.0)
            x *= scale
            y *= scale
        return x, y

    def _update_keyboard_sticks(self):
        self._select_active_backend()
        if not self.active_backend or not getattr(self.active_backend, "connected", False):
            return

        backend = self.active_backend
        inner_backend = backend.backend if isinstance(backend, SerialController.SerialController) else backend

        if inner_backend and hasattr(inner_backend, "set_left_stick"):
            x, y = self._stick_dirs_to_xy(self.kb_left_stick_dirs, "Left Stick")
            backend.set_left_stick(x, y)

        if inner_backend and hasattr(inner_backend, "set_right_stick"):
            x, y = self._stick_dirs_to_xy(self.kb_right_stick_dirs, "Right Stick")
            backend.set_right_stick(x, y)

    def _manual_control_allowed(self):
        """
        Check if manual keyboard control is allowed.
        Allowed when:
        - kb_enabled is True (checkbox) OR kb_camera_focused is True (auto-focus)
        - A backend is connected (serial or 3DS)
        - Script is NOT running
        """
        if self.engine.running:
            return False

        # Check if keyboard control is enabled (either by checkbox or camera focus)
        if not (self.kb_enabled.get() or self.kb_camera_focused):
            return False

        # Check if any backend is connected
        self._select_active_backend()
        if self.active_backend and getattr(self.active_backend, "connected", False):
            return True

        return False

    def _on_keyboard_toggle(self):
        # Turning off: release everything
        if not self.kb_enabled.get():
            self._release_all_keyboard_buttons()
            self.kb_camera_focused = False  # Also clear camera focus
            self.set_status("Keyboard Control: OFF")
            return

        # Turning on: if script running, disallow
        if self.engine.running:
            self.kb_enabled.set(False)
            messagebox.showwarning("Keyboard Control", "Stop the script before enabling keyboard control.")
            return

        # Check if any backend is connected
        self._select_active_backend()
        backend_connected = self.active_backend and getattr(self.active_backend, "connected", False)

        if not backend_connected:
            # allow toggling on, but it won't do anything until connected
            self.set_status("Keyboard Control: ON (connect a backend to use)")
        else:
            self.set_status("Keyboard Control: ON")

    def _release_all_keyboard_buttons(self):
        self.kb_down.clear()
        self.kb_buttons_held.clear()
        self.kb_left_stick_dirs.clear()
        self.kb_right_stick_dirs.clear()
        # go neutral only if script not running
        self._select_active_backend()
        if not self.engine.running and self.active_backend and getattr(self.active_backend, "connected", False):
            self.active_backend.set_buttons([])
            self._update_keyboard_sticks()

    def _enable_camera_keyboard_focus(self, event=None):
        """Enable keyboard control focus when camera frame is clicked."""
        if self.engine.running:
            return  # Don't enable during script execution

        was_focused = self.kb_camera_focused
        self.kb_camera_focused = True

        # Check if backend is connected
        self._select_active_backend()
        backend_connected = self.active_backend and getattr(self.active_backend, "connected", False)

        if not was_focused:
            if backend_connected:
                self.set_status("Keyboard Control: Active (click elsewhere to deactivate)")
            else:
                self.set_status("Keyboard Control: Focused (connect a backend to use)")

    def _disable_camera_keyboard_focus(self, event=None):
        """Disable keyboard control focus when clicking outside camera frame."""
        if not self.kb_camera_focused:
            return  # Already not focused

        self.kb_camera_focused = False

        # Only release buttons if the checkbox is not enabled
        # (if checkbox is enabled, keyboard control continues)
        if not self.kb_enabled.get():
            self._release_all_keyboard_buttons()
            self.set_status("Keyboard Control: Inactive")

    def _check_focus_loss(self, event):
        """
        Check if a click event is outside all camera frames.
        If so, disable camera keyboard focus.
        """
        if not self.kb_camera_focused:
            return  # Already not focused

        # Get the widget that was clicked
        widget = event.widget

        # Check if the click was on the main video label
        if widget == self.video_label:
            return  # Click on main camera - keep focus

        # Check if we have a popout window
        if self.popout_window is not None:
            # Check if click was on the popout video label
            if widget == self.popout_window.video_label:
                return  # Click on popout camera - keep focus

            # Check if click was anywhere inside the popout window
            # This allows clicking coord bar etc. without losing focus
            try:
                if widget.winfo_toplevel() == self.popout_window.window:
                    return  # Click inside popout window - keep focus
            except (tk.TclError, AttributeError):
                pass  # Widget may have been destroyed

        # Click was elsewhere - disable focus
        self._disable_camera_keyboard_focus()

    def _on_key_press(self, event):
        if not self._manual_control_allowed():
            return

        ks = self._normalize_keysym(event)
        if not ks:
            return

        # Check if this key is mapped to a controller control
        target = self.kb_bindings.get(ks)

        # Prevent repeat spamming (Tk sends repeats while held)
        if ks in self.kb_down:
            if self._key_debug and ("shift" in (event.keysym or "").lower() or ks.startswith("shift")):
                self._debug_key_event("repeat", event, ks)
            # Return "break" to prevent key from triggering GUI buttons
            return "break" if target else None
        self.kb_down.add(ks)
        if self._key_debug and ("shift" in (event.keysym or "").lower() or ks.startswith("shift")):
            self._debug_key_event("press", event, ks)

        if not target:
            return

        if target in SerialController.LEFT_STICK_BINDINGS:
            self.kb_left_stick_dirs.add(target)
            self._update_keyboard_sticks()
            return "break"

        if target in SerialController.RIGHT_STICK_BINDINGS:
            self.kb_right_stick_dirs.add(target)
            self._update_keyboard_sticks()
            return "break"

        self.kb_buttons_held.add(target)
        self._select_active_backend()
        if self.active_backend and getattr(self.active_backend, "connected", False):
            self.active_backend.set_buttons(sorted(self.kb_buttons_held))

        # Return "break" to prevent the key event from propagating to GUI widgets
        # This prevents Enter/Space from activating focused buttons while controlling
        return "break"


    def _on_key_release(self, event):
        if not self._manual_control_allowed():
            return

        ks = self._normalize_keysym(event)
        if not ks:
            return
        if self._key_debug and ("shift" in (event.keysym or "").lower() or ks.startswith("shift")):
            self._debug_key_event("release", event, ks)
        if ks in ("shift_l", "shift_r") and ks not in self.kb_down:
            # If Tk reports the wrong shift key on release, fall back to the one that's down.
            other = "shift_r" if ks == "shift_l" else "shift_l"
            if other in self.kb_down:
                ks = other
        if ks == "shift":
            released = False
            for shift_key in ("shift_l", "shift_r"):
                if shift_key in self.kb_down:
                    released = self._release_keyboard_binding(shift_key) or released
            if released:
                return "break"
            return

        if self._release_keyboard_binding(ks):
            return "break"



    def _copy_to_clipboard(self, text: str):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            # ensures clipboard persists after app closes on some systems
            self.root.update_idletasks()
        except Exception as e:
            self.set_status(f"Clipboard error: {e}")

    def _on_video_click(self, event):
        """Handle single click on video - enables keyboard focus."""
        self._enable_camera_keyboard_focus(event)


    def _event_to_frame_xy(self, event):
        with self.frame_lock:
            frame = self.latest_frame_bgr
            if frame is None:
                return None
            fh, fw, _ = frame.shape

        iw = getattr(self, "_disp_img_w", fw) or fw
        ih = getattr(self, "_disp_img_h", fh) or fh

        # Because video_label is anchor="nw", the image is pinned to top-left
        off_x, off_y = 0, 0

        x_img = int(event.x) - off_x
        y_img = int(event.y) - off_y

        if not (0 <= x_img < iw and 0 <= y_img < ih):
            return None

        # If later you scale the image, this keeps working:
        x = int(x_img * fw / iw)
        y = int(y_img * fh / ih)

        if 0 <= x < fw and 0 <= y < fh:
            return (x, y)
        return None


    def _select_active_backend(self):
        if self.backend_var.get() == "3DS Input Redirection":
            if self.input_redirection_backend and self.input_redirection_backend.connected:
                self.active_backend = self.input_redirection_backend
            else:
                # Not enabled yet; still point to a disconnected backend if it exists
                self.active_backend = self.input_redirection_backend or self.serial
        else:
            self.active_backend = self.serial



    def _on_first_configure(self, event):
        # Run only once
        if self._did_initial_split:
            return
        self._did_initial_split = True
        # Give Tk a moment to finish geometry
        self.root.after(50, self._set_default_split_retry)

    def _set_default_split_retry(self, tries=20):
        """
        Tries multiple times because PanedWindow often reports width=1 until fully laid out.
        """
        try:
            self.main_pane.update_idletasks()
            total = self.main_pane.winfo_width()

            # If not ready, retry
            if total is None or total < 200:
                if tries > 0:
                    self.root.after(50, lambda: self._set_default_split_retry(tries - 1))
                return

            # Make video bigger by default
            x = int(total * 0.0)  #

            # ttk.PanedWindow uses sashpos; tk.PanedWindow uses sash_place
            if hasattr(self.main_pane, "sashpos"):
                self.main_pane.sashpos(0, x)
            else:
                self.main_pane.sash_place(0, x, 0)

        except Exception:
            if tries > 0:
                self.root.after(50, lambda: self._set_default_split_retry(tries - 1))


    def _build_context_menu(self):
        self.ctx = tk.Menu(self.root, tearoff=0)
        self.ctx_add_menu = self._build_add_context_menu()
        self.ctx.add_cascade(label="Add", menu=self.ctx_add_menu)
        self.ctx.add_command(label="Edit", command=self.edit_command)
        self.ctx.add_command(label="Copy", command=self.copy_command)
        self.ctx.add_command(label="Paste", command=self.paste_command)
        self.ctx.add_command(label="Delete", command=self.delete_command)
        self.ctx.add_separator()
        self.ctx.add_command(label="Move Up", command=lambda: self.move_command(-1))
        self.ctx.add_command(label="Move Down", command=lambda: self.move_command(1))
        self.ctx.add_separator()
        self.ctx.add_command(label="Add Comment", command=self.add_comment)
        self.ctx.add_separator()
        self.ctx.add_command(label="Save", command=self.save_script)
        self.ctx.add_command(label="Save As...", command=self.save_script_as)

    def _build_add_context_menu(self):
        add_menu = tk.Menu(self.ctx, tearoff=0)
        group_menus = {}

        for name, spec in self.engine.ordered_specs():
            group = spec.group or "Other"
            if group not in group_menus:
                group_menus[group] = tk.Menu(add_menu, tearoff=0)
                add_menu.add_cascade(label=group, menu=group_menus[group])

            group_menus[group].add_command(
                label=name,
                command=lambda n=name: self._add_command_by_name(n)
            )

        self.ctx_add_group_menus = group_menus
        return add_menu

    def _on_script_click(self, event):
        # Select line on click
        self.script_text.focus_set()
        line_num = int(self.script_text.index(f"@{event.x},{event.y}").split('.')[0])
        self._select_script_line(line_num - 1)  # -1 because Text widget is 1-indexed

    def _on_script_right_click(self, event):
        # Select line and show context menu
        self.script_text.focus_set()
        line_num = int(self.script_text.index(f"@{event.x},{event.y}").split('.')[0])
        self._select_script_line(line_num - 1)  # -1 because Text widget is 1-indexed
        try:
            self.ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx.grab_release()

    def _on_script_double_click(self, event):
        # Edit command on double-click
        line_num = int(self.script_text.index(f"@{event.x},{event.y}").split('.')[0])
        self._select_script_line(line_num - 1)  # -1 because Text widget is 1-indexed
        self.edit_command()

    def _on_script_delete_key(self, event):
        self.delete_command()
        return "break"

    def _on_script_copy_key(self, event):
        self.copy_command()
        return "break"

    def _on_script_paste_key(self, event):
        self.paste_command()
        return "break"

    def _select_script_line(self, idx):
        """Select a script line by index."""
        if idx < 0 or idx >= len(self.engine.commands):
            return

        # Clear previous selection
        self.script_text.tag_remove("selected", "1.0", "end")

        # Select new line (Text widget is 1-indexed)
        line_start = f"{idx + 1}.0"
        line_end = f"{idx + 1}.end"
        self.script_text.tag_add("selected", line_start, line_end)
        self.selected_script_line = idx

    def _on_video_mouse_leave(self, event):
        self._last_video_xy = None
        self.video_mouse_xy_var.set("x: -, y: -")

    def _on_video_mouse_move(self, event):
        xy = self._event_to_frame_xy(event)
        if xy is None:
            self._last_video_xy = None
            self.video_mouse_xy_var.set("x: -, y: -")
            return
        x, y = xy
        self._last_video_xy = (x, y)
        self.video_mouse_xy_var.set(f"x: {x}, y: {y}")

    def _on_video_double_click(self, event):
        """Double-click to pop out camera window"""
        # Only pop out if camera is running
        if not self.cam_running:
            return
        self.popout_camera()

    def popout_camera(self):
        """Pop out the camera to a separate window"""
        if self.popout_window is not None:
            return  # Already popped out

        # Create popout window
        self.popout_window = CameraPopoutWindow(self, self._on_popout_close)
        self.set_status("Camera popped out (double-click for fullscreen)")

        # Hide main video display (keep label but clear image)
        self.video_label.configure(image="")
        self.video_label.imgtk = None

    def _on_popout_close(self):
        """Handle popout window closing - return to embedded mode"""
        if self.popout_window is not None:
            self.popout_window.close()
            self.popout_window = None
            # Disable camera focus since the focused window is closing
            self._disable_camera_keyboard_focus()
            self.set_status("Camera returned to main window")

    def _on_camera_selected(self, event=None):
        device = self.cam_var.get().strip()
        if device:
            self._persist_setting_value("default_camera_device", device)

    def _on_audio_input_selected(self, event=None):
        name = self.audio_input_var.get().strip()
        if name and name not in ("No input devices", "PyAudio not installed"):
            self._persist_setting_value("default_audio_input_device", name)

    def _on_audio_output_selected(self, event=None):
        name = self.audio_output_var.get().strip()
        if name and name not in ("No output devices", "PyAudio not installed"):
            self._persist_setting_value("default_audio_output_device", name)

    def _on_com_selected(self, event=None):
        port = self.com_var.get().strip()
        if port:
            self._persist_setting_value("default_com_port", port)



    # ---- camera
    def refresh_cameras(self):
        cams = list_dshow_video_devices()
        self.cam_combo["values"] = cams
        current = self.cam_var.get().strip()
        saved = (self._settings.get("default_camera_device") or "").strip()
        selection = None
        if current in cams:
            selection = current
        elif saved and saved in cams:
            selection = saved
        elif cams:
            selection = cams[0]
        if selection:
            self.cam_var.set(selection)

    def toggle_camera(self):
        if self.cam_running:
            self.stop_camera()
        else:
            self.start_camera()

    def start_camera(self):
        device = self.cam_var.get().strip()
        if not device:
            messagebox.showwarning("No camera", "Select a camera.")
            return
        device_spec = f"video={device}"
        cmd = [
            ffmpeg_path(), "-f", "dshow", "-i", device_spec,
            "-s", f"{self.cam_width}x{self.cam_height}",
            "-r", str(self.cam_fps),
            "-pix_fmt", "bgr24",
            "-vcodec", "rawvideo",
            "-f", "rawvideo",
            "-"
        ]
        try:
            self.cam_proc = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
                bufsize=10**7,
                creationflags=_SUBPROCESS_FLAGS
            )
        except FileNotFoundError:
            messagebox.showerror("ffmpeg not found", "ffmpeg was not found on PATH.")
            return
        except Exception as e:
            messagebox.showerror("Camera error", str(e))
            return

        self.cam_running = True
        self.cam_toggle_btn.configure(text="Stop Cam")
        self.set_status(f"Camera streaming: {device}")
        self._persist_setting_value("default_camera_device", device)

        self.cam_thread = threading.Thread(target=self._camera_reader_loop, daemon=True)
        self.cam_thread.start()
        # Ensure visible when camera starts
        self.show_camera_panel()


    def stop_camera(self):
        self.cam_running = False
        self.cam_toggle_btn.configure(text="Start Cam")
        if self.cam_proc:
            try:
                self.cam_proc.kill()
                self.cam_proc.wait(timeout=1)
            except Exception:
                pass
        self.cam_proc = None
        with self.frame_lock:
            self.latest_frame_bgr = None

        # Close popout window if open
        if self.popout_window is not None:
            self.popout_window.close()
            self.popout_window = None

        # Disable camera focus if it was active
        self._disable_camera_keyboard_focus()

        self.set_status("Camera stopped.")
        # Auto-hide if camera stopped
        self.hide_camera_panel()


    def toggle_camera_panel(self):
        if self.camera_panel_hidden:
            self.show_camera_panel()
        else:
            self.hide_camera_panel()

    def hide_camera_panel(self):
        # Save current sash position so we can restore it
        try:
            self.main_pane.update_idletasks()
            if hasattr(self.main_pane, "sashpos"):
                self._saved_sash_x = self.main_pane.sashpos(0)
            else:
                # tk.PanedWindow doesn't provide a getter; store a reasonable default
                self._saved_sash_x = int(self.main_pane.winfo_width() * 0.45)
        except Exception:
            self._saved_sash_x = int(self.main_pane.winfo_width() * 0.45)

        # Collapse left pane
        try:
            if hasattr(self.main_pane, "sashpos"):
                self.main_pane.sashpos(0, 0)
            else:
                self.main_pane.sash_place(0, 0, 0)
        except Exception:
            pass

        # Hide audio controls
        if hasattr(self, "audio_input_frame"):
            self.audio_input_frame.grid_remove()
        if hasattr(self, "audio_output_frame"):
            self.audio_output_frame.grid_remove()

        self.camera_panel_hidden = True
        self.set_status("Camera panel hidden.")
        self.cam_display_btn.configure(text="Show Cam")

    def show_camera_panel(self):
        try:
            self.main_pane.update_idletasks()
            total = self.main_pane.winfo_width()
            pane_h = self.main_pane.winfo_height()

            # Calculate minimum width needed to show camera at reasonable size
            # Account for coord bar and audio controls (~80px overhead)
            available_h = max(100, pane_h - 80)

            # Calculate ideal width based on camera aspect ratio
            cam_aspect = self.cam_width / self.cam_height if self.cam_height > 0 else 1.5
            ideal_cam_width = int(available_h * cam_aspect)

            # Use saved position if available and reasonable, otherwise calculate based on camera
            x = self._saved_sash_x
            if x is None or x < 150:
                # Default: use ideal camera width, but cap at 50% of total width
                x = min(ideal_cam_width, int(total * 0.5))

            # Ensure camera pane gets at least minimum viable space
            min_cam_width = min(300, int(total * 0.3))
            x = max(min_cam_width, min(x, total - 220))  # Keep right pane usable too

            if hasattr(self.main_pane, "sashpos"):
                self.main_pane.sashpos(0, x)
            else:
                self.main_pane.sash_place(0, x, 0)
        except Exception:
            pass

        # Show audio controls
        if hasattr(self, "audio_input_frame"):
            self.audio_input_frame.grid()
        if hasattr(self, "audio_output_frame"):
            self.audio_output_frame.grid()

        self.camera_panel_hidden = False
        self.set_status("Camera panel shown.")
        self.cam_display_btn.configure(text="Hide Cam")

    def apply_video_ratio(self, persist=True):
        ratio = (self.ratio_var.get()).strip()

        # choose width based on base_video_width, compute height from ratio
        w = int(self.base_video_width)

        if ratio == "4:3 (DS Single Screen)":
            h = int(round(w * 3 / 4))
        elif ratio == "3:2 (GBA)":
            h = int(round(w * 2 / 3))
        elif ratio == "16:9 (Standard)":
            h = int(round(w * 9 / 16))
        elif ratio == "2:3 (DS Dual Screen)":
            h = int(round(w * 3 / 2))
        elif ratio == "5:3 (3DS Top Screen)":
            h = int(round(w * 3 / 5))

        elif ratio == "5:6 (3DS Dual Screen)":
            h = int(round(w * 6 / 5))
        else:
            messagebox.showerror("Ratio", f"Unknown ratio: {ratio}")
            return

        # Force even dims (some devices/filters behave better)
        w = max(160, (w // 2) * 2)
        h = max(120, (h // 2) * 2)

        # Apply to camera pipeline
        was_running = self.cam_running
        if was_running:
            self.stop_camera()

        self.cam_width = w
        self.cam_height = h

        # Clear stored display size so coordinate mapper stays correct
        self._disp_img_w = 0
        self._disp_img_h = 0

        status_msg = f"Video ratio set to {ratio} ({w}x{h})."
        if persist:
            self._settings["camera_ratio"] = ratio
            if not save_settings(self._settings):
                status_msg = f"{status_msg} Settings save failed."
        self.set_status(status_msg)

        if was_running:
            self.start_camera()



    def _camera_reader_loop(self):
        if not self.cam_proc or not self.cam_proc.stdout:
            return
        frame_size = self.cam_width * self.cam_height * 3
        while self.cam_running and self.cam_proc and self.cam_proc.stdout:
            try:
                raw = self.cam_proc.stdout.read(frame_size)
                if not raw:
                    # Process ended or pipe closed
                    break
                if len(raw) != frame_size:
                    # Incomplete frame, skip it
                    continue
                frame = np.frombuffer(raw, dtype=np.uint8).reshape((self.cam_height, self.cam_width, 3))
                with self.frame_lock:
                    self.latest_frame_bgr = frame
            except Exception:
                # Handle any read errors (broken pipe, etc.)
                break

        # If we exited due to error, ensure camera state is updated
        if self.cam_running:
            self.root.after(0, lambda: self.set_status("Camera disconnected unexpectedly"))

    def _schedule_frame_update(self):
        self._update_video_frame()
        self.root.after(15, self._schedule_frame_update)

    def _update_video_frame(self):
        with self.frame_lock:
            frame = None if self.latest_frame_bgr is None else self.latest_frame_bgr.copy()
        if frame is None:
            return
        rgb = frame[:, :, ::-1]
        img = Image.fromarray(rgb)

        # Route to popout window if active, otherwise to main window
        if self.popout_window is not None:
            # Update popout window with PIL image (it does its own scaling)
            self.popout_window.update_frame(img)
        else:
            # Update main window - scale to fit available space
            self.video_label.update_idletasks()
            available_w = self.video_label.winfo_width()
            available_h = self.video_label.winfo_height()

            # Fallback: if label not yet sized, try to get size from main pane
            if available_w <= 1 or available_h <= 1:
                try:
                    self.main_pane.update_idletasks()
                    # Get sash position to determine left pane width
                    if hasattr(self.main_pane, "sashpos"):
                        available_w = self.main_pane.sashpos(0)
                    else:
                        available_w = self.cam_width  # Use camera dimensions as fallback
                    # Estimate available height from main pane height minus controls
                    pane_h = self.main_pane.winfo_height()
                    available_h = max(100, pane_h - 80)  # Reserve space for coord bar and audio controls
                except Exception:
                    pass

            # Scale if we have valid dimensions, otherwise show at native size
            if available_w > 1 and available_h > 1:
                scaled_img = scale_image_to_fit(img, available_w, available_h)
            else:
                scaled_img = img

            tk_img = ImageTk.PhotoImage(scaled_img)
            self._disp_img_w = tk_img.width()
            self._disp_img_h = tk_img.height()
            self.video_label.imgtk = tk_img
            self.video_label.configure(image=tk_img)

    # ---- audio
    def refresh_audio_devices(self):
        if not PYAUDIO_AVAILABLE:
            self.audio_input_combo["values"] = ["PyAudio not installed"]
            self.audio_output_combo["values"] = ["PyAudio not installed"]
            self.audio_input_var.set("PyAudio not installed")
            self.audio_output_var.set("PyAudio not installed")
            if hasattr(self, "audio_toggle_btn"):
                self.audio_toggle_btn.configure(state="disabled")
            return

        inputs, outputs = list_audio_devices()
        self.audio_input_devices = inputs
        self.audio_output_devices = outputs

        input_names = [name for idx, name in inputs]
        output_names = [name for idx, name in outputs]

        self.audio_input_combo["values"] = input_names if input_names else ["No input devices"]
        self.audio_output_combo["values"] = output_names if output_names else ["No output devices"]

        current_input = self.audio_input_var.get().strip()
        saved_input = (self._settings.get("default_audio_input_device") or "").strip()
        input_selection = None
        if current_input in input_names:
            input_selection = current_input
        elif saved_input and saved_input in input_names:
            input_selection = saved_input
        elif input_names:
            input_selection = input_names[0]
        if input_selection:
            self.audio_input_var.set(input_selection)

        current_output = self.audio_output_var.get().strip()
        saved_output = (self._settings.get("default_audio_output_device") or "").strip()
        output_selection = None
        if current_output in output_names:
            output_selection = current_output
        elif saved_output and saved_output in output_names:
            output_selection = saved_output
        elif output_names:
            output_selection = output_names[0]
        if output_selection:
            self.audio_output_var.set(output_selection)

    def toggle_audio(self):
        if self.audio_running:
            self.stop_audio()
        else:
            self.start_audio()

    def start_audio(self):
        if not PYAUDIO_AVAILABLE:
            messagebox.showwarning("Audio", "PyAudio is not installed.\nInstall with: pip install pyaudio")
            return

        input_name = self.audio_input_var.get().strip()
        output_name = self.audio_output_var.get().strip()

        if not input_name or input_name == "No input devices" or input_name == "PyAudio not installed":
            messagebox.showwarning("Audio", "Select an audio input device.")
            return
        if not output_name or output_name == "No output devices" or output_name == "PyAudio not installed":
            messagebox.showwarning("Audio", "Select an audio output device.")
            return

        # Find device indices
        input_idx = None
        output_idx = None

        for idx, name in self.audio_input_devices:
            if name == input_name:
                input_idx = idx
                break
        for idx, name in self.audio_output_devices:
            if name == output_name:
                output_idx = idx
                break

        if input_idx is None or output_idx is None:
            messagebox.showerror("Audio", "Could not find selected devices.\nTry refreshing the device list.")
            return

        try:
            self.audio_pyaudio = pyaudio.PyAudio()

            # Get device info to determine format
            input_info = self.audio_pyaudio.get_device_info_by_index(input_idx)
            output_info = self.audio_pyaudio.get_device_info_by_index(output_idx)

            # Use device default rates and channels (WASAPI requires exact match)
            input_rate = int(input_info.get('defaultSampleRate', 48000))
            output_rate = int(output_info.get('defaultSampleRate', 48000))
            input_channels = min(2, int(input_info.get('maxInputChannels', 1)))
            output_channels = min(2, int(output_info.get('maxOutputChannels', 2)))

            # Use larger buffer for WASAPI stability and reduce callback frequency
            input_chunk = 2048 * 4
            output_chunk = 4096 * 2

            # Create a larger queue with more buffering to prevent underruns
            import queue
            self.audio_queue = queue.Queue(maxsize=200)

            # Resampling state for smoother conversion
            self.audio_resample_buffer = np.array([], dtype=np.int16)
            self.audio_prebuffer_ready = False
            self.audio_underruns = 0
            self.audio_overruns = 0

            # Check if we can use simple decimation/interpolation (for integer ratios)
            rate_ratio = input_rate / output_rate
            use_simple_resample = abs(rate_ratio - round(rate_ratio)) < 0.01  # Close to integer ratio

            # Try to import scipy for high-quality resampling (optional)
            try:
                from scipy import signal as scipy_signal
                use_scipy = True
            except ImportError:
                use_scipy = False
                scipy_signal = None

            # Input stream callback - captures audio and puts in queue
            def input_callback(in_data, frame_count, time_info, status):
                try:
                    self.audio_queue.put_nowait(in_data)
                except queue.Full:
                    self.audio_overruns += 1
                return (None, pyaudio.paContinue)

            # Output stream callback - gets audio from queue and plays it
            def output_callback(in_data, frame_count, time_info, status):
                # Wait for queue to fill up a bit before starting (pre-buffering)
                if not self.audio_prebuffer_ready:
                    if self.audio_queue.qsize() >= 3:
                        self.audio_prebuffer_ready = True
                    else:
                        # Still pre-buffering, output silence
                        silence = np.zeros(frame_count * output_channels, dtype=np.int16)
                        return (silence.tobytes(), pyaudio.paContinue)

                try:
                    # Accumulate enough input data for conversion
                    accumulated_data = []

                    # Get chunks more aggressively to empty queue
                    chunks_to_get = max(2, min(8, self.audio_queue.qsize()))
                    for _ in range(chunks_to_get):
                        try:
                            data = self.audio_queue.get_nowait()
                            accumulated_data.append(np.frombuffer(data, dtype=np.int16))
                        except queue.Empty:
                            break

                    if not accumulated_data:
                        # No data available - underrun
                        self.audio_underruns += 1
                        silence = np.zeros(frame_count * output_channels, dtype=np.int16)
                        return (silence.tobytes(), pyaudio.paContinue)

                    # Combine all accumulated data
                    audio_data = np.concatenate(accumulated_data)

                    # Reshape based on input channels
                    if input_channels > 1:
                        audio_data = audio_data.reshape(-1, input_channels)

                    # Add to resample buffer
                    if len(self.audio_resample_buffer) > 0:
                        if input_channels > 1:
                            audio_data = np.vstack([self.audio_resample_buffer.reshape(-1, input_channels), audio_data])
                        else:
                            audio_data = np.concatenate([self.audio_resample_buffer, audio_data])

                    # Handle sample rate conversion (optimized)
                    if input_rate != output_rate:
                        ratio = input_rate / output_rate  # e.g., 96000/48000 = 2.0
                        input_len = len(audio_data)
                        output_len_needed = frame_count

                        # Calculate how many input samples we need
                        input_samples_needed = int(output_len_needed * ratio) + int(ratio) + 1

                        if input_len >= input_samples_needed:
                            # We have enough data for conversion
                            if use_simple_resample and abs(ratio - 2.0) < 0.01:
                                # Fast decimation by 2 (96kHz -> 48kHz)
                                samples_needed = output_len_needed * 2
                                if input_channels == 1:
                                    audio_data_converted = audio_data[:samples_needed:2]  # Take every 2nd sample
                                else:
                                    audio_data_converted = audio_data[:samples_needed:2, :]
                                samples_used = samples_needed
                            elif use_scipy and scipy_signal is not None:
                                # High-quality scipy resampling
                                samples_to_use = int(output_len_needed * ratio)
                                if input_channels == 1:
                                    audio_data_converted = scipy_signal.resample(
                                        audio_data[:samples_to_use], output_len_needed
                                    ).astype(np.int16)
                                else:
                                    audio_data_converted = np.column_stack([
                                        scipy_signal.resample(audio_data[:samples_to_use, ch], output_len_needed).astype(np.int16)
                                        for ch in range(input_channels)
                                    ])
                                samples_used = samples_to_use
                            else:
                                # Simple nearest-neighbor (fastest fallback)
                                indices = (np.arange(output_len_needed) * ratio).astype(int)
                                indices = np.clip(indices, 0, input_len - 1)
                                if input_channels == 1:
                                    audio_data_converted = audio_data[indices]
                                else:
                                    audio_data_converted = audio_data[indices, :]
                                samples_used = int(output_len_needed * ratio)

                            # Store remaining samples for next callback
                            if samples_used < input_len:
                                if input_channels > 1:
                                    self.audio_resample_buffer = audio_data[samples_used:].flatten()
                                else:
                                    self.audio_resample_buffer = audio_data[samples_used:]
                            else:
                                self.audio_resample_buffer = np.array([], dtype=np.int16)

                            audio_data = audio_data_converted
                        else:
                            # Not enough data, save for next time and output silence
                            self.audio_resample_buffer = audio_data.flatten()
                            silence = np.zeros(frame_count * output_channels, dtype=np.int16)
                            return (silence.tobytes(), pyaudio.paContinue)
                    else:
                        # No rate conversion needed
                        self.audio_resample_buffer = np.array([], dtype=np.int16)

                    # Handle channel conversion
                    if input_channels == 1 and output_channels == 2:
                        # Mono to stereo: duplicate channel
                        audio_data = np.column_stack([audio_data, audio_data])
                    elif input_channels == 2 and output_channels == 1:
                        # Stereo to mono: average channels
                        audio_data = audio_data.mean(axis=1).astype(np.int16)

                    # Ensure correct shape and size
                    audio_data = audio_data.flatten()
                    expected_samples = frame_count * output_channels

                    if len(audio_data) < expected_samples:
                        # Pad with last value to avoid clicks
                        last_value = audio_data[-1] if len(audio_data) > 0 else 0
                        padding = np.full(expected_samples - len(audio_data), last_value, dtype=np.int16)
                        audio_data = np.concatenate([audio_data, padding])
                    elif len(audio_data) > expected_samples:
                        # Trim excess
                        audio_data = audio_data[:expected_samples]

                    return (audio_data.tobytes(), pyaudio.paContinue)

                except Exception:
                    # Output silence on error
                    silence = np.zeros(frame_count * output_channels, dtype=np.int16)
                    return (silence.tobytes(), pyaudio.paContinue)

            # Open input stream
            self.audio_input_stream = self.audio_pyaudio.open(
                format=pyaudio.paInt16,
                channels=input_channels,
                rate=input_rate,
                input=True,
                input_device_index=input_idx,
                frames_per_buffer=input_chunk,
                stream_callback=input_callback
            )

            # Open output stream
            self.audio_output_stream = self.audio_pyaudio.open(
                format=pyaudio.paInt16,
                channels=output_channels,
                rate=output_rate,
                output=True,
                output_device_index=output_idx,
                frames_per_buffer=output_chunk,
                stream_callback=output_callback
            )

            # Start both streams
            self.audio_input_stream.start_stream()
            self.audio_output_stream.start_stream()

            self.audio_running = True
            self.audio_toggle_btn.configure(text="Stop Audio")
            self._persist_setting_value("default_audio_input_device", input_name)
            self._persist_setting_value("default_audio_output_device", output_name)

            conversion_info = ""
            if input_rate != output_rate:
                conversion_info += f" [rate: {input_rate}→{output_rate} Hz]"
            if input_channels != output_channels:
                conversion_info += f" [ch: {input_channels}→{output_channels}]"

            self.set_status(f"Audio streaming: {input_name} → {output_name}{conversion_info}")

        except Exception as e:
            messagebox.showerror("Audio error", f"Failed to start audio:\n{e}")
            self.stop_audio()

    def stop_audio(self):
        self.audio_running = False
        self.audio_toggle_btn.configure(text="Start Audio")

        # Stop and close input stream
        try:
            if hasattr(self, 'audio_input_stream') and self.audio_input_stream:
                self.audio_input_stream.stop_stream()
                self.audio_input_stream.close()
                self.audio_input_stream = None
        except Exception:
            pass

        # Stop and close output stream
        try:
            if hasattr(self, 'audio_output_stream') and self.audio_output_stream:
                self.audio_output_stream.stop_stream()
                self.audio_output_stream.close()
                self.audio_output_stream = None
        except Exception:
            pass

        # Legacy support for old single stream
        try:
            if hasattr(self, 'audio_stream') and self.audio_stream:
                self.audio_stream.stop_stream()
                self.audio_stream.close()
                self.audio_stream = None
        except Exception:
            pass

        # Clear queue
        try:
            if hasattr(self, 'audio_queue') and self.audio_queue:
                while not self.audio_queue.empty():
                    try:
                        self.audio_queue.get_nowait()
                    except:
                        break
                self.audio_queue = None
        except Exception:
            pass

        # Clear resampling state
        try:
            if hasattr(self, 'audio_resample_buffer'):
                self.audio_resample_buffer = None
            if hasattr(self, 'audio_prebuffer_ready'):
                self.audio_prebuffer_ready = False
            if hasattr(self, 'audio_underruns'):
                self.audio_underruns = 0
                self.audio_overruns = 0
        except Exception:
            pass

        # Terminate PyAudio
        try:
            if hasattr(self, 'audio_pyaudio') and self.audio_pyaudio:
                self.audio_pyaudio.terminate()
                self.audio_pyaudio = None
        except Exception:
            pass

        self.set_status("Audio stopped.")

    # ---- serial
    def refresh_ports(self):
        ports = list_com_ports()
        self.com_combo["values"] = ports
        current = self.com_var.get().strip()
        saved = (self._settings.get("default_com_port") or "").strip()
        selection = None
        if current in ports:
            selection = current
        elif saved and saved in ports:
            selection = saved
        elif ports:
            selection = ports[0]
        if selection:
            self.com_var.set(selection)

    def toggle_serial(self):
        if self.serial.connected:
            self.stop_script()
            self.serial.disconnect()
            self.ser_btn.configure(text="Connect")
        else:
            port = self.com_var.get().strip()
            if not port:
                messagebox.showwarning("No COM", "Select a COM port.")
                return
            try:
                self.serial.connect(port)
                self.ser_btn.configure(text="Disconnect")
                self._persist_setting_value("default_com_port", port)
            except Exception as e:
                messagebox.showerror("Serial error", str(e))
    def set_channel(self):
        if not self.serial.connected:
            messagebox.showwarning("Not connected", "Connect to the COM device first.")
            return

        ch_str = (self.chan_var.get() or "").strip()
        try:
            ch = int(ch_str)
        except ValueError:
            messagebox.showerror("Invalid channel", "Channel must be a number.")
            return

        if not (1 <= ch <= 255):
            messagebox.showerror("Invalid channel", "Channel must be between 1 and 255.")
            return

        # If your device expects 0x01.. and you want UI to show 1.., this maps directly.
        ch_byte = ch & 0xFF

        if not messagebox.askyesno(
            "Change Channel",
            f"Send channel set to {ch} (0x{ch_byte:02X})?\n\n"
            "Note: receiver must be power-cycled after changing the channel.",
            parent=self.root
        ):
            return

        try:
            self.serial.send_channel_set(ch_byte)  # sends: 0x43, ch, 0x00
            self.set_status(f"Channel set sent: {ch} (0x{ch_byte:02X}). Power-cycle receiver.")
        except Exception as e:
            messagebox.showerror("Channel error", str(e))

    def open_settings_dialog(self):
        """Open the settings dialog for keybinds and 3DS configuration."""
        try:
            port = int(self.threeds_port_var.get())
        except ValueError:
            port = 4950

        def on_apply(theme_settings):
            theme_setting = theme_settings.get("theme", "auto")
            self._settings["custom_theme"] = theme_settings.get("custom_theme", {})
            self._settings["theme"] = theme_setting
            self.apply_theme_setting(theme_setting)

        def on_save(settings):
            # Update keybindings
            self.kb_bindings = settings["keybindings"]

            # Update 3DS settings
            threeds = settings["threeds"]
            self.threeds_ip_var.set(threeds["ip"])
            self.threeds_port_var.set(str(threeds["port"]))

            # Apply theme
            theme_setting = settings.get("theme", "auto")
            self._settings["custom_theme"] = settings.get("custom_theme", {})
            self.apply_theme_setting(theme_setting)

            # Save to file
            self._settings["keybindings"] = settings["keybindings"]
            self._settings["threeds"] = settings["threeds"]
            self._settings["discord"] = settings["discord"]
            self._settings["theme"] = theme_setting
            self._settings["custom_theme"] = settings.get("custom_theme", {})
            self._settings["confirm_delete"] = settings.get("confirm_delete", True)
            if save_settings(self._settings):
                self.set_status("Settings saved.")
            else:
                self.set_status("Settings updated (could not save to file).")

        dialog = SettingsDialog(
            self.root,
            keybindings=self.kb_bindings,
            threeds_ip=self.threeds_ip_var.get(),
            threeds_port=port,
            theme_mode=self._theme_setting,
            discord_settings=self._settings.get("discord", {}),
            custom_theme=self._settings.get("custom_theme", {}),
            theme_colors=THEME_COLORS,
            confirm_delete=self._settings.get("confirm_delete", True),
            on_save_callback=on_save,
            on_apply_callback=on_apply
        )
        self.root.wait_window(dialog)

    def on_backend_changed(self):
        desired_3ds = (self.backend_var.get() == "3DS Input Redirection")

        if desired_3ds and self.engine.running:
            messagebox.showwarning("3DS", "Stop the script before enabling/changing 3DS backend.")
            self.backend_var.set("USB Serial")
            desired_3ds = False

        if desired_3ds:
            try:
                self._enable_input_redirection_backend()
            except Exception as e:
                messagebox.showerror("3DS", str(e))
                self._disable_input_redirection_backend()
                self.backend_var.set("USB Serial")
                desired_3ds = False
        else:
            self._disable_input_redirection_backend()

        # Disable serial-specific UI when using 3DS
        if hasattr(self, "com_combo"):
            try:
                state = "disabled" if desired_3ds else "readonly"
                self.com_combo.configure(state=state)
            except Exception:
                pass
        if hasattr(self, "connect_btn"):
            try:
                self.connect_btn.configure(state=("disabled" if desired_3ds else "normal"))
            except Exception:
                pass

        # Switch active backend pointer
        self._select_active_backend()

        self.set_status(f"Output backend set to: {self.backend_var.get()}")

    def _enable_input_redirection_backend(self):
        ip = (self.threeds_ip_var.get() or "").strip()
        if not ip:
            raise RuntimeError("Please configure the 3DS IP address in Settings first.")
        try:
            port = int((self.threeds_port_var.get() or "4950").strip())
        except ValueError:
            raise RuntimeError("Invalid port. Please configure in Settings.")

        self._disable_input_redirection_backend()
        self.input_redirection_backend = InputRedirection.InputRedirectionBackend(ip=ip, port=port)
        self.input_redirection_backend.connect()
        self.set_status(f"3DS backend enabled: {ip}:{port}")

    def _disable_input_redirection_backend(self):
        try:
            if self.input_redirection_backend:
                self.input_redirection_backend.disconnect()
        except Exception:
            pass
        self.input_redirection_backend = None

    def reset_output_neutral(self):
        self._select_active_backend()
        b = self.active_backend
        if b and getattr(b, "connected", False):
            try:
                if hasattr(b, "reset_neutral"):
                    b.reset_neutral()
                else:
                    b.set_buttons([])
            except Exception:
                pass

    # ---- scripts: new/load/save
    def refresh_scripts(self):
        files = list_script_files()
        self.script_combo["values"] = files
        if files and self.script_var.get() not in files:
            self.script_var.set(files[0])

    def _confirm_discard_if_dirty(self):
        if not self.dirty:
            return True
        return messagebox.askyesno("Unsaved changes", "You have unsaved changes. Discard them?")

    def new_script(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before creating a new script.")
            return
        if not self._confirm_discard_if_dirty():
            return

        # Ask for script name
        raw = simpledialog.askstring("New Script", "Enter script name (will be saved in ./scripts):", parent=self.root)
        if raw is None:
            return  # cancelled

        filename = safe_script_filename(raw)
        if not filename:
            messagebox.showerror("Invalid name", "Please enter a valid script name.")
            return

        os.makedirs("scripts", exist_ok=True)
        path = os.path.join("scripts", filename)

        if os.path.exists(path):
            if not messagebox.askyesno("File exists", f"'{filename}' already exists. Overwrite?", parent=self.root):
                return

        # Create/overwrite file immediately (so it appears in the list)
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump([], f, indent=2, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Create error", str(e), parent=self.root)
            return

        # Load the new (empty) script into the app
        self.engine.commands = []
        self.engine.vars = {}
        self.engine.ip = 0
        # Tolerant rebuild so editor doesn't complain about blocks
        try:
            self.engine.rebuild_indexes(strict=False)
        except Exception:
            pass

        self.script_path = path
        self.mark_dirty(False)  # already saved as empty
        self.populate_script_view()
        self.refresh_vars_view()

        # Refresh dropdown and select new script
        self.refresh_scripts()
        self.script_var.set(filename)

        self.set_status(f"New script created: {filename}")


    def load_script_from_dropdown(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before loading another script.")
            return
        if not self._confirm_discard_if_dirty():
            return

        name = self.script_var.get().strip()
        if not name:
            messagebox.showwarning("No script", "Select a script JSON from ./scripts.")
            return
        path = os.path.join("scripts", name)

        try:
            self.engine.load_script(path)
            self.script_path = path
            self.mark_dirty(False)
            self.populate_script_view()
            self.refresh_vars_view()
        except Exception as e:
            messagebox.showerror("Load error", str(e))

    def save_script(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before saving.")
            return
        if self.engine.commands is None:
            messagebox.showinfo("Nothing to save", "No script loaded/created yet.")
            return

        if not self.script_path:
            return self.save_script_as()

        try:
            os.makedirs(os.path.dirname(self.script_path) or ".", exist_ok=True)
            with open(self.script_path, "w", encoding="utf-8") as f:
                json.dump(self.engine.commands, f, indent=2, ensure_ascii=False)
            self.mark_dirty(False)
            self.set_status(f"Saved: {self.script_path}")
            self.refresh_scripts()
        except Exception as e:
            messagebox.showerror("Save error", str(e))

    def save_script_as(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before saving.")
            return

        os.makedirs("scripts", exist_ok=True)
        initial = os.path.basename(self.script_path) if self.script_path else "new_script.json"
        path = filedialog.asksaveasfilename(
            title="Save Script As",
            initialdir=os.path.abspath("scripts"),
            initialfile=initial,
            defaultextension=".json",
            filetypes=[("JSON scripts", "*.json"), ("All files", "*.*")]
        )
        if not path:
            return

        self.script_path = path
        self.save_script()

    # ---- script viewer
    def populate_script_view(self, preserve_view=False):
        yview = None
        xview = None
        if preserve_view:
            yview = self.script_text.yview()
            xview = self.script_text.xview()

        # Enable editing to modify content
        self.script_text.config(state="normal")
        self.script_text.delete("1.0", "end")

        indent_on = bool(self.indent_var.get()) if hasattr(self, "indent_var") else True
        depth = 0

        def _collect_math_exprs(value, out):
            if isinstance(value, str):
                if value.startswith("=") and len(value) > 1:
                    out.append(value)
                return
            if isinstance(value, dict):
                for v in value.values():
                    _collect_math_exprs(v, out)
            elif isinstance(value, (list, tuple)):
                for v in value:
                    _collect_math_exprs(v, out)

        for i, c in enumerate(self.engine.commands):
            cmd = c.get("cmd")
            spec = self.engine.registry.get(cmd)
            pretty = spec.format_fn(c) if spec else f"(unknown) {cmd}"

            # Decrease indent BEFORE printing for closing blocks
            if indent_on and cmd in ("end_if", "end_while"):
                depth = max(0, depth - 1)

            if indent_on:
                pretty = ("      " * depth) + pretty  # 6 spaces per level

            # Format line with index number (right-aligned in 4 chars)
            # 1 char for marker (">" when current), then 4 for index, then 2 spaces
            line_text = f" {i:4}  {pretty}\n"

            # Insert line
            line_start = self.script_text.index("end-1c")
            self.script_text.insert("end", line_text)
            line_end = self.script_text.index("end-1c")

            # Apply syntax highlighting
            if cmd == "comment":
                # Color entire comment line green
                self.script_text.tag_add("comment", line_start, line_end)
            else:
                # Find and color all $variable references in blue
                # Search after the line number (skip first 7 chars: "    0  " including marker area)
                content_start_col = 7
                line_num = i + 1  # Text widget is 1-indexed

                math_exprs = []
                _collect_math_exprs(c, math_exprs)
                if math_exprs:
                    haystack = line_text[content_start_col:]
                    for expr in dict.fromkeys(math_exprs):
                        for match in re.finditer(re.escape(expr), haystack):
                            expr_start = f"{line_num}.{content_start_col + match.start()}"
                            expr_end = f"{line_num}.{content_start_col + match.end()}"
                            self.script_text.tag_add("math", expr_start, expr_end)

                # Find all variable references in the line
                for match in re.finditer(r'\$\w+', line_text[content_start_col:]):
                    var_start = f"{line_num}.{content_start_col + match.start()}"
                    var_end = f"{line_num}.{content_start_col + match.end()}"
                    self.script_text.tag_add("variable", var_start, var_end)

                # Highlight file paths for run_python and discord_status commands
                if cmd == "run_python":
                    filepath = c.get("file", "")
                    if filepath:
                        haystack = line_text[content_start_col:]
                        idx = haystack.find(filepath)
                        if idx >= 0:
                            fp_start = f"{line_num}.{content_start_col + idx}"
                            fp_end = f"{line_num}.{content_start_col + idx + len(filepath)}"
                            self.script_text.tag_add("filepath", fp_start, fp_end)
                elif cmd == "discord_status":
                    image_path = c.get("image", "")
                    if image_path:
                        haystack = line_text[content_start_col:]
                        idx = haystack.find(image_path)
                        if idx >= 0:
                            fp_start = f"{line_num}.{content_start_col + idx}"
                            fp_end = f"{line_num}.{content_start_col + idx + len(image_path)}"
                            self.script_text.tag_add("filepath", fp_start, fp_end)

            # Increase indent AFTER printing for opening blocks
            if indent_on and cmd in ("if", "while"):
                depth += 1

        # Disable editing to make it read-only
        self.script_text.config(state="disabled")
        # Reset IP marker tracking since content was rebuilt
        self._prev_ip = None
        self.highlight_ip(-1)
        if preserve_view:
            if yview is not None:
                self.script_text.yview_moveto(yview[0])
            if xview is not None:
                self.script_text.xview_moveto(xview[0])





    def refresh_vars_view(self):
        self.vars_tree.delete(*self.vars_tree.get_children())
        for k, v in sorted(self.engine.vars.items(), key=lambda kv: kv[0]):
            self.vars_tree.insert("", "end", values=(k, json.dumps(v, ensure_ascii=False)))

    def run_script(self):
        # Turn off keyboard control so we don't fight the script
        if self.kb_enabled.get():
            self.kb_enabled.set(False)
        if self.kb_camera_focused:
            self.kb_camera_focused = False
        self._release_all_keyboard_buttons()

        # Check if script contains run_python commands and Python is available
        if not is_python_available():
            has_run_python = any(
                c.get("cmd") == "run_python"
                for c in self.engine.commands
                if isinstance(c, dict)
            )
            if has_run_python:
                result = messagebox.askyesno(
                    "Python Required",
                    "This script contains run_python commands but Python is not installed.\n\n"
                    "Would you like to download Python now?\n\n"
                    "(This is a one-time ~11 MB download)",
                    parent=self.root
                )
                if result:
                    dialog = PythonDownloadDialog(self.root)
                    self.root.wait_window(dialog)
                    if not dialog.result:
                        return  # Download failed or cancelled, don't run script
                else:
                    return  # User declined, don't run script

        try:
            self.engine.rebuild_indexes(strict=True)  # strict only when running
            self.engine.run()
        except Exception as e:
            messagebox.showerror("Run error", str(e))


    def stop_script(self):
        if self.engine.running:
            self.engine.stop()
            self.reset_output_neutral()
            self.highlight_ip(-1)

    # ---- ip highlight
    def on_ip_update(self, ip):
        self.root.after(0, lambda: self.highlight_ip(ip))

    def highlight_ip(self, ip):
        # Clear IP highlight (syntax highlighting is preserved in tags)
        self.script_text.tag_remove("ip", "1.0", "end")

        # Temporarily enable editing to update the marker
        self.script_text.config(state="normal")

        # Remove ">" marker from the previous IP line
        prev_ip = getattr(self, "_prev_ip", None)
        if prev_ip is not None and prev_ip >= 0 and prev_ip < len(self.engine.commands):
            prev_line_marker = f"{prev_ip + 1}.0"
            prev_line_marker_end = f"{prev_ip + 1}.1"
            self.script_text.delete(prev_line_marker, prev_line_marker_end)
            self.script_text.insert(prev_line_marker, " ")

        if ip is None or ip < 0:
            self._prev_ip = None
            self.script_text.config(state="disabled")
            return

        if ip < len(self.engine.commands):
            # Add ">" marker to the current IP line
            line_marker = f"{ip + 1}.0"
            line_marker_end = f"{ip + 1}.1"
            self.script_text.delete(line_marker, line_marker_end)
            self.script_text.insert(line_marker, ">")

            # Highlight the instruction pointer line (Text widget is 1-indexed)
            line_start = f"{ip + 1}.0"
            line_end = f"{ip + 1}.end"
            self.script_text.tag_add("ip", line_start, line_end)
            self.script_text.see(line_start)

        self._prev_ip = ip
        self.script_text.config(state="disabled")

    # ---- editor actions
    def _get_selected_index(self):
        """Get the index of the currently selected script line."""
        return self.selected_script_line

    def _reindex_after_edit(self):
        try:
            self.engine.rebuild_indexes(strict=False)  # tolerant during editing
        except Exception as e:
            # This should be rare now; but don't crash UI
            self.set_status(f"Index warning: {e}")
        self.populate_script_view(preserve_view=True)
        self.mark_dirty(True)
        self._update_structure_warning()

    def _update_structure_warning(self):
        msgs = []
        if getattr(self.engine, "_unclosed_ifs", []):
            msgs.append(f"unclosed if: {len(self.engine._unclosed_ifs)}")
        if getattr(self.engine, "_unclosed_whiles", []):
            msgs.append(f"unclosed while: {len(self.engine._unclosed_whiles)}")
        if msgs:
            self.set_status("Script structure incomplete (" + ", ".join(msgs) + "). Add end_if / end_while.")




    def _get_script_clipboard_payload(self):
        try:
            raw = self.root.clipboard_get()
        except tk.TclError:
            return None
        raw = raw.strip()
        if not raw:
            return None
        try:
            return json.loads(raw)
        except Exception:
            return None

    def _normalize_command_payload(self, payload):
        if isinstance(payload, dict) and payload.get("cmd"):
            return [payload]
        if isinstance(payload, list) and payload:
            if all(isinstance(item, dict) and item.get("cmd") for item in payload):
                return payload
        return None

    def copy_command(self):
        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("Copy", "Select a command to copy.")
            return

        payload = [self.engine.commands[idx]]
        self._script_cmd_clipboard = copy.deepcopy(payload)
        self._copy_to_clipboard(json.dumps(payload, ensure_ascii=False, indent=2))
        self.set_status("Command copied to clipboard.")

    def paste_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        payload = self._get_script_clipboard_payload()
        commands = self._normalize_command_payload(payload)
        if commands is None:
            commands = copy.deepcopy(getattr(self, "_script_cmd_clipboard", None) or [])
        else:
            commands = copy.deepcopy(commands)

        if not commands:
            messagebox.showinfo("Paste", "Clipboard does not contain a command.")
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)
        last_cmd_idx = None

        for cmd in commands:
            self.engine.commands.insert(insert_at, cmd)
            last_cmd_idx = insert_at
            insert_at += 1
            if cmd.get("cmd") == "if":
                self.engine.commands.insert(insert_at, {"cmd": "end_if"})
                insert_at += 1
            elif cmd.get("cmd") == "while":
                self.engine.commands.insert(insert_at, {"cmd": "end_while"})
                insert_at += 1

        self._reindex_after_edit()
        if last_cmd_idx is not None:
            self._select_script_line(last_cmd_idx)


    def add_command(self):
        self._open_add_command_dialog()

    def _add_command_by_name(self, name):
        if name not in self.engine.registry:
            messagebox.showwarning("Unknown command", f"Command '{name}' not found.")
            return
        self._open_add_command_dialog(initial_cmd={"cmd": name})

    def _open_add_command_dialog(self, initial_cmd=None):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)

        dlg = CommandEditorDialog(
            self.root, self.engine.registry,
            initial_cmd=initial_cmd, title="Add Command",
            test_callback=self._dialog_test_callback,
            select_area_callback=self._open_region_selector,
            select_color_callback=self._open_color_picker,
            select_area_color_callback=self._open_area_color_picker
        )
        self.root.wait_window(dlg)
        if dlg.result is None:
            return

        self.engine.commands.insert(insert_at, dlg.result)
        if dlg.result["cmd"] == "if":
            self.engine.commands.insert(insert_at + 1, {"cmd": "end_if"})
        elif dlg.result["cmd"] == "while":
            self.engine.commands.insert(insert_at + 1, {"cmd": "end_while"})

        self._reindex_after_edit()

        self._select_script_line(insert_at)

    def edit_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("Edit", "Select a command to edit.")
            return

        initial = self.engine.commands[idx]
        dlg = CommandEditorDialog(
            self.root, self.engine.registry,
            initial_cmd=initial, title="Edit Command",
            test_callback=self._dialog_test_callback,
            select_area_callback=self._open_region_selector,
            select_color_callback=self._open_color_picker,
            select_area_color_callback=self._open_area_color_picker
        )
        self.root.wait_window(dlg)
        if dlg.result is None:
            return

        self.engine.commands[idx] = dlg.result
        self._reindex_after_edit()
        self._select_script_line(idx)

    def _confirm_delete_command(self):
        if not self._settings.get("confirm_delete", True):
            return True

        result = {"ok": False}
        dont_ask_var = tk.BooleanVar(value=False)

        dlg = tk.Toplevel(self.root)
        dlg.title("Delete")
        dlg.transient(self.root)
        dlg.grab_set()

        body = ttk.Frame(dlg, padding=12)
        body.grid(row=0, column=0, sticky="nsew")
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        ttk.Label(body, text="Delete selected command?").grid(
            row=0, column=0, columnspan=2, sticky="w"
        )
        ttk.Checkbutton(
            body,
            text="Don't ask again",
            variable=dont_ask_var
        ).grid(row=1, column=0, columnspan=2, sticky="w", pady=(8, 0))

        btn_row = ttk.Frame(body)
        btn_row.grid(row=2, column=0, columnspan=2, sticky="e", pady=(12, 0))

        def on_delete():
            result["ok"] = True
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        ttk.Button(btn_row, text="Cancel", command=on_cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btn_row, text="Delete", command=on_delete).pack(side="right")

        dlg.bind("<Escape>", lambda e: on_cancel())
        dlg.bind("<Return>", lambda e: on_delete())

        dlg.update_idletasks()
        x = self.root.winfo_rootx() + 120
        y = self.root.winfo_rooty() + 120
        dlg.geometry(f"+{x}+{y}")

        self.root.wait_window(dlg)

        if result["ok"] and dont_ask_var.get():
            self._settings["confirm_delete"] = False
            if save_settings(self._settings):
                self.set_status("Delete confirmation disabled.")
            else:
                self.set_status("Delete confirmation disabled (could not save).")

        return result["ok"]

    def delete_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        if idx is None:
            return

        if not self._confirm_delete_command():
            return

        del self.engine.commands[idx]
        self._reindex_after_edit()

        new_idx = min(idx, len(self.engine.commands) - 1)
        if new_idx >= 0:
            self._select_script_line(new_idx)

    def move_command(self, delta):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        if idx is None:
            return
        j = idx + delta
        if not (0 <= j < len(self.engine.commands)):
            return

        self.engine.commands[idx], self.engine.commands[j] = self.engine.commands[j], self.engine.commands[idx]
        self._reindex_after_edit()

        self._select_script_line(j)

    def add_comment(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)

        self.engine.commands.insert(insert_at, {"cmd": "comment", "text": "New comment"})
        self._reindex_after_edit()

        self._select_script_line(insert_at)

    def _resolve_test_value(self, v):
        """
        Resolve '$var' in editor tests using current engine vars.
        """
        if isinstance(v, str) and v.startswith("$"):
            return self.engine.vars.get(v[1:], None)
        return v

    def _open_region_selector(self, initial_region, on_select_callback, on_close_callback=None):
        """
        Open the region selector window for selecting an area on the camera.

        Args:
            initial_region: Optional tuple (x, y, width, height) to show initially
            on_select_callback: Callback with (x, y, width, height) when confirmed
            on_close_callback: Optional callback called when window closes (for any reason)

        Returns:
            True if the selector window was opened, False otherwise
        """
        if not self.cam_running:
            messagebox.showwarning(
                "Camera Required",
                "Please start the camera first to select a region.",
                parent=self.root
            )
            return False

        # Open the region selector window
        RegionSelectorWindow(self, on_select_callback, initial_region=initial_region, on_close_callback=on_close_callback)
        return True

    def _open_color_picker(self, initial_x, initial_y, on_select_callback, on_close_callback=None):
        """
        Open the color picker window for selecting a color from the camera.

        Args:
            initial_x: Optional initial X coordinate
            initial_y: Optional initial Y coordinate
            on_select_callback: Callback with (x, y, r, g, b) when confirmed
            on_close_callback: Optional callback called when window closes (for any reason)

        Returns:
            True if the picker window was opened, False otherwise
        """
        if not self.cam_running:
            messagebox.showwarning(
                "Camera Required",
                "Please start the camera first to pick a color.",
                parent=self.root
            )
            return False

        # Open the color picker window
        ColorPickerWindow(self, on_select_callback, initial_x=initial_x, initial_y=initial_y, on_close_callback=on_close_callback)
        return True

    def _open_area_color_picker(self, initial_region, initial_rgb, on_select_callback, on_close_callback=None):
        """
        Open the area color picker window for selecting an area and color from the camera.

        Args:
            initial_region: Optional tuple (x, y, width, height) to show initially
            initial_rgb: Optional tuple/list (r, g, b) for initial target color
            on_select_callback: Callback with (x, y, width, height, r, g, b) when confirmed
            on_close_callback: Optional callback called when window closes (for any reason)

        Returns:
            True if the picker window was opened, False otherwise
        """
        if not self.cam_running:
            messagebox.showwarning(
                "Camera Required",
                "Please start the camera first to select an area and color.",
                parent=self.root
            )
            return False

        # Import here to avoid circular import
        from camera import AreaColorPickerWindow

        # Open the area color picker window
        AreaColorPickerWindow(self, on_select_callback, initial_region=initial_region,
                            initial_rgb=initial_rgb, on_close_callback=on_close_callback)
        return True

    def test_command_dialog(self, cmd_obj):
        """
        Returns (title, message) for a given cmd_obj.
        Currently supports: find_color, read_text
        """
        cmd = cmd_obj.get("cmd")
        match cmd:
            case "find_color":
                frame = self.get_latest_frame()
                if frame is None:
                    return ("find_color Test", "No camera frame available.\nStart the camera first.")

                # Read args
                x = int(self._resolve_test_value(cmd_obj.get("x", 0)))
                y = int(self._resolve_test_value(cmd_obj.get("y", 0)))
                rgb = cmd_obj.get("rgb", [0, 0, 0])
                tol = float(self._resolve_test_value(cmd_obj.get("tol", 10)))
                out = (cmd_obj.get("out") or "match").strip()

                h, w, _ = frame.shape
                if not (0 <= x < w and 0 <= y < h):
                    return ("find_color Test",
                            f"Point out of bounds.\n"
                            f"Requested: ({x},{y})\n"
                            f"Frame size: {w}x{h}")

                # Sample pixel (frame is BGR)
                b, g, r = frame[y, x].tolist()
                sampled_rgb = (int(r), int(g), int(b))
                target = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

                # Calculate CIE76 Delta E
                delta_e = ScriptEngine.delta_e_cie76(sampled_rgb, target)
                ok = delta_e <= tol

                # Interpretation of Delta E values
                if delta_e <= 1:
                    perception = "imperceptible"
                elif delta_e <= 2:
                    perception = "barely perceptible"
                elif delta_e <= 10:
                    perception = "noticeable"
                elif delta_e <= 49:
                    perception = "obvious"
                else:
                    perception = "very different"

                msg = (
                    f"Point: ({x},{y})\n"
                    f"Sampled RGB: {list(sampled_rgb)}\n"
                    f"Target RGB:  {list(target)}\n\n"
                    f"Delta E (CIE76): {delta_e:.2f} ({perception})\n"
                    f"Tolerance: {tol}\n\n"
                    f"Result (would set ${out}): {ok}"
                )
                return ("find_color Test", msg)

            case "find_area_color":
                frame = self.get_latest_frame()
                if frame is None:
                    return ("find_area_color Test", "No camera frame available.\nStart the camera first.")

                # Read args
                x = int(self._resolve_test_value(cmd_obj.get("x", 0)))
                y = int(self._resolve_test_value(cmd_obj.get("y", 0)))
                width = int(self._resolve_test_value(cmd_obj.get("width", 10)))
                height = int(self._resolve_test_value(cmd_obj.get("height", 10)))
                rgb = cmd_obj.get("rgb", [0, 0, 0])
                tol = float(self._resolve_test_value(cmd_obj.get("tol", 10)))
                out = (cmd_obj.get("out") or "match").strip()

                h_frame, w_frame, _ = frame.shape

                # Clamp region to frame bounds
                x = max(0, min(x, w_frame - 1))
                y = max(0, min(y, h_frame - 1))
                x2 = max(x + 1, min(x + width, w_frame))
                y2 = max(y + 1, min(y + height, h_frame))

                # Check bounds
                if x >= w_frame or y >= h_frame:
                    return ("find_area_color Test",
                            f"Region out of bounds.\n"
                            f"Top-left: ({x},{y})\n"
                            f"Frame size: {w_frame}x{h_frame}")

                # Extract region (BGR)
                region_bgr = frame[y:y2, x:x2]

                if region_bgr.size == 0:
                    return ("find_area_color Test", "Region is empty (size is 0).")

                # Calculate average color
                avg_b = float(np.mean(region_bgr[:, :, 0]))
                avg_g = float(np.mean(region_bgr[:, :, 1]))
                avg_r = float(np.mean(region_bgr[:, :, 2]))

                avg_rgb = (int(avg_r), int(avg_g), int(avg_b))
                target = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

                # Calculate CIE76 Delta E
                delta_e = ScriptEngine.delta_e_cie76(avg_rgb, target)
                ok = delta_e <= tol

                # Interpretation of Delta E values
                if delta_e <= 1:
                    perception = "imperceptible"
                elif delta_e <= 2:
                    perception = "barely perceptible"
                elif delta_e <= 10:
                    perception = "noticeable"
                elif delta_e <= 49:
                    perception = "obvious"
                else:
                    perception = "very different"

                actual_w = x2 - x
                actual_h = y2 - y

                msg = (
                    f"Region: ({x},{y}) {actual_w}x{actual_h}\n"
                    f"Pixels sampled: {region_bgr.shape[0] * region_bgr.shape[1]}\n\n"
                    f"Average RGB: {list(avg_rgb)}\n"
                    f"Target RGB:  {list(target)}\n\n"
                    f"Delta E (CIE76): {delta_e:.2f} ({perception})\n"
                    f"Tolerance: {tol}\n\n"
                    f"Result (would set ${out}): {ok}"
                )
                return ("find_area_color Test", msg)

            case "wait_for_color":
                frame = self.get_latest_frame()
                if frame is None:
                    return ("wait_for_color Test", "No camera frame available.\nStart the camera first.")

                # Read args
                x = int(self._resolve_test_value(cmd_obj.get("x", 0)))
                y = int(self._resolve_test_value(cmd_obj.get("y", 0)))
                rgb = cmd_obj.get("rgb", [0, 0, 0])
                tol = float(self._resolve_test_value(cmd_obj.get("tol", 10)))
                interval = float(self._resolve_test_value(cmd_obj.get("interval", 0.1)))
                timeout = float(self._resolve_test_value(cmd_obj.get("timeout", 0)))
                wait_for = bool(self._resolve_test_value(cmd_obj.get("wait_for", True)))
                out = (cmd_obj.get("out") or "match").strip()

                h, w, _ = frame.shape
                if not (0 <= x < w and 0 <= y < h):
                    return ("wait_for_color Test",
                            f"Point out of bounds.\n"
                            f"Requested: ({x},{y})\n"
                            f"Frame size: {w}x{h}")

                # Sample pixel (frame is BGR)
                b, g, r = frame[y, x].tolist()
                sampled_rgb = (int(r), int(g), int(b))
                target = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

                # Calculate CIE76 Delta E
                delta_e = ScriptEngine.delta_e_cie76(sampled_rgb, target)
                matches = delta_e <= tol

                # Interpretation of Delta E values
                if delta_e <= 1:
                    perception = "imperceptible"
                elif delta_e <= 2:
                    perception = "barely perceptible"
                elif delta_e <= 10:
                    perception = "noticeable"
                elif delta_e <= 49:
                    perception = "obvious"
                else:
                    perception = "very different"

                wait_mode = "match" if wait_for else "no match"
                condition_met = matches == wait_for
                timeout_str = f"{timeout}s" if timeout > 0 else "none"

                msg = (
                    f"Point: ({x},{y})\n"
                    f"Sampled RGB: {list(sampled_rgb)}\n"
                    f"Target RGB:  {list(target)}\n\n"
                    f"Delta E (CIE76): {delta_e:.2f} ({perception})\n"
                    f"Tolerance: {tol}\n"
                    f"Currently matches: {matches}\n\n"
                    f"Wait mode: wait for {wait_mode}\n"
                    f"Check interval: {interval}s\n"
                    f"Timeout: {timeout_str}\n\n"
                    f"Condition met now: {condition_met}\n"
                    f"Would set ${out}: {condition_met}"
                )
                return ("wait_for_color Test", msg)

            case "wait_for_color_area":
                frame = self.get_latest_frame()
                if frame is None:
                    return ("wait_for_color_area Test", "No camera frame available.\nStart the camera first.")

                # Read args
                x = int(self._resolve_test_value(cmd_obj.get("x", 0)))
                y = int(self._resolve_test_value(cmd_obj.get("y", 0)))
                width = int(self._resolve_test_value(cmd_obj.get("width", 10)))
                height = int(self._resolve_test_value(cmd_obj.get("height", 10)))
                rgb = cmd_obj.get("rgb", [0, 0, 0])
                tol = float(self._resolve_test_value(cmd_obj.get("tol", 10)))
                interval = float(self._resolve_test_value(cmd_obj.get("interval", 0.1)))
                timeout = float(self._resolve_test_value(cmd_obj.get("timeout", 0)))
                wait_for = bool(self._resolve_test_value(cmd_obj.get("wait_for", True)))
                out = (cmd_obj.get("out") or "match").strip()

                h_frame, w_frame, _ = frame.shape

                # Clamp region to frame bounds
                x = max(0, min(x, w_frame - 1))
                y = max(0, min(y, h_frame - 1))
                x2 = max(x + 1, min(x + width, w_frame))
                y2 = max(y + 1, min(y + height, h_frame))

                # Check bounds
                if x >= w_frame or y >= h_frame:
                    return ("wait_for_color_area Test",
                            f"Region out of bounds.\n"
                            f"Top-left: ({x},{y})\n"
                            f"Frame size: {w_frame}x{h_frame}")

                # Extract region (BGR)
                region_bgr = frame[y:y2, x:x2]

                if region_bgr.size == 0:
                    return ("wait_for_color_area Test", "Region is empty (size is 0).")

                # Calculate average color
                avg_b = float(np.mean(region_bgr[:, :, 0]))
                avg_g = float(np.mean(region_bgr[:, :, 1]))
                avg_r = float(np.mean(region_bgr[:, :, 2]))

                avg_rgb = (int(avg_r), int(avg_g), int(avg_b))
                target = (int(rgb[0]), int(rgb[1]), int(rgb[2]))

                # Calculate CIE76 Delta E
                delta_e = ScriptEngine.delta_e_cie76(avg_rgb, target)
                matches = delta_e <= tol

                # Interpretation of Delta E values
                if delta_e <= 1:
                    perception = "imperceptible"
                elif delta_e <= 2:
                    perception = "barely perceptible"
                elif delta_e <= 10:
                    perception = "noticeable"
                elif delta_e <= 49:
                    perception = "obvious"
                else:
                    perception = "very different"

                actual_w = x2 - x
                actual_h = y2 - y
                wait_mode = "match" if wait_for else "no match"
                condition_met = matches == wait_for
                timeout_str = f"{timeout}s" if timeout > 0 else "none"

                msg = (
                    f"Region: ({x},{y}) {actual_w}x{actual_h}\n"
                    f"Pixels sampled: {region_bgr.shape[0] * region_bgr.shape[1]}\n\n"
                    f"Average RGB: {list(avg_rgb)}\n"
                    f"Target RGB:  {list(target)}\n\n"
                    f"Delta E (CIE76): {delta_e:.2f} ({perception})\n"
                    f"Tolerance: {tol}\n"
                    f"Currently matches: {matches}\n\n"
                    f"Wait mode: wait for {wait_mode}\n"
                    f"Check interval: {interval}s\n"
                    f"Timeout: {timeout_str}\n\n"
                    f"Condition met now: {condition_met}\n"
                    f"Would set ${out}: {condition_met}"
                )
                return ("wait_for_color_area Test", msg)

            case "read_text":
                # Check if pytesseract is available
                if not ScriptEngine.PYTESSERACT_AVAILABLE:
                    return ("read_text Test",
                            "pytesseract is not installed.\n\n"
                            "Install with:\n"
                            "  pip install pytesseract\n\n"
                            "Also install Tesseract OCR:\n"
                            "  Windows: https://github.com/UB-Mannheim/tesseract/wiki\n"
                            "  Linux: sudo apt install tesseract-ocr")

                frame = self.get_latest_frame()
                if frame is None:
                    return ("read_text Test", "No camera frame available.\nStart the camera first.")

                # Read args
                x = int(self._resolve_test_value(cmd_obj.get("x", 0)))
                y = int(self._resolve_test_value(cmd_obj.get("y", 0)))
                width = int(self._resolve_test_value(cmd_obj.get("width", 100)))
                height = int(self._resolve_test_value(cmd_obj.get("height", 20)))
                scale = int(self._resolve_test_value(cmd_obj.get("scale", 4)))
                threshold = int(self._resolve_test_value(cmd_obj.get("threshold", 0)))
                invert = bool(self._resolve_test_value(cmd_obj.get("invert", False)))
                psm = int(self._resolve_test_value(cmd_obj.get("psm", 7)))
                whitelist = str(self._resolve_test_value(cmd_obj.get("whitelist", "")))
                out = (cmd_obj.get("out") or "text").strip()

                h_frame, w_frame, _ = frame.shape

                # Check bounds
                if x < 0 or y < 0 or x >= w_frame or y >= h_frame:
                    return ("read_text Test",
                            f"Region out of bounds.\n"
                            f"Top-left: ({x},{y})\n"
                            f"Frame size: {w_frame}x{h_frame}")

                # Perform OCR
                try:
                    text = ScriptEngine.ocr_region(
                        frame, x, y, width, height,
                        scale=scale, threshold=threshold, invert=invert,
                        psm=psm, whitelist=whitelist
                    )
                except Exception as e:
                    return ("read_text Test", f"OCR Error:\n{e}")

                # Build result message
                msg = (
                    f"Region: ({x},{y}) {width}x{height}\n"
                    f"Settings:\n"
                    f"  Scale: {scale}x\n"
                    f"  Threshold: {threshold}\n"
                    f"  Invert: {invert}\n"
                    f"  PSM: {psm}\n"
                    f"  Whitelist: '{whitelist}'\n\n"
                    f"Recognized text (would set ${out}):\n"
                    f"───────────────────────\n"
                    f"{text if text else '(empty)'}\n"
                    f"───────────────────────"
                )
                return ("read_text Test", msg)

            case "play_sound":
                sound = cmd_obj.get("sound", "")
                volume_raw = self._resolve_test_value(cmd_obj.get("volume", 80))
                wait = bool(self._resolve_test_value(cmd_obj.get("wait", False)))

                try:
                    volume = int(round(float(volume_raw)))
                except (TypeError, ValueError):
                    volume = 80
                volume = max(0, min(100, volume))

                ok, msg = ScriptEngine.play_sound_file(sound, volume=volume, wait=wait)
                self.set_status(msg)
                return (None, None)

            case "prompt_input":
                title_raw = cmd_obj.get("title", "Input")
                message_raw = cmd_obj.get("message", "Enter value:")
                default_raw = cmd_obj.get("default", "")
                confirm_raw = cmd_obj.get("confirm", False)
                out = (cmd_obj.get("out") or "input").strip()

                title_val = self._resolve_test_value(title_raw)
                message_val = self._resolve_test_value(message_raw)
                default_val = self._resolve_test_value(default_raw)
                confirm_val = bool(self._resolve_test_value(confirm_raw))

                title = str(title_val) if title_val is not None else "Input"
                prompt = str(message_val) if message_val is not None else "Enter value:"
                default_display = "" if default_val is None else str(default_val)
                result = self.on_prompt_input(title, prompt, default_display, confirm_val)

                stored = result
                status_prefix = "Would store"
                if result is None:
                    stored = default_val if default_val is not None else ""
                    status_prefix = "Canceled prompt. Would store default"

                display_value = "" if stored is None else str(stored)
                if out:
                    self.set_status(f"{status_prefix} ${out} = {display_value!r}")
                else:
                    self.set_status(f"{status_prefix} {display_value!r}")
                return (None, None)

            case "prompt_choice":
                title_raw = cmd_obj.get("title", "Choose")
                message_raw = cmd_obj.get("message", "Select a value:")
                choices_raw = cmd_obj.get("choices", [])
                default_raw = cmd_obj.get("default", None)
                confirm_raw = cmd_obj.get("confirm", False)
                display_raw = cmd_obj.get("display", "dropdown")
                out = (cmd_obj.get("out") or "choice").strip()

                title_val = self._resolve_test_value(title_raw)
                message_val = self._resolve_test_value(message_raw)
                default_val = self._resolve_test_value(default_raw)
                confirm_val = bool(self._resolve_test_value(confirm_raw))
                display_val = self._resolve_test_value(display_raw)

                title = str(title_val) if title_val is not None else "Choose"
                prompt = str(message_val) if message_val is not None else "Select a value:"
                display = str(display_val) if display_val is not None else "dropdown"
                display = display.strip().lower()
                if display not in ("dropdown", "buttons"):
                    display = "dropdown"

                if isinstance(choices_raw, str) and choices_raw.strip().startswith("$"):
                    choices_val = self._resolve_test_value(choices_raw.strip())
                else:
                    choices_val = choices_raw
                if isinstance(choices_val, str):
                    try:
                        parsed = json.loads(choices_val)
                        choices_val = parsed
                    except Exception:
                        pass
                if not isinstance(choices_val, list):
                    return ("prompt_choice Test", "choices must be a list.")
                if not choices_val:
                    return ("prompt_choice Test", "choices list is empty.")

                default_index = None
                if default_val is not None:
                    try:
                        default_index = choices_val.index(default_val)
                    except ValueError:
                        default_index = None
                if default_index is None:
                    default_index = 0

                result_index = self.on_prompt_choice(title, prompt, choices_val, default_index, confirm_val, display)
                stored = default_val if result_index is None else None
                if result_index is not None and 0 <= result_index < len(choices_val):
                    stored = choices_val[result_index]
                if stored is None and result_index is not None:
                    stored = result_index

                status_prefix = "Would store"
                if result_index is None:
                    status_prefix = "Canceled prompt. Would store default"

                display_value = "" if stored is None else str(stored)
                if out:
                    self.set_status(f"{status_prefix} ${out} = {display_value!r}")
                else:
                    self.set_status(f"{status_prefix} {display_value!r}")
                return (None, None)

            case _:
                raise ValueError("No tester implemented for this command.")

    def _dialog_test_callback(self, cmd_obj):
        # Enable for commands with test support
        cmd = cmd_obj.get("cmd")
        match cmd:
            case "find_color" | "find_area_color" | "wait_for_color" | "wait_for_color_area" | "read_text" | "play_sound" | "prompt_input" | "prompt_choice":
                return self.test_command_dialog(cmd_obj)
            case _:
                raise ValueError("No test available for this command.")



    # ---- close
    def on_close(self):
        if self.engine.running:
            self.stop_script()

        if self.dirty:
            if not messagebox.askyesno("Unsaved changes", "You have unsaved changes. Exit anyway?"):
                return

        try:
            if self.serial.connected:
                self.serial.disconnect()
        except Exception:
            pass
        try:
            if self.cam_running:
                self.stop_camera()
        except Exception:
            pass
        try:
            if self.audio_running:
                self.stop_audio()
        except Exception:
            pass

        self._stop_theme_poll()
        self.root.destroy()


if __name__ == "__main__":
    os.makedirs("scripts", exist_ok=True)
    os.makedirs("py_scripts", exist_ok=True)
    root = tk.Tk()
    app = App(root)
    root.mainloop()
