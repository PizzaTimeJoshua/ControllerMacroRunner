"""
Controller Macro Runner (Camera + Serial + Script Engine + Editor)

Install:
  pip install numpy pillow pyserial pyaudio

FFmpeg:
  Ensure ffmpeg is on PATH:
    ffmpeg -version
"""
import os
import json
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import re
import numpy as np
from PIL import Image, ImageTk

from serial.tools import list_ports
import sys
from typing import Optional
import ThreeDSClasses
import SerialController
import ScriptEngine
import ScriptToPy

# Audio support (optional)
try:
    import pyaudio
    PYAUDIO_AVAILABLE = True
except ImportError:
    PYAUDIO_AVAILABLE = False
    pyaudio = None


def resource_path(rel_path: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)

def ffmpeg_path() -> str:
    # Prefer bundled ffmpeg.exe
    bundled = resource_path("bin/ffmpeg.exe")
    if os.path.exists(bundled):
        return bundled
    return "ffmpeg"  # fallback to PATH





# ----------------------------
# Camera (FFmpeg pipe -> Tk)
# ----------------------------

def list_dshow_video_devices():
    try:
        p = subprocess.run(
            [ffmpeg_path(), "-list_devices", "true", "-f", "dshow", "-i", "dummy"],
            capture_output=True
        )

    except FileNotFoundError:
        return []

    raw = p.stderr or b""
    try:
        text = raw.decode("mbcs", errors="replace")
    except Exception:
        text = raw.decode("utf-8", errors="replace")

    devices = []
    for line in text.splitlines():
        s = line.strip()
        if "Alternative name" in s:
            continue
        if "(video)" not in s:
            continue
        first = s.find('"')
        last = s.rfind('"')
        if first != -1 and last != -1 and last > first:
            name = s[first + 1:last].strip()
            if name and name not in devices:
                devices.append(name)
    return devices

def safe_script_filename(name: str) -> str:
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)  # Windows-illegal chars
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        return ""
    if not name.lower().endswith(".json"):
        name += ".json"
    return name

def list_python_files():
    folder = "py_scripts"
    if not os.path.isdir(folder):
        return []
    return sorted([n for n in os.listdir(folder) if n.lower().endswith(".py")])


# ----------------------------
# Audio (PyAudio)
# ----------------------------

def list_audio_devices():
    """Returns (input_devices, output_devices) as lists of (index, name) tuples."""
    if not PYAUDIO_AVAILABLE or pyaudio is None:
        return [], []

    try:
        p = pyaudio.PyAudio()
        inputs = []
        outputs = []

        for i in range(p.get_device_count()):
            try:
                info = p.get_device_info_by_index(i)
                name = info.get('name', f'Device {i}')
                max_in = info.get('maxInputChannels', 0)
                max_out = info.get('maxOutputChannels', 0)

                if max_in > 0:
                    inputs.append((i, name))
                if max_out > 0:
                    outputs.append((i, name))
            except Exception:
                continue

        p.terminate()
        return inputs, outputs
    except Exception:
        return [], []


# ----------------------------
# Editor Dialog (schema-driven)
# ----------------------------

class CommandEditorDialog(tk.Toplevel):
    def __init__(self, parent, registry, initial_cmd=None, title="Edit Command",test_callback = None):
        super().__init__(parent)
        self.parent = parent
        self.registry = registry
        self.result = None
        self.test_callback = test_callback


        self.title(title)
        self.transient(parent)
        self.grab_set()

        self.cmd_name_var = tk.StringVar()
        self.field_vars = {}
        self.widgets = {}

        top = ttk.Frame(self, padding=10)
        top.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        ttk.Label(top, text="Command:").grid(row=0, column=0, sticky="w")
        self.cmd_combo = ttk.Combobox(
            top, textvariable=self.cmd_name_var, state="readonly",
            values=self._ordered_command_names(), width=30
        )
        self.cmd_combo.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        top.columnconfigure(1, weight=1)

        self.doc_var = tk.StringVar(value="")
        ttk.Label(top, textvariable=self.doc_var, foreground="gray").grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 8)
        )

        self.fields_frame = ttk.Frame(top)
        self.fields_frame.grid(row=2, column=0, columnspan=2, sticky="nsew")
        top.rowconfigure(2, weight=1)

        bottom = ttk.Frame(top)
        bottom.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Button(bottom, text="Cancel", command=self._cancel).pack(side="right", padx=(6, 0))
        ttk.Button(bottom, text="Save", command=self._save).pack(side="right")
        self.test_btn = ttk.Button(bottom, text="Test", command=self._test)
        self.test_btn.pack(side="right", padx=(0, 6))

        self.cmd_combo.bind("<<ComboboxSelected>>", lambda e: self._render_fields())

        if initial_cmd and "cmd" in initial_cmd:
            self.cmd_name_var.set(initial_cmd["cmd"])
        else:
            keys = self._ordered_command_names()
            self.cmd_name_var.set("press" if "press" in keys else (keys[0] if keys else ""))


        self._render_fields(initial_cmd=initial_cmd)

        self.update_idletasks()
        x = parent.winfo_rootx() + 80
        y = parent.winfo_rooty() + 80
        self.geometry(f"+{x}+{y}")

    def _ordered_command_names(self):
        def keyfn(name):
            s = self.registry[name]
            return (s.group, s.order, s.name)
        return [n for n in sorted(self.registry.keys(), key=keyfn)]

    def _clear_fields(self):
        for child in self.fields_frame.winfo_children():
            child.destroy()
        self.field_vars.clear()
        self.widgets.clear()

    def _render_fields(self, initial_cmd=None):
        self._clear_fields()
        name = self.cmd_name_var.get()
        if not name:
            return

        spec = self.registry[name]
        self.doc_var.set(spec.doc or "")

        for r, field in enumerate(spec.arg_schema):
            key = field["key"]
            ftype = field["type"]
            help_text = field.get("help", "")
            default = field.get("default", "")

            ttk.Label(self.fields_frame, text=key + ":").grid(row=r, column=0, sticky="w", pady=3)

            init_val = default
            if initial_cmd and initial_cmd.get("cmd") == name and key in initial_cmd:
                init_val = initial_cmd[key]

            match ftype:
                case "int":
                    var = tk.StringVar(value=str(init_val))
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case "float":
                    var = tk.StringVar(value=str(init_val))
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case "str":
                    var = tk.StringVar(value=str(init_val))
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case "bool":
                    var = tk.BooleanVar(value=bool(init_val))
                    cb = ttk.Checkbutton(self.fields_frame, variable=var)
                    cb.grid(row=r, column=1, sticky="w", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = cb

                case "choice":
                    var = tk.StringVar(value=str(init_val))
                    combo = ttk.Combobox(
                        self.fields_frame, textvariable=var, state="readonly",
                        values=field.get("choices", []), width=28
                    )
                    combo.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = combo

                case "json":
                    var = tk.StringVar(value=json.dumps(init_val))
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case "rgb":
                    if isinstance(init_val, (list, tuple)) and len(init_val) == 3:
                        init_text = ",".join(str(int(x)) for x in init_val)
                    else:
                        init_text = str(init_val)
                    var = tk.StringVar(value=init_text)
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case "buttons":
                    frame = ttk.Frame(self.fields_frame)
                    frame.grid(row=r, column=1, sticky="ew", pady=3)
                    lb = tk.Listbox(frame, selectmode="multiple", height=6, exportselection=False)
                    sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
                    lb.configure(yscrollcommand=sb.set)
                    lb.pack(side="left", fill="both", expand=True)
                    sb.pack(side="left", fill="y")

                    for b in SerialController.ALL_BUTTONS:
                        lb.insert("end", b)

                    init_buttons = init_val if isinstance(init_val, list) else []
                    for i, b in enumerate(SerialController.ALL_BUTTONS):
                        if b in init_buttons:
                            lb.selection_set(i)

                    self.widgets[key] = lb
                    self.field_vars[key] = None

                case "pyfile":
                    # Dropdown of ./py_scripts/*.py, but allow typing arbitrary text too (absolute path)
                    var = tk.StringVar(value=str(init_val) if init_val is not None else "")
                    files = list_python_files()

                    combo = ttk.Combobox(
                        self.fields_frame,
                        textvariable=var,
                        values=files,
                        state="normal",   # allow typing too
                        width=28
                    )
                    combo.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = combo

                case "expr":
                    var = tk.StringVar(value=str(init_val))
                    ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                    ent.grid(row=r, column=1, sticky="ew", pady=3)
                    self.field_vars[key] = var
                    self.widgets[key] = ent

                case _:
                    ttk.Label(self.fields_frame, text=f"(unsupported type: {ftype})").grid(row=r, column=1, sticky="w")
                

            ttk.Label(self.fields_frame, text=help_text, foreground="gray").grid(row=r, column=2, sticky="w", padx=(8, 0))

        self.fields_frame.columnconfigure(1, weight=1)
        # Enable Test only if provided and command is supported
        if self.test_callback is None or spec.test == False:
            self.test_btn.pack_forget()
        else:
            self.test_btn.state(["!disabled"])
            self.test_btn.pack(side="right", padx=(0, 6))


    def _parse_field(self, key, field):
        ftype = field["type"]

        if ftype == "buttons":
            lb = self.widgets[key]
            return [lb.get(i) for i in lb.curselection()]

        var = self.field_vars[key]
        raw = var.get() if var is not None else None
        match ftype:
            case "int":
                return int(raw.strip())
            
            case "float":
                return float(raw.strip())

            case "str":
                return str(raw)

            case "bool":
                return bool(var.get())

            case "choice":
                return raw

            case "json":
                s = raw.strip()
                try:
                    return json.loads(s)
                except Exception:
                    return s

            case "rgb":
                s = raw.strip()
                if s.startswith("["):
                    v = json.loads(s)
                    if not (isinstance(v, list) and len(v) == 3):
                        raise ValueError("rgb must be [R,G,B]")
                    return [int(v[0]), int(v[1]), int(v[2])]
                parts = [p.strip() for p in s.split(",")]
                if len(parts) != 3:
                    raise ValueError("rgb must be 'R,G,B'")
                return [int(parts[0]), int(parts[1]), int(parts[2])]
            
            case "expr":
                    return raw

        raise ValueError(f"Unsupported type: {ftype}")

    def _save(self):
        name = self.cmd_name_var.get()
        if not name:
            return
        spec = self.registry[name]

        cmd_obj = {"cmd": name}
        try:
            for field in spec.arg_schema:
                key = field["key"]
                cmd_obj[key] = self._parse_field(key, field)

            for k in spec.required_keys:
                if k not in cmd_obj:
                    raise ValueError(f"Missing required key: {k}")

        except Exception as e:
            messagebox.showerror("Invalid input", str(e), parent=self)
            return

        self.result = cmd_obj
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()

    def _test(self):
        if self.test_callback is None:
            return

        # Build a cmd object from the current UI fields without closing the dialog
        name = self.cmd_name_var.get()
        spec = self.registry.get(name)
        if not spec:
            messagebox.showerror("Test", "Unknown command.", parent=self)
            return

        cmd_obj = {"cmd": name}
        try:
            for field in spec.arg_schema:
                key = field["key"]
                cmd_obj[key] = self._parse_field(key, field)
        except Exception as e:
            messagebox.showerror("Test", f"Invalid inputs:\n{e}", parent=self)
            return

        try:
            title, msg = self.test_callback(cmd_obj)
            messagebox.showinfo(title, msg, parent=self)
        except Exception as e:
            messagebox.showerror("Test error", str(e), parent=self)



# ----------------------------
# Helpers
# ----------------------------

def list_com_ports():
    return [p.device for p in list_ports.comports()]


def list_script_files():
    folder = "scripts"
    if not os.path.isdir(folder):
        return []
    return sorted([n for n in os.listdir(folder) if n.lower().endswith(".json")])


# ----------------------------
# Tkinter App
# ----------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.script_path = None
        self.dirty = False

        self.root.title("Controller Macro Runner")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.root.geometry("1200x700")

        # --- keyboard controller mode (manual control)
        self.kb_enabled = tk.BooleanVar(value=False)
        self.kb_bindings = {  # key -> controller button name (must match ALL_BUTTONS entries)
            "w": "Up",
            "a": "Left",
            "s": "Down",
            "d": "Right",
            "j": "A",
            "k": "B",
            "u": "X",
            "i": "Y",
            "enter": "Start",
            "space": "Select",
            "q": "L",
            "e": "R",
        }
        self.kb_down = set()         # set of pressed Tk keysyms (normalized)
        self.kb_buttons_held = set() # controller buttons currently held due to keyboard
        self._rebinding_target = None  # button being rebound in UI, or None

        # Global key events (manual controller)
        self.root.bind_all("<KeyPress>", self._on_key_press)
        self.root.bind_all("<KeyRelease>", self._on_key_release)

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

        # audio state
        self.audio_input_var = tk.StringVar()
        self.audio_output_var = tk.StringVar()
        self.audio_pyaudio = None
        self.audio_stream = None
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
        )
        
        # Output backend selection
        self.backend_var = tk.StringVar(value="USB Serial")  # "USB Serial" or "3DS Input Redirection"
        self.threeds_ip_var = tk.StringVar(value="192.168.1.1")
        self.threeds_port_var = tk.StringVar(value="4950")
        self.threeds_backend: Optional[ThreeDSClasses.ThreeDSBackend] = None

        # Active backend points to either self.serial or self.threeds_backend
        self.active_backend = self.serial

        self.engine.set_backend_getter(lambda: self.active_backend)

        self._build_ui()
        self._build_context_menu()
        self._schedule_frame_update()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.refresh_cameras()
        self.refresh_ports()
        self.refresh_scripts()
        self.refresh_audio_devices()
        self._update_title()

    # ---- title/dirty
    def mark_dirty(self, dirty=True):
        self.dirty = dirty
        self._update_title()

    def _update_title(self):
        name = os.path.basename(self.script_path) if self.script_path else "(unsaved script)"
        star = " *" if self.dirty else ""
        self.root.title(f"Controller Macro Runner - {name}{star}")

    # ---- status
    def set_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    # ---- engine tick (live vars)
    def on_engine_tick(self):
        self.root.after(0, self.refresh_vars_view)

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
        top.columnconfigure(1, weight=1)
        top.columnconfigure(13, weight=1)

        ttk.Checkbutton(top, text="Keyboard Control", variable=self.kb_enabled,
                        command=self._on_keyboard_toggle).grid(row=1, column=4, columnspan=2, padx=(10, 6), sticky="w")
        ttk.Button(top, text="Keybinds…", command=self.open_keybinds_window).grid(row=1, column=6, padx=(0, 6))

        # Camera controls
        ttk.Label(top, text="Camera:").grid(row=0, column=0, sticky="w")
        self.cam_var = tk.StringVar()
        self.cam_combo = ttk.Combobox(top, textvariable=self.cam_var, state="readonly", width=28)
        self.cam_combo.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_cameras).grid(row=0, column=2, padx=(0, 6))
        self.cam_toggle_btn = ttk.Button(top, text="Start Cam", command=self.toggle_camera)
        self.cam_toggle_btn.grid(row=0, column=3, padx=(6, 6))

        self.cam_display_btn = ttk.Button(top, text="Show Cam", command=self.toggle_camera_panel)
        self.cam_display_btn.grid(row=1, column=3, padx=(6,6))

        ttk.Label(top, text="Cam Ratio:").grid(row=1, column=0, sticky="w")

        self.ratio_var = tk.StringVar(value="3:2 (GBA)")
        self.ratio_combo = ttk.Combobox(
            top, textvariable=self.ratio_var, state="readonly",
            values=["3:2 (GBA)","16:9 (Standard)","4:3 (DS Single Screen)", "2:3 (DS Dual Screen)", "5:3 (3DS Top Screen)", "5:6 (3DS Dual Screen)"], width=6
        )
        self.ratio_combo.grid(row=1, column=1, sticky="ew", padx=(6, 6))

        ttk.Button(top, text="Apply", command=self.apply_video_ratio).grid(row=1, column=2, padx=(0, 6))



        # Serial controls
        ttk.Label(top, text="COM:").grid(row=0, column=4, sticky="w")
        self.com_var = tk.StringVar()
        self.com_combo = ttk.Combobox(top, textvariable=self.com_var, state="readonly", width=10)
        self.com_combo.grid(row=0, column=5, sticky="w", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_ports).grid(row=0, column=6, padx=(0, 6))
        self.ser_btn = ttk.Button(top, text="Connect", command=self.toggle_serial)
        self.ser_btn.grid(row=0, column=7, sticky="w", padx=(0, 18))

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

        ttk.Button(top, text="Set Channel", command=self.set_channel).grid(row=0, column=10, padx=(0, 18))

        # Backend Selectors
        ttk.Label(top, text="Output:").grid(row=2, column=0, sticky="w")

        self.backend_combo = ttk.Combobox(
            top, textvariable=self.backend_var, state="readonly",
            values=["USB Serial", "3DS Input Redirection"], width=20
        )
        self.backend_combo.grid(row=2, column=1, sticky="ew", padx=(6, 6),pady=(4,0))
        self.backend_combo.bind("<<ComboboxSelected>>", lambda e: self.on_backend_changed())

        # 3DS config (shown/hidden)
        self.threeds_ip_label = ttk.Label(top, text="3DS IP:")
        self.threeds_ip_entry = ttk.Entry(top, textvariable=self.threeds_ip_var, width=12)

        self.threeds_port_label = ttk.Label(top, text="Port:")
        self.threeds_port_entry = ttk.Entry(top, textvariable=self.threeds_port_var, width=6)

        self.threeds_enable_btn = ttk.Button(top, text="Enable 3DS", command=self.enable_threeds_backend)
        self.threeds_disable_btn = ttk.Button(top, text="Disable 3DS", command=self.disable_threeds_backend)

        # place them (we'll hide/show in on_backend_changed)
        self.threeds_ip_label.grid(row=2, column=2, sticky="w")
        self.threeds_ip_entry.grid(row=2, column=3, sticky="w", padx=(4, 8))

        self.threeds_port_label.grid(row=2, column=4, sticky="w")
        self.threeds_port_entry.grid(row=2, column=5, sticky="w", padx=(4, 8))

        self.threeds_enable_btn.grid(row=2, column=6, sticky="w", padx=(0, 6))
        self.threeds_disable_btn.grid(row=2, column=7, sticky="w")

        # initialize visibility
        self.on_backend_changed()



        # Script file controls
        ttk.Label(top, text="Script:").grid(row=0, column=12, sticky="w")
        self.script_var = tk.StringVar()
        self.script_combo = ttk.Combobox(top, textvariable=self.script_var, state="readonly", width=26)
        self.script_combo.grid(row=0, column=13, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_scripts).grid(row=0, column=14, padx=(0, 6))
        ttk.Button(top, text="Load", command=self.load_script_from_dropdown).grid(row=0, column=15, padx=(0, 6))
        ttk.Button(top, text="New", command=self.new_script).grid(row=0, column=16, padx=(0, 6))

        ttk.Button(top, text="Export .py", command= lambda: ScriptToPy.export_script_to_python(self)).grid(row=1, column=14, padx=(0, 6))
        ttk.Button(top, text="Save", command=self.save_script).grid(row=1, column=15, padx=(0, 6))
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
        self.video_label.bind("<Button-1>", self._on_video_click_copy)
        self.video_label.bind("<Shift-Button-1>", self._on_video_click_copy_json)


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

        # Initially hide audio controls (camera panel starts hidden)
        self.audio_input_frame.grid_remove()
        self.audio_output_frame.grid_remove()

        self.video_label.grid(row=0, column=0, sticky="nsew")

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



        # Right pane: vertically split Insert/Script/Vars
        right_split = tk.PanedWindow(right, orient=tk.VERTICAL, sashrelief=tk.RAISED, sashwidth=6, showhandle=False, bd=0)
        right_split.grid(row=0, column=0, sticky="nsew")
        right.rowconfigure(0, weight=1)

        # --- Insert panel
        insert_box = ttk.LabelFrame(right_split, text="Insert Command")
        insert_box.columnconfigure(0, weight=1)
        insert_box.rowconfigure(1, weight=1)

        self.cmd_search_var = tk.StringVar(value="")
        search_row = ttk.Frame(insert_box)
        search_row.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 2))
        search_row.columnconfigure(1, weight=1)

        ttk.Label(search_row, text="Search:").grid(row=0, column=0, sticky="w")
        search_entry = ttk.Entry(search_row, textvariable=self.cmd_search_var)
        search_entry.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        self.auto_close_block_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            search_row,
            text="Auto add end_if/end_while",
            variable=self.auto_close_block_var
        ).grid(row=0, column=2, sticky="w", padx=(10, 0))

        # Tree of commands with group headers
        self.insert_tree = ttk.Treeview(insert_box, columns=("cmd", "desc"), show="tree headings", height=8)
        self.insert_tree.heading("#0", text="Group")
        self.insert_tree.heading("cmd", text="Command")
        self.insert_tree.heading("desc", text="Description")
        self.insert_tree.column("#0", width=140, stretch=False)
        self.insert_tree.column("cmd", width=120, stretch=False)
        self.insert_tree.column("desc", width=420, stretch=True)
        self.insert_tree.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))

        ins_scr = ttk.Scrollbar(insert_box, orient="vertical", command=self.insert_tree.yview)
        self.insert_tree.configure(yscrollcommand=ins_scr.set)
        ins_scr.grid(row=1, column=1, sticky="ns", pady=(0, 6))

        # Buttons
        ins_btnrow = ttk.Frame(insert_box)
        ins_btnrow.grid(row=2, column=0, sticky="ew", padx=6, pady=(0, 6))
        ttk.Button(ins_btnrow, text="Insert", command=self.insert_selected_command).pack(side="left")
        ttk.Button(ins_btnrow, text="Insert + Edit", command=self.insert_selected_command_and_edit).pack(side="left", padx=(6, 0))

        # --- Script viewer
        script_box = ttk.LabelFrame(right_split, text="Script Commands (right-click menu, double-click edit)")
        script_box.rowconfigure(0, weight=1)
        script_box.columnconfigure(0, weight=1)

        self.script_tree = ttk.Treeview(
            script_box,
            columns=("idx", "pretty"),
            show="headings",
            height=12
        )
        self.script_tree.heading("idx", text="#")
        self.script_tree.heading("pretty", text="Command")
        self.script_tree.column("idx", width=40, anchor="e", stretch=False)
        self.script_tree.column("pretty", width=520, anchor="w", stretch=True)
        self.script_tree.grid(row=0, column=0, sticky="nsew")

        scr = ttk.Scrollbar(script_box, orient="vertical", command=self.script_tree.yview)
        self.script_tree.configure(yscrollcommand=scr.set)
        scr.grid(row=0, column=1, sticky="ns")

        btnrow = ttk.Frame(script_box)
        btnrow.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(btnrow, text="Add", command=self.add_command).pack(side="left", padx=2)
        ttk.Button(btnrow, text="Edit", command=self.edit_command).pack(side="left", padx=2)
        ttk.Button(btnrow, text="Delete", command=self.delete_command).pack(side="left", padx=2)
        ttk.Button(btnrow, text="Up", command=lambda: self.move_command(-1)).pack(side="left", padx=2)
        ttk.Button(btnrow, text="Down", command=lambda: self.move_command(1)).pack(side="left", padx=2)
        ttk.Button(btnrow, text="Comment", command=self.add_comment).pack(side="left", padx=2)

        # Indent view toggle (if you already have it, keep yours)
        self.indent_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(btnrow, text="Indent view", variable=self.indent_var,
                        command=self.populate_script_view).pack(side="right", padx=6)

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
        right_split.add(insert_box, minsize=170)
        right_split.add(script_box, minsize=220)
        right_split.add(vars_box, minsize=140)

        # Tags / bindings
        self.script_tree.tag_configure("ip", background="#dbeafe")
        self.script_tree.bind("<Button-3>", self._on_script_right_click)
        self.script_tree.bind("<Double-1>", self._on_script_double_click)

        # Insert panel bindings
        self.cmd_search_var.trace_add("write", lambda *args: self.populate_insert_panel())
        self.insert_tree.bind("<Double-1>", lambda e: self.insert_selected_command_and_edit())
        self.insert_tree.bind("<Return>", lambda e: self.insert_selected_command_and_edit())

        # Build initial insert list
        self.populate_insert_panel()


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
        if ks == "return":
            ks = "enter"
        return ks

    def _manual_control_allowed(self):
        # Only allow manual keyboard control when:
        # - enabled
        # - serial connected
        # - script NOT running
        return bool(self.kb_enabled.get()) and self.serial.connected and (not self.engine.running)

    def _on_keyboard_toggle(self):
        # Turning off: release everything
        if not self.kb_enabled.get():
            self._release_all_keyboard_buttons()
            self.set_status("Keyboard Control: OFF")
            return

        # Turning on: if script running, disallow
        if self.engine.running:
            self.kb_enabled.set(False)
            messagebox.showwarning("Keyboard Control", "Stop the script before enabling keyboard control.")
            return

        if not self.serial.connected:
            # allow toggling on, but it won't do anything until connected
            self.set_status("Keyboard Control: ON (connect serial to use)")
        else:
            self.set_status("Keyboard Control: ON")

    def _release_all_keyboard_buttons(self):
        self.kb_down.clear()
        self.kb_buttons_held.clear()
        # go neutral only if script not running
        if not self.engine.running and self.serial.connected:
            self.serial.set_state(0, 0)

    def _on_key_press(self, event):
        if not self._manual_control_allowed():
            return

        ks = self._normalize_keysym(event)
        if not ks:
            return

        # Prevent repeat spamming (Tk sends repeats while held)
        if ks in self.kb_down:
            return
        self.kb_down.add(ks)

        btn = self.kb_bindings.get(ks)
        if not btn:
            return

        self.kb_buttons_held.add(btn)
        self._select_active_backend()
        if self.active_backend and getattr(self.active_backend, "connected", False):
            self.active_backend.set_buttons(sorted(self.kb_buttons_held))


    def _on_key_release(self, event):
        if not self._manual_control_allowed():
            return

        ks = self._normalize_keysym(event)
        if not ks:
            return

        if ks in self.kb_down:
            self.kb_down.remove(ks)

        btn = self.kb_bindings.get(ks)
        if not btn:
            return

        if btn in self.kb_buttons_held:
            self.kb_buttons_held.remove(btn)

        self._select_active_backend()
        if self.active_backend and getattr(self.active_backend, "connected", False):
            self.active_backend.set_buttons(sorted(self.kb_buttons_held))



    def _copy_to_clipboard(self, text: str):
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            # ensures clipboard persists after app closes on some systems
            self.root.update_idletasks()
        except Exception as e:
            self.set_status(f"Clipboard error: {e}")

    def _on_video_click_copy(self, event):
        xy = self._event_to_frame_xy(event) or self._last_video_xy
        if xy is None:
            self.set_status("No coords to copy.")
            return
        x, y = xy
        s = f"{x},{y}"
        self._copy_to_clipboard(s)
        self.set_status(f"Copied coords: {s}")

    def _on_video_click_copy_json(self, event):
        xy = self._event_to_frame_xy(event) or self._last_video_xy
        if xy is None:
            self.set_status("No coords to copy.")
            return
        x, y = xy
        s = json.dumps({"x": x, "y": y})
        self._copy_to_clipboard(s)
        self.set_status(f"Copied coords JSON: {s}")


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
            if self.threeds_backend and self.threeds_backend.connected:
                self.active_backend = self.threeds_backend
            else:
                # Not enabled yet; still point to a disconnected backend if it exists
                self.active_backend = self.threeds_backend or self.serial
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
        self.ctx.add_command(label="Add", command=self.add_command)
        self.ctx.add_command(label="Edit", command=self.edit_command)
        self.ctx.add_command(label="Delete", command=self.delete_command)
        self.ctx.add_separator()
        self.ctx.add_command(label="Move Up", command=lambda: self.move_command(-1))
        self.ctx.add_command(label="Move Down", command=lambda: self.move_command(1))
        self.ctx.add_separator()
        self.ctx.add_command(label="Add Comment", command=self.add_comment)
        self.ctx.add_separator()
        self.ctx.add_command(label="Save", command=self.save_script)
        self.ctx.add_command(label="Save As...", command=self.save_script_as)

    def _on_script_right_click(self, event):
        iid = self.script_tree.identify_row(event.y)
        if iid:
            self.script_tree.selection_set(iid)
        try:
            self.ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self.ctx.grab_release()

    def _on_script_double_click(self, event):
        # Only edit if a row is double-clicked (not empty area)
        iid = self.script_tree.identify_row(event.y)
        if iid:
            self.script_tree.selection_set(iid)
            self.edit_command()

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



    # ---- camera
    def refresh_cameras(self):
        cams = list_dshow_video_devices()
        self.cam_combo["values"] = cams
        if cams and self.cam_var.get() not in cams:
            self.cam_var.set(cams[0])

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
            self.cam_proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.DEVNULL, bufsize=10**7)
        except FileNotFoundError:
            messagebox.showerror("ffmpeg not found", "ffmpeg was not found on PATH.")
            return
        except Exception as e:
            messagebox.showerror("Camera error", str(e))
            return

        self.cam_running = True
        self.cam_toggle_btn.configure(text="Stop Cam")
        self.set_status(f"Camera streaming: {device}")

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
            x = self._saved_sash_x
            if x is None or x < 150:
                x = int(total * 0.45)  # default restore
            x = max(220, min(x, total - 220))  # keep both panes usable

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

    def apply_video_ratio(self):
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

        self.set_status(f"Video ratio set to {ratio} ({w}x{h}).")

        if was_running:
            self.start_camera()



    def _camera_reader_loop(self):
        if not self.cam_proc or not self.cam_proc.stdout:
            return
        frame_size = self.cam_width * self.cam_height * 3
        while self.cam_running and self.cam_proc and self.cam_proc.stdout:
            raw = self.cam_proc.stdout.read(frame_size)
            if not raw or len(raw) != frame_size:
                continue
            frame = np.frombuffer(raw, dtype=np.uint8).reshape((self.cam_height, self.cam_width, 3))
            with self.frame_lock:
                self.latest_frame_bgr = frame

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
        tk_img = ImageTk.PhotoImage(img)
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

        if input_names and self.audio_input_var.get() not in input_names:
            self.audio_input_var.set(input_names[0])
        if output_names and self.audio_output_var.get() not in output_names:
            self.audio_output_var.set(output_names[0])

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

            # Use common settings
            sample_rate = int(min(input_info.get('defaultSampleRate', 44100),
                                  output_info.get('defaultSampleRate', 44100)))
            channels = min(int(input_info.get('maxInputChannels', 2)),
                          int(output_info.get('maxOutputChannels', 2)))
            chunk = 1024

            # Open stream in callback mode for passthrough
            def audio_callback(in_data, frame_count, time_info, status):
                return (in_data, pyaudio.paContinue)

            self.audio_stream = self.audio_pyaudio.open(
                format=pyaudio.paInt16,
                channels=channels,
                rate=sample_rate,
                input=True,
                output=True,
                input_device_index=input_idx,
                output_device_index=output_idx,
                frames_per_buffer=chunk,
                stream_callback=audio_callback
            )

            self.audio_stream.start_stream()
            self.audio_running = True
            self.audio_toggle_btn.configure(text="Stop Audio")
            self.set_status(f"Audio streaming: {input_name} → {output_name}")

        except Exception as e:
            messagebox.showerror("Audio error", f"Failed to start audio:\n{e}")
            self.stop_audio()

    def stop_audio(self):
        self.audio_running = False
        self.audio_toggle_btn.configure(text="Start Audio")

        try:
            if self.audio_stream:
                self.audio_stream.stop_stream()
                self.audio_stream.close()
                self.audio_stream = None
        except Exception:
            pass

        try:
            if self.audio_pyaudio:
                self.audio_pyaudio.terminate()
                self.audio_pyaudio = None
        except Exception:
            pass

        self.set_status("Audio stopped.")

    # ---- serial
    def refresh_ports(self):
        ports = list_com_ports()
        self.com_combo["values"] = ports
        if ports and self.com_var.get() not in ports:
            self.com_var.set(ports[0])

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

    def open_keybinds_window(self):
        win = tk.Toplevel(self.root)
        win.title("Keybinds")
        win.transient(self.root)
        win.grab_set()

        # info
        info = ttk.Label(win, text="Keyboard Control is active only when no script is running.\n"
                                "Click Rebind, then press a key. Keys are shown as Tk keysyms.")
        info.grid(row=0, column=0, columnspan=4, sticky="w", padx=10, pady=(10, 6))

        # tree
        tree = ttk.Treeview(win, columns=("button", "key"), show="headings", height=12)
        tree.heading("button", text="Controller Button")
        tree.heading("key", text="Key")
        tree.column("button", width=160, anchor="w")
        tree.column("key", width=120, anchor="w")
        tree.grid(row=1, column=0, columnspan=4, sticky="nsew", padx=10)

        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

        def refresh():
            tree.delete(*tree.get_children())
            # invert mapping: button -> keys (allow multiple keys if desired)
            inv = {b: [] for b in SerialController.ALL_BUTTONS}
            for k, b in self.kb_bindings.items():
                if b in inv:
                    inv[b].append(k)
            for b in SerialController.ALL_BUTTONS:
                keys = ", ".join(sorted(inv[b])) if inv[b] else ""
                tree.insert("", "end", values=(b, keys))

        refresh()

        def get_selected_button():
            sel = tree.selection()
            if not sel:
                return None
            vals = tree.item(sel[0], "values")
            return vals[0] if vals else None

        def rebind():
            b = get_selected_button()
            if not b:
                messagebox.showinfo("Rebind", "Select a controller button first.", parent=win)
                return
            self._rebinding_target = b
            status_var.set(f"Press a key to bind to {b} (Esc cancels)…")

        def clear_binding():
            b = get_selected_button()
            if not b:
                return
            # remove all keys mapping to this button
            to_del = [k for k, btn in self.kb_bindings.items() if btn == b]
            for k in to_del:
                del self.kb_bindings[k]
            refresh()

        def restore_defaults():
            self.kb_bindings = {
                "w": "Up", "a": "Left", "s": "Down", "d": "Right",
                "j": "A", "k": "B", "u": "X", "i": "Y",
                "enter": "Start", "space": "Select",
                "q": "L", "e": "R",
            }
            refresh()

        status_var = tk.StringVar(value="Select a button and click Rebind.")
        ttk.Label(win, textvariable=status_var, foreground="gray").grid(row=2, column=0, columnspan=4, sticky="w", padx=10, pady=(6, 0))

        btnrow = ttk.Frame(win)
        btnrow.grid(row=3, column=0, columnspan=4, sticky="ew", padx=10, pady=10)

        ttk.Button(btnrow, text="Rebind…", command=rebind).pack(side="left")
        ttk.Button(btnrow, text="Clear", command=clear_binding).pack(side="left", padx=(6, 0))
        ttk.Button(btnrow, text="Restore defaults", command=restore_defaults).pack(side="left", padx=(6, 0))
        ttk.Button(btnrow, text="Close", command=win.destroy).pack(side="right")

        # Capture key presses while rebinding
        def on_key(event):
            if self._rebinding_target is None:
                return
            ks = (event.keysym or "").lower()
            if ks == "escape":
                self._rebinding_target = None
                status_var.set("Rebind cancelled.")
                return

            if ks == "return":
                ks = "enter"

            # Ensure uniqueness: remove this key if already bound
            self.kb_bindings[ks] = self._rebinding_target
            status_var.set(f"Bound {ks} -> {self._rebinding_target}")
            self._rebinding_target = None
            refresh()

        win.bind("<KeyPress>", on_key)

    def on_backend_changed(self):
        is_3ds = (self.backend_var.get() == "3DS Input Redirection")

        # Show/hide 3DS widgets
        widgets = [
            self.threeds_ip_label, self.threeds_ip_entry,
            self.threeds_port_label, self.threeds_port_entry,
            self.threeds_enable_btn, self.threeds_disable_btn
        ]
        for w in widgets:
            if is_3ds:
                w.grid()  # show
            else:
                w.grid_remove()  # hide

        # Optional: disable serial-specific UI when using 3DS
        # If you have serial COM widgets like self.com_combo, self.connect_btn, disable them:
        if hasattr(self, "com_combo"):
            try:
                state = "disabled" if is_3ds else "readonly"
                self.com_combo.configure(state=state)
            except Exception:
                pass
        if hasattr(self, "connect_btn"):
            try:
                self.connect_btn.configure(state=("disabled" if is_3ds else "normal"))
            except Exception:
                pass

        # Switch active backend pointer
        self._select_active_backend()

        self.set_status(f"Output backend set to: {self.backend_var.get()}")

    def enable_threeds_backend(self):
        if self.engine.running:
            messagebox.showwarning("3DS", "Stop the script before enabling/changing 3DS backend.")
            return

        ip = (self.threeds_ip_var.get() or "").strip()
        if not ip:
            messagebox.showerror("3DS", "Please enter a 3DS IP address.")
            return
        try:
            port = int((self.threeds_port_var.get() or "4950").strip())
        except ValueError:
            messagebox.showerror("3DS", "Port must be a number.")
            return

        try:
            self.threeds_backend = ThreeDSClasses.ThreeDSBackend(ip=ip, port=port)
            self.threeds_backend.connect()
            self.backend_var.set("3DS Input Redirection")
            self._select_active_backend()
            self.set_status(f"3DS backend enabled: {ip}:{port}")
        except Exception as e:
            messagebox.showerror("3DS", str(e))

    def disable_threeds_backend(self):
        try:
            if self.threeds_backend:
                self.threeds_backend.disconnect()
        except Exception:
            pass
        self.threeds_backend = None
        self.backend_var.set("USB Serial")
        self._select_active_backend()
        self.set_status("3DS backend disabled.")
        self.on_backend_changed()

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
    def populate_script_view(self):
        self.script_tree.delete(*self.script_tree.get_children())

        indent_on = bool(self.indent_var.get()) if hasattr(self, "indent_var") else True
        depth = 0

        for i, c in enumerate(self.engine.commands):
            cmd = c.get("cmd")
            spec = self.engine.registry.get(cmd)
            pretty = spec.format_fn(c) if spec else f"(unknown) {cmd}"

            # Decrease indent BEFORE printing for closing blocks
            if indent_on and cmd in ("end_if", "end_while"):
                depth = max(0, depth - 1)

            if indent_on:
                pretty = ("      " * depth) + pretty  # 6 spaces per level

            # Insert row
            self.script_tree.insert("", "end", iid=str(i), values=(i, pretty))

            # Increase indent AFTER printing for opening blocks
            if indent_on and cmd in ("if", "while"):
                depth += 1

        self.highlight_ip(-1)

    def populate_insert_panel(self):
        """
        Builds a grouped command list with group headers as parent nodes.
        Filters by search text across name/group/doc.
        """
        if not hasattr(self, "insert_tree"):
            return

        search = (self.cmd_search_var.get() or "").strip().lower()

        self.insert_tree.delete(*self.insert_tree.get_children())

        # group -> parent iid
        group_nodes = {}

        for name, spec in self.engine.ordered_specs():
            hay = f"{spec.group} {name} {spec.doc}".lower()
            if search and search not in hay:
                continue

            if spec.group not in group_nodes:
                gid = self.insert_tree.insert("", "end", text=spec.group, values=("", ""), open=True)
                group_nodes[spec.group] = gid

            parent = group_nodes[spec.group]
            # store command name in iid so we can retrieve easily
            iid = f"cmd::{name}"
            self.insert_tree.insert(parent, "end", iid=iid, text="", values=(name, (spec.doc or "")))

        # If there's only one group and one command visible, select it
        all_cmd_items = []
        for g in self.insert_tree.get_children(""):
            for child in self.insert_tree.get_children(g):
                all_cmd_items.append(child)
        if all_cmd_items:
            self.insert_tree.selection_set(all_cmd_items[0])
            self.insert_tree.see(all_cmd_items[0])

    def insert_selected_command(self):
        self._insert_selected_command(edit_after=False)

    def insert_selected_command_and_edit(self):
        self._insert_selected_command(edit_after=True)

    def _insert_selected_command(self, edit_after: bool):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        cmd_name = self._get_selected_insert_command()
        if not cmd_name:
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)

        # Build default command object based on schema defaults
        spec = self.engine.registry[cmd_name]
        cmd_obj = {"cmd": cmd_name}
        for f in spec.arg_schema:
            k = f["key"]
            if "default" in f:
                cmd_obj[k] = f["default"]

        # Insert command
        self.engine.commands.insert(insert_at, cmd_obj)

        # Auto-insert end markers for blocks
        if self.auto_close_block_var.get():
            if cmd_name == "if":
                self.engine.commands.insert(insert_at + 1, {"cmd": "end_if"})
            elif cmd_name == "while":
                self.engine.commands.insert(insert_at + 1, {"cmd": "end_while"})

        # Tolerant rebuild during editing
        try:
            self.engine.rebuild_indexes(strict=False)
        except Exception as e:
            self.set_status(f"Index warning: {e}")

        self.populate_script_view()
        self.mark_dirty(True)

        # Select inserted row and optionally open editor
        self.script_tree.selection_set(str(insert_at))
        self.script_tree.see(str(insert_at))

        if edit_after:
            # If we auto-inserted end_if/end_while, edit the opening line
            self.edit_command()


    def _get_selected_insert_command(self):
        sel = self.insert_tree.selection()
        if not sel:
            return None
        iid = sel[0]
        if not iid.startswith("cmd::"):
            # group header selected, try first child
            kids = self.insert_tree.get_children(iid)
            if kids:
                iid = kids[0]
            else:
                return None
        return iid.split("cmd::", 1)[1]



    def refresh_vars_view(self):
        self.vars_tree.delete(*self.vars_tree.get_children())
        for k, v in sorted(self.engine.vars.items(), key=lambda kv: kv[0]):
            self.vars_tree.insert("", "end", values=(k, json.dumps(v, ensure_ascii=False)))

    def run_script(self):
        if self.kb_enabled.get():
            # Turn off manual so we don't fight the script
            self.kb_enabled.set(False)
            self._release_all_keyboard_buttons()
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
        for item in self.script_tree.get_children():
            self.script_tree.item(item, tags=())
        if ip is None or ip < 0:
            return
        iid = str(ip)
        if self.script_tree.exists(iid):
            self.script_tree.item(iid, tags=("ip",))
            self.script_tree.see(iid)

    # ---- editor actions
    def _get_selected_index(self):
        sel = self.script_tree.selection()
        if not sel:
            return None
        try:
            return int(sel[0])
        except Exception:
            return None

    def _reindex_after_edit(self):
        try:
            self.engine.rebuild_indexes(strict=False)  # tolerant during editing
        except Exception as e:
            # This should be rare now; but don't crash UI
            self.set_status(f"Index warning: {e}")
        self.populate_script_view()
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



    def add_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)

        dlg = CommandEditorDialog(self.root, self.engine.registry, initial_cmd=None, title="Add Command", test_callback=self._dialog_test_callback)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return

        self.engine.commands.insert(insert_at, dlg.result)
        if dlg.result["cmd"] == "if":
            self.engine.commands.insert(insert_at + 1, {"cmd": "end_if"})
        elif dlg.result["cmd"] == "while":
            self.engine.commands.insert(insert_at + 1, {"cmd": "end_while"})

        self._reindex_after_edit()

        self.script_tree.selection_set(str(insert_at))
        self.script_tree.see(str(insert_at))

    def edit_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        if idx is None:
            messagebox.showinfo("Edit", "Select a command to edit.")
            return

        initial = self.engine.commands[idx]
        dlg = CommandEditorDialog(self.root, self.engine.registry, initial_cmd=initial, title="Edit Command", test_callback=self._dialog_test_callback)
        self.root.wait_window(dlg)
        if dlg.result is None:
            return

        self.engine.commands[idx] = dlg.result
        self._reindex_after_edit()
        self.script_tree.selection_set(str(idx))
        self.script_tree.see(str(idx))

    def delete_command(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        if idx is None:
            return

        if not messagebox.askyesno("Delete", "Delete selected command?"):
            return

        del self.engine.commands[idx]
        self._reindex_after_edit()

        new_idx = min(idx, len(self.engine.commands) - 1)
        if new_idx >= 0:
            self.script_tree.selection_set(str(new_idx))
            self.script_tree.see(str(new_idx))

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

        self.script_tree.selection_set(str(j))
        self.script_tree.see(str(j))

    def add_comment(self):
        if self.engine.running:
            messagebox.showwarning("Running", "Stop the script before editing.")
            return

        idx = self._get_selected_index()
        insert_at = (idx + 1) if idx is not None else len(self.engine.commands)

        self.engine.commands.insert(insert_at, {"cmd": "comment", "text": "New comment"})
        self._reindex_after_edit()

        self.script_tree.selection_set(str(insert_at))
        self.script_tree.see(str(insert_at))

    def _resolve_test_value(self, v):
        """
        Resolve '$var' in editor tests using current engine vars.
        """
        if isinstance(v, str) and v.startswith("$"):
            return self.engine.vars.get(v[1:], None)
        return v
    
    def test_command_dialog(self, cmd_obj):
        """
        Returns (title, message) for a given cmd_obj.
        Currently supports: find_color
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
                tol = int(self._resolve_test_value(cmd_obj.get("tol", 0)))
                out = (cmd_obj.get("out") or "match").strip()

                h, w, _ = frame.shape
                if not (0 <= x < w and 0 <= y < h):
                    return ("find_color Test",
                            f"Point out of bounds.\n"
                            f"Requested: ({x},{y})\n"
                            f"Frame size: {w}x{h}")

                # Sample pixel (frame is BGR)
                b, g, r = frame[y, x].tolist()
                sampled_rgb = [int(r), int(g), int(b)]

                target = [int(rgb[0]), int(rgb[1]), int(rgb[2])]
                tol = max(0, tol)

                ok = all(abs(sampled_rgb[i] - target[i]) <= tol for i in range(3))

                msg = (
                    f"Point: ({x},{y})\n"
                    f"Sampled RGB: {sampled_rgb}\n"
                    f"Target RGB:  {target}\n"
                    f"Tolerance:   {tol}\n\n"
                    f"Result (would set ${out}): {ok}"
                )
                return ("find_color Test", msg)
            case _:
                raise ValueError("No tester implemented for this command.")

    def _dialog_test_callback(self, cmd_obj):
        # Only enable for find_color (for now)
        cmd = cmd_obj.get("cmd")
        match cmd:
            case "find_color":
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

        self.root.destroy()


if __name__ == "__main__":
    os.makedirs("scripts", exist_ok=True)
    os.makedirs("py_scripts", exist_ok=True)
    root = tk.Tk()
    app = App(root)
    root.mainloop()
