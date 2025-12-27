"""
Controller Macro Runner (Camera + Serial + Script Engine + Editor)

Install:
  pip install numpy pillow pyserial

FFmpeg:
  Ensure ffmpeg is on PATH:
    ffmpeg -version
"""

import os
import json
import time
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import re

import numpy as np
from PIL import Image, ImageTk

import serial
from serial.tools import list_ports
import sys
import base64


def resource_path(rel_path: str) -> str:
    """
    Works for:
    - normal python run (rel path from current dir)
    - PyInstaller onefile/onedir (MEIPASS temp dir)
    """
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)

def ffmpeg_path() -> str:
    # Prefer bundled ffmpeg.exe
    bundled = resource_path("ffmpeg.exe")
    if os.path.exists(bundled):
        return bundled
    return "ffmpeg"  # fallback to PATH

def frame_to_json_payload(frame_bgr: np.ndarray):
    """
    Convert BGR frame (H,W,3 uint8) to a JSON-serializable payload (PNG base64).
    """
    if frame_bgr is None:
        return None
    # Convert to RGB for PIL
    rgb = frame_bgr[:, :, ::-1]
    img = Image.fromarray(rgb)
    import io
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return {"__frame__": "png_base64", "data_b64": b64}


def run_python_main(script_path, args, timeout_s=10):
    """
    Runs a python file in a subprocess and calls main(*args).
    Returns the Python object that main returns (decoded from JSON).
    """
    runner = r"""
import json, sys, importlib.util, traceback

def load_module_from_path(path):
    spec = importlib.util.spec_from_file_location("user_module", path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Could not load module from path: " + path)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod

def main():
    path = sys.argv[1]
    args_json = sys.argv[2] if len(sys.argv) > 2 else "[]"
    args = json.loads(args_json)

    mod = load_module_from_path(path)
    if not hasattr(mod, "main"):
        raise RuntimeError("Script does not define main(...).")

    res = mod.main(*args)
    # Always print JSON so the host can decode it reliably
    print(json.dumps(res, ensure_ascii=False))

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        traceback.print_exc()
        sys.exit(1)
"""
    # args must be JSON-serializable
    args_json = json.dumps(args, ensure_ascii=False)

    cp = subprocess.run(
        [sys.executable, "-c", runner, script_path, args_json],
        capture_output=True,
        text=True,
        timeout=timeout_s
    )

    if cp.returncode != 0:
        raise RuntimeError(f"run_python failed:\n{cp.stderr.strip() or cp.stdout.strip()}")

    out_text = (cp.stdout or "").strip()
    if not out_text:
        return None
    return json.loads(out_text)



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
# Serial controller
# ----------------------------

BUTTON_MAP = {
    "L": ("high", 0x01),
    "R": ("high", 0x02),
    "X": ("high", 0x04),
    "Y": ("high", 0x08),
    "A": ("low", 0x01),
    "B": ("low", 0x02),
    "Right": ("low", 0x04),
    "Left": ("low", 0x08),
    "Up": ("low", 0x10),
    "Down": ("low", 0x20),
    "Select": ("low", 0x40),
    "Start": ("low", 0x80),
}

ALL_BUTTONS = ["A", "B", "X", "Y", "Up", "Down", "Left", "Right", "Start", "Select", "L", "R"]


def buttons_to_bytes(buttons):
    high, low = 0, 0
    for b in buttons:
        if b not in BUTTON_MAP:
            raise ValueError(f"Unknown button: {b}")
        which, mask = BUTTON_MAP[b]
        if which == "high":
            high |= mask
        else:
            low |= mask
    return high & 0xFF, low & 0xFF


class SerialController:
    def __init__(self, status_cb=None):
        self.status_cb = status_cb or (lambda s: None)
        self.ser = None
        self.interval_s = 0.05

        self._lock = threading.Lock()
        self._running = False
        self._thread = None

        self._high = 0
        self._low = 0

    @property
    def connected(self):
        return self.ser is not None and self.ser.is_open

    def set_state(self, high, low):
        with self._lock:
            self._high = high & 0xFF
            self._low = low & 0xFF

    def set_buttons(self, buttons):
        high, low = buttons_to_bytes(buttons)
        self.set_state(high, low)

    def connect(self, port, baud=1_000_000):
        if self.connected:
            self.disconnect()

        self.ser = serial.Serial(port, baud, timeout=1)
        self.status_cb(f"Serial connected: {port} @ {baud}")

        self._running = True
        self._thread = threading.Thread(target=self._keepalive_loop, daemon=True)
        self._thread.start()

        # Pairing warm-up (neutral for ~3s)
        self.status_cb("Pairing warm-up: neutral for ~3 seconds...")
        self.set_state(0, 0)
        time.sleep(3.0)
        self.status_cb("Pairing warm-up done.")

    def disconnect(self):
        self._running = False
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)

        if self.ser:
            try:
                self.ser.close()
            except Exception:
                pass
        self.ser = None
        self.status_cb("Serial disconnected.")

    def send_channel_set(self, channel_byte):
        if not self.connected:
            raise RuntimeError("Not connected.")
        ch = int(channel_byte) & 0xFF
        pkt = bytearray([0x43, ch, 0x00])
        self.ser.write(pkt)
        self.ser.flush()
        self.status_cb(f"Sent channel set: 0x{ch:02X} (power cycle receiver required)")

    def _keepalive_loop(self):
        while self._running:
            if not self.connected:
                break
            with self._lock:
                high = self._high
                low = self._low
            try:
                self.ser.write(bytearray([0x54, high, low]))
            except Exception as e:
                self.status_cb(f"Serial write error: {e}")
                break
            time.sleep(self.interval_s)


# ----------------------------
# Script Command Spec
# ----------------------------

class CommandSpec:
    def __init__(self, name, required_keys, fn, doc="", arg_schema=None, format_fn=None,
                 group="Other", order=999):
        self.name = name
        self.required_keys = required_keys
        self.fn = fn
        self.doc = doc
        self.arg_schema = arg_schema or []
        self.format_fn = format_fn or (lambda c: f"{name} " + " ".join(f"{k}={c.get(k)!r}" for k in c if k != "cmd"))
        self.group = group
        self.order = order



# ----------------------------
# Script Engine helpers
# ----------------------------

def resolve_value(ctx, v):
    if isinstance(v, str) and v.startswith("$"):
        return ctx["vars"].get(v[1:], None)
    return v

def resolve_vars_deep(ctx, obj):
    """
    Recursively resolve $var strings inside lists/dicts/strings.
    Special tokens:
      "$frame" -> JSON payload for the latest camera frame (PNG base64)
      "$frame_bgr" -> NOT supported across subprocess (keep for future in-proc commands)
    """
    if isinstance(obj, str):
        if obj == "$frame":
            frame = ctx["get_frame"]()
            return frame_to_json_payload(frame)
        return resolve_value(ctx, obj)

    if isinstance(obj, list):
        return [resolve_vars_deep(ctx, x) for x in obj]

    if isinstance(obj, dict):
        return {k: resolve_vars_deep(ctx, v) for k, v in obj.items()}

    return obj




def eval_condition(ctx, left, op, right):
    L = resolve_value(ctx, left)
    R = resolve_value(ctx, right)
    if op == "==": return L == R
    if op == "!=": return L != R
    if op == "<":  return L < R
    if op == "<=": return L <= R
    if op == ">":  return L > R
    if op == ">=": return L >= R
    raise ValueError(f"Unknown op: {op}")


def build_label_index(commands):
    labels = {}
    for i, c in enumerate(commands):
        if c.get("cmd") == "label":
            name = c.get("name")
            if not name:
                raise ValueError(f"label missing name at index {i}")
            labels[name] = i
    return labels


def build_if_matching(commands, strict=True):
    stack = []
    m = {}
    for i, c in enumerate(commands):
        if c.get("cmd") == "if":
            stack.append(i)
        elif c.get("cmd") == "end_if":
            if not stack:
                if strict:
                    raise ValueError(f"end_if without if at index {i}")
                continue
            j = stack.pop()
            m[j] = i
    if stack and strict:
        raise ValueError(f"Unclosed if at index {stack[-1]}")
    return m, stack  # return leftovers for warnings


def build_while_matching(commands, strict=True):
    stack = []
    while_to_end = {}
    end_to_while = {}
    for i, c in enumerate(commands):
        if c.get("cmd") == "while":
            stack.append(i)
        elif c.get("cmd") == "end_while":
            if not stack:
                if strict:
                    raise ValueError(f"end_while without while at index {i}")
                continue
            w = stack.pop()
            while_to_end[w] = i
            end_to_while[i] = w
    if stack and strict:
        raise ValueError(f"Unclosed while at index {stack[-1]}")
    return while_to_end, end_to_while, stack



# ----------------------------
# Script Engine
# ----------------------------

class ScriptEngine:
    def __init__(self, serial_ctrl, get_frame_fn, status_cb=None, on_ip_update=None, on_tick=None):
        self.serial = serial_ctrl
        self.get_frame = get_frame_fn
        self.status_cb = status_cb or (lambda s: None)
        self.on_ip_update = on_ip_update or (lambda ip: None)
        self.on_tick = on_tick or (lambda: None)

        self.vars = {}
        self.commands = []
        self.labels = {}
        self.if_map = {}
        self.while_to_end = {}
        self.end_to_while = {}
        self._unclosed_ifs = []
        self._unclosed_whiles = []


        self.registry = self._build_default_registry()

        self._stop = threading.Event()
        self._thread = None
        self.running = False
        self.ip = 0

    def ordered_specs(self):
        """
        Returns list of (name, spec) sorted by (group, order, name).
        """
        items = list(self.registry.items())
        items.sort(key=lambda kv: (kv[1].group, kv[1].order, kv[0]))
        return items


    def rebuild_indexes(self, strict=True):
        self.labels = build_label_index(self.commands)
        self.if_map, self._unclosed_ifs = build_if_matching(self.commands, strict=strict)
        self.while_to_end, self.end_to_while, self._unclosed_whiles = build_while_matching(self.commands, strict=strict)


    def list_available_commands(self):
        return sorted(self.registry.keys())

    def load_script(self, path):
        with open(path, "r", encoding="utf-8") as f:
            cmds = json.load(f)
        if not isinstance(cmds, list):
            raise ValueError("Script must be a list of command objects.")

        for i, c in enumerate(cmds):
            if not isinstance(c, dict) or "cmd" not in c:
                raise ValueError(f"Command at index {i} must be an object with 'cmd'.")
            name = c["cmd"]
            if name not in self.registry:
                raise ValueError(f"Unknown cmd '{name}' at index {i}.")
            spec = self.registry[name]
            for k in spec.required_keys:
                if k not in c:
                    raise ValueError(f"'{name}' missing required key '{k}' at index {i}.")

        self.commands = cmds
        self.rebuild_indexes()

        self.vars = {}
        self.ip = 0
        self.status_cb(f"Loaded script: {os.path.basename(path)} ({len(cmds)} commands)")

    def stop(self):
        self._stop.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        self._stop.clear()
        self.running = False
        self.serial.set_state(0, 0)
        self.status_cb("Script stopped.")
        self.on_ip_update(-1)

    def run(self):
        if not self.serial.connected:
            raise RuntimeError("Serial not connected.")
        if not self.commands:
            raise RuntimeError("No script loaded.")
        if self.running:
            return

        self.running = True
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _loop(self):
        ctx = {
            "vars": self.vars,
            "labels": self.labels,
            "if_map": self.if_map,
            "while_to_end": self.while_to_end,
            "end_to_while": self.end_to_while,
            "stop": self._stop,
            "get_frame": self.get_frame,
            "ip": self.ip,
        }

        try:
            while not self._stop.is_set() and 0 <= self.ip < len(self.commands):
                self.on_ip_update(self.ip)

                c = self.commands[self.ip]
                name = c["cmd"]
                spec = self.registry[name]

                ctx["ip"] = self.ip
                spec.fn(ctx, c)
                self.ip = ctx["ip"]

                self.on_tick()
                self.ip += 1

            self.serial.set_state(0, 0)
            if not self._stop.is_set():
                self.status_cb("Script completed.")
        except Exception as e:
            self.serial.set_state(0, 0)
            self.status_cb(f"Script error: {e}")
        finally:
            self.running = False
            self.on_ip_update(-1)

    def _build_default_registry(self):
        reg = {}

        # ---- pretty formatters
        def fmt_wait(c): return f"Wait {c.get('ms')} ms"
        def fmt_press(c):
            btns = c.get("buttons", [])
            return f"Press {', '.join(btns) if btns else '(none)'} for {c.get('ms')} ms"
        def fmt_hold(c):
            btns = c.get("buttons", [])
            return f"Hold {', '.join(btns) if btns else '(none)'}"
        def fmt_label(c): return f"Label: {c.get('name')}"
        def fmt_goto(c): return f"Goto '{c.get('label')}'"
        def fmt_set(c): return f"Set ${c.get('var')} = {c.get('value')}"
        def fmt_add(c): return f"Add {c.get('value')} to ${c.get('var')}"
        def fmt_if(c): return f"If {c.get('left')} {c.get('op')} {c.get('right')}"
        def fmt_end_if(c): return "End If"
        def fmt_while(c): return f"While {c.get('left')} {c.get('op')} {c.get('right')}"
        def fmt_end_while(c): return "End While"
        def fmt_find_color(c):
            return f"FindColor ({c.get('x')},{c.get('y')}) ~ {c.get('rgb')} tol={c.get('tol',0)} -> ${c.get('out')}"
        def fmt_comment(c): return f"// {c.get('text','')}"
        def fmt_run_python(c):
            out = c.get("out")
            args = c.get("args", [])
            a = "" if not args else f" args={args}"
            o = "" if not out else f" -> ${out}"
            return f"RunPython {c.get('file')}{a}{o}"


        # ---- execution fns
        def cmd_wait(ctx, c):
            ms = int(resolve_value(ctx, c["ms"]))
            end_t = time.monotonic() + ms / 1000.0
            while time.monotonic() < end_t:
                if ctx["stop"].is_set():
                    break
                time.sleep(0.01)

        def cmd_press(ctx, c):
            buttons = c.get("buttons", [])
            ms = int(c.get("ms", 50))
            self.serial.set_buttons(buttons)
            cmd_wait(ctx, {"ms": ms})
            self.serial.set_state(0, 0)

        def cmd_hold(ctx, c):
            self.serial.set_buttons(c.get("buttons", []))

        def cmd_label(ctx, c):
            pass

        def cmd_goto(ctx, c):
            label = c["label"]
            if label not in ctx["labels"]:
                raise ValueError(f"Unknown label: {label}")
            ctx["ip"] = ctx["labels"][label]

        def cmd_set(ctx, c):
            ctx["vars"][c["var"]] = resolve_value(ctx, c.get("value"))

        def cmd_add(ctx, c):
            var = c["var"]
            cur = ctx["vars"].get(var, 0)
            ctx["vars"][var] = cur + resolve_value(ctx, c.get("value", 0))

        def cmd_if(ctx, c):
            ok = eval_condition(ctx, c["left"], c["op"], c["right"])
            if not ok:
                end_idx = ctx["if_map"][ctx["ip"]]
                ctx["ip"] = end_idx  # loop will +1 -> after end_if

        def cmd_end_if(ctx, c):
            pass

        def cmd_while(ctx, c):
            ok = eval_condition(ctx, c["left"], c["op"], c["right"])
            if not ok:
                end_idx = ctx["while_to_end"][ctx["ip"]]
                ctx["ip"] = end_idx  # +1 -> after end_while

        def cmd_end_while(ctx, c):
            w = ctx["end_to_while"][ctx["ip"]]
            ctx["ip"] = w - 1  # +1 -> while line to re-evaluate

        def cmd_find_color(ctx, c):
            frame = ctx["get_frame"]()
            out = c.get("out", "match")

            if frame is None:
                ctx["vars"][out] = False
                return

            x = int(resolve_value(ctx, c["x"]))
            y = int(resolve_value(ctx, c["y"]))
            h, w, _ = frame.shape
            if not (0 <= x < w and 0 <= y < h):
                ctx["vars"][out] = False
                return

            b, g, r = frame[y, x].tolist()
            sample_rgb = [int(r), int(g), int(b)]
            target = c["rgb"]
            tol = int(c.get("tol", 0))

            ok = all(abs(sample_rgb[i] - int(target[i])) <= tol for i in range(3))
            ctx["vars"][out] = ok

        def cmd_comment(ctx, c):
            pass

        def cmd_run_python(ctx, c):
            file_name = str(c["file"]).strip()
            if not file_name:
                raise ValueError("run_python: file is empty")

            if os.path.isabs(file_name):
                script_path = file_name
            else:
                script_path = os.path.join("py_scripts", file_name)

            if not os.path.exists(script_path):
                raise ValueError(f"run_python: file not found: {script_path}")

            args = c.get("args", [])
            if args is None:
                args = []
            if not isinstance(args, list):
                raise ValueError("run_python: args must be a JSON list")

            # NEW: allow $var references (and nested structures)
            args = resolve_vars_deep(ctx, args)

            timeout_s = int(c.get("timeout_s", 10))
            res = run_python_main(script_path, args, timeout_s=timeout_s)

            outvar = (c.get("out") or "").strip()
            if outvar:
                ctx["vars"][outvar] = res



        cond_schema = [
            {"key": "left", "type": "str", "default": "$flag", "help": "Left operand (literal or $var)"},
            {"key": "op", "type": "choice", "choices": ["==", "!=", "<", "<=", ">", ">="], "default": "==", "help": "Comparison operator"},
            {"key": "right", "type": "json", "default": True, "help": "Right operand (literal or $var)"},
        ]

        specs = [
            CommandSpec(
                "comment", ["text"], cmd_comment,
                doc="Comment line. Does nothing at runtime.",
                arg_schema=[{"key": "text", "type": "str", "default": "", "help": "Comment text"}],
                format_fn=fmt_comment,
                group="Meta",
                order=10
            ),
            CommandSpec(
                "wait", ["ms"], cmd_wait,
                doc="Wait for a number of milliseconds.",
                arg_schema=[{"key": "ms", "type": "int", "default": 100, "help": "Milliseconds to wait"}],
                format_fn=fmt_wait,
                group="Timing",
                order=10
            ),
            CommandSpec(
                "press", ["buttons", "ms"], cmd_press,
                doc="Press buttons for ms, then release to neutral.",
                arg_schema=[
                    {"key": "buttons", "type": "buttons", "default": ["A"], "help": "Buttons to press"},
                    {"key": "ms", "type": "int", "default": 80, "help": "Hold duration in milliseconds"},
                ],
                format_fn=fmt_press,
                group="Controller",
                order=10
            ),
            CommandSpec(
                "hold", ["buttons"], cmd_hold,
                doc="Hold buttons until changed by another command.",
                arg_schema=[{"key": "buttons", "type": "buttons", "default": [], "help": "Buttons to hold"}],
                format_fn=fmt_hold,
                group="Controller",
                order=20
            ),
            CommandSpec(
                "label", ["name"], cmd_label,
                doc="Define a label for goto.",
                arg_schema=[{"key": "name", "type": "str", "default": "start", "help": "Label name"}],
                format_fn=fmt_label,
                group="Control Flow",
                order=50
            ),
            CommandSpec(
                "goto", ["label"], cmd_goto,
                doc="Jump to a label.",
                arg_schema=[{"key": "label", "type": "str", "default": "start", "help": "Label name to jump to"}],
                format_fn=fmt_goto,
                group="Control Flow",
                order=60
            ),
            CommandSpec(
                "set", ["var", "value"], cmd_set,
                doc="Set a variable. Use $var in other commands.",
                arg_schema=[
                    {"key": "var", "type": "str", "default": "flag", "help": "Variable name (without $)"},
                    {"key": "value", "type": "json", "default": 0, "help": "Any JSON value (number/string/bool/list/object)"},
                ],
                format_fn=fmt_set,
                group="Variables",
                order=10
            ),
            CommandSpec(
                "add", ["var", "value"], cmd_add,
                doc="Add a numeric value to a variable.",
                arg_schema=[
                    {"key": "var", "type": "str", "default": "counter", "help": "Variable name (without $)"},
                    {"key": "value", "type": "int", "default": 1, "help": "Amount to add"},
                ],
                format_fn=fmt_add,
                group="Variables",
                order=20
            ),
            CommandSpec(
                "if", ["left", "op", "right"], cmd_if,
                doc="Conditional block. If condition is false, skip to matching end_if.",
                arg_schema=cond_schema,
                format_fn=fmt_if,
                group="Control Flow",
                order=10
            ),
            CommandSpec(
                "end_if", [], cmd_end_if,
                doc="Ends an if block.",
                arg_schema=[],
                format_fn=fmt_end_if,
                group="Control Flow",
                order=20
            ),
            CommandSpec(
                "while", ["left", "op", "right"], cmd_while,
                doc="Loop block. While condition is true, run the block. Re-evaluated at 'while' line.",
                arg_schema=cond_schema,
                format_fn=fmt_while,
                group="Control Flow",
                order=30
            ),
            CommandSpec(
                "end_while", [], cmd_end_while,
                doc="Ends a while block (jumps back to its while).",
                arg_schema=[],
                format_fn=fmt_end_while,
                group="Control Flow",
                order=40
            ),
            CommandSpec(
                "find_color", ["x", "y", "rgb", "out"], cmd_find_color,
                doc="Sample pixel at (x,y) from latest camera frame and compare to rgb within tolerance. Stores bool in $out.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate"},
                    {"key": "rgb", "type": "rgb", "default": [255, 0, 0], "help": "Target RGB as [R,G,B]"},
                    {"key": "tol", "type": "int", "default": 20, "help": "Tolerance per channel"},
                    {"key": "out", "type": "str", "default": "match", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_find_color,
                group="Image",
                order=10
            ),
            CommandSpec(
                "run_python",
                ["file"],
                cmd_run_python,
                doc="Run a Python file from ./py_scripts (or absolute path). Calls main(*args). "
                    "If 'out' is set, stores main's return value into $out. Return value must be JSON-serializable.",
                arg_schema=[
                    {"key":"file","type":"pyfile","default":"myscript.py", "help":"Filename in ./py_scripts (or absolute path)"},
                    {"key":"args","type":"json","default":[], "help":"JSON list of arguments passed to main(*args)"},
                    {"key":"out","type":"str","default":"", "help":"Variable name to store return value (no $). Leave blank to ignore."},
                    {"key":"timeout_s","type":"int","default":10, "help":"Timeout in seconds"},
                ],
                format_fn=fmt_run_python,
                group="Custom",
                order=10
            ),

        ]

        for s in specs:
            reg[s.name] = s
        return reg


# ----------------------------
# Editor Dialog (schema-driven)
# ----------------------------

class CommandEditorDialog(tk.Toplevel):
    def __init__(self, parent, registry, initial_cmd=None, title="Edit Command"):
        super().__init__(parent)
        self.parent = parent
        self.registry = registry
        self.result = None

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

            if ftype == "int":
                var = tk.StringVar(value=str(init_val))
                ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                ent.grid(row=r, column=1, sticky="ew", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = ent

            elif ftype == "str":
                var = tk.StringVar(value=str(init_val))
                ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                ent.grid(row=r, column=1, sticky="ew", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = ent

            elif ftype == "bool":
                var = tk.BooleanVar(value=bool(init_val))
                cb = ttk.Checkbutton(self.fields_frame, variable=var)
                cb.grid(row=r, column=1, sticky="w", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = cb

            elif ftype == "choice":
                var = tk.StringVar(value=str(init_val))
                combo = ttk.Combobox(
                    self.fields_frame, textvariable=var, state="readonly",
                    values=field.get("choices", []), width=28
                )
                combo.grid(row=r, column=1, sticky="ew", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = combo

            elif ftype == "json":
                var = tk.StringVar(value=json.dumps(init_val))
                ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                ent.grid(row=r, column=1, sticky="ew", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = ent

            elif ftype == "rgb":
                if isinstance(init_val, (list, tuple)) and len(init_val) == 3:
                    init_text = ",".join(str(int(x)) for x in init_val)
                else:
                    init_text = str(init_val)
                var = tk.StringVar(value=init_text)
                ent = ttk.Entry(self.fields_frame, textvariable=var, width=30)
                ent.grid(row=r, column=1, sticky="ew", pady=3)
                self.field_vars[key] = var
                self.widgets[key] = ent

            elif ftype == "buttons":
                frame = ttk.Frame(self.fields_frame)
                frame.grid(row=r, column=1, sticky="ew", pady=3)
                lb = tk.Listbox(frame, selectmode="multiple", height=6, exportselection=False)
                sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
                lb.configure(yscrollcommand=sb.set)
                lb.pack(side="left", fill="both", expand=True)
                sb.pack(side="left", fill="y")

                for b in ALL_BUTTONS:
                    lb.insert("end", b)

                init_buttons = init_val if isinstance(init_val, list) else []
                for i, b in enumerate(ALL_BUTTONS):
                    if b in init_buttons:
                        lb.selection_set(i)

                self.widgets[key] = lb
                self.field_vars[key] = None
            elif ftype == "pyfile":
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


            else:
                ttk.Label(self.fields_frame, text=f"(unsupported type: {ftype})").grid(row=r, column=1, sticky="w")

            ttk.Label(self.fields_frame, text=help_text, foreground="gray").grid(row=r, column=2, sticky="w", padx=(8, 0))

        self.fields_frame.columnconfigure(1, weight=1)

    def _parse_field(self, key, field):
        ftype = field["type"]

        if ftype == "buttons":
            lb = self.widgets[key]
            return [lb.get(i) for i in lb.curselection()]

        var = self.field_vars[key]
        raw = var.get() if var is not None else None

        if ftype == "int":
            return int(raw.strip())

        if ftype == "str":
            return str(raw)

        if ftype == "bool":
            return bool(var.get())

        if ftype == "choice":
            return raw

        if ftype == "json":
            return json.loads(raw.strip())

        if ftype == "rgb":
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

        # camera state
        self.cam_width = 640
        self.cam_height = 480
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



        # serial + engine
        self.serial = SerialController(status_cb=self.set_status)
        self.engine = ScriptEngine(
            self.serial,
            get_frame_fn=self.get_latest_frame,
            status_cb=self.set_status,
            on_ip_update=self.on_ip_update,
            on_tick=self.on_engine_tick,
        )

        self._build_ui()
        self._build_context_menu()
        self._schedule_frame_update()

        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.refresh_cameras()
        self.refresh_ports()
        self.refresh_scripts()
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
        top.columnconfigure(9, weight=1)

        # Camera controls
        ttk.Label(top, text="Camera:").grid(row=0, column=0, sticky="w")
        self.cam_var = tk.StringVar()
        self.cam_combo = ttk.Combobox(top, textvariable=self.cam_var, state="readonly", width=28)
        self.cam_combo.grid(row=0, column=1, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_cameras).grid(row=0, column=2, padx=(0, 6))
        self.cam_toggle_btn = ttk.Button(top, text="Start Cam", command=self.toggle_camera)
        self.cam_toggle_btn.grid(row=0, column=3, padx=(0, 18))

        # Serial controls
        ttk.Label(top, text="COM:").grid(row=0, column=4, sticky="w")
        self.com_var = tk.StringVar()
        self.com_combo = ttk.Combobox(top, textvariable=self.com_var, state="readonly", width=10)
        self.com_combo.grid(row=0, column=5, sticky="w", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_ports).grid(row=0, column=6, padx=(0, 6))
        self.ser_btn = ttk.Button(top, text="Connect", command=self.toggle_serial)
        self.ser_btn.grid(row=0, column=7, sticky="w", padx=(0, 18))

        # Script file controls
        ttk.Label(top, text="Script:").grid(row=0, column=8, sticky="w")
        self.script_var = tk.StringVar()
        self.script_combo = ttk.Combobox(top, textvariable=self.script_var, state="readonly", width=26)
        self.script_combo.grid(row=0, column=9, sticky="ew", padx=(6, 6))
        ttk.Button(top, text="Refresh", command=self.refresh_scripts).grid(row=0, column=10, padx=(0, 6))
        ttk.Button(top, text="Load", command=self.load_script_from_dropdown).grid(row=0, column=11, padx=(0, 6))
        ttk.Button(top, text="New", command=self.new_script).grid(row=0, column=12, padx=(0, 18))

        ttk.Button(top, text="Save", command=self.save_script).grid(row=0, column=13, padx=(0, 6))
        ttk.Button(top, text="Save As", command=self.save_script_as).grid(row=0, column=14, padx=(0, 6))

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

        self.video_label.grid(row=0, column=0, sticky="nsew")

        # Right: script viewer + vars
        right = ttk.Frame(main)
        right.rowconfigure(0, weight=1)
        right.rowconfigure(1, weight=1)
        right.columnconfigure(0, weight=1)

        main.add(left, minsize=320)
        main.add(right, minsize=380)

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
            x = int(total * 0.50)  # 68% left video, 32% right scripts

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

        # Auto-insert end markers for blocks (recommended)
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
        try:
            self.engine.rebuild_indexes(strict=True)  # strict only when running
            self.engine.run()
        except Exception as e:
            messagebox.showerror("Run error", str(e))


    def stop_script(self):
        if self.engine.running:
            self.engine.stop()
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

        dlg = CommandEditorDialog(self.root, self.engine.registry, initial_cmd=None, title="Add Command")
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
        dlg = CommandEditorDialog(self.root, self.engine.registry, initial_cmd=initial, title="Edit Command")
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

        self.root.destroy()


if __name__ == "__main__":
    os.makedirs("scripts", exist_ok=True)
    os.makedirs("py_scripts", exist_ok=True)
    root = tk.Tk()
    app = App(root)
    root.mainloop()
