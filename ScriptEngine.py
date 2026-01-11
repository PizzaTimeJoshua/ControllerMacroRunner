import time
import base64
from PIL import Image, ImageFilter, ImageOps, ImageEnhance
import numpy as np
import threading
import json
import os
import subprocess
import sys
import ast
import math
import re
from tkinter import messagebox

# Optional OCR support via pytesseract
try:
    import pytesseract
    PYTESSERACT_AVAILABLE = True
except ImportError:
    PYTESSERACT_AVAILABLE = False

# Optional OpenCV support for advanced image processing
try:
    import cv2
    CV2_AVAILABLE = True
except ImportError:
    CV2_AVAILABLE = False

# ----------------------------
# Pokemon Name Typer Keyboard Layouts
# Compatible with Pokemon FRLG and RSE naming screens
# ----------------------------
NAME_PAGE_UPPER = [
    'A','B','C','D','E','F',' ','.','next',
    'G','H','I','J','K','L',' ',',','back',
    'M','N','O','P','Q','R','S',' ','back',
    'T','U','V','W','X','Y','Z',' ','OK'
]
NAME_PAGE_LOWER = [
    'a','b','c','d','e','f',' ','.','next',
    'g','h','i','j','k','l',' ',',','back',
    'm','n','o','p','q','r','s',' ','back',
    't','u','v','w','x','y','z',' ','OK'
]
NAME_PAGE_OTHER = [
    '0','1','2','3','4',' ','next',
    '5','6','7','8','9',' ','back',
    '!','?','♂','♀','/','-','back',
    '…','"','"',''',''',' ','OK'
]

def get_page_width(page):
    """Returns the column width for a given name page."""
    return 7 if page is NAME_PAGE_OTHER else 9

# ----------------------------
# High-Precision Timing Utilities
# ----------------------------

def precise_sleep(duration_sec):
    """
    High-precision sleep using a hybrid approach:
    - Use time.sleep() for most of the duration
    - Busy-wait for the final ~2ms for precision

    This minimizes CPU usage while maintaining sub-millisecond precision.
    """
    if duration_sec <= 0:
        return

    # For very short durations (< 2ms), use pure busy-wait
    if duration_sec < 0.002:
        end = time.perf_counter() + duration_sec
        while time.perf_counter() < end:
            pass  # Busy-wait for precision
        return

    # For longer durations, sleep most of the time, then busy-wait
    # Sleep until 2ms before target to avoid oversleeping
    sleep_until = time.perf_counter() + duration_sec - 0.002

    # Sleep in small chunks, checking periodically
    while time.perf_counter() < sleep_until:
        remaining = sleep_until - time.perf_counter()
        if remaining > 0.005:  # If more than 5ms remaining, sleep
            time.sleep(min(remaining * 0.5, 0.001))  # Sleep conservatively
        else:
            break

    # Busy-wait for the final ~2ms for precision
    end = time.perf_counter() + duration_sec - (time.perf_counter() - (sleep_until - duration_sec + 0.002))
    while time.perf_counter() < end:
        pass

def precise_sleep_interruptible(duration_sec, stop_event):
    """
    High-precision interruptible sleep that checks stop_event.
    Returns True if interrupted, False if completed.
    """
    if duration_sec <= 0:
        return False

    end_time = time.perf_counter() + duration_sec

    # For very short durations, just busy-wait with stop checks
    if duration_sec < 0.002:
        while time.perf_counter() < end_time:
            if stop_event.is_set():
                return True
        return False

    # For longer durations, check stop event periodically
    while time.perf_counter() < end_time:
        if stop_event.is_set():
            return True

        remaining = end_time - time.perf_counter()
        if remaining <= 0:
            break

        # Sleep or busy-wait depending on remaining time
        if remaining > 0.002:
            # Sleep in small chunks for most of the duration
            time.sleep(min(remaining * 0.5, 0.001))
        else:
            # Busy-wait for final precision
            pass

    return False

# ----------------------------
# Script Command Spec
# ----------------------------

class CommandSpec:
    def __init__(self, name, required_keys, fn, doc="", arg_schema=None, format_fn=None,
                 group="Other", order=999,
                 test=False,exportable=True, export_note =""):
        self.name = name
        self.required_keys = required_keys
        self.fn = fn
        self.doc = doc
        self.arg_schema = arg_schema or []
        self.format_fn = format_fn or (lambda c: f"{name} " + " ".join(f"{k}={c.get(k)!r}" for k in c if k != "cmd"))
        self.group = group
        self.order = order
        self.test = test
        self.exportable = exportable
        self.export_note = export_note



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
    messagebox.showerror(f"Unknown op: {op}")
    return
def resolve_number(ctx, raw):
    """
    Resolves a numeric-ish value:
      - int/float pass through
      - "$var" becomes ctx["vars"]["var"]
      - "=expr" is evaluated via eval_expr (supports $var inside)
      - numeric strings like "123" are cast
    """
    if isinstance(raw, (int, float)):
        return raw

    if isinstance(raw, str):
        s = raw.strip()
        if s.startswith("="):
            return eval_expr(ctx, s[1:])
        if s.startswith("$"):
            return resolve_value(ctx, s)  # $var
        # fall back to parsing numeric string
        try:
            return int(s)
        except ValueError:
            return float(s)

    return raw



def build_label_index(commands):
    labels = {}
    for i, c in enumerate(commands):
        if c.get("cmd") == "label":
            name = c.get("name")
            if not name:
                messagebox.showerror(f"label missing name at index {i}")
                return
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
                    messagebox.showerror(f"end_if without if at index {i}")
                    return
                continue
            j = stack.pop()
            m[j] = i
    if stack and strict:
        messagebox.showerror(f"Unclosed if at index {stack[-1]}")
        return
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
                    messagebox.showerror(f"end_while without while at index {i}")
                    return
                continue
            w = stack.pop()
            while_to_end[w] = i
            end_to_while[i] = w
    if stack and strict:
        messagebox.showerror(f"Unclosed while at index {stack[-1]}")
        return
    return while_to_end, end_to_while, stack

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


# ----------------------------
# OCR / Text Recognition Helpers
# ----------------------------

def preprocess_for_ocr(img: Image.Image, scale: int = 4, threshold: int = 0, invert: bool = False):
    """
    Preprocess an image region for better OCR on pixel fonts.

    Args:
        img: PIL Image (RGB or RGBA)
        scale: Upscale factor (higher = better for small pixel fonts, default 4)
        threshold: Binary threshold 0-255 (0 = use automatic OTSU thresholding)
        invert: If True, invert colors (useful for light text on dark background)

    Returns:
        Tuple of (primary_image, fallback_image) - both preprocessed for OCR
    """
    # Convert to RGB if needed
    if img.mode != 'RGB':
        img = img.convert('RGB')

    # Upscale using nearest neighbor to preserve pixel edges
    w, h = img.size
    img = img.resize((w * scale, h * scale), Image.Resampling.NEAREST)

    # Invert if needed (light text on dark background)
    if invert:
        img = ImageOps.invert(img)

    # Convert to grayscale for processing
    grayscale = np.array(img.convert('L'))

    # Use OpenCV if available for better preprocessing
    if CV2_AVAILABLE:
        # Apply Gaussian blur to reduce noise
        blurred = cv2.GaussianBlur(grayscale, (5, 5), 0)

        # Apply binary thresholding (OTSU for automatic threshold selection)
        if threshold > 0:
            _, binary = cv2.threshold(blurred, threshold, 255, cv2.THRESH_BINARY)
        else:
            _, binary = cv2.threshold(blurred, 128, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)

        # Convert back to PIL Image
        binary_pil = Image.fromarray(binary)
    else:
        # Fallback without OpenCV - simple thresholding
        binary_pil = img.convert('L')
        if threshold > 0:
            binary_pil = binary_pil.point(lambda x: 255 if x > threshold else 0)

    # Create enhanced version (primary) with increased contrast and brightness
    img_enhanced = ImageEnhance.Brightness(binary_pil).enhance(2)
    img_enhanced = ImageEnhance.Contrast(img_enhanced).enhance(2)

    # Create blurred fallback version for difficult cases
    img_fallback = binary_pil.filter(ImageFilter.BLUR)

    return img_enhanced, img_fallback


def ocr_region(frame_bgr: np.ndarray, x: int, y: int, width: int, height: int,
               scale: int = 4, threshold: int = 0, invert: bool = False,
               psm: int = 7, whitelist: str = "") -> str:
    """
    Perform OCR on a region of the frame.

    Args:
        frame_bgr: BGR numpy array from camera
        x, y: Top-left corner of region
        width, height: Size of region
        scale: Upscale factor for preprocessing
        threshold: Binary threshold (0 = use automatic OTSU)
        invert: Invert colors for light-on-dark text
        psm: Tesseract Page Segmentation Mode (7 = single line, 6 = single block)
        whitelist: Characters to restrict OCR to (e.g., "0123456789" for numbers only)

    Returns:
        Recognized text string (stripped), or numeric string if whitelist is digits only
    """
    if not PYTESSERACT_AVAILABLE:
        raise RuntimeError("pytesseract is not installed. Install with: pip install pytesseract")

    if frame_bgr is None:
        return ""

    h_frame, w_frame, _ = frame_bgr.shape

    # Clamp region to frame bounds
    x = max(0, min(x, w_frame - 1))
    y = max(0, min(y, h_frame - 1))
    x2 = max(x + 1, min(x + width, w_frame))
    y2 = max(y + 1, min(y + height, h_frame))

    # Extract region (BGR)
    region_bgr = frame_bgr[y:y2, x:x2]

    # Convert to RGB PIL Image
    region_rgb = region_bgr[:, :, ::-1]
    img = Image.fromarray(region_rgb)

    # Preprocess for OCR (returns primary and fallback images)
    img_primary, img_fallback = preprocess_for_ocr(
        img, scale=scale, threshold=threshold, invert=invert
    )

    # Check if we're in numeric-only mode
    numeric_mode = (whitelist == "0123456789")

    # Build tesseract config
    config_parts = [f"--psm {psm}"]
    if whitelist:
        if numeric_mode:
            # Include commonly confused characters for better recognition
            config_parts.append("-c tessedit_char_whitelist=0123456789BSZOI")
        else:
            config_parts.append(f"-c tessedit_char_whitelist={whitelist}")

    config = " ".join(config_parts)

    def fix_numeric_confusions(text: str) -> str:
        """Fix common OCR confusions for numeric text."""
        return (text
                .replace('O', '0')
                .replace('I', '1')
                .replace('S', '5')
                .replace('Z', '2')
                .replace('B', '8'))

    # Run OCR
    try:
        text = pytesseract.image_to_string(img_primary, config=config).strip()

        if numeric_mode:
            text = fix_numeric_confusions(text)

            # If result isn't valid digits, try fallback image
            if not text.isdigit():
                fallback_text = pytesseract.image_to_string(
                    img_fallback, config=config
                ).strip()
                fallback_text = fix_numeric_confusions(fallback_text)

                # Use fallback if it's valid digits
                if fallback_text.isdigit():
                    text = fallback_text

        return text

    except Exception as e:
        raise RuntimeError(f"OCR failed: {e}")


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

_EXPR_VAR_RE = re.compile(r"\$([A-Za-z_]\w*)")

def eval_expr(ctx, expr: str):
    """
    Evaluate a simple math expression safely.
    Variables are referenced as $name inside the expression.
    Example: "9/$value*1000"
    """
    if not isinstance(expr, str):
        return expr

    # Replace $var with a python identifier var
    used = set(_EXPR_VAR_RE.findall(expr))
    py_expr = _EXPR_VAR_RE.sub(r"\1", expr)

    # Build locals for used vars (missing vars become error)
    local_vars = {}
    for name in used:
        if name not in ctx["vars"]:
            messagebox.showerror(f"Expression references undefined variable: ${name}")
            return
        local_vars[name] = ctx["vars"][name]

    allowed = {
        # math funcs
        "abs": abs,
        "round": round,
        "min": min,
        "max": max,
        "int": int,
        "float": float,
        # math module
        "math": math,
        "pi": math.pi,
        "e": math.e,
    }
    allowed.update(local_vars)

    node = ast.parse(py_expr, mode="eval")

    # Validate AST: allow only safe nodes
    for n in ast.walk(node):
        if isinstance(n, (ast.Expression, ast.Load, ast.Constant, ast.Name)):
            continue
        if isinstance(n, (ast.BinOp, ast.UnaryOp)):
            continue
        if isinstance(n, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.Pow)):
            continue
        if isinstance(n, (ast.UAdd, ast.USub)):
            continue
        if isinstance(n, ast.Call):
            # Allow calls only to whitelisted functions or math.<fn>
            if isinstance(n.func, ast.Name):
                if n.func.id not in allowed:
                    messagebox.showerror(f"Function not allowed: {n.func.id}")
                    return
            elif isinstance(n.func, ast.Attribute):
                # allow math.xxx only
                if not (isinstance(n.func.value, ast.Name) and n.func.value.id == "math"):
                    messagebox.showerror("Only math.<fn>(...) calls are allowed")
                    return
            else:
                messagebox.showerror("Invalid function call")
                return
            continue
        if isinstance(n, ast.Attribute):
            # allow math.<attr> only
            if not (isinstance(n.value, ast.Name) and n.value.id == "math"):
                messagebox.showerror("Only math.<attr> is allowed")
                return
            continue

        # Block everything else (comparisons, subscripts, lambdas, etc.)
        messagebox.showerror(f"Disallowed expression element: {type(n).__name__}")
        return

    return eval(compile(node, "<expr>", "eval"), {"__builtins__": {}}, allowed)

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

        self._backend_getter = None

    def set_backend_getter(self, fn):
        self._backend_getter = fn

    def get_backend(self):
        if self._backend_getter:
            return self._backend_getter()
        return None


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
            messagebox.showerror("Script must be a list of command objects.")
            return

        for i, c in enumerate(cmds):
            if not isinstance(c, dict) or "cmd" not in c:
                messagebox.showerror(f"Command at index {i} must be an object with 'cmd'.")
                return
            name = c["cmd"]
            if name not in self.registry:
                messagebox.showerror(f"Unknown cmd '{name}' at index {i}.")
                return
            spec = self.registry[name]
            for k in spec.required_keys:
                if k not in c:
                    messagebox.showerror(f"'{name}' missing required key '{k}' at index {i}.")
                    return

        self.commands = cmds
        self.rebuild_indexes()

        self.vars = {}
        self.ip = 0
        self.status_cb(f"Loaded script: {os.path.basename(path)} ({len(cmds)} commands)")

    def stop(self):
        self._stop.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=1.0)
        self.running = False
        self.serial.set_state(0, 0)
        self.status_cb("Script stopped.")
        self.on_ip_update(-1)
        self.ip = 0

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
            "get_backend": self.get_backend,
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
            raise(e)
        finally:
            self.running = False
            self.on_ip_update(-1)
            self.ip = 0
            self._stop.clear()

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
        def fmt_tap_touch(c):
            return f"TapTouch x={c.get('x')} y={c.get('y')} down={c.get('down_time', 0.1)} settle={c.get('settle', 0.1)}"
        def fmt_mash(c):
            btns = c.get("buttons", [])
            duration = c.get("duration_ms", 1000)
            hold = c.get("hold_ms", 25)
            wait = c.get("wait_ms", 25)
            return f"Mash {', '.join(btns) if btns else '(none)'} for {duration}ms (hold:{hold}ms wait:{wait}ms)"

        def fmt_type_name(c):
            name = c.get("name", "Red")
            confirm = c.get("confirm", True)
            confirm_str = " + confirm" if confirm else ""
            return f"TypeName \"{name}\"{confirm_str}"

        def fmt_read_text(c):
            x = c.get("x", 0)
            y = c.get("y", 0)
            w = c.get("width", 100)
            h = c.get("height", 20)
            out = c.get("out", "text")
            return f"ReadText ({x},{y}) {w}x{h} -> ${out}"

        # ---- execution fns
        def cmd_wait(ctx, c):
            ms_raw = c.get("ms", 0)
            ms = float(resolve_number(ctx, ms_raw))

            # Use high-precision interruptible sleep
            precise_sleep_interruptible(ms / 1000.0, ctx["stop"])


        def cmd_press(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("press: buttons must be a list")
                return

            # Press buttons with precise timing
            backend.set_buttons(buttons)

            ms_raw = c.get("ms", 50)
            ms = float(resolve_number(ctx, ms_raw))
            if ms > 0:
                # Use high-precision interruptible sleep
                precise_sleep_interruptible(ms / 1000.0, ctx["stop"])

            # Release buttons
            backend.set_buttons([])


        def cmd_hold(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("hold: buttons must be a list")
                return

            backend.set_buttons(buttons)


        def cmd_label(ctx, c):
            pass

        def cmd_goto(ctx, c):
            label = c["label"]
            if label not in ctx["labels"]:
                messagebox.showerror(f"Unknown label: {label}")
                return
            ctx["ip"] = ctx["labels"][label]

        def cmd_set(ctx, c):
            raw = c.get("value")
            if isinstance(raw, str) and raw.strip().startswith("="):
                ctx["vars"][c["var"]] = eval_expr(ctx, raw.strip()[1:])
            else:
                ctx["vars"][c["var"]] = resolve_value(ctx, raw)

        def cmd_add(ctx, c):
            var = c["var"]
            cur = ctx["vars"].get(var, 0)
            raw = c.get("value", 0)
            if isinstance(raw, str) and raw.strip().startswith("="):
                val = eval_expr(ctx, raw.strip()[1:])
            else:
                val = resolve_value(ctx, raw)
            ctx["vars"][var] = cur + val


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

        def cmd_read_text(ctx, c):
            """OCR a region of the camera frame and store the text in a variable."""
            frame = ctx["get_frame"]()
            out = c.get("out", "text")

            if frame is None:
                ctx["vars"][out] = ""
                return

            x = int(resolve_value(ctx, c.get("x", 0)))
            y = int(resolve_value(ctx, c.get("y", 0)))
            width = int(resolve_value(ctx, c.get("width", 100)))
            height = int(resolve_value(ctx, c.get("height", 20)))
            scale = int(resolve_value(ctx, c.get("scale", 4)))
            threshold = int(resolve_value(ctx, c.get("threshold", 0)))
            invert = bool(resolve_value(ctx, c.get("invert", False)))
            psm = int(resolve_value(ctx, c.get("psm", 7)))
            whitelist = str(resolve_value(ctx, c.get("whitelist", "")))

            try:
                text = ocr_region(
                    frame, x, y, width, height,
                    scale=scale, threshold=threshold, invert=invert,
                    psm=psm, whitelist=whitelist
                )
                ctx["vars"][out] = text
            except Exception as e:
                ctx["vars"][out] = ""
                messagebox.showerror(f"read_text error: {e}")

        def cmd_comment(ctx, c):
            pass

        def cmd_run_python(ctx, c):
            file_name = str(c["file"]).strip()
            if not file_name:
                messagebox.showerror("run_python: file is empty")
                return

            if os.path.isabs(file_name):
                script_path = file_name
            else:
                script_path = os.path.join("py_scripts", file_name)

            if not os.path.exists(script_path):
                messagebox.showerror(f"run_python: file not found: {script_path}")
                return

            args = c.get("args", [])
            if args is None:
                args = []
            if not isinstance(args, list):
                messagebox.showerror("run_python: args must be a JSON list")
                return

            # NEW: allow $var references (and nested structures)
            args = resolve_vars_deep(ctx, args)

            timeout_s = int(c.get("timeout_s", 10))
            res = run_python_main(script_path, args, timeout_s=timeout_s)

            outvar = (c.get("out") or "").strip()
            if outvar:
                ctx["vars"][outvar] = res
        def cmd_tap_touch(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            if not hasattr(backend, "tap_touch"):
                raise RuntimeError("tap_touch is only supported by the 3DS backend.")

            x = int(resolve_value(ctx, c.get("x")))
            y = int(resolve_value(ctx, c.get("y")))
            down_time = float(resolve_value(ctx, c.get("down_time", 0.1)))
            settle = float(resolve_value(ctx, c.get("settle", 0.1)))

            backend.tap_touch(x, y, down_time=down_time, settle=settle)

        def cmd_mash(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("mash: buttons must be a list")
                return

            duration_ms = float(resolve_number(ctx, c.get("duration_ms", 1000)))
            hold_ms = float(resolve_number(ctx, c.get("hold_ms", 25)))
            wait_ms = float(resolve_number(ctx, c.get("wait_ms", 25)))

            # Convert to seconds for precise timing
            hold_sec = hold_ms / 1000.0
            wait_sec = wait_ms / 1000.0

            # Calculate end time
            total_duration = duration_ms / 1000.0
            end_time = time.perf_counter() + total_duration

            # Main mashing loop with precise timing
            while time.perf_counter() < end_time:
                if ctx["stop"].is_set():
                    break

                # Press buttons
                backend.set_buttons(buttons)

                # Hold for precise duration
                if precise_sleep_interruptible(hold_sec, ctx["stop"]):
                    break  # Interrupted

                # Release buttons
                backend.set_buttons([])

                # Check if we have time for a full wait cycle
                time_remaining = end_time - time.perf_counter()
                if time_remaining <= 0:
                    break

                # Wait for precise duration, but not longer than remaining time
                wait_duration = min(wait_sec, time_remaining)
                if precise_sleep_interruptible(wait_duration, ctx["stop"]):
                    break  # Interrupted

            # Ensure buttons are released at the end
            backend.set_buttons([])

        def cmd_type_name(ctx, c):
            """
            Types a name on Pokemon FRLG/RSE naming screens.
            Navigates the on-screen keyboard using D-pad, Select to switch pages,
            A to select letters, and Start+A to confirm.
            """
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            name = str(resolve_value(ctx, c.get("name", "Red")))
            confirm = bool(resolve_value(ctx, c.get("confirm", True)))

            # Timing settings (in milliseconds, converted to seconds)
            move_delay_ms = float(resolve_number(ctx, c.get("move_delay_ms", 200)))
            select_delay_ms = float(resolve_number(ctx, c.get("select_delay_ms", 600)))
            press_delay_ms = float(resolve_number(ctx, c.get("press_delay_ms", 400)))
            button_hold_ms = float(resolve_number(ctx, c.get("button_hold_ms", 50)))

            move_delay = move_delay_ms / 1000.0
            select_delay = select_delay_ms / 1000.0
            press_delay = press_delay_ms / 1000.0
            button_hold = button_hold_ms / 1000.0

            name_pages = [NAME_PAGE_UPPER, NAME_PAGE_LOWER, NAME_PAGE_OTHER]
            current_position = [0, 0]  # [x, y] cursor position
            current_page = 0
            current_page_width = get_page_width(name_pages[current_page])

            def press_button(btn):
                """Press a button with proper timing. Returns True if interrupted."""
                if ctx["stop"].is_set():
                    return True
                backend.set_buttons([btn])
                if precise_sleep_interruptible(button_hold, ctx["stop"]):
                    backend.set_buttons([])
                    return True
                backend.set_buttons([])
                return False

            def check_stop():
                """Returns True if stop was requested."""
                return ctx["stop"].is_set()

            try:
                for letter in name:
                    if check_stop():
                        break

                    # Check if the letter exists in any page
                    letter_found = any(letter in page for page in name_pages)
                    if not letter_found:
                        # Skip characters that don't exist in the keyboard
                        continue

                    # Find the page containing this letter
                    while letter not in name_pages[current_page]:
                        if check_stop():
                            break

                        # Press Select to switch pages
                        if press_button("Select"):
                            break
                        if precise_sleep_interruptible(select_delay, ctx["stop"]):
                            break

                        # Move to next page
                        current_page = (current_page + 1) % len(name_pages)
                        new_page_width = get_page_width(name_pages[current_page])

                        # Adjust cursor position when switching pages
                        # If on rightmost column, stay on rightmost column of new page
                        if current_position[0] >= current_page_width - 1:
                            current_position[0] = new_page_width - 1
                        else:
                            # Otherwise, clamp to the new page width
                            # (leave room for the control column on the right)
                            current_position[0] = min(current_position[0], new_page_width - 2)
                        current_page_width = new_page_width

                    if check_stop():
                        break

                    # Locate letter position in the current page
                    pos = name_pages[current_page].index(letter)
                    target_x = pos % current_page_width
                    target_y = pos // current_page_width

                    # Navigate horizontally to the letter
                    while current_position[0] != target_x:
                        if check_stop():
                            break
                        if target_x > current_position[0]:
                            if press_button("Right"):
                                break
                            current_position[0] += 1
                        elif target_x < current_position[0]:
                            if press_button("Left"):
                                break
                            current_position[0] -= 1
                        if precise_sleep_interruptible(move_delay, ctx["stop"]):
                            break

                    if check_stop():
                        break

                    # Navigate vertically to the letter
                    while current_position[1] != target_y:
                        if check_stop():
                            break
                        if target_y > current_position[1]:
                            if press_button("Down"):
                                break
                            current_position[1] += 1
                        elif target_y < current_position[1]:
                            if press_button("Up"):
                                break
                            current_position[1] -= 1
                        if precise_sleep_interruptible(move_delay, ctx["stop"]):
                            break

                    if check_stop():
                        break

                    # Select the letter
                    if press_button("A"):
                        break
                    if precise_sleep_interruptible(press_delay, ctx["stop"]):
                        break

                # Press Start to finish name entry
                if not check_stop():
                    press_button("Start")
                    precise_sleep_interruptible(move_delay, ctx["stop"])

                # Confirm if requested
                if confirm and not check_stop():
                    press_button("A")
                    precise_sleep_interruptible(move_delay, ctx["stop"])

            finally:
                # Ensure buttons are released
                backend.set_buttons([])


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
                arg_schema=[{"key": "ms", "type": "json", "default": 100, "help": "Milliseconds to wait"}],
                format_fn=fmt_wait,
                group="Timing",
                order=10
            ),
            CommandSpec(
                "press", ["buttons", "ms"], cmd_press,
                doc="Press buttons for ms, then release to neutral.",
                arg_schema=[
                    {"key": "buttons", "type": "buttons", "default": ["A"], "help": "Buttons to press"},
                    {"key": "ms", "type": "json", "default": 80, "help": "Hold duration in milliseconds"},
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
                "mash", ["buttons", "duration_ms"], cmd_mash,
                doc="Rapidly mash buttons for a duration. Default: ~20 presses/second.",
                arg_schema=[
                    {"key": "buttons", "type": "buttons", "default": ["A"], "help": "Buttons to mash"},
                    {"key": "duration_ms", "type": "json", "default": 1000, "help": "Total mashing duration in milliseconds"},
                    {"key": "hold_ms", "type": "json", "default": 25, "help": "How long to hold each press (default: 25ms)"},
                    {"key": "wait_ms", "type": "json", "default": 25, "help": "Wait between presses (default: 25ms)"},
                ],
                format_fn=fmt_mash,
                group="Controller",
                order=15
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
                order=10,
                test = True,
                exportable=False,
                export_note="Requires camera frame processing which is unsupported at the moment."
            ),
            CommandSpec(
                "read_text", ["x", "y", "width", "height", "out"], cmd_read_text,
                doc="OCR a region of the camera frame to extract text. Uses pytesseract with preprocessing for pixel fonts.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate (top-left corner)"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate (top-left corner)"},
                    {"key": "width", "type": "int", "default": 100, "help": "Width of region"},
                    {"key": "height", "type": "int", "default": 20, "help": "Height of region"},
                    {"key": "out", "type": "str", "default": "text", "help": "Variable name to store result (no $)"},
                    {"key": "scale", "type": "int", "default": 4, "help": "Upscale factor (higher = better for small fonts)"},
                    {"key": "threshold", "type": "int", "default": 0, "help": "Binary threshold 0-255 (0 = disabled)"},
                    {"key": "invert", "type": "bool", "default": False, "help": "Invert colors (for light text on dark)"},
                    {"key": "psm", "type": "int", "default": 7, "help": "Tesseract PSM (7=single line, 6=block, 13=raw)"},
                    {"key": "whitelist", "type": "str", "default": "", "help": "Allowed chars (e.g. '0123456789' for numbers)"},
                ],
                format_fn=fmt_read_text,
                group="Image",
                order=20,
                test=True,
                exportable=False,
                export_note="Requires camera frame and pytesseract which are unsupported in standalone export."
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
            CommandSpec(
                "tap_touch",
                ["x", "y"],
                cmd_tap_touch,
                doc="3DS only: taps the touchscreen at (x,y).",
                arg_schema=[
                    {"key":"x","type":"int","default":160, "help":"X pixel (0..319)"},
                    {"key":"y","type":"int","default":120, "help":"Y pixel (0..239)"},
                    {"key":"down_time","type":"float","default":0.1, "help":"Seconds to hold touch down"},
                    {"key":"settle","type":"float","default":0.1, "help":"Seconds to wait after release"},
                ],
                format_fn=fmt_tap_touch,
                group="3DS",
                order=10
            ),
            CommandSpec(
                "type_name",
                ["name"],
                cmd_type_name,
                doc="Type a name on Pokemon FRLG/RSE naming screens. Navigates the on-screen keyboard using D-pad, Select to switch pages, A to select letters, and Start+A to confirm.",
                arg_schema=[
                    {"key": "name", "type": "str", "default": "Red", "help": "The name to type (supports uppercase, lowercase, numbers, and some symbols)"},
                    {"key": "confirm", "type": "bool", "default": True, "help": "Press A after Start to confirm the name"},
                    {"key": "move_delay_ms", "type": "int", "default": 200, "help": "Delay after each D-pad move (ms)"},
                    {"key": "select_delay_ms", "type": "int", "default": 600, "help": "Delay after Select to switch pages (ms) Only adjust if overclocking gamespeed"},
                    {"key": "press_delay_ms", "type": "int", "default": 400, "help": "Delay after pressing A to select a letter (ms)"},
                    {"key": "button_hold_ms", "type": "int", "default": 50, "help": "How long to hold each button press (ms)"},
                ],
                format_fn=fmt_type_name,
                group="Pokemon",
                order=10,
                exportable=False,
                export_note="Complex keyboard navigation not supported in standalone export."
            ),

        ]

        for s in specs:
            reg[s.name] = s
        return reg

