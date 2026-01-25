"""
Script engine for macro execution.

Parses JSON script files containing commands (button presses, waits,
conditionals, loops) and executes them against a connected controller backend.
Also provides image analysis commands (color matching, OCR) using the camera feed.
"""
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
import random
import io
import mimetypes
import urllib.request
import urllib.error
import uuid
from tkinter import messagebox
from utils import exe_dir_path, python_path, is_python_available, ffplay_path, find_sound_file, list_sound_files

# Optional OCR support via pytesseract
try:
    import pytesseract
    from utils import tesseract_path
    # Configure pytesseract to use bundled binary if available
    _tesseract_cmd = tesseract_path()
    if _tesseract_cmd != "tesseract":
        pytesseract.pytesseract.tesseract_cmd = _tesseract_cmd
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
# Custom Exceptions
# ----------------------------

class PythonNotAvailableError(Exception):
    """Raised when run_python is called but no Python interpreter is available."""
    pass


# ----------------------------
# CIE76 Color Difference (Delta E) Functions
# ----------------------------

def rgb_to_xyz(r, g, b):
    """
    Convert sRGB to CIE XYZ color space.
    Uses D65 white point.
    """
    # Normalize RGB to 0-1 range
    r = r / 255.0
    g = g / 255.0
    b = b / 255.0

    # Apply sRGB gamma correction (inverse companding)
    def linearize(c):
        if c <= 0.04045:
            return c / 12.92
        return ((c + 0.055) / 1.055) ** 2.4

    r = linearize(r)
    g = linearize(g)
    b = linearize(b)

    # sRGB to XYZ matrix (D65 white point)
    x = r * 0.4124564 + g * 0.3575761 + b * 0.1804375
    y = r * 0.2126729 + g * 0.7151522 + b * 0.0721750
    z = r * 0.0193339 + g * 0.1191920 + b * 0.9503041

    return x, y, z


def xyz_to_lab(x, y, z):
    """
    Convert CIE XYZ to CIELAB color space.
    Uses D65 white point (Xn=0.95047, Yn=1.0, Zn=1.08883).
    """
    # D65 white point
    xn, yn, zn = 0.95047, 1.0, 1.08883

    # Normalize by white point
    x = x / xn
    y = y / yn
    z = z / zn

    # Apply f(t) function
    def f(t):
        delta = 6.0 / 29.0
        if t > delta ** 3:
            return t ** (1.0 / 3.0)
        return t / (3.0 * delta ** 2) + 4.0 / 29.0

    fx = f(x)
    fy = f(y)
    fz = f(z)

    L = 116.0 * fy - 16.0
    a = 500.0 * (fx - fy)
    b = 200.0 * (fy - fz)

    return L, a, b


def rgb_to_lab(r, g, b):
    """Convert RGB (0-255) to CIELAB color space."""
    x, y, z = rgb_to_xyz(r, g, b)
    return xyz_to_lab(x, y, z)


def delta_e_cie76(rgb1, rgb2):
    """
    Calculate CIE76 color difference (Delta E) between two RGB colors.

    Delta E values interpretation:
    - 0-1: Not perceptible by human eyes
    - 1-2: Perceptible through close observation
    - 2-10: Perceptible at a glance
    - 11-49: Colors are more similar than opposite
    - 100: Colors are exact opposite

    Args:
        rgb1: First color as (R, G, B) with values 0-255
        rgb2: Second color as (R, G, B) with values 0-255

    Returns:
        Delta E value (0 = identical, higher = more different)
    """
    # Clamp RGB values to valid range
    r1 = max(0, min(255, int(rgb1[0])))
    g1 = max(0, min(255, int(rgb1[1])))
    b1 = max(0, min(255, int(rgb1[2])))
    r2 = max(0, min(255, int(rgb2[0])))
    g2 = max(0, min(255, int(rgb2[1])))
    b2 = max(0, min(255, int(rgb2[2])))

    L1, a1, b1_lab = rgb_to_lab(r1, g1, b1)
    L2, a2, b2_lab = rgb_to_lab(r2, g2, b2)

    return math.sqrt((L2 - L1) ** 2 + (a2 - a1) ** 2 + (b2_lab - b1_lab) ** 2)

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
    """
    Resolve a value that may be a variable reference.
    Supports:
      - Simple variable: "$varname" -> ctx["vars"]["varname"]
      - Indexed access: "$list[0]" or "$list[0][1]" for nested access
    """
    if isinstance(v, str) and v.startswith("$"):
        var_str = v[1:]  # Remove leading $

        # Check for index access: varname[index] or varname[i][j]...
        bracket_pos = var_str.find("[")
        if bracket_pos == -1:
            # Simple variable lookup
            return ctx["vars"].get(var_str, None)

        # Extract variable name and indices
        var_name = var_str[:bracket_pos]
        indices_str = var_str[bracket_pos:]

        # Get the base variable value
        value = ctx["vars"].get(var_name, None)
        if value is None:
            return None

        # Parse and apply indices: [0][1][2] etc.
        import re
        index_pattern = re.compile(r'\[([^\]]+)\]')
        indices = index_pattern.findall(indices_str)

        for idx_str in indices:
            idx_str = idx_str.strip()
            # Check if the index itself is a variable reference
            if idx_str.startswith("$"):
                idx = resolve_value(ctx, idx_str)
            else:
                # Try to parse as integer, otherwise use as string key
                try:
                    idx = int(idx_str)
                except ValueError:
                    # Could be a string key for dict access
                    idx = idx_str.strip('"\'')

            # Apply the index
            try:
                value = value[idx]
            except (IndexError, KeyError, TypeError):
                return None

        return value
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
    messagebox.showerror("Comparison Error", f"Unknown op: {op}")
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
                messagebox.showerror("Label Error", f"label missing name at index {i}")
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
                    messagebox.showerror("If Statement Error", f"end_if without if at index {i}")
                    return
                continue
            j = stack.pop()
            m[j] = i
    if stack and strict:
        messagebox.showerror("If Statement Error", f"Unclosed if at index {stack[-1]}")
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
                    messagebox.showerror("While Loop Error", f"end_while without while at index {i}")
                    return
                continue
            w = stack.pop()
            while_to_end[w] = i
            end_to_while[i] = w
    if stack and strict:
        messagebox.showerror("While Loop Error", f"Unclosed while at index {stack[-1]}")
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
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    b64 = base64.b64encode(buf.getvalue()).decode("ascii")
    return {"__frame__": "png_base64", "data_b64": b64}

def frame_to_png_bytes(frame_bgr: np.ndarray):
    """
    Convert BGR frame (H,W,3 uint8) to PNG bytes.
    """
    if frame_bgr is None:
        return None
    rgb = frame_bgr[:, :, ::-1]
    img = Image.fromarray(rgb)
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()

def _encode_multipart_form(payload_json: str, file_name: str, file_bytes: bytes, content_type: str):
    boundary = f"----CMRBoundary{uuid.uuid4().hex}"
    buf = io.BytesIO()

    def write_line(line: str = ""):
        buf.write(line.encode("utf-8"))
        buf.write(b"\r\n")

    write_line(f"--{boundary}")
    write_line('Content-Disposition: form-data; name="payload_json"')
    write_line("Content-Type: application/json")
    write_line()
    buf.write(payload_json.encode("utf-8"))
    buf.write(b"\r\n")

    write_line(f"--{boundary}")
    write_line(f'Content-Disposition: form-data; name="files[0]"; filename="{file_name}"')
    write_line(f"Content-Type: {content_type}")
    write_line()
    buf.write(file_bytes)
    buf.write(b"\r\n")
    write_line(f"--{boundary}--")

    body = buf.getvalue()
    content_type_header = f"multipart/form-data; boundary={boundary}"
    return body, content_type_header

def send_discord_webhook(url: str, payload: dict, file_tuple=None, timeout_s: int = 10):
    """
    Send a Discord webhook with optional image attachment.
    file_tuple is (filename, bytes, content_type).
    """
    payload_json = json.dumps(payload, ensure_ascii=False)

    if file_tuple:
        file_name, file_bytes, content_type = file_tuple
        body, content_type_header = _encode_multipart_form(
            payload_json, file_name, file_bytes, content_type
        )
    else:
        body = payload_json.encode("utf-8")
        content_type_header = "application/json"

    req = urllib.request.Request(
        url,
        data=body,
        headers={
            "Content-Type": content_type_header,
            "User-Agent": "ControllerMacroRunner",
        },
    )

    try:
        with urllib.request.urlopen(req, timeout=timeout_s) as resp:
            resp.read()
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = e.read().decode("utf-8", "replace").strip()
        except Exception:
            pass
        msg = f"Discord webhook error: HTTP {e.code} {e.reason}"
        if detail:
            msg = f"{msg} - {detail}"
        raise RuntimeError(msg) from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Discord webhook error: {e.reason}") from e


# ----------------------------
# Sound Playback Helpers
# ----------------------------

def play_sound_file(sound_name: str, volume: int = 80, wait: bool = False,
                    stop_event: threading.Event | None = None) -> tuple[bool, str]:
    """
    Play a sound from bin/sounds using ffplay.
    Returns (ok, message).
    """
    name = os.path.basename(str(sound_name or "")).strip()
    if not name:
        return False, "play_sound: sound is empty"

    sound_path = find_sound_file(name)
    if not sound_path:
        return False, f"play_sound: sound not found: {name}"

    try:
        volume_val = int(round(float(volume)))
    except (TypeError, ValueError):
        volume_val = 100
    volume_val = max(0, min(100, volume_val))

    ffplay = ffplay_path()
    args = [ffplay, "-nodisp", "-autoexit", "-loglevel", "quiet", "-volume", str(volume_val), sound_path]
    creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)

    try:
        if wait:
            proc = subprocess.Popen(
                args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags
            )
            stopped = False
            while proc.poll() is None:
                if stop_event is not None and stop_event.is_set():
                    stopped = True
                    proc.terminate()
                    try:
                        proc.wait(timeout=0.5)
                    except subprocess.TimeoutExpired:
                        proc.kill()
                    break
                time.sleep(0.05)
            if stopped:
                return True, f"play_sound: stopped {name}"
        else:
            subprocess.Popen(
                args,
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
                creationflags=creationflags
            )
    except FileNotFoundError:
        return False, "play_sound: ffplay not found. Install FFmpeg or place it in bin/ffmpeg."
    except Exception as e:
        return False, f"play_sound: {e}"

    action = "Played" if wait else "Playing"
    return True, f"{action} {name} at volume {volume_val}."


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
            config_parts.append("-c tessedit_char_whitelist=0123456789BSZOIl[]")
        else:
            config_parts.append(f"-c tessedit_char_whitelist={whitelist}")

    config = " ".join(config_parts)

    def fix_numeric_confusions(text: str) -> str:
        """Fix common OCR confusions for numeric text."""
        return (text
                .replace('O', '0')
                .replace('I', '1')
                .replace('l', '1')
                .replace('[', '1')
                .replace(']', '1')
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

    # Use embedded Python if available, otherwise fall back to sys.executable
    python_exe = python_path()

    cp = subprocess.run(
        [python_exe, "-c", runner, script_path, args_json],
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

# Safe list/dict methods that don't require external state
_SAFE_LIST_METHODS = frozenset({
    # Non-mutating methods (return new values)
    "copy", "count", "index",
    # Mutating methods (safe because we operate on copies)
    "append", "clear", "extend", "insert", "pop", "remove", "reverse", "sort",
})

# List methods that mutate in-place and return None - we wrap these to return the list instead
_LIST_METHODS_RETURN_SELF = frozenset({
    "append", "extend", "insert", "remove", "reverse", "sort", "clear",
})

class _ListWrapper(list):
    """List wrapper that returns self for mutating methods instead of None."""
    def append(self, item):
        super().append(item)
        return self
    def extend(self, items):
        super().extend(items)
        return self
    def insert(self, index, item):
        super().insert(index, item)
        return self
    def pop(self, index=-1):
        super().pop(index)
        return self
    def remove(self, item):
        super().remove(item)
        return self
    def reverse(self):
        super().reverse()
        return self
    def sort(self, *, key=None, reverse=False):
        super().sort(key=key, reverse=reverse)
        return self
    def clear(self):
        super().clear()
        return self

class _DictWrapper(dict):
    """Dict wrapper that returns self for mutating methods instead of None."""
    def clear(self):
        super().clear()
        return self
    def update(self, *args, **kwargs):
        super().update(*args, **kwargs)
        return self

_SAFE_DICT_METHODS = frozenset({
    "copy", "get", "items", "keys", "values",
    "pop", "popitem", "clear", "update", "setdefault",
})

_SAFE_STR_METHODS = frozenset({
    "lower", "upper", "strip", "lstrip", "rstrip", "split", "join",
    "replace", "find", "rfind", "index", "rindex", "count",
    "startswith", "endswith", "isdigit", "isalpha", "isalnum",
    "capitalize", "title", "swapcase", "center", "ljust", "rjust",
    "zfill", "partition", "rpartition", "format",
})

def eval_expr(ctx, expr: str):
    """
    Evaluate a simple math expression safely.
    Variables are referenced as $name inside the expression.
    Supports:
      - Math operations: "9/$value*1000"
      - Index access: "$list[0]" or "$dict['key']"
      - List methods: "$list.pop()" (operates on a copy, original unchanged)
      - String methods: "$str.upper()"
    """
    import copy

    if not isinstance(expr, str):
        return expr

    # Replace $var with a python identifier var
    used = set(_EXPR_VAR_RE.findall(expr))
    py_expr = _EXPR_VAR_RE.sub(r"\1", expr)

    # Build locals for used vars (missing vars become error)
    # Use deep copies for mutable types to prevent mutation of originals
    # Lists are wrapped with _ListWrapper so mutating methods return the list
    local_vars = {}
    for name in used:
        if name not in ctx["vars"]:
            messagebox.showerror("Variable Error", f"Expression references undefined variable: ${name}")
            return
        val = ctx["vars"][name]
        # Deep copy mutable types to prevent mutation
        # Use wrappers so mutating methods return the container instead of None
        if isinstance(val, list):
            local_vars[name] = _ListWrapper(copy.deepcopy(val))
        elif isinstance(val, dict):
            local_vars[name] = _DictWrapper(copy.deepcopy(val))
        else:
            local_vars[name] = val

    # Built-in functions allowed in expressions
    allowed_funcs = {
        # math funcs
        "abs": abs,
        "round": round,
        "min": min,
        "max": max,
        "int": int,
        "float": float,
        "len": len,
        "str": str,
        "list": list,
        "dict": dict,
        "tuple": tuple,
        "sorted": sorted,
        "reversed": reversed,
        "enumerate": enumerate,
        "zip": zip,
        "range": range,
        "sum": sum,
        "any": any,
        "all": all,
        # math module
        "math": math,
        "pi": math.pi,
        "e": math.e,
    }
    allowed = dict(allowed_funcs)
    allowed.update(local_vars)

    node = ast.parse(py_expr, mode="eval")

    def is_safe_method_call(call_node):
        """Check if a method call is on a safe method of a known type."""
        if not isinstance(call_node.func, ast.Attribute):
            return False
        method_name = call_node.func.attr
        # Allow safe list/dict/str methods
        return method_name in _SAFE_LIST_METHODS | _SAFE_DICT_METHODS | _SAFE_STR_METHODS

    def is_var_or_subscript(node):
        """Check if node is a variable name or subscript of a variable."""
        if isinstance(node, ast.Name):
            return node.id in local_vars
        if isinstance(node, ast.Subscript):
            return is_var_or_subscript(node.value)
        if isinstance(node, ast.Attribute):
            return is_var_or_subscript(node.value)
        return False

    # Validate AST: allow only safe nodes
    for n in ast.walk(node):
        if isinstance(n, (ast.Expression, ast.Load, ast.Store, ast.Constant, ast.Name)):
            continue
        if isinstance(n, (ast.BinOp, ast.UnaryOp)):
            continue
        if isinstance(n, (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.FloorDiv, ast.Mod, ast.Pow)):
            continue
        if isinstance(n, (ast.UAdd, ast.USub)):
            continue
        # Allow subscript for index access
        if isinstance(n, ast.Subscript):
            continue
        # Allow slice for slicing operations
        if isinstance(n, ast.Slice):
            continue
        # Allow list/tuple/dict literals
        if isinstance(n, (ast.List, ast.Tuple, ast.Dict)):
            continue
        # Allow comparisons for expressions like "x if a > b else y"
        if isinstance(n, (ast.Compare, ast.Eq, ast.NotEq, ast.Lt, ast.LtE, ast.Gt, ast.GtE, ast.In, ast.NotIn)):
            continue
        # Allow ternary expressions
        if isinstance(n, ast.IfExp):
            continue
        # Allow boolean operations
        if isinstance(n, (ast.BoolOp, ast.And, ast.Or, ast.Not)):
            continue
        if isinstance(n, ast.Call):
            # Allow calls to whitelisted functions
            if isinstance(n.func, ast.Name):
                if n.func.id not in allowed_funcs:
                    messagebox.showerror("Expression Error", f"Function not allowed: {n.func.id}")
                    return
            elif isinstance(n.func, ast.Attribute):
                # Allow math.xxx
                if isinstance(n.func.value, ast.Name) and n.func.value.id == "math":
                    pass
                # Allow safe methods on variables (list.pop(), str.upper(), etc.)
                elif is_var_or_subscript(n.func.value) and is_safe_method_call(n):
                    pass
                else:
                    method_name = n.func.attr
                    if method_name not in _SAFE_LIST_METHODS | _SAFE_DICT_METHODS | _SAFE_STR_METHODS:
                        messagebox.showerror("Expression Error", f"Method not allowed: {method_name}")
                    else:
                        messagebox.showerror("Expression Error", "Method calls only allowed on variables")
                    return
            else:
                messagebox.showerror("Expression Error", "Invalid function call")
                return
            continue
        if isinstance(n, ast.Attribute):
            # Allow math.<attr>
            if isinstance(n.value, ast.Name) and n.value.id == "math":
                continue
            # Allow attribute access on variables for method calls
            if is_var_or_subscript(n.value):
                continue
            messagebox.showerror("Expression Error", f"Attribute access not allowed: {n.attr}")
            return

        # Block everything else (lambdas, etc.)
        messagebox.showerror("Expression Error", f"Disallowed expression element: {type(n).__name__}")
        return

    return eval(compile(node, "<expr>", "eval"), {"__builtins__": {}}, allowed)

# ----------------------------
# Script Engine
# ----------------------------

class ScriptEngine:
    def __init__(self, serial_ctrl, get_frame_fn, status_cb=None, on_ip_update=None, on_tick=None,
                 settings_getter=None, on_python_needed=None, on_error=None):
        self.serial = serial_ctrl
        self.get_frame = get_frame_fn
        self.status_cb = status_cb or (lambda s: None)
        self.on_ip_update = on_ip_update or (lambda ip: None)
        self.on_tick = on_tick or (lambda: None)
        self.on_python_needed = on_python_needed or (lambda: None)
        self.on_error = on_error or (lambda title, msg: None)

        self.vars = {}
        self.commands = []
        self.labels = {}
        self.if_map = {}
        self.while_to_end = {}
        self.end_to_while = {}
        self._unclosed_ifs = []
        self._unclosed_whiles = []

        # Initialize random seed to current time
        random.seed(time.time())

        self.registry = self._build_default_registry()

        self._stop = threading.Event()
        self._thread = None
        self.running = False
        self.ip = 0

        self._backend_getter = None
        self._settings_getter = settings_getter or (lambda: {})

    def set_backend_getter(self, fn):
        self._backend_getter = fn

    def get_backend(self):
        if self._backend_getter:
            return self._backend_getter()
        return None

    def set_settings_getter(self, fn):
        self._settings_getter = fn

    def get_settings(self):
        if self._settings_getter:
            return self._settings_getter()
        return {}


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
            messagebox.showerror("Script Error", "Script must be a list of command objects.")
            return

        rename_map = {
            "set_circle_pad": "set_left_stick",
            "reset_circle_pad": "reset_left_stick",
            "set_c_stick": "set_right_stick",
            "reset_c_stick": "reset_right_stick",
        }
        for c in cmds:
            if isinstance(c, dict):
                name = c.get("cmd")
                if name in rename_map:
                    c["cmd"] = rename_map[name]

        for i, c in enumerate(cmds):
            if not isinstance(c, dict) or "cmd" not in c:
                messagebox.showerror("Script Error", f"Command at index {i} must be an object with 'cmd'.")
                return
            name = c["cmd"]
            if name not in self.registry:
                messagebox.showerror("Script Error", f"Unknown cmd '{name}' at index {i}.")
                return
            spec = self.registry[name]
            for k in spec.required_keys:
                if k not in c:
                    messagebox.showerror("Script Error", f"'{name}' missing required key '{k}' at index {i}.")
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
        self._reset_backend_neutral()
        self.status_cb("Script stopped.")
        self.on_ip_update(-1)
        self.ip = 0

    def run(self):
        backend = self.get_backend() or self.serial
        if backend is None or not getattr(backend, "connected", False):
            raise RuntimeError("No output backend connected.")
        if not self.commands:
            raise RuntimeError("No script loaded.")
        if self.running:
            return

        unsupported = self._find_unsupported_commands(backend)
        if unsupported:
            backend_name = getattr(backend, "backend_name", None)
            if backend_name is None:
                inner = getattr(backend, "backend", None)
                if inner is not None:
                    backend_name = getattr(inner, "backend_name", inner.__class__.__name__)
                else:
                    backend_name = backend.__class__.__name__
            msg_lines = [
                f"The current backend ({backend_name}) does not support:",
                "",
            ]
            for i, name in unsupported[:12]:
                msg_lines.append(f"- Line {i + 1}: {name}")
            if len(unsupported) > 12:
                msg_lines.append(f"...and {len(unsupported) - 12} more.")
            msg_lines.append("")
            msg_lines.append("The script will not run until unsupported commands are removed.")
            messagebox.showwarning("Unsupported commands", "\n".join(msg_lines))
            return

        # Reset variables and instruction pointer before running
        self.vars = {}
        self.ip = 0

        self.running = True
        self._stop.clear()
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()

    def _find_unsupported_commands(self, backend):
        if backend is None:
            return []

        requirements = {
            "tap_touch": "tap_touch",
            "set_left_stick": "set_left_stick",
            "reset_left_stick": "reset_left_stick",
            "set_right_stick": "set_right_stick",
            "reset_right_stick": "reset_right_stick",
            "press_ir": "set_ir_buttons",
            "hold_ir": "set_ir_buttons",
            "press_interface": "set_interface_buttons",
            "hold_interface": "set_interface_buttons",
        }

        def supports(method_name):
            if hasattr(backend, "backend"):
                inner = getattr(backend, "backend", None)
                return inner is not None and hasattr(inner, method_name)
            return hasattr(backend, method_name)

        unsupported = []
        for i, c in enumerate(self.commands):
            name = c.get("cmd")
            required = requirements.get(name)
            if required and not supports(required):
                unsupported.append((i, name))
        return unsupported

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
            "get_settings": self.get_settings,
            "on_python_needed": self.on_python_needed,
            "_timing_reference": None,  # perf_counter at start_timing for cumulative timing
        }

        try:
            self.status_cb("Running Script.")
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

            self._reset_backend_neutral()
            if not self._stop.is_set():
                self.status_cb("Script completed.")
        except Exception as e:
            self._reset_backend_neutral()
            self.status_cb(f"Script error: {e}")
            self.on_error("Script Error", str(e))
        finally:
            self.running = False
            self.on_ip_update(-1)
            self.ip = 0
            self._stop.clear()

    def _reset_backend_neutral(self):
        backend = self.get_backend() or self.serial
        if backend is None or not getattr(backend, "connected", False):
            return
        if hasattr(backend, "reset_neutral"):
            backend.reset_neutral()
            return
        if hasattr(backend, "set_buttons"):
            backend.set_buttons([])

    def _build_default_registry(self):
        reg = {}

        # ---- pretty formatters
        def fmt_wait(c): return f"Wait {c.get('ms')} ms"
        def fmt_start_timing(c): return "Start Timing Reference"
        def fmt_wait_until(c): return f"Wait Until {c.get('ms')} ms elapsed"
        def fmt_get_elapsed(c): return f"Get Elapsed -> ${c.get('out', 'elapsed')}"
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
            return f"FindColor ({c.get('x')},{c.get('y')}) ~ {c.get('rgb')} ΔE≤{c.get('tol',10)} -> ${c.get('out')}"
        def fmt_find_area_color(c):
            x = c.get('x', 0)
            y = c.get('y', 0)
            w = c.get('width', 10)
            h = c.get('height', 10)
            return f"FindAreaColor ({x},{y}) {w}x{h} ~ {c.get('rgb')} ΔE≤{c.get('tol',10)} -> ${c.get('out')}"
        def fmt_wait_for_color(c):
            wait_str = "match" if c.get('wait_for', True) else "no match"
            interval = c.get('interval', 0.1)
            timeout = c.get('timeout', 0)
            timeout_str = f" timeout={timeout}s" if timeout > 0 else ""
            return f"WaitForColor ({c.get('x')},{c.get('y')}) ~ {c.get('rgb')} ΔE≤{c.get('tol',10)} until {wait_str} (check every {interval}s{timeout_str}) -> ${c.get('out')}"
        def fmt_wait_for_color_area(c):
            x = c.get('x', 0)
            y = c.get('y', 0)
            w = c.get('width', 10)
            h = c.get('height', 10)
            wait_str = "match" if c.get('wait_for', True) else "no match"
            interval = c.get('interval', 0.1)
            timeout = c.get('timeout', 0)
            timeout_str = f" timeout={timeout}s" if timeout > 0 else ""
            return f"WaitForAreaColor ({x},{y}) {w}x{h} ~ {c.get('rgb')} ΔE≤{c.get('tol',10)} until {wait_str} (check every {interval}s{timeout_str}) -> ${c.get('out')}"
        def fmt_comment(c): return f"// {c.get('text','')}"
        def fmt_run_python(c):
            out = c.get("out")
            args = c.get("args", [])
            a = "" if not args else f" args={args}"
            o = "" if not out else f" -> ${out}"
            return f"RunPython {c.get('file')}{a}{o}"
        def fmt_discord_status(c):
            msg = c.get("message", "")
            ping = bool(c.get("ping", False))
            image = c.get("image", "")
            flags = []
            if ping:
                flags.append("ping")
            if image:
                flags.append(f"image={image}")
            flag_str = f" ({', '.join(flags)})" if flags else ""
            return f"DiscordStatus {msg!r}{flag_str}"
        def fmt_play_sound(c):
            sound = c.get("sound", "")
            wait = bool(c.get("wait", False))
            volume = c.get("volume", 80)
            suffix = " (wait)" if wait else ""
            return f"PlaySound {sound!r} vol={volume}{suffix}"
        def fmt_save_frame(c):
            name = c.get("filename", "")
            out = (c.get("out") or "").strip()
            suffix = f" -> ${out}" if out else ""
            return f"SaveFrame {name!r}{suffix}"
        def fmt_tap_touch(c):
            return f"TapTouch x={c.get('x')} y={c.get('y')} down={c.get('down_time', 0.1)} settle={c.get('settle', 0.1)}"
        def fmt_set_left_stick(c):
            return f"Left Stick x={c.get('x')} y={c.get('y')}"
        def fmt_reset_left_stick(c):
            return "Left Stick Reset"
        def fmt_set_right_stick(c):
            return f"Right Stick x={c.get('x')} y={c.get('y')}"
        def fmt_reset_right_stick(c):
            return "Right Stick Reset"
        def fmt_press_ir(c):
            btns = c.get("buttons", [])
            return f"Press IR {', '.join(btns) if btns else '(none)'} for {c.get('ms')} ms"
        def fmt_hold_ir(c):
            btns = c.get("buttons", [])
            return f"Hold IR {', '.join(btns) if btns else '(none)'}"
        def fmt_press_interface(c):
            btns = c.get("buttons", [])
            return f"Press Interface {', '.join(btns) if btns else '(none)'} for {c.get('ms')} ms"
        def fmt_hold_interface(c):
            btns = c.get("buttons", [])
            return f"Hold Interface {', '.join(btns) if btns else '(none)'}"
        def fmt_mash(c):
            btns = c.get("buttons", [])
            until_ms = c.get("until_ms")
            hold = c.get("hold_ms", 25)
            wait = c.get("wait_ms", 25)
            if until_ms is not None:
                return f"Mash {', '.join(btns) if btns else '(none)'} until {until_ms} ms (hold:{hold} ms wait:{wait} ms)"
            duration = c.get("duration_ms", 1000)
            return f"Mash {', '.join(btns) if btns else '(none)'} for {duration} ms (hold:{hold} ms wait:{wait} ms)"

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

        def fmt_contains(c):
            needle = c.get("needle", "")
            haystack = c.get("haystack", "")
            out = c.get("out", "found")
            return f"Contains {needle!r} in {haystack!r} -> ${out}"

        def fmt_random(c):
            choices = c.get("choices", [])
            out = c.get("out", "random_value")
            return f"Random choice from {choices!r} -> ${out}"

        def fmt_random_range(c):
            min_val = c.get("min", 0)
            max_val = c.get("max", 100)
            integer = c.get("integer", False)
            out = c.get("out", "random_value")
            type_str = "int" if integer else "float"
            return f"Random {type_str} from {min_val} to {max_val} -> ${out}"

        def fmt_random_value(c):
            out = c.get("out", "random_value")
            return f"Random float [0.0, 1.0) -> ${out}"

        def fmt_export_json(c):
            vars_list = c.get("vars", [])
            filename = c.get("filename", "export.json")
            if vars_list:
                return f"ExportJSON {vars_list!r} -> {filename!r}"
            return f"ExportJSON (all vars) -> {filename!r}"

        def fmt_import_json(c):
            filename = c.get("filename", "import.json")
            return f"ImportJSON from {filename!r}"

        # ---- execution fns
        def cmd_wait(ctx, c):
            ms_raw = c.get("ms", 0)
            ms = float(resolve_number(ctx, ms_raw))

            # Use high-precision interruptible sleep
            precise_sleep_interruptible(ms / 1000.0, ctx["stop"])

        def cmd_start_timing(ctx, c):
            """Set timing reference point for cumulative timing commands."""
            ctx["_timing_reference"] = time.perf_counter()

        def cmd_wait_until(ctx, c):
            """Wait until specified ms have elapsed since start_timing.

            Uses hybrid approach for maximum precision:
            - Sleep until close to target (interruptible, low CPU)
            - Busy-wait for final approach (high precision, removes scheduler variance)

            Dynamic busy-wait threshold based on wait duration:
            - Short waits (<1s): 100ms threshold
            - Medium waits (1-10s): 500ms threshold
            - Long waits (>10s): 1000ms threshold (1 second)
            """
            ref = ctx.get("_timing_reference")
            target_ms = float(resolve_number(ctx, c.get("ms", 0)))

            if ref is None:
                # No reference set - fall back to regular wait
                precise_sleep_interruptible(target_ms / 1000.0, ctx["stop"])
                return

            target_sec = target_ms / 1000.0

            # Compensate for post-wait_until overhead (GUI callbacks, command dispatch)
            # This is the time between wait_until returning and the next command executing
            POST_CALLBACK_COMPENSATION = 0.003  # 3ms typical overhead

            target_time = ref + target_sec - POST_CALLBACK_COMPENSATION

            # Check if already past target
            now = time.perf_counter()
            if now >= target_time:
                overrun_ms = (now - target_time) * 1000
                print(f"[TIMING WARNING] wait_until {target_ms}ms: already {overrun_ms:.2f}ms past target")
                ctx["vars"]["_wait_until_actual_ms"] = (now - ref) * 1000
                return

            # Dynamic busy-wait threshold based on total wait duration
            # Longer waits need larger thresholds because Windows sleep() can be very
            # unpredictable - a single sleep call can take 50-200ms longer than requested
            #
            # For short waits (<1s): 100ms threshold
            # For medium waits (1-10s): 500ms threshold
            # For long waits (>10s): 1000ms threshold (1 second)
            total_wait = target_time - now
            if total_wait > 10.0:
                BUSY_WAIT_THRESHOLD = 1.000  # 1 second for long waits
            elif total_wait > 1.0:
                BUSY_WAIT_THRESHOLD = 0.500  # 500ms for medium waits
            else:
                BUSY_WAIT_THRESHOLD = 0.100  # 100ms for short waits

            busy_wait_start = target_time - BUSY_WAIT_THRESHOLD

            # Phase 1: Interruptible sleep until close to target
            # Check stop flag every 100ms to allow interruption
            stop_check_interval = 0.100  # Check stop every 100ms
            last_stop_check = now

            while True:
                now = time.perf_counter()
                remaining_to_busy = busy_wait_start - now

                if remaining_to_busy <= 0:
                    break  # Time to switch to busy-wait

                # Check stop flag periodically (not every iteration)
                if now - last_stop_check >= stop_check_interval:
                    if ctx["stop"].is_set():
                        return  # Interrupted
                    last_stop_check = now

                # Sleep in small chunks
                # Use conservative sleep to avoid overshooting
                sleep_chunk = min(remaining_to_busy, 0.010)  # 10ms chunks max
                if sleep_chunk > 0.005:
                    time.sleep(sleep_chunk * 0.5)  # Sleep for half the chunk
                else:
                    break  # Close enough, switch to busy-wait

            # Phase 2: Tight busy-wait for final precision
            # No stop checks here - we're in the critical timing window
            target = target_time  # Local variable for faster access
            perf_counter = time.perf_counter  # Local reference for speed
            while perf_counter() < target:
                pass  # Pure busy-wait, no overhead

        def cmd_get_elapsed(ctx, c):
            """Get milliseconds elapsed since start_timing."""
            ref = ctx.get("_timing_reference")
            out = c.get("out", "elapsed")

            if ref is None:
                ctx["vars"][out] = 0
                return

            elapsed_ms = (time.perf_counter() - ref) * 1000.0
            ctx["vars"][out] = elapsed_ms

        def cmd_press(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "press: buttons must be a list")
                return

            ms_raw = c.get("ms", 50)
            ms = float(resolve_number(ctx, ms_raw))

            # Pause keepalive loop to prevent threading conflicts
            if hasattr(backend, "pause_keepalive"):
                backend.pause_keepalive()

            try:
                use_timed_press = bool(getattr(backend, "supports_timed_press", False))
                if use_timed_press:
                    backend.press_buttons(buttons, ms)
                    if ms > 0:
                        interrupted = precise_sleep_interruptible(ms / 1000.0, ctx["stop"])
                        if interrupted:
                            backend.set_buttons([])
                    return

                # Press buttons with precise timing
                backend.set_buttons(buttons)
                if ms > 0:
                    # Use high-precision interruptible sleep
                    precise_sleep_interruptible(ms / 1000.0, ctx["stop"])

                # Release buttons
                backend.set_buttons([])
            finally:
                # Resume keepalive loop
                if hasattr(backend, "resume_keepalive"):
                    backend.resume_keepalive()


        def cmd_hold(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "hold: buttons must be a list")
                return

            # Pause keepalive loop to prevent threading conflicts
            if hasattr(backend, "pause_keepalive"):
                backend.pause_keepalive()

            try:
                backend.set_buttons(buttons)
            finally:
                # Resume keepalive loop
                if hasattr(backend, "resume_keepalive"):
                    backend.resume_keepalive()


        def cmd_label(ctx, c):
            pass

        def cmd_goto(ctx, c):
            label = c["label"]
            if label not in ctx["labels"]:
                messagebox.showerror("Label Error", f"Unknown label: {label}")
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
            sample_rgb = (int(r), int(g), int(b))
            target = (int(c["rgb"][0]), int(c["rgb"][1]), int(c["rgb"][2]))
            tol = float(c.get("tol", 10))

            # Use CIE76 Delta E for perceptually accurate color comparison
            delta_e = delta_e_cie76(sample_rgb, target)
            ok = delta_e <= tol
            ctx["vars"][out] = ok

        def cmd_find_area_color(ctx, c):
            """Find average color in an area and compare to target."""
            frame = ctx["get_frame"]()
            out = c.get("out", "match")

            if frame is None:
                ctx["vars"][out] = False
                return

            x = int(resolve_value(ctx, c.get("x", 0)))
            y = int(resolve_value(ctx, c.get("y", 0)))
            width = int(resolve_value(ctx, c.get("width", 10)))
            height = int(resolve_value(ctx, c.get("height", 10)))

            h_frame, w_frame, _ = frame.shape

            # Clamp region to frame bounds
            x = max(0, min(x, w_frame - 1))
            y = max(0, min(y, h_frame - 1))
            x2 = max(x + 1, min(x + width, w_frame))
            y2 = max(y + 1, min(y + height, h_frame))

            # Extract region (BGR)
            region_bgr = frame[y:y2, x:x2]

            # Calculate average color
            if region_bgr.size == 0:
                ctx["vars"][out] = False
                return

            # Compute mean color across all pixels in the region
            avg_b = float(np.mean(region_bgr[:, :, 0]))
            avg_g = float(np.mean(region_bgr[:, :, 1]))
            avg_r = float(np.mean(region_bgr[:, :, 2]))

            avg_rgb = (int(avg_r), int(avg_g), int(avg_b))
            target = (int(c["rgb"][0]), int(c["rgb"][1]), int(c["rgb"][2]))
            tol = float(c.get("tol", 10))

            # Use CIE76 Delta E for perceptually accurate color comparison
            delta_e = delta_e_cie76(avg_rgb, target)
            ok = delta_e <= tol
            ctx["vars"][out] = ok

        def cmd_wait_for_color(ctx, c):
            """Wait until pixel at (x,y) matches/doesn't match target color."""
            import time

            out = c.get("out", "match")
            x = int(resolve_value(ctx, c["x"]))
            y = int(resolve_value(ctx, c["y"]))
            target = (int(c["rgb"][0]), int(c["rgb"][1]), int(c["rgb"][2]))
            tol = float(c.get("tol", 10))
            interval = float(c.get("interval", 0.1))
            timeout = float(c.get("timeout", 0))  # 0 = no timeout
            wait_for = bool(c.get("wait_for", True))  # True = wait for match, False = wait for no match

            start_time = time.time()

            while True:
                # Check stop flag
                if ctx.get("stop_requested", False):
                    ctx["vars"][out] = False
                    return

                frame = ctx["get_frame"]()

                if frame is not None:
                    h, w, _ = frame.shape
                    if 0 <= x < w and 0 <= y < h:
                        b, g, r = frame[y, x].tolist()
                        sample_rgb = (int(r), int(g), int(b))

                        delta_e = delta_e_cie76(sample_rgb, target)
                        matches = delta_e <= tol

                        # Check if condition is met
                        if matches == wait_for:
                            ctx["vars"][out] = True
                            return

                # Check timeout
                if timeout > 0:
                    elapsed = time.time() - start_time
                    if elapsed >= timeout:
                        ctx["vars"][out] = False
                        return

                # Wait before next check
                time.sleep(interval)

        def cmd_wait_for_color_area(ctx, c):
            """Wait until average color in area matches/doesn't match target color."""
            import time

            out = c.get("out", "match")
            x = int(resolve_value(ctx, c.get("x", 0)))
            y = int(resolve_value(ctx, c.get("y", 0)))
            width = int(resolve_value(ctx, c.get("width", 10)))
            height = int(resolve_value(ctx, c.get("height", 10)))
            target = (int(c["rgb"][0]), int(c["rgb"][1]), int(c["rgb"][2]))
            tol = float(c.get("tol", 10))
            interval = float(c.get("interval", 0.1))
            timeout = float(c.get("timeout", 0))  # 0 = no timeout
            wait_for = bool(c.get("wait_for", True))  # True = wait for match, False = wait for no match

            start_time = time.time()

            while True:
                # Check stop flag
                if ctx.get("stop_requested", False):
                    ctx["vars"][out] = False
                    return

                frame = ctx["get_frame"]()

                if frame is not None:
                    h_frame, w_frame, _ = frame.shape

                    # Clamp region to frame bounds
                    x_clamped = max(0, min(x, w_frame - 1))
                    y_clamped = max(0, min(y, h_frame - 1))
                    x2 = max(x_clamped + 1, min(x_clamped + width, w_frame))
                    y2 = max(y_clamped + 1, min(y_clamped + height, h_frame))

                    # Extract region (BGR)
                    region_bgr = frame[y_clamped:y2, x_clamped:x2]

                    if region_bgr.size > 0:
                        # Calculate average color
                        avg_b = float(np.mean(region_bgr[:, :, 0]))
                        avg_g = float(np.mean(region_bgr[:, :, 1]))
                        avg_r = float(np.mean(region_bgr[:, :, 2]))

                        avg_rgb = (int(avg_r), int(avg_g), int(avg_b))

                        delta_e = delta_e_cie76(avg_rgb, target)
                        matches = delta_e <= tol

                        # Check if condition is met
                        if matches == wait_for:
                            ctx["vars"][out] = True
                            return

                # Check timeout
                if timeout > 0:
                    elapsed = time.time() - start_time
                    if elapsed >= timeout:
                        ctx["vars"][out] = False
                        return

                # Wait before next check
                time.sleep(interval)

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
                messagebox.showerror("error", f"read_text error: {e}")

        def cmd_comment(ctx, c):
            pass

        def cmd_run_python(ctx, c):
            # Check if Python is available (important for frozen exe builds)
            if not is_python_available():
                # Stop the script and notify that Python is needed
                ctx["stop"].set()
                ctx["on_python_needed"]()
                return

            file_name = str(resolve_value(ctx, c["file"]) or c["file"]).strip()
            if not file_name:
                messagebox.showerror("error", "run_python: file is empty")
                return

            if os.path.isabs(file_name):
                script_path = file_name
            else:
                script_path = os.path.join("py_scripts", file_name)

            if not os.path.exists(script_path):
                messagebox.showerror("error",f"run_python: file not found: {script_path}")
                return

            args = c.get("args", [])
            # Handle $variable as entire args value
            if isinstance(args, str) and args.strip().startswith("$"):
                args = resolve_value(ctx, args.strip())
            elif isinstance(args, str):
                args = ast.literal_eval(args)
            if args is None:
                args = []
            if not isinstance(args, list):
                messagebox.showerror("error", "run_python: args must be a JSON list")
                return

            # Resolve $var references inside list elements
            args = resolve_vars_deep(ctx, args)

            timeout_s = int(resolve_value(ctx, c.get("timeout_s", 10)) or 10)
            res = run_python_main(script_path, args, timeout_s=timeout_s)

            outvar = (c.get("out") or "").strip()
            if outvar:
                ctx["vars"][outvar] = res

        def cmd_discord_status(ctx, c):
            settings = ctx.get("get_settings", lambda: {})()
            discord_settings = settings.get("discord", {}) if isinstance(settings, dict) else {}
            webhook_url = (discord_settings.get("webhook_url") or "").strip()
            if not webhook_url:
                raise RuntimeError("discord_status: webhook URL is not configured in Settings.")

            user_id = str(discord_settings.get("user_id", "") or "").strip()
            if user_id and not user_id.isdigit():
                digits = "".join(ch for ch in user_id if ch.isdigit())
                if digits:
                    user_id = digits
            ping = bool(resolve_value(ctx, c.get("ping", False)))
            if ping and not user_id:
                raise RuntimeError("discord_status: ping requested but no Discord user ID is configured.")

            message_raw = c.get("message", "")
            message_val = resolve_value(ctx, message_raw)
            message = "" if message_val is None else str(message_val)

            image_raw = c.get("image", "")
            image_value = image_raw
            use_frame = False
            if isinstance(image_raw, str) and image_raw.strip() == "$frame":
                use_frame = True
            else:
                image_value = resolve_value(ctx, image_raw)
                if isinstance(image_value, str) and image_value.strip() == "$frame":
                    use_frame = True

            file_tuple = None
            if use_frame:
                frame = ctx["get_frame"]()
                if frame is None:
                    raise RuntimeError("discord_status: no camera frame available for $frame.")
                file_bytes = frame_to_png_bytes(frame)
                if not file_bytes:
                    raise RuntimeError("discord_status: failed to encode frame.")
                file_tuple = ("frame.png", file_bytes, "image/png")
            else:
                image_path = ""
                if image_value is not None:
                    image_path = str(image_value).strip()
                if image_path:
                    if not os.path.isabs(image_path):
                        image_path = exe_dir_path(image_path)
                    if not os.path.exists(image_path):
                        raise RuntimeError(f"discord_status: image not found: {image_path}")
                    with open(image_path, "rb") as f:
                        file_bytes = f.read()
                    mime = mimetypes.guess_type(image_path)[0] or "application/octet-stream"
                    file_tuple = (os.path.basename(image_path), file_bytes, mime)

            payload = {}
            content = message.strip()
            if ping:
                mention = f"<@{user_id}>"
                content = f"{mention} {content}".strip()
                payload["allowed_mentions"] = {"users": [str(user_id)]}
            else:
                payload["allowed_mentions"] = {"parse": []}

            if not content and not file_tuple:
                raise RuntimeError("discord_status: message is empty and no image provided.")

            if content:
                payload["content"] = content

            if file_tuple:
                payload["embeds"] = [{"image": {"url": f"attachment://{file_tuple[0]}"}}]

            send_discord_webhook(webhook_url, payload, file_tuple=file_tuple)
            self.status_cb("Discord status sent.")

        def cmd_play_sound(ctx, c):
            sound_raw = c.get("sound", default_sound)
            sound_val = resolve_value(ctx, sound_raw)
            sound_name = "" if sound_val is None else str(sound_val).strip()
            if not sound_name:
                sound_name = default_sound
            wait = bool(resolve_value(ctx, c.get("wait", False)))
            volume_raw = c.get("volume", 80)
            volume_val = resolve_number(ctx, volume_raw)

            ok, msg = play_sound_file(sound_name, volume=volume_val, wait=wait, stop_event=ctx.get("stop"))
            if not ok:
                messagebox.showerror("Command Error", msg)

        def cmd_save_frame(ctx, c):
            frame = ctx["get_frame"]()
            if frame is None:
                raise RuntimeError("save_frame: no camera frame available.")

            out_dir = exe_dir_path("saved_images")
            os.makedirs(out_dir, exist_ok=True)

            filename_raw = c.get("filename", "")
            filename_val = resolve_value(ctx, filename_raw)
            filename = "" if filename_val is None else str(filename_val).strip()
            if filename:
                name = os.path.basename(filename)
                if not name:
                    name = ""
                if name and not os.path.splitext(name)[1]:
                    name = f"{name}.png"
                if not name:
                    filename = ""
                else:
                    filename = name

            if filename:
                out_path = os.path.join(out_dir, filename)
            else:
                ts = time.strftime("%Y%m%d_%H%M%S")
                ms = int((time.time() % 1) * 1000)
                out_path = os.path.join(out_dir, f"frame_{ts}_{ms:03d}.png")

            if os.path.exists(out_path):
                base, ext = os.path.splitext(out_path)
                suffix = 1
                while True:
                    candidate = f"{base}_{suffix}{ext}"
                    if not os.path.exists(candidate):
                        out_path = candidate
                        break
                    suffix += 1

            rgb = frame[:, :, ::-1]
            img = Image.fromarray(rgb)
            img.save(out_path, format="PNG")

            outvar = (c.get("out") or "").strip()
            if outvar:
                ctx["vars"][outvar] = out_path

            self.status_cb(f"Saved frame: {out_path}")
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

        def cmd_set_left_stick(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_left_stick"):
                raise RuntimeError("set_left_stick is not supported by this backend.")

            x = float(resolve_number(ctx, c.get("x", 0.0)))
            y = float(resolve_number(ctx, c.get("y", 0.0)))
            backend.set_left_stick(x, y)

        def cmd_reset_left_stick(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if hasattr(backend, "reset_left_stick"):
                backend.reset_left_stick()
                return
            if hasattr(backend, "set_left_stick"):
                backend.set_left_stick(0.0, 0.0)
                return
            raise RuntimeError("reset_left_stick is not supported by this backend.")

        def cmd_set_right_stick(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_right_stick"):
                raise RuntimeError("set_right_stick is not supported by this backend.")

            x = float(resolve_number(ctx, c.get("x", 0.0)))
            y = float(resolve_number(ctx, c.get("y", 0.0)))
            backend.set_right_stick(x, y)

        def cmd_reset_right_stick(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if hasattr(backend, "reset_right_stick"):
                backend.reset_right_stick()
                return
            if hasattr(backend, "set_right_stick"):
                backend.set_right_stick(0.0, 0.0)
                return
            raise RuntimeError("reset_right_stick is not supported by this backend.")

        def cmd_press_ir(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_ir_buttons"):
                raise RuntimeError("press_ir is only supported by the 3DS backend.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "press_ir: buttons must be a list")
                return

            backend.set_ir_buttons(buttons)

            ms_raw = c.get("ms", 50)
            ms = float(resolve_number(ctx, ms_raw))
            if ms > 0:
                precise_sleep_interruptible(ms / 1000.0, ctx["stop"])

            backend.set_ir_buttons([])

        def cmd_hold_ir(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_ir_buttons"):
                raise RuntimeError("hold_ir is only supported by the 3DS backend.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "hold_ir: buttons must be a list")
                return

            backend.set_ir_buttons(buttons)

        def cmd_press_interface(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_interface_buttons"):
                raise RuntimeError("press_interface is only supported by the 3DS backend.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "press_interface: buttons must be a list")
                return

            backend.set_interface_buttons(buttons)

            ms_raw = c.get("ms", 50)
            ms = float(resolve_number(ctx, ms_raw))
            if ms > 0:
                precise_sleep_interruptible(ms / 1000.0, ctx["stop"])

            backend.set_interface_buttons([])

        def cmd_hold_interface(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")
            if not hasattr(backend, "set_interface_buttons"):
                raise RuntimeError("hold_interface is only supported by the 3DS backend.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "hold_interface: buttons must be a list")
                return

            backend.set_interface_buttons(buttons)

        def cmd_mash(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                messagebox.showerror("Command Error", "mash: buttons must be a list")
                return

            hold_ms = float(resolve_number(ctx, c.get("hold_ms", 25)))
            wait_ms = float(resolve_number(ctx, c.get("wait_ms", 25)))
            use_timed_press = bool(getattr(backend, "supports_timed_press", False))

            # Convert to seconds for precise timing
            hold_sec = hold_ms / 1000.0
            wait_sec = wait_ms / 1000.0

            # Calculate end time - support reference-based timing with until_ms
            until_ms = c.get("until_ms")
            ref = ctx.get("_timing_reference")

            if until_ms is not None and ref is not None:
                # Reference-based timing: mash until X ms from start_timing
                target_sec = float(resolve_number(ctx, until_ms)) / 1000.0
                end_time = ref + target_sec
            else:
                # Original duration-based timing
                duration_ms = float(resolve_number(ctx, c.get("duration_ms", 1000)))
                total_duration = duration_ms / 1000.0
                end_time = time.perf_counter() + total_duration

            # Pause keepalive loop to prevent threading conflicts
            if hasattr(backend, "pause_keepalive"):
                backend.pause_keepalive()

            try:
                # Main mashing loop with precise timing
                # For reference-based timing, we need to be precise about when we stop
                cycle_time = hold_sec + wait_sec

                while True:
                    if ctx["stop"].is_set():
                        break

                    time_remaining = end_time - time.perf_counter()

                    # Exit if we don't have enough time for even a minimal press
                    # (need at least hold_sec to do a meaningful button press)
                    if time_remaining < hold_sec:
                        break

                    # Press buttons
                    if use_timed_press:
                        backend.press_buttons(buttons, hold_ms)
                    else:
                        backend.set_buttons(buttons)

                    # Hold for precise duration, but truncate if needed
                    actual_hold = min(hold_sec, time_remaining)
                    if precise_sleep_interruptible(actual_hold, ctx["stop"]):
                        break  # Interrupted

                    if not use_timed_press:
                        # Release buttons
                        backend.set_buttons([])

                    # Recalculate remaining time after hold
                    time_remaining = end_time - time.perf_counter()
                    if time_remaining <= 0:
                        break

                    # Wait for precise duration, but not longer than remaining time
                    wait_duration = min(wait_sec, time_remaining)
                    if precise_sleep_interruptible(wait_duration, ctx["stop"]):
                        break  # Interrupted

                # Ensure buttons are released at the end
                backend.set_buttons([])
            finally:
                # Resume keepalive loop
                if hasattr(backend, "resume_keepalive"):
                    backend.resume_keepalive()

        def cmd_contains(ctx, c):
            """
            Check if a value is contained in another value (like Python's 'in' operator).
            Works with strings (substring check) and lists (membership check).
            Stores boolean result in the output variable.
            """
            needle = resolve_value(ctx, c.get("needle"))
            haystack = resolve_value(ctx, c.get("haystack"))
            out = c.get("out", "found")

            try:
                result = needle in haystack
            except TypeError:
                # If 'in' operator fails (e.g., incompatible types), return False
                result = False

            ctx["vars"][out] = result

        def cmd_random(ctx, c):
            """
            Randomly select a value from the provided choices list.
            Stores the selected value in the output variable.
            """
            choices_raw = c.get("choices", [])

            # Resolve variables in the choices list
            if isinstance(choices_raw, str) and choices_raw.startswith("$"):
                # If choices is a variable reference, resolve it
                choices = resolve_value(ctx, choices_raw)
            else:
                # Otherwise, resolve each element in the list
                choices = resolve_vars_deep(ctx, choices_raw)

            out = c.get("out", "random_value")

            # Validate that choices is a list
            if not isinstance(choices, list):
                messagebox.showerror("Command Error", "random: choices must be a list")
                return

            # Validate that choices is not empty
            if len(choices) == 0:
                messagebox.showerror("Command Error", "random: choices list cannot be empty")
                return

            # Select a random choice
            selected = random.choice(choices)
            ctx["vars"][out] = selected

        def cmd_random_range(ctx, c):
            """
            Generate a random number between min and max (inclusive).
            If integer is True, returns an integer; otherwise returns a float.
            """
            min_raw = c.get("min", 0)
            max_raw = c.get("max", 100)
            integer = bool(resolve_value(ctx, c.get("integer", False)))
            out = c.get("out", "random_value")

            # Resolve min and max values (support variables and expressions)
            min_val = resolve_number(ctx, min_raw)
            max_val = resolve_number(ctx, max_raw)

            # Validate range
            if min_val > max_val:
                messagebox.showerror("Command Error", f"random_range: min ({min_val}) cannot be greater than max ({max_val})")
                return

            # Generate random value
            if integer:
                # For integers, use randint (inclusive on both ends)
                selected = random.randint(int(min_val), int(max_val))
            else:
                # For floats, use uniform
                selected = random.uniform(min_val, max_val)

            ctx["vars"][out] = selected

        def cmd_random_value(ctx, c):
            """
            Generate a random float between 0.0 and 1.0 (exclusive of 1.0).
            """
            out = c.get("out", "random_value")
            selected = random.random()
            ctx["vars"][out] = selected

        def cmd_export_json(ctx, c):
            """
            Export specified variables to a JSON file.
            If vars list is empty or not specified, exports all variables.
            """
            vars_to_export = c.get("vars", [])
            filename = resolve_value(ctx, c.get("filename", "export.json"))

            # If vars_to_export is a string (variable reference), resolve it
            if isinstance(vars_to_export, str) and vars_to_export.startswith("$"):
                vars_to_export = resolve_value(ctx, vars_to_export)

            if not isinstance(vars_to_export, list):
                vars_to_export = []

            # Validate that all items in vars list are strings
            if vars_to_export:
                invalid_items = [v for v in vars_to_export if not isinstance(v, str)]
                if invalid_items:
                    messagebox.showerror(
                        "Export Error",
                        f"Invalid vars list: all items must be strings (variable names).\n"
                        f"Invalid items: {invalid_items}\n\n"
                        f"Example: [\"var1\", \"var2\"] is valid\n"
                        f"[var1, var2] is not valid"
                    )
                    return

            # Build the data to export
            if vars_to_export:
                # Export only specified variables
                data = {}
                for var_name in vars_to_export:
                    if var_name in ctx["vars"]:
                        data[var_name] = ctx["vars"][var_name]
            else:
                # Export all variables
                data = dict(ctx["vars"])

            # Write to file
            try:
                with open(filename, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            except Exception as e:
                messagebox.showerror("Export Error", f"Failed to export JSON: {e}")

        def cmd_import_json(ctx, c):
            """
            Import variables from a JSON file.
            The JSON file should contain an object with key-value pairs.
            """
            filename = resolve_value(ctx, c.get("filename", "import.json"))

            if not os.path.exists(filename):
                messagebox.showerror("Import Error", f"File not found: {filename}")
                return

            try:
                with open(filename, "r", encoding="utf-8") as f:
                    data = json.load(f)

                if not isinstance(data, dict):
                    messagebox.showerror("Import Error", "JSON file must contain an object (dictionary)")
                    return

                # Merge imported variables into context
                for key, value in data.items():
                    ctx["vars"][key] = value
            except json.JSONDecodeError as e:
                messagebox.showerror("Import Error", f"Invalid JSON: {e}")
            except Exception as e:
                messagebox.showerror("Import Error", f"Failed to import JSON: {e}")

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

            # Pause keepalive loop to prevent threading conflicts
            if hasattr(backend, "pause_keepalive"):
                backend.pause_keepalive()

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
                # Resume keepalive loop
                if hasattr(backend, "resume_keepalive"):
                    backend.resume_keepalive()


        cond_schema = [
            {"key": "left", "type": "str", "default": "$flag", "help": "Left operand (literal or $var)"},
            {"key": "op", "type": "choice", "choices": ["==", "!=", "<", "<=", ">", ">="], "default": "==", "help": "Comparison operator"},
            {"key": "right", "type": "json", "default": True, "help": "Right operand (literal or $var)"},
        ]

        sound_choices = list_sound_files()
        if not sound_choices:
            sound_choices = ["alert1.mp3"]
        default_sound = sound_choices[0]

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
                "start_timing", [], cmd_start_timing,
                doc="Set timing reference point for cumulative timing commands (wait_until, mash until_ms).",
                arg_schema=[],
                format_fn=fmt_start_timing,
                group="Timing",
                order=5,
                exportable=True
            ),
            CommandSpec(
                "wait_until", ["ms"], cmd_wait_until,
                doc="Wait until specified ms have elapsed since start_timing. Compensates for all overhead.",
                arg_schema=[
                    {"key": "ms", "type": "json", "default": 1000, "help": "Target elapsed time in milliseconds from start_timing"}
                ],
                format_fn=fmt_wait_until,
                group="Timing",
                order=15,
                exportable=True
            ),
            CommandSpec(
                "get_elapsed", ["out"], cmd_get_elapsed,
                doc="Store milliseconds elapsed since start_timing in a variable.",
                arg_schema=[
                    {"key": "out", "type": "str", "default": "elapsed", "help": "Variable name to store result (no $)"}
                ],
                format_fn=fmt_get_elapsed,
                group="Timing",
                order=20,
                exportable=True
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
                "mash", ["buttons"], cmd_mash,
                doc="Rapidly mash buttons for a duration. Use until_ms with start_timing for reference-based timing. Default: ~20 presses/second.",
                arg_schema=[
                    {"key": "buttons", "type": "buttons", "default": ["A"], "help": "Buttons to mash"},
                    {"key": "duration_ms", "type": "json", "default": 1000, "help": "Total mashing duration in milliseconds (ignored if until_ms is set)"},
                    {"key": "until_ms", "type": "json", "default": None, "help": "Target elapsed time from start_timing (overrides duration_ms)"},
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
                "contains", ["needle", "haystack", "out"], cmd_contains,
                doc="Check if needle is contained in haystack (like Python's 'in'). Works with strings (substring) and lists (membership). Stores boolean in $out.",
                arg_schema=[
                    {"key": "needle", "type": "json", "default": "abc", "help": "Value to search for (literal or $var)"},
                    {"key": "haystack", "type": "json", "default": "abcdef", "help": "Container to search in (literal or $var)"},
                    {"key": "out", "type": "str", "default": "found", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_contains,
                group="Variables",
                order=30,
                exportable=True
            ),
            CommandSpec(
                "random", ["choices", "out"], cmd_random,
                doc="Randomly select one value from a list of choices. Stores selected value in $out.",
                arg_schema=[
                    {"key": "choices", "type": "json", "default": [1, 2, 3, 4, 5], "help": "List of values to choose from (literal list or $var)"},
                    {"key": "out", "type": "str", "default": "random_value", "help": "Variable name to store selected value (no $)"},
                ],
                format_fn=fmt_random,
                group="Variables",
                order=40,
                exportable=True
            ),
            CommandSpec(
                "random_range", ["min", "max", "out"], cmd_random_range,
                doc="Generate a random number between min and max (inclusive). Use integer=true for whole numbers.",
                arg_schema=[
                    {"key": "min", "type": "json", "default": 0, "help": "Minimum value (inclusive, supports $var and =expr)"},
                    {"key": "max", "type": "json", "default": 100, "help": "Maximum value (inclusive, supports $var and =expr)"},
                    {"key": "integer", "type": "bool", "default": False, "help": "If true, returns integer; if false, returns float"},
                    {"key": "out", "type": "str", "default": "random_value", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_random_range,
                group="Variables",
                order=41,
                exportable=True
            ),
            CommandSpec(
                "random_value", ["out"], cmd_random_value,
                doc="Generate a random float between 0.0 and 1.0 (exclusive of 1.0).",
                arg_schema=[
                    {"key": "out", "type": "str", "default": "random_value", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_random_value,
                group="Variables",
                order=42,
                exportable=True
            ),
            CommandSpec(
                "export_json", ["filename"], cmd_export_json,
                doc="Export variables to a JSON file. If vars is empty, exports all variables.",
                arg_schema=[
                    {"key": "filename", "type": "str", "default": "export.json", "help": "Output filename (supports $var)"},
                    {"key": "vars", "type": "json", "default": [], "help": "List of variable names (no $) to export (empty = all)"},
                ],
                format_fn=fmt_export_json,
                group="Variables",
                order=50,
                exportable=False,
                export_note="File I/O not supported in export"
            ),
            CommandSpec(
                "import_json", ["filename"], cmd_import_json,
                doc="Import variables from a JSON file. The file must contain a JSON object.",
                arg_schema=[
                    {"key": "filename", "type": "str", "default": "import.json", "help": "Input filename (supports $var)"},
                ],
                format_fn=fmt_import_json,
                group="Variables",
                order=51,
                exportable=False,
                export_note="File I/O not supported in export"
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
                doc="Sample pixel at (x,y) and compare to rgb using CIE76 Delta E (perceptual). Stores bool in $out.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate"},
                    {"key": "rgb", "type": "rgb", "default": [255, 0, 0], "help": "Target RGB as [R,G,B]"},
                    {"key": "tol", "type": "float", "default": 10, "help": "Delta E tolerance (0-1: imperceptible, 2-10: noticeable, 10+: obvious)"},
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
                "find_area_color", ["x", "y", "width", "height", "rgb", "out"], cmd_find_area_color,
                doc="Calculate average color in an area and compare to target rgb using CIE76 Delta E. Stores bool in $out.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate (top-left corner)"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate (top-left corner)"},
                    {"key": "width", "type": "int", "default": 10, "help": "Width of region"},
                    {"key": "height", "type": "int", "default": 10, "help": "Height of region"},
                    {"key": "rgb", "type": "rgb", "default": [255, 0, 0], "help": "Target RGB as [R,G,B]"},
                    {"key": "tol", "type": "float", "default": 10, "help": "Delta E tolerance (0-1: imperceptible, 2-10: noticeable, 10+: obvious)"},
                    {"key": "out", "type": "str", "default": "match", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_find_area_color,
                group="Image",
                order=11,
                test=True,
                exportable=False,
                export_note="Requires camera frame processing which is unsupported at the moment."
            ),
            CommandSpec(
                "wait_for_color", ["x", "y", "rgb"], cmd_wait_for_color,
                doc="Wait until pixel at (x,y) matches/doesn't match target color. Polls at regular intervals until condition is met or timeout.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate"},
                    {"key": "rgb", "type": "rgb", "default": [255, 0, 0], "help": "Target RGB as [R,G,B]"},
                    {"key": "tol", "type": "float", "default": 10, "help": "Delta E tolerance (0-1: imperceptible, 2-10: noticeable, 10+: obvious)"},
                    {"key": "interval", "type": "float", "default": 0.1, "help": "Check interval in seconds"},
                    {"key": "timeout", "type": "float", "default": 0, "help": "Timeout in seconds (0 = no timeout)"},
                    {"key": "wait_for", "type": "bool", "default": True, "help": "True = wait for match, False = wait for no match"},
                    {"key": "out", "type": "str", "default": "match", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_wait_for_color,
                group="Image",
                order=12,
                test=True,
                exportable=False,
                export_note="Requires camera frame processing which is unsupported at the moment."
            ),
            CommandSpec(
                "wait_for_color_area", ["x", "y", "width", "height", "rgb"], cmd_wait_for_color_area,
                doc="Wait until average color in area matches/doesn't match target color. Polls at regular intervals until condition is met or timeout.",
                arg_schema=[
                    {"key": "x", "type": "int", "default": 0, "help": "X coordinate (top-left corner)"},
                    {"key": "y", "type": "int", "default": 0, "help": "Y coordinate (top-left corner)"},
                    {"key": "width", "type": "int", "default": 10, "help": "Width of region"},
                    {"key": "height", "type": "int", "default": 10, "help": "Height of region"},
                    {"key": "rgb", "type": "rgb", "default": [255, 0, 0], "help": "Target RGB as [R,G,B]"},
                    {"key": "tol", "type": "float", "default": 10, "help": "Delta E tolerance (0-1: imperceptible, 2-10: noticeable, 10+: obvious)"},
                    {"key": "interval", "type": "float", "default": 0.1, "help": "Check interval in seconds"},
                    {"key": "timeout", "type": "float", "default": 0, "help": "Timeout in seconds (0 = no timeout)"},
                    {"key": "wait_for", "type": "bool", "default": True, "help": "True = wait for match, False = wait for no match"},
                    {"key": "out", "type": "str", "default": "match", "help": "Variable name to store result (no $)"},
                ],
                format_fn=fmt_wait_for_color_area,
                group="Image",
                order=13,
                test=True,
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
                "save_frame",
                [],
                cmd_save_frame,
                doc="Save the current camera frame to ./saved_images.",
                arg_schema=[
                    {"key": "filename", "type": "str", "default": "", "help": "Optional filename (saved under saved_images)."},
                    {"key": "out", "type": "str", "default": "", "help": "Variable name to store saved path (no $)."},
                ],
                format_fn=fmt_save_frame,
                group="Image",
                order=30,
                exportable=False,
                export_note="Requires camera frame and filesystem access."
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
                "discord_status",
                [],
                cmd_discord_status,
                doc="Send a Discord webhook status update (optional ping/image). Uses webhook URL and user ID from Settings.",
                arg_schema=[
                    {"key": "message", "type": "str", "default": "", "help": "Message text (supports $var)."},
                    {"key": "ping", "type": "bool", "default": False, "help": "Ping the configured user ID."},
                    {"key": "image", "type": "str", "default": "", "help": "Image file path or $frame."},
                ],
                format_fn=fmt_discord_status,
                group="Custom",
                order=20,
                exportable=False,
                export_note="Uses Discord webhooks and local files, unsupported in standalone export."
            ),
            CommandSpec(
                "play_sound",
                [],
                cmd_play_sound,
                doc="Play an alert sound from bin/sounds using ffplay at the specified volume.",
                arg_schema=[
                    {"key": "sound", "type": "choice", "choices": sound_choices, "default": default_sound, "help": "Sound filename in bin/sounds"},
                    {"key": "volume", "type": "volume", "default": 80, "help": "Volume 0-100"},
                    {"key": "wait", "type": "bool", "default": False, "help": "Wait until the sound finishes playing"},
                ],
                format_fn=fmt_play_sound,
                group="Custom",
                order=30,
                test=True,
                exportable=False,
                export_note="Audio playback is not supported in standalone export."
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
                order=10,
                exportable=False,
                export_note="3DS-specific command not supported in standalone export."
            ),
            CommandSpec(
                "set_left_stick",
                ["x", "y"],
                cmd_set_left_stick,
                doc="Set left stick position until changed or reset (3DS/PABotBase only).",
                arg_schema=[
                    {"key":"x","type":"json","default":0.0, "help":"X axis (-1.0..1.0)"},
                    {"key":"y","type":"json","default":0.0, "help":"Y axis (-1.0..1.0)"},
                ],
                format_fn=fmt_set_left_stick,
                group="Controller",
                order=20,
                exportable=False,
                export_note="Not supported in standalone export."
            ),
            CommandSpec(
                "reset_left_stick",
                [],
                cmd_reset_left_stick,
                doc="Reset left stick to center (3DS/PABotBase only).",
                arg_schema=[],
                format_fn=fmt_reset_left_stick,
                group="Controller",
                order=21,
                exportable=False,
                export_note="Not supported in standalone export."
            ),
            CommandSpec(
                "set_right_stick",
                ["x", "y"],
                cmd_set_right_stick,
                doc="Set right stick position until changed or reset (3DS/PABotBase only).",
                arg_schema=[
                    {"key":"x","type":"json","default":0.0, "help":"X axis (-1.0..1.0)"},
                    {"key":"y","type":"json","default":0.0, "help":"Y axis (-1.0..1.0)"},
                ],
                format_fn=fmt_set_right_stick,
                group="Controller",
                order=22,
                exportable=False,
                export_note="Not supported in standalone export."
            ),
            CommandSpec(
                "reset_right_stick",
                [],
                cmd_reset_right_stick,
                doc="Reset right stick to center (3DS/PABotBase only).",
                arg_schema=[],
                format_fn=fmt_reset_right_stick,
                group="Controller",
                order=23,
                exportable=False,
                export_note="Not supported in standalone export."
            ),
            CommandSpec(
                "press_ir",
                ["buttons", "ms"],
                cmd_press_ir,
                doc="3DS only: press ZL/ZR for ms, then release.",
                arg_schema=[
                    {"key":"buttons","type":"buttons","choices":["ZL", "ZR"], "default":["ZL"], "help":"IR buttons to press"},
                    {"key":"ms","type":"json","default":80, "help":"Hold duration in milliseconds"},
                ],
                format_fn=fmt_press_ir,
                group="3DS",
                order=30,
                exportable=False,
                export_note="3DS-specific command not supported in standalone export."
            ),
            CommandSpec(
                "hold_ir",
                ["buttons"],
                cmd_hold_ir,
                doc="3DS only: hold ZL/ZR until changed by another command.",
                arg_schema=[
                    {"key":"buttons","type":"buttons","choices":["ZL", "ZR"], "default":[], "help":"IR buttons to hold"},
                ],
                format_fn=fmt_hold_ir,
                group="3DS",
                order=31,
                exportable=False,
                export_note="3DS-specific command not supported in standalone export."
            ),
            CommandSpec(
                "press_interface",
                ["buttons", "ms"],
                cmd_press_interface,
                doc="3DS only: press interface buttons (Home/Power). PowerLong triggers power dialog.",
                arg_schema=[
                    {"key":"buttons","type":"buttons","choices":["Home", "Power", "PowerLong"], "default":["Home"], "help":"Interface buttons to press"},
                    {"key":"ms","type":"json","default":80, "help":"Hold duration in milliseconds"},
                ],
                format_fn=fmt_press_interface,
                group="3DS",
                order=40,
                exportable=False,
                export_note="3DS-specific command not supported in standalone export."
            ),
            CommandSpec(
                "hold_interface",
                ["buttons"],
                cmd_hold_interface,
                doc="3DS only: hold interface buttons (Home/Power) until changed. PowerLong triggers power dialog.",
                arg_schema=[
                    {"key":"buttons","type":"buttons","choices":["Home", "Power", "PowerLong"], "default":[], "help":"Interface buttons to hold"},
                ],
                format_fn=fmt_hold_interface,
                group="3DS",
                order=41,
                exportable=False,
                export_note="3DS-specific command not supported in standalone export."
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
