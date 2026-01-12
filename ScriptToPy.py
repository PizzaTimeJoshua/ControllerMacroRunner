import re
from tkinter import filedialog, messagebox
import os

def _py_ident(name: str) -> str:
    """
    Convert script var name to a safe Python identifier.
    """
    s = re.sub(r"\W+", "_", str(name).strip())
    if not s:
        s = "var"
    if s[0].isdigit():
        s = "_" + s
    return s

def _py_literal(v):
    """
    Convert JSON-ish values into Python literals.
    """
    # repr() is fine for numbers/bools/None/strings/lists/dicts
    return repr(v)

def _is_frame_payload(obj) -> bool:
    """
    Detect special camera frame placeholder in args.
    """
    return obj == "$frame"  # your chosen token

def export_script_to_python(self):
    if self.engine.running:
        messagebox.showwarning("Export", "Stop the script before exporting.")
        return
    if not self.engine.commands:
        messagebox.showinfo("Export", "No script loaded.")
        return

    commands = self.engine.commands

    # -------------------------
    # 1) Validate exportability
    # -------------------------
    incompatible = []  # list of (index, cmd, reason)

    # Structured export cannot represent goto/label nicely.
    # Vision/camera commands are not included in export runtime.
    disallow_cmds = {"find_color", "read_text", "tap_touch", "type_name", "label", "goto"}

    for i, c in enumerate(commands):
        cmd = c.get("cmd")

        if cmd in disallow_cmds:
            incompatible.append((i, cmd, "Not supported for structured Python export."))

        # Disallow camera frame payload usage
        for k, v in c.items():
            if isinstance(v, str) and _is_frame_payload(v):
                incompatible.append((i, cmd, "Uses $frame (camera). Export does not include camera runtime."))
            if isinstance(v, list) and any(_is_frame_payload(x) for x in v):
                incompatible.append((i, cmd, "Uses $frame (camera). Export does not include camera runtime."))

        # Unknown commands
        if cmd not in self.engine.registry:
            incompatible.append((i, cmd, "Unknown command (not registered)."))

    if incompatible:
        msg = "Cannot export due to incompatible commands:\n\n"
        for i, cmd, reason in incompatible[:30]:
            msg += f"Line {i}: {cmd} — {reason}\n"
        if len(incompatible) > 30:
            msg += f"\n...and {len(incompatible) - 30} more."
        messagebox.showerror("Export failed", msg)
        return

    # ------------------------------------------
    # 2) Build variable name mapping ($var -> py)
    # ------------------------------------------
    used_vars = set()

    for c in commands:
        cmd = c.get("cmd")
        if cmd in ("set", "add"):
            used_vars.add(c.get("var"))

        if cmd in ("if", "while"):
            for side in ("left", "right"):
                v = c.get(side)
                if isinstance(v, str) and v.startswith("$"):
                    used_vars.add(v[1:])

        if cmd == "contains":
            for key in ("needle", "haystack"):
                v = c.get(key)
                if isinstance(v, str) and v.startswith("$"):
                    used_vars.add(v[1:])
            outv = (c.get("out") or "").strip()
            if outv:
                used_vars.add(outv)

        if cmd == "random":
            # Check if choices is a variable reference
            choices = c.get("choices")
            if isinstance(choices, str) and choices.startswith("$"):
                used_vars.add(choices[1:])
            # Also check for variables inside the choices list
            elif isinstance(choices, list):
                for item in choices:
                    if isinstance(item, str) and item.startswith("$"):
                        used_vars.add(item[1:])
            outv = (c.get("out") or "").strip()
            if outv:
                used_vars.add(outv)

        if cmd == "run_python":
            args = c.get("args", [])
            if isinstance(args, list):
                for a in args:
                    if isinstance(a, str) and a.startswith("$"):
                        used_vars.add(a[1:])

            outv = (c.get("out") or "").strip()
            if outv:
                used_vars.add(outv)

    used_vars = {v for v in used_vars if v}

    var_map = {}
    taken = set()
    for v in sorted(used_vars):
        base = _py_ident(v)
        name = base
        n = 2
        while name in taken:
            name = f"{base}_{n}"
            n += 1
        var_map[v] = name
        taken.add(name)

    def op_to_py(x):
        if isinstance(x, str) and x.startswith("$"):
            vn = x[1:]
            return var_map.get(vn, _py_ident(vn))
        if isinstance(x, str) and x.startswith("="):
            vn = x.replace("=","").replace("$","")
            return str(vn)
        return _py_literal(x)

    def args_to_py(args):
        out = []
        for a in (args or []):
            if isinstance(a, str) and a.startswith("$"):
                vn = a[1:]
                out.append(var_map.get(vn, _py_ident(vn)))
            else:
                out.append(_py_literal(a))
        return "[" + ", ".join(out) + "]"

    # -------------------------
    # 3) Ask output location
    # -------------------------
    default_name = (os.path.splitext(os.path.basename(self.script_path))[0] if self.script_path else "exported_macro") + ".py"
    out_path = filedialog.asksaveasfilename(
        title="Export Script to Python",
        defaultextension=".py",
        initialfile=default_name,
        filetypes=[("Python files", "*.py"), ("All files", "*.*")]
    )
    if not out_path:
        return

    # -------------------------
    # 4) Generate macro body
    # -------------------------
    body_lines = []
    indent = 0

    def emit(s=""):
        body_lines.append(("    " * indent) + s)

    # Variables (native Python vars)
    if var_map:
        emit("# Variables")
        for script_var, py_var in var_map.items():
            emit(f"{py_var} = 0  # from ${script_var}")
        emit("")

    for i, c in enumerate(commands):
        cmd = c.get("cmd")

        # closing blocks reduce indent first
        if cmd in ("end_if", "end_while"):
            indent = max(0, indent - 1)

        if cmd == "comment":
            emit(f"# {c.get('text','')}".rstrip())

        elif cmd == "wait":
            raw = str(c.get("ms", 0))
            if raw.strip().startswith("="):
                ms = str(raw.replace("=","").replace("$",""))
            else:
                ms = float(raw)
            emit(f"wait_with_keepalive({ms}/1000.0)")

        elif cmd == "hold":
            btns = c.get("buttons", [])
            emit(f"send_buttons({btns!r})")

        elif cmd == "press":
            btns = c.get("buttons", [])
            raw = str(c.get("ms", 0))
            if raw.strip().startswith("="):
                ms = str(raw.replace("=","").replace("$",""))
            else:
                ms = float(raw)
            emit(f"press({btns!r}, {ms})")

        elif cmd == "mash":
            btns = c.get("buttons", [])
            duration_ms = float(c.get("duration_ms", 1000))
            hold_ms = float(c.get("hold_ms", 25))
            wait_ms = float(c.get("wait_ms", 25))
            emit(f"mash({btns!r}, {duration_ms}, {hold_ms}, {wait_ms})")

        elif cmd == "set":
            sv = c.get("var")
            pyv = var_map.get(sv, _py_ident(sv))
            emit(f"{pyv} = {op_to_py(c.get('value'))}")

        elif cmd == "add":
            sv = c.get("var")
            pyv = var_map.get(sv, _py_ident(sv))
            emit(f"{pyv} += {op_to_py(c.get('value', 0))}")

        elif cmd == "contains":
            needle = op_to_py(c.get("needle"))
            haystack = op_to_py(c.get("haystack"))
            outv = (c.get("out") or "found").strip()
            pyv = var_map.get(outv, _py_ident(outv))
            emit(f"try:")
            emit(f"    {pyv} = {needle} in {haystack}")
            emit(f"except TypeError:")
            emit(f"    {pyv} = False")

        elif cmd == "random":
            choices = op_to_py(c.get("choices"))
            outv = (c.get("out") or "random_value").strip()
            pyv = var_map.get(outv, _py_ident(outv))
            emit(f"{pyv} = random.choice({choices})")

        elif cmd == "if":
            left = op_to_py(c.get("left"))
            op = c.get("op")
            right = op_to_py(c.get("right"))
            emit(f"if {left} {op} {right}:")
            indent += 1

        elif cmd == "end_if":
            pass

        elif cmd == "while":
            left = op_to_py(c.get("left"))
            op = c.get("op")
            right = op_to_py(c.get("right"))
            emit(f"while {left} {op} {right}:")
            indent += 1

        elif cmd == "end_while":
            pass

        elif cmd == "run_python":
            file_ = c.get("file")
            args = c.get("args", [])
            outv = (c.get("out") or "").strip()
            timeout_s = int(c.get("timeout_s", 10))
            emit(f"res = run_python_main({file_!r}, {args_to_py(args)}, timeout_s={timeout_s})")
            if outv:
                pyv = var_map.get(outv, _py_ident(outv))
                emit(f"{pyv} = res")

        else:
            emit(f"# Unsupported command at line {i}: {cmd}")

        emit("")

    uses_run_python = any(c.get("cmd") == "run_python" for c in commands)
    uses_random = any(c.get("cmd") == "random" for c in commands)

    # -------------------------
    # 5) Exported file header
    # -------------------------
    #
    default_port = "COM4"
    try:
        if hasattr(self, "com_var"):
            pv = (self.com_var.get() or "").strip()
            if pv:
                default_port = pv
    except Exception:
        pass

    exported = []
    exported.append("import time")
    exported.append("import serial")
    exported.append("import math")
    if uses_random:
        exported.append("import random")
    exported.append("")
    exported.append("# =========================")
    exported.append("# User-editable settings")
    exported.append("# =========================")
    exported.append(f"PORT = \"{default_port}\"")
    exported.append("BAUD = 1_000_000")
    exported.append("KEEPALIVE_INTERVAL_S = 0.1")
    exported.append("LONG_WAIT_THRESHOLD_S = 0.5")
    exported.append("")
    exported.append("# Button mapping")
    exported.append("BUTTON_MAP = {")
    exported.append("    'L': ('high', 0x01), 'R': ('high', 0x02), 'X': ('high', 0x04), 'Y': ('high', 0x08),")
    exported.append("    'A': ('low', 0x01), 'B': ('low', 0x02), 'Right': ('low', 0x04), 'Left': ('low', 0x08),")
    exported.append("    'Up': ('low', 0x10), 'Down': ('low', 0x20), 'Select': ('low', 0x40), 'Start': ('low', 0x80),")
    exported.append("}")
    exported.append("")
    exported.append("def buttons_to_bytes(buttons):")
    exported.append("    high = 0")
    exported.append("    low = 0")
    exported.append("    for b in buttons:")
    exported.append("        which, mask = BUTTON_MAP[b]")
    exported.append("        if which == 'high':")
    exported.append("            high |= mask")
    exported.append("        else:")
    exported.append("            low |= mask")
    exported.append("    return high & 0xFF, low & 0xFF")
    exported.append("")
    if uses_run_python:
        exported.append("import sys")
        exported.append("import json")
        exported.append("import subprocess")
        exported.append("")
        exported.append("def run_python_main(script_path, args, timeout_s=10):")
        exported.append("    runner = r'''")
        exported.append("import json, sys, importlib.util, traceback")
        exported.append("")
        exported.append("def load_module_from_path(path):")
        exported.append("    spec = importlib.util.spec_from_file_location('user_module', path)")
        exported.append("    if spec is None or spec.loader is None:")
        exported.append("        raise RuntimeError('Could not load module: ' + path)")
        exported.append("    mod = importlib.util.module_from_spec(spec)")
        exported.append("    spec.loader.exec_module(mod)")
        exported.append("    return mod")
        exported.append("")
        exported.append("def main():")
        exported.append("    path = sys.argv[1]")
        exported.append("    args = json.loads(sys.argv[2]) if len(sys.argv) > 2 else []")
        exported.append("    mod = load_module_from_path(path)")
        exported.append("    if not hasattr(mod, 'main'):")
        exported.append("        raise RuntimeError('Script does not define main(...)')")
        exported.append("    res = mod.main(*args)")
        exported.append("    print(json.dumps(res, ensure_ascii=False))")
        exported.append("")
        exported.append("if __name__ == '__main__':")
        exported.append("    try:")
        exported.append("        main()")
        exported.append("    except Exception:")
        exported.append("        traceback.print_exc()")
        exported.append("        sys.exit(1)")
        exported.append("'''")
        exported.append("    cp = subprocess.run([sys.executable, '-c', runner, script_path, json.dumps(args, ensure_ascii=False)],")
        exported.append("                        capture_output=True, text=True, timeout=timeout_s)")
        exported.append("    if cp.returncode != 0:")
        exported.append("        raise RuntimeError((cp.stderr or cp.stdout or '').strip())")
        exported.append("    out = (cp.stdout or '').strip()")
        exported.append("    return json.loads(out) if out else None")
        exported.append("")

    # Main
    exported.append("def main():")
    if uses_random:
        exported.append("    # Initialize random seed to current time")
        exported.append("    random.seed(time.time())")
        exported.append("")
    exported.append("    ser = serial.Serial(PORT, BAUD, timeout=1)")
    exported.append("    current_buttons = []  # held buttons")
    exported.append("")
    exported.append("    def send_buttons(buttons):")
    exported.append("        nonlocal current_buttons")
    exported.append("        current_buttons = list(buttons)")
    exported.append("        high, low = buttons_to_bytes(current_buttons)")
    exported.append("        ser.write(bytes([0x54, high, low]))")
    exported.append("")
    exported.append("    def wait_with_keepalive(seconds):")
    exported.append("        if seconds <= 0:")
    exported.append("            return")
    exported.append("        if seconds <= LONG_WAIT_THRESHOLD_S:")
    exported.append("            time.sleep(seconds)")
    exported.append("            return")
    exported.append("        end_t = time.monotonic() + seconds")
    exported.append("        next_send = time.monotonic()")
    exported.append("        while True:")
    exported.append("            now = time.monotonic()")
    exported.append("            if now >= end_t:")
    exported.append("                break")
    exported.append("            if now >= next_send:")
    exported.append("                high, low = buttons_to_bytes(current_buttons)")
    exported.append("                ser.write(bytes([0x54, high, low]))")
    exported.append("                next_send += KEEPALIVE_INTERVAL_S")
    exported.append("            sleep_for = min(0.01, end_t - now, max(0.0, next_send - now))")
    exported.append("            if sleep_for > 0:")
    exported.append("                time.sleep(sleep_for)")
    exported.append("")
    exported.append("    def press(buttons, ms):")
    exported.append("        send_buttons(buttons)")
    exported.append("        wait_with_keepalive(ms/1000.0)")
    exported.append("        send_buttons([])")
    exported.append("")
    exported.append("    def mash(buttons, duration_ms, hold_ms=25, wait_ms=25):")
    exported.append("        end_time = time.perf_counter() + duration_ms / 1000.0")
    exported.append("        hold_sec = hold_ms / 1000.0")
    exported.append("        wait_sec = wait_ms / 1000.0")
    exported.append("        while time.perf_counter() < end_time:")
    exported.append("            send_buttons(buttons)")
    exported.append("            time.sleep(hold_sec)")
    exported.append("            send_buttons([])")
    exported.append("            time_remaining = end_time - time.perf_counter()")
    exported.append("            if time_remaining <= 0:")
    exported.append("                break")
    exported.append("            wait_duration = min(wait_sec, time_remaining)")
    exported.append("            time.sleep(wait_duration)")
    exported.append("        send_buttons([])")
    exported.append("")
    exported.append("    try:")
    exported.append("        # Pairing warm-up (~3 seconds neutral)")
    exported.append("        send_buttons([])")
    exported.append("        wait_with_keepalive(3.0)")
    exported.append("")
    for l in body_lines:
        exported.append("        " + l if l else "")
    exported.append("    finally:")
    exported.append("        try:")
    exported.append("            send_buttons([])")
    exported.append("        except Exception:")
    exported.append("            pass")
    exported.append("        try:")
    exported.append("            ser.close()")
    exported.append("        except Exception:")
    exported.append("            pass")
    exported.append("")
    exported.append("if __name__ == '__main__':")
    exported.append("    main()")
    exported.append("")

    # -------------------------
    # 6) Write file
    # -------------------------
    try:
        with open(out_path, "w", encoding="utf-8") as f:
            f.write("\n".join(exported))
        self.set_status(f"Exported to Python: {out_path}")
        messagebox.showinfo("Export", f"Exported Python file:\n{out_path}")
    except Exception as e:
        messagebox.showerror("Export error", str(e))
