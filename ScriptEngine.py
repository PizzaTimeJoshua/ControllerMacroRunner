import time
import base64
from PIL import Image
import numpy as np
import threading
import json
import os
import subprocess
import sys
# ----------------------------
# Script Command Spec
# ----------------------------

class CommandSpec:
    def __init__(self, name, required_keys, fn, doc="", arg_schema=None, format_fn=None,
                 group="Other", order=999,test=False):
        self.name = name
        self.required_keys = required_keys
        self.fn = fn
        self.doc = doc
        self.arg_schema = arg_schema or []
        self.format_fn = format_fn or (lambda c: f"{name} " + " ".join(f"{k}={c.get(k)!r}" for k in c if k != "cmd"))
        self.group = group
        self.order = order
        self.test = test



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


        # ---- execution fns
        def cmd_wait(ctx, c):
            ms = int(resolve_value(ctx, c["ms"]))
            end_t = time.monotonic() + ms / 1000.0
            while time.monotonic() < end_t:
                if ctx["stop"].is_set():
                    break
                time.sleep(0.01)

        def cmd_press(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                raise ValueError("press: buttons must be a list")

            # Resolve variable refs in buttons list if you support it; otherwise leave:
            backend.set_buttons(buttons)

            ms = int(resolve_value(ctx, c.get("ms", 60)))
            if ms > 0:
                time.sleep(ms / 1000.0)

            # release
            backend.set_buttons([])


        def cmd_hold(ctx, c):
            backend = ctx["get_backend"]()
            if backend is None or not getattr(backend, "connected", False):
                raise RuntimeError("No output backend connected.")

            buttons = c.get("buttons", [])
            if not isinstance(buttons, list):
                raise ValueError("hold: buttons must be a list")

            backend.set_buttons(buttons)


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
                order=10,
                test = True
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

        ]

        for s in specs:
            reg[s.name] = s
        return reg

