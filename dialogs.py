"""
Dialog classes for Controller Macro Runner.
Contains the CommandEditorDialog for editing script commands.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import json

import SerialController
from utils import list_python_files


class CommandEditorDialog(tk.Toplevel):
    """Dialog for editing script commands with schema-driven fields."""

    def __init__(self, parent, registry, initial_cmd=None, title="Edit Command", test_callback=None):
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
