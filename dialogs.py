"""
Dialog classes for Controller Macro Runner.

SettingsDialog: Keybindings and 3DS Input Redirection configuration.
CommandEditorDialog: Schema-driven editor for script commands.
"""
import tkinter as tk
from tkinter import ttk, messagebox
import json

import SerialController
from utils import list_python_files, get_default_keybindings


class SettingsDialog(tk.Toplevel):
    """Dialog for editing application settings including keybinds and 3DS IP."""

    def __init__(self, parent, keybindings: dict, threeds_ip: str, threeds_port: int,
                 on_save_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.result = None
        self.on_save_callback = on_save_callback

        # Store initial values
        self.keybindings = keybindings.copy()
        self.threeds_ip = threeds_ip
        self.threeds_port = threeds_port
        self._rebinding_target = None

        self.title("Settings")
        self.transient(parent)
        self.grab_set()

        # Main container
        main = ttk.Frame(self, padding=10)
        main.grid(row=0, column=0, sticky="nsew")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # Notebook for tabs
        self.notebook = ttk.Notebook(main)
        self.notebook.grid(row=0, column=0, sticky="nsew")
        main.columnconfigure(0, weight=1)
        main.rowconfigure(0, weight=1)

        # Create tabs
        self._create_keybinds_tab()
        self._create_threeds_tab()

        # Bottom buttons
        btn_frame = ttk.Frame(main)
        btn_frame.grid(row=1, column=0, sticky="ew", pady=(10, 0))
        ttk.Button(btn_frame, text="Cancel", command=self._cancel).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Save", command=self._save).pack(side="right")

        # Position dialog
        self.update_idletasks()
        x = parent.winfo_rootx() + 100
        y = parent.winfo_rooty() + 50
        self.geometry(f"+{x}+{y}")
        self.minsize(450, 400)

    def _create_keybinds_tab(self):
        """Create the keybindings tab."""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="Keybinds")

        # Info label
        info = ttk.Label(tab, text="Keyboard control is active when connected to a backend.\n"
                                   "Click Rebind, then press a key. Press Esc to cancel.")
        info.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 10))

        # Treeview for keybindings
        tree_frame = ttk.Frame(tab)
        tree_frame.grid(row=1, column=0, columnspan=4, sticky="nsew")
        tab.rowconfigure(1, weight=1)
        tab.columnconfigure(0, weight=1)

        self.keybinds_tree = ttk.Treeview(tree_frame, columns=("button", "key"),
                                          show="headings", height=12)
        self.keybinds_tree.heading("button", text="Controller Input")
        self.keybinds_tree.heading("key", text="Key")
        self.keybinds_tree.column("button", width=160, anchor="w")
        self.keybinds_tree.column("key", width=160, anchor="w")
        self.keybinds_tree.pack(side="left", fill="both", expand=True)

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical",
                                  command=self.keybinds_tree.yview)
        self.keybinds_tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")

        # Status label
        self.keybind_status_var = tk.StringVar(value="Select an input and click Rebind.")
        ttk.Label(tab, textvariable=self.keybind_status_var, foreground="gray").grid(
            row=2, column=0, columnspan=4, sticky="w", pady=(10, 0))

        # Buttons
        btn_row = ttk.Frame(tab)
        btn_row.grid(row=3, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Button(btn_row, text="Rebind...", command=self._start_rebind).pack(side="left")
        ttk.Button(btn_row, text="Clear", command=self._clear_binding).pack(side="left", padx=(6, 0))
        ttk.Button(btn_row, text="Restore Defaults", command=self._restore_default_keybinds).pack(
            side="left", padx=(6, 0))

        # Populate tree
        self._refresh_keybinds_tree()

        # Bind key press for rebinding
        self.bind("<KeyPress>", self._on_rebind_key)

    def _create_threeds_tab(self):
        """Create the 3DS settings tab."""
        tab = ttk.Frame(self.notebook, padding=10)
        self.notebook.add(tab, text="3DS Input Redirection")

        # Info label
        info = ttk.Label(tab, text="Configure the 3DS IP address for Input Redirection.\n"
                                   "The 3DS must be running InputRedirectionNTR or similar.")
        info.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))

        # IP address
        ttk.Label(tab, text="3DS IP Address:").grid(row=1, column=0, sticky="w", pady=5)
        self.ip_var = tk.StringVar(value=self.threeds_ip)
        ip_entry = ttk.Entry(tab, textvariable=self.ip_var, width=20)
        ip_entry.grid(row=1, column=1, sticky="w", padx=(10, 0), pady=5)

        # Port
        ttk.Label(tab, text="Port:").grid(row=2, column=0, sticky="w", pady=5)
        self.port_var = tk.StringVar(value=str(self.threeds_port))
        port_entry = ttk.Entry(tab, textvariable=self.port_var, width=10)
        port_entry.grid(row=2, column=1, sticky="w", padx=(10, 0), pady=5)

        # Default port hint
        ttk.Label(tab, text="Default port is 4950", foreground="gray").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(0, 10))

    def _refresh_keybinds_tree(self):
        """Refresh the keybindings treeview."""
        self.keybinds_tree.delete(*self.keybinds_tree.get_children())

        # Build inverse mapping: button -> list of keys
        inv = {b: [] for b in SerialController.KEYBIND_TARGETS}
        for key, btn in self.keybindings.items():
            if btn in inv:
                inv[btn].append(key)

        # Insert rows
        for btn in SerialController.KEYBIND_TARGETS:
            keys = ", ".join(sorted(inv[btn])) if inv[btn] else ""
            self.keybinds_tree.insert("", "end", values=(btn, keys))

    def _get_selected_button(self):
        """Get the currently selected controller button."""
        sel = self.keybinds_tree.selection()
        if not sel:
            return None
        vals = self.keybinds_tree.item(sel[0], "values")
        return vals[0] if vals else None

    def _start_rebind(self):
        """Start rebinding a key."""
        btn = self._get_selected_button()
        if not btn:
            messagebox.showinfo("Rebind", "Select a controller input first.", parent=self)
            return
        self._rebinding_target = btn
        self.keybind_status_var.set(f"Press a key to bind to {btn} (Esc cancels)...")

    def _on_rebind_key(self, event):
        """Handle key press during rebinding."""
        if self._rebinding_target is None:
            return

        ks = (event.keysym or "").lower()
        if not ks:
            return

        if ks == "escape":
            self._rebinding_target = None
            self.keybind_status_var.set("Rebind cancelled.")
            return

        if ks == "return":
            ks = "enter"

        # Assign the key to the button
        self.keybindings[ks] = self._rebinding_target
        self.keybind_status_var.set(f"Bound '{ks}' to {self._rebinding_target}")
        self._rebinding_target = None
        self._refresh_keybinds_tree()

    def _clear_binding(self):
        """Clear all keybindings for the selected button."""
        btn = self._get_selected_button()
        if not btn:
            return

        # Remove all keys that map to this button
        keys_to_remove = [k for k, b in self.keybindings.items() if b == btn]
        for k in keys_to_remove:
            del self.keybindings[k]

        self._refresh_keybinds_tree()
        self.keybind_status_var.set(f"Cleared bindings for {btn}")

    def _restore_default_keybinds(self):
        """Restore default keybindings."""
        self.keybindings = get_default_keybindings()
        self._refresh_keybinds_tree()
        self.keybind_status_var.set("Restored default keybindings")

    def _save(self):
        """Save settings and close dialog."""
        # Validate 3DS settings
        ip = self.ip_var.get().strip()
        if not ip:
            messagebox.showerror("Invalid IP", "Please enter a valid IP address.", parent=self)
            self.notebook.select(1)  # Switch to 3DS tab
            return

        try:
            port = int(self.port_var.get().strip())
            if port < 1 or port > 65535:
                raise ValueError("Port out of range")
        except ValueError:
            messagebox.showerror("Invalid Port", "Port must be a number between 1 and 65535.",
                                 parent=self)
            self.notebook.select(1)  # Switch to 3DS tab
            return

        # Build result
        self.result = {
            "keybindings": self.keybindings.copy(),
            "threeds": {
                "ip": ip,
                "port": port,
            }
        }

        # Call save callback if provided
        if self.on_save_callback:
            self.on_save_callback(self.result)

        self.destroy()

    def _cancel(self):
        """Cancel and close dialog."""
        self.result = None
        self.destroy()


class CommandEditorDialog(tk.Toplevel):
    """Dialog for editing script commands with schema-driven fields."""

    def __init__(self, parent, registry, initial_cmd=None, title="Edit Command",
                 test_callback=None, select_area_callback=None, select_color_callback=None,
                 select_area_color_callback=None):
        super().__init__(parent)
        self.parent = parent
        self.registry = registry
        self.result = None
        self.test_callback = test_callback
        self.select_area_callback = select_area_callback  # Callback for region selection
        self.select_color_callback = select_color_callback  # Callback for color picker
        self.select_area_color_callback = select_area_color_callback  # Callback for area+color picker

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
            # Find first non-header command (prefer "press" if available)
            if "press" in keys:
                self.cmd_name_var.set("press")
            else:
                # Find first non-header item
                for k in keys:
                    if not k.startswith("───"):
                        self.cmd_name_var.set(k)
                        break

        self._render_fields(initial_cmd=initial_cmd)

        self.update_idletasks()
        x = parent.winfo_rootx() + 80
        y = parent.winfo_rooty() + 80
        self.geometry(f"+{x}+{y}")

    def _ordered_command_names(self):
        """Build a list of command names with category headers."""
        # Sort commands by group, then order, then name
        def keyfn(name):
            s = self.registry[name]
            return (s.group, s.order, s.name)

        sorted_names = sorted(self.registry.keys(), key=keyfn)

        # Group commands by category
        result = []
        current_group = None

        for name in sorted_names:
            spec = self.registry[name]
            if spec.group != current_group:
                # Add category header (prefixed with special marker)
                result.append(f"─── {spec.group} ───")
                current_group = spec.group
            result.append(name)

        return result

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

        # Skip category headers (they start with "───")
        if name.startswith("───"):
            # Find first command after this header
            all_names = self._ordered_command_names()
            try:
                idx = all_names.index(name)
                # Find next non-header item
                for i in range(idx + 1, len(all_names)):
                    if not all_names[i].startswith("───"):
                        self.cmd_name_var.set(all_names[i])
                        self._render_fields(initial_cmd)
                        return
            except (ValueError, IndexError):
                pass
            return

        spec = self.registry.get(name)
        if not spec:
            return
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

                    # Create a frame to hold entry and color preview
                    rgb_frame = ttk.Frame(self.fields_frame)
                    rgb_frame.grid(row=r, column=1, sticky="ew", pady=3)

                    ent = ttk.Entry(rgb_frame, textvariable=var, width=20)
                    ent.pack(side="left", fill="x", expand=True)

                    # Color preview swatch
                    swatch = tk.Canvas(rgb_frame, width=30, height=22, bg="#808080",
                                       highlightthickness=1, highlightbackground="#555")
                    swatch.pack(side="left", padx=(6, 0))

                    # Store swatch reference for updates
                    self.widgets[f"{key}_swatch"] = swatch

                    # Update swatch when value changes
                    def update_swatch(var=var, swatch=swatch):
                        try:
                            text = var.get().strip()
                            if text.startswith("["):
                                rgb = json.loads(text)
                            else:
                                parts = [p.strip() for p in text.split(",")]
                                rgb = [int(p) for p in parts]
                            if len(rgb) == 3:
                                r_val = max(0, min(255, int(rgb[0])))
                                g_val = max(0, min(255, int(rgb[1])))
                                b_val = max(0, min(255, int(rgb[2])))
                                hex_color = f"#{r_val:02x}{g_val:02x}{b_val:02x}"
                                swatch.configure(bg=hex_color)
                                return
                        except Exception:
                            pass
                        swatch.configure(bg="#808080")

                    # Initial update
                    update_swatch()
                    # Trace changes
                    var.trace_add("write", lambda *args, fn=update_swatch: fn())

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

                    choices = field.get("choices")
                    if not isinstance(choices, list) or not choices:
                        choices = SerialController.ALL_BUTTONS

                    for b in choices:
                        lb.insert("end", b)

                    init_buttons = init_val if isinstance(init_val, list) else []
                    for i, b in enumerate(choices):
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

        # Add "Select Area" button for read_text command
        if name == "read_text" and self.select_area_callback:
            # Add button row after all fields
            next_row = len(spec.arg_schema)
            select_frame = ttk.Frame(self.fields_frame)
            select_frame.grid(row=next_row, column=0, columnspan=3, sticky="ew", pady=(10, 0))

            ttk.Button(
                select_frame, text="Select Area on Camera",
                command=self._open_region_selector
            ).pack(side="left")

            ttk.Label(
                select_frame,
                text="Click to visually select region on camera feed",
                foreground="gray"
            ).pack(side="left", padx=(10, 0))

        # Add "Pick Color from Camera" button for find_color command
        if name == "find_color" and self.select_color_callback:
            next_row = len(spec.arg_schema)
            picker_frame = ttk.Frame(self.fields_frame)
            picker_frame.grid(row=next_row, column=0, columnspan=3, sticky="ew", pady=(10, 0))

            ttk.Button(
                picker_frame, text="Pick Color from Camera",
                command=self._open_color_picker
            ).pack(side="left")

            ttk.Label(
                picker_frame,
                text="Click to pick a color and position from camera feed",
                foreground="gray"
            ).pack(side="left", padx=(10, 0))

        # Add "Select Area & Color from Camera" button for find_area_color command
        if name == "find_area_color" and self.select_area_color_callback:
            next_row = len(spec.arg_schema)
            area_color_frame = ttk.Frame(self.fields_frame)
            area_color_frame.grid(row=next_row, column=0, columnspan=3, sticky="ew", pady=(10, 0))

            ttk.Button(
                area_color_frame, text="Select Area & Color from Camera",
                command=self._open_area_color_picker
            ).pack(side="left")

            ttk.Label(
                area_color_frame,
                text="Click to visually select area and target color from camera feed",
                foreground="gray"
            ).pack(side="left", padx=(10, 0))

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
        if not name or name.startswith("───"):
            messagebox.showerror("Invalid selection", "Please select a command (not a category header).", parent=self)
            return
        spec = self.registry.get(name)
        if not spec:
            messagebox.showerror("Invalid command", f"Command '{name}' not found in registry.", parent=self)
            return

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
        if not name or name.startswith("───"):
            messagebox.showerror("Test", "Please select a command (not a category header).", parent=self)
            return
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

    def _open_region_selector(self):
        """Open the region selector for read_text command."""
        if not self.select_area_callback:
            return

        # Get current values to use as initial selection
        initial_region = None
        try:
            x = int(self.field_vars.get("x", tk.StringVar(value="0")).get())
            y = int(self.field_vars.get("y", tk.StringVar(value="0")).get())
            w = int(self.field_vars.get("width", tk.StringVar(value="100")).get())
            h = int(self.field_vars.get("height", tk.StringVar(value="20")).get())
            if w > 0 and h > 0:
                initial_region = (x, y, w, h)
        except (ValueError, TypeError):
            pass

        # Release grab temporarily so the selector window can work
        self.grab_release()

        # Call the callback to open the selector
        # It returns True if selector was opened, False otherwise
        # Pass _restore_grab as on_close callback to ensure grab is restored when selector closes
        opened = self.select_area_callback(initial_region, self._on_region_selected, self._restore_grab)
        if not opened:
            # Selector wasn't opened (e.g., camera not running), restore grab
            self._restore_grab()

    def _on_region_selected(self, x, y, width, height):
        """Callback when a region is selected."""
        # Update the field values
        if "x" in self.field_vars:
            self.field_vars["x"].set(str(x))
        if "y" in self.field_vars:
            self.field_vars["y"].set(str(y))
        if "width" in self.field_vars:
            self.field_vars["width"].set(str(width))
        if "height" in self.field_vars:
            self.field_vars["height"].set(str(height))

        # Re-grab focus
        self.grab_set()
        self.focus_set()

    def _restore_grab(self):
        """Restore grab on this dialog after picker/selector closes."""
        try:
            if self.winfo_exists():
                self.grab_set()
                self.focus_set()
        except Exception:
            pass

    def _open_color_picker(self):
        """Open the color picker for find_color command."""
        if not self.select_color_callback:
            return

        # Get current values to use as initial selection
        initial_x = None
        initial_y = None
        try:
            x_var = self.field_vars.get("x")
            y_var = self.field_vars.get("y")
            if x_var and y_var:
                initial_x = int(x_var.get())
                initial_y = int(y_var.get())
        except (ValueError, TypeError):
            pass

        # Release grab temporarily so the picker window can work
        self.grab_release()

        # Call the callback to open the picker
        # It returns True if picker was opened, False otherwise
        # Pass _restore_grab as on_close callback to ensure grab is restored when picker closes
        opened = self.select_color_callback(initial_x, initial_y, self._on_color_selected, self._restore_grab)
        if not opened:
            # Picker wasn't opened (e.g., camera not running), restore grab
            self._restore_grab()

    def _on_color_selected(self, x, y, r, g, b):
        """Callback when a color is selected from camera."""
        # Update the field values
        if "x" in self.field_vars:
            self.field_vars["x"].set(str(x))
        if "y" in self.field_vars:
            self.field_vars["y"].set(str(y))
        if "rgb" in self.field_vars:
            self.field_vars["rgb"].set(f"{r},{g},{b}")

        # Re-grab focus
        self.grab_set()
        self.focus_set()

    def _open_area_color_picker(self):
        """Open the area color picker for find_area_color command."""
        if not self.select_area_color_callback:
            return

        # Get current values to use as initial selection
        initial_region = None
        initial_rgb = None
        try:
            x = int(self.field_vars.get("x", tk.StringVar(value="0")).get())
            y = int(self.field_vars.get("y", tk.StringVar(value="0")).get())
            w = int(self.field_vars.get("width", tk.StringVar(value="10")).get())
            h = int(self.field_vars.get("height", tk.StringVar(value="10")).get())
            if w > 0 and h > 0:
                initial_region = (x, y, w, h)

            # Get RGB if available
            rgb_var = self.field_vars.get("rgb")
            if rgb_var:
                rgb_text = rgb_var.get().strip()
                if rgb_text.startswith("["):
                    initial_rgb = json.loads(rgb_text)
                else:
                    parts = [p.strip() for p in rgb_text.split(",")]
                    initial_rgb = [int(parts[0]), int(parts[1]), int(parts[2])]
        except (ValueError, TypeError, IndexError):
            pass

        # Release grab temporarily so the picker window can work
        self.grab_release()

        # Call the callback to open the picker
        # It returns True if picker was opened, False otherwise
        # Pass _restore_grab as on_close callback to ensure grab is restored when picker closes
        opened = self.select_area_color_callback(initial_region, initial_rgb, self._on_area_color_selected, self._restore_grab)
        if not opened:
            # Picker wasn't opened (e.g., camera not running), restore grab
            self._restore_grab()

    def _on_area_color_selected(self, x, y, width, height, r, g, b):
        """Callback when an area and color are selected."""
        # Update the field values
        if "x" in self.field_vars:
            self.field_vars["x"].set(str(x))
        if "y" in self.field_vars:
            self.field_vars["y"].set(str(y))
        if "width" in self.field_vars:
            self.field_vars["width"].set(str(width))
        if "height" in self.field_vars:
            self.field_vars["height"].set(str(height))
        if "rgb" in self.field_vars:
            self.field_vars["rgb"].set(f"{r},{g},{b}")

        # Re-grab focus
        self.grab_set()
        self.focus_set()
