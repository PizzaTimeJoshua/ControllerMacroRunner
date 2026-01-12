# CLAUDE.md - AI Assistant Guide for Controller Macro Runner

## Project Overview

**Controller Macro Runner** is a Windows desktop application for automating game controller inputs through macro scripting. It enables users to:

- Capture and display live video from DirectShow cameras (via FFmpeg)
- Send game controller button presses continuously over USB serial connections
- Create, edit, and execute macro scripts with variables and control flow (if/while loops)
- Extend functionality with custom Python tools
- Support multiple output backends (USB serial wireless transmitter or 3DS input redirection)

**Target Use Cases:**
- Gaming automation for consoles (Nintendo 3DS, NES Classic, etc.)
- Repetitive task automation requiring controller inputs
- Vision-based automation (pixel detection and response)

**Status:** Currently in active development

---

## Repository Structure

```
ControllerMacroRunner/
├── main.py                 # Primary Tkinter GUI application (2,039 lines)
├── ScriptEngine.py         # Script execution engine with command registry (821 lines)
├── SerialController.py     # USB serial communication (119 lines)
├── ThreeDSClasses.py       # 3DS Input Redirection backend (180 lines)
├── ScriptToPy.py           # Script-to-Python exporter (392 lines)
├── bin/                    # FFmpeg binaries (bundled)
│   ├── ffmpeg.exe
│   └── *.dll
├── py_scripts/             # User-provided Python scripts (auto-created)
│   ├── example_frame_input.py
│   └── example_math.py
├── scripts/                # Macro script files in JSON format (auto-created)
├── .git/                   # Git repository
├── .gitignore
├── .gitattributes
├── README.md               # User-facing documentation
└── CLAUDE.md               # This file - AI assistant guide
```

### File Purposes

- **main.py**: Tkinter GUI with camera management, backend control, script editor UI, and event handling
- **ScriptEngine.py**: Core execution engine with command registry, variable resolution, control flow, and extensible command system
- **SerialController.py**: USB serial backend for wireless controller transmitter (1,000,000 baud)
- **ThreeDSClasses.py**: Alternative backend for 3DS Input Redirection via UDP
- **ScriptToPy.py**: Exports compatible scripts to standalone Python files

---

## Technology Stack

| Component | Technology | Purpose |
|-----------|-----------|---------|
| Language | Python 3.10+ | Core application |
| GUI Framework | Tkinter + ttk | Cross-platform GUI |
| Image Processing | Pillow, NumPy | Frame handling and pixel operations |
| Serial Communication | pyserial | Controller transmission |
| Video Capture | FFmpeg (DirectShow) | Windows camera capture |
| Networking | Socket (UDP) | 3DS communication |

**Dependencies:**
```bash
py -m pip install numpy pillow pyserial
```

**FFmpeg:** Place `ffmpeg.exe` + DLLs in `./bin/` or ensure on PATH

---

## Architecture and Design Patterns

### Architectural Overview

```
┌─────────────────────────────────────────────────────┐
│                  main.py (App)                      │
│  ┌──────────────┐  ┌──────────────┐  ┌───────────┐ │
│  │Camera Thread │  │Script Thread │  │GUI Thread │ │
│  └──────────────┘  └──────────────┘  └───────────┘ │
└─────────────────────────────────────────────────────┘
                      ▼           ▼
         ┌────────────────────────────────┐
         │   ScriptEngine (Command Registry)  │
         └────────────────────────────────┘
                      ▼           ▼
    ┌─────────────────────┐   ┌──────────────────┐
    │ SerialController    │   │ ThreeDSBackend   │
    │ (USB Wireless TX)   │   │ (3DS UDP)        │
    └─────────────────────┘   └──────────────────┘
```

### Design Patterns

1. **Registry Pattern**: Commands registered via `CommandSpec` for extensibility
2. **Backend Strategy**: Pluggable serial/3DS backends with common `set_buttons()` interface
3. **MVC-like**: GUI (view/controller) + ScriptEngine (model) + backends (services)
4. **Command Pattern**: Script commands as dictionaries with registered execution handlers
5. **Thread-safe**: Locks for frame buffers, event-based coordination

### Threading Model

- **Main thread**: Tkinter event loop (GUI updates)
- **Camera thread**: Continuously pulls frames from FFmpeg subprocess
- **Serial thread**: Keep-alive loop sending ~20 Hz packets
- **Script thread**: Executes user scripts with instruction pointer (IP) based stepping
- **Frame lock**: Protects `latest_frame_bgr` access

---

## Key Conventions

### Naming Conventions

| Type | Convention | Examples |
|------|-----------|----------|
| Functions | `snake_case` | `cmd_press`, `fmt_wait`, `resolve_value` |
| Classes | `PascalCase` | `CommandSpec`, `SerialController`, `ScriptEngine` |
| Variables | `snake_case` | `latest_frame_bgr`, `_stop` (threading.Event) |
| JSON Keys | `lowercase_underscore` | `cmd`, `buttons`, `rgb`, `out` |
| Script Variables | `$` prefix | `$counter`, `$match`, `$frame` |

### Code Organization

- **Top-level modules**: No packages, flat structure
- **Section delimiters**: Use `# ---- Title` comments for major sections
- **Helper functions**: Grouped before classes
- **Type annotations**: Minimal (mostly untyped for flexibility)

### Script Format (JSON)

Scripts are JSON arrays of command objects:

```json
[
  {"cmd": "comment", "text": "Example script"},
  {"cmd": "set", "var": "counter", "value": 0},
  {"cmd": "while", "left": "$counter", "op": "<", "right": 10},
    {"cmd": "press", "buttons": ["A"], "ms": 50},
    {"cmd": "wait", "ms": 100},
    {"cmd": "add", "var": "counter", "value": 1},
  {"cmd": "end_while"}
]
```

### Variable References

- Use `$varname` (as string) to reference variables in conditions and arguments
- Engine resolves via `resolve_value(ctx, "$varname")` → actual value
- Special: `$frame` passes current camera frame as PNG base64 payload

---

## Important Subsystems

### 1. Script Engine Context (`ctx`)

Commands receive a context dictionary:

```python
{
  "vars": {...},              # Script variables (dict)
  "labels": {...},            # Label name → IP mapping
  "if_map": {...},            # if IP → end_if IP
  "while_to_end": {...},      # while IP → end_while IP
  "end_to_while": {...},      # Reverse mapping
  "stop": threading.Event(),  # Stop flag
  "get_frame": callable,      # Returns latest BGR numpy array or None
  "ip": int,                  # Current instruction pointer
  "get_backend": callable,    # Returns active backend (serial/3DS)
}
```

### 2. High-Precision Timing System

The engine uses a hybrid timing approach for sub-millisecond precision:

**Timing Functions:**
- `precise_sleep(duration_sec)`: Non-interruptible high-precision sleep
- `precise_sleep_interruptible(duration_sec, stop_event)`: Interruptible with stop checking

**Precision Strategy:**
- Uses `time.perf_counter()` for sub-millisecond accuracy
- **Short durations (< 2ms)**: Pure busy-wait for maximum precision
- **Long durations (≥ 2ms)**: Hybrid approach:
  - Sleep for most of the duration (conservatively, in 0.5-1ms chunks)
  - Busy-wait for the final ~2ms for precision
  - Periodic stop event checking for interruptibility

**Timing Accuracy:**
- Sub-millisecond precision for all commands
- Supports fractional milliseconds (e.g., `"ms": 3.5`)
- Minimizes CPU usage while maintaining accuracy
- Windows-compatible (avoids time.sleep() precision issues)

**Commands Using Precise Timing:**
- `wait`: Uses `precise_sleep_interruptible()`
- `press`: Uses `precise_sleep_interruptible()` for hold duration
- `mash`: Uses `precise_sleep_interruptible()` for hold and wait phases
- `tap_touch` (3DS): Uses `precise_sleep()` in ThreeDSClasses

**Performance Characteristics:**
- Minimal CPU overhead for waits > 5ms
- Increased CPU during final 2ms of each wait (busy-wait)
- Eliminates timing drift in rapid sequences
- Accurate for mashing rates up to 200+ presses/second

### 3. Command Registry System

Commands defined in `ScriptEngine._build_default_registry()`:

```python
CommandSpec(
    name="press",                           # JSON cmd value
    required_keys=["buttons", "ms"],        # Validation
    fn=cmd_press,                           # Runtime executor
    doc="Press buttons for duration",       # Help text
    arg_schema=[                            # UI editor schema
        {"key": "buttons", "type": "list", "default": ["A"], "help": "..."},
        {"key": "ms", "type": "int", "default": 50, "help": "..."}
    ],
    format_fn=fmt_press,                    # Pretty display
    group="Controller",                     # Insert panel grouping
    order=10,                               # Sort order
    test=False,                             # Has test button
    exportable=True                         # Python export compatible
)
```

### 3. Serial Communication

**Packet Format** (3 bytes):
```
Byte 0: 0x54 (header)
Byte 1: High byte (L=0x01, R=0x02, X=0x04, Y=0x08)
Byte 2: Low byte (A=0x01, B=0x02, Right=0x04, Left=0x08, Up=0x10, Down=0x20, Select=0x40, Start=0x80)
```

**Immediate Send Architecture**:
- `set_buttons()` and `set_state()` **immediately** send packets to the serial device
- No delay waiting for keepalive thread
- Ensures accurate transmission even for rapid button changes (< 50ms intervals)
- `ser.flush()` called to guarantee immediate transmission

**Keep-alive Thread**:
- Runs at ~20 Hz (50 ms interval) as backup
- Re-transmits current state periodically
- Prevents receiver timeout
- Acts as safety net, not primary transmission mechanism

**Why This Matters**:
- Enables button changes faster than 50ms (e.g., 1ms, 5ms, 10ms)
- No missed button presses due to transmission delays
- Critical for ultra-fast mashing (200+ presses/sec)
- Ensures frame-perfect combos and rapid sequences work correctly

### 4. 3DS Backend

**UDP Protocol**: 20-byte packets containing HID state, touch coordinates, circle pad, etc.

**Touch Mapping**: `px_to_hid(x, y)` converts 320×240 pixel coords to HID values

### 5. Camera Integration

**FFmpeg Pipeline**:
1. Enumerate devices via `-list_devices true -f dshow -i dummy`
2. Capture via `ffmpeg -f dshow -i video="DeviceName" -f rawvideo -pix_fmt bgr24 -`
3. Read raw BGR24 bytes from stdout
4. Reshape to numpy array: `frame_bgr = np.frombuffer(bytes, np.uint8).reshape(h, w, 3)`
5. Display via PIL → ImageTk.PhotoImage

**Coordinate System**:
- Mouse hover shows pixel coordinates
- Click copies `x,y` to clipboard
- Shift+Click copies `{"x":123,"y":45}` (JSON format)

---

## Development Workflows

### Running the Application

```bash
# Install dependencies
py -m pip install numpy pillow pyserial

# Run application
py main.py
```

### Creating a New Script

1. Click "New" in GUI
2. Enter script name
3. Use "Insert Command" panel to add commands
4. Double-click commands to edit parameters
5. Click "Save" to write to `./scripts/scriptname.json`

### Adding a Custom Command

**Example: Adding a `beep` command**

Edit `ScriptEngine.py` in `_build_default_registry()`:

```python
# 1. Define formatter (optional but recommended)
def fmt_beep(c):
    return f"Beep {c.get('freq', 440)} Hz for {c.get('ms', 100)} ms"

# 2. Implement runtime function
def cmd_beep(ctx, c):
    freq = int(resolve_value(ctx, c.get("freq", 440)))
    ms = int(resolve_value(ctx, c.get("ms", 100)))

    # Your implementation here
    # Example: Store result in variable
    ctx["vars"]["last_beep_freq"] = freq

    # Use existing wait if needed
    cmd_wait(ctx, {"ms": ms})

# 3. Register CommandSpec
specs = [
    # ... existing commands ...
    CommandSpec(
        "beep",
        ["freq", "ms"],
        cmd_beep,
        doc="Play a beep sound at specified frequency",
        arg_schema=[
            {"key": "freq", "type": "int", "default": 440, "help": "Frequency in Hz"},
            {"key": "ms", "type": "int", "default": 100, "help": "Duration in milliseconds"}
        ],
        format_fn=fmt_beep,
        group="Custom",
        order=20,
        exportable=True
    ),
]
```

### Adding a Python Extension

**Create `py_scripts/my_tool.py`:**

```python
def main(*args):
    # args contain resolved values from script
    # Example: args = [10, 20, {"key": "value"}]

    result = process_data(args)

    # Return JSON-serializable value
    return result
```

**Use in script:**

```json
{
  "cmd": "run_python",
  "file": "my_tool.py",
  "args": [10, "$counter", {"x": "$x", "y": "$y"}],
  "out": "result"
}
```

**Access frame data:**

```python
import base64
from io import BytesIO
from PIL import Image

def decode_frame(payload):
    if not payload or payload.get("__frame__") != "png_base64":
        return None
    data = base64.b64decode(payload["data_b64"])
    return Image.open(BytesIO(data)).convert("RGB")

def main(frame_payload, x, y):
    img = decode_frame(frame_payload)
    if img:
        pixel = img.getpixel((x, y))
        return {"r": pixel[0], "g": pixel[1], "b": pixel[2]}
    return None
```

### Packaging for Distribution

```bash
# Install PyInstaller
py -m pip install pyinstaller

# Build
pyinstaller --noconsole --onedir --clean --name ControllerMacroRunner \
  --add-data "scripts;scripts" \
  --add-data "py_scripts;py_scripts" \
  main.py

# Distribute: zip dist/ControllerMacroRunner/ folder
```

---

## Testing and Quality Assurance

### Current Testing Status

⚠️ **Minimal formal test suite** - No dedicated test files

**Available Testing:**
- **Manual GUI testing**: Editor dialog test buttons for some commands (e.g., `find_color`)
- **Validation**: Script loading checks required keys, command existence
- **Block structure**: Editor warns about unclosed if/while blocks (non-strict during editing, strict at runtime)

### Testing Approach

1. **Command testing**: Use "Test" button in editor dialog for commands that support it
2. **Script validation**: Load scripts to check JSON structure and required keys
3. **Manual testing**: Run scripts with actual hardware/camera to verify behavior

### Safety Features

- **AST-validated math expressions**: Whitelist-based expression evaluation
- **Subprocess isolation**: Python scripts run in separate processes
- **Button validation**: Checks against `ALL_BUTTONS` list
- **Boundary checking**: Pixel operations validate coordinates
- **Frame payload limits**: PNG base64 encoding controls size

---

## Common Tasks and How-Tos

### How to: Debug a Script

1. **Check variables**: Use variable viewer pane while script is running
2. **Add comments**: Insert `{"cmd": "comment", "text": "Debug checkpoint"}` to track execution
3. **Use set commands**: Store intermediate values in variables for inspection
4. **Check IP**: Script engine shows current instruction pointer during execution

### How to: Find Pixel Coordinates

1. Start camera feed
2. Hover mouse over video preview
3. Coordinates appear as `(x,y)` near video
4. Click to copy to clipboard
5. Shift+Click for JSON format: `{"x":123,"y":45}`

### How to: Use find_color Command

```json
{
  "cmd": "find_color",
  "x": 100,           # Pixel X coordinate
  "y": 200,           # Pixel Y coordinate
  "rgb": [255, 0, 0], # Target color [R, G, B]
  "tol": 20,          # Tolerance (0-255 per channel)
  "out": "match"      # Output variable name
}
```

Result stored in `$match` (boolean true/false)

### How to: Use contains Command

```json
{
  "cmd": "contains",
  "needle": "abc",         # Value to search for
  "haystack": "abcdefgh",  # Container to search in
  "out": "found"           # Output variable name
}
```

Result stored in `$found` (boolean true/false)

**Works with strings (substring check):**
```json
{"cmd": "contains", "needle": "world", "haystack": "hello world", "out": "found"}
```
Result: `$found` = `true`

**Works with lists (membership check):**
```json
{"cmd": "set", "var": "items", "value": ["apple", "banana", "cherry"]}
{"cmd": "contains", "needle": "banana", "haystack": "$items", "out": "has_banana"}
```
Result: `$has_banana` = `true`

**Use with variables:**
```json
{"cmd": "set", "var": "search", "value": "target"}
{"cmd": "set", "var": "text", "value": "find the target here"}
{"cmd": "contains", "needle": "$search", "haystack": "$text", "out": "matched"}
```
Result: `$matched` = `true`

### How to: Use mash Command

```json
{
  "cmd": "mash",
  "buttons": ["A"],   # Buttons to mash
  "duration_ms": 1000, # Mash for 1 second
  "hold_ms": 25,      # Hold each press for 25ms (optional)
  "wait_ms": 25       # Wait 25ms between presses (optional)
}
```

**Default behavior**: 20 presses per second (25ms hold + 25ms wait = 50ms per cycle)

**Example - Fast mashing** (40 presses/second):
```json
{
  "cmd": "mash",
  "buttons": ["A", "B"],
  "duration_ms": 2000,
  "hold_ms": 12,
  "wait_ms": 13
}
```

**Example - Slower mashing** (10 presses/second):
```json
{
  "cmd": "mash",
  "buttons": ["A"],
  "duration_ms": 5000,
  "hold_ms": 50,
  "wait_ms": 50
}
```

### How to: Control Script Flow

**Infinite loop:**
```json
[
  {"cmd": "label", "name": "start"},
  {"cmd": "press", "buttons": ["A"], "ms": 50},
  {"cmd": "wait", "ms": 1000},
  {"cmd": "goto", "label": "start"}
]
```

**Conditional execution:**
```json
[
  {"cmd": "find_color", "x": 100, "y": 200, "rgb": [255, 0, 0], "tol": 20, "out": "found"},
  {"cmd": "if", "left": "$found", "op": "==", "right": true},
    {"cmd": "press", "buttons": ["A"], "ms": 50},
  {"cmd": "end_if"}
]
```

**Counter loop:**
```json
[
  {"cmd": "set", "var": "i", "value": 0},
  {"cmd": "while", "left": "$i", "op": "<", "right": 10},
    {"cmd": "press", "buttons": ["A"], "ms": 50},
    {"cmd": "wait", "ms": 100},
    {"cmd": "add", "var": "i", "value": 1},
  {"cmd": "end_while"}
]
```

### How to: Export Script to Python

Use `ScriptToPy.py` to convert compatible scripts to standalone Python files:

**Supported Commands:**
- ✅ `comment`, `wait`, `press`, `hold`, `mash`
- ✅ `set`, `add`, `contains` (variables)
- ✅ `if/end_if`, `while/end_while` (control flow)
- ✅ `run_python` (Python script execution)

**Unsupported Commands (Export Limitations):**
- ❌ `find_color`, `read_text` - Camera/vision commands (no camera runtime in export)
- ❌ `label`, `goto` - Not compatible with structured Python export
- ❌ `$frame` references - Camera frame payload not supported
- ❌ `tap_touch` - 3DS-specific command
- ❌ `type_name` - Complex keyboard navigation not supported

**Benefits:**
- **No timing delays** - Direct execution without engine overhead
- **Standalone** - Single Python file with all dependencies
- **Performance** - Native serial communication
- **Distribution** - Easy to share and deploy

**Usage (programmatically):**
```python
from ScriptToPy import export_script_to_python

# From main.py App instance
export_script_to_python(self)
```

**Generated Code Includes:**
- Serial communication functions
- Button mapping
- Keep-alive mechanism
- All command implementations (press, mash, wait, etc.)
- Variable declarations and control flow

---

## Important Gotchas and Notes

### Critical Information for AI Assistants

1. **Always read files before modifying**: Never propose changes to code you haven't read
2. **Avoid over-engineering**: Make only requested changes, don't add unnecessary features
3. **No backwards-compatibility hacks**: Delete unused code completely, don't comment it out
4. **Security awareness**: Watch for command injection, XSS in script execution
5. **Thread safety**: Be aware of frame lock when accessing `latest_frame_bgr`
6. **Verify Code Integrity**: Check all code and fix any potential bugs

### File Operation Patterns

**ALWAYS read before editing:**
```python
# Good: Read, then edit
Read file → Analyze → Edit specific section

# Bad: Edit without reading
Edit file with guessed content
```

**Prefer Edit over Write for existing files:**
- Use `Edit` tool for modifications to existing files
- Use `Write` tool only for new files
- Never use Write to overwrite existing files without reading first

### Common Pitfalls

1. **FFmpeg not found**: Ensure `bin/ffmpeg.exe` exists or FFmpeg is on PATH
2. **Camera not listing**: Check DirectShow enumeration with `ffmpeg -list_devices true -f dshow -i dummy`
3. **Script won't run**: Verify if/while blocks are properly closed
4. **Serial timeout**: Ensure keep-alive thread is running (~20 Hz)
5. **Frame access race**: Use frame lock when accessing camera frames
6. **Variable resolution**: Remember to use `$` prefix for variable references in JSON

### Windows-Specific Notes

- **Line endings**: Git configured for auto normalization (`text=auto` in `.gitattributes`)
- **File paths**: Use `resource_path()` helper for PyInstaller compatibility
- **COM ports**: Serial ports appear as `COM1`, `COM2`, etc.
- **DirectShow**: Camera capture only works on Windows (FFmpeg limitation)

### Script Execution Model

- **Instruction Pointer (IP)**: Commands execute sequentially by IP
- **Control flow**: `goto`, `if`, `while` modify IP
- **Stop flag**: Check `ctx["stop"].is_set()` for graceful termination
- **Block structure**: Must be properly nested (validated at runtime, warned in editor)

### Extension Guidelines

**When adding custom commands:**
- ✅ Use `resolve_value()` for variable support
- ✅ Use `resolve_vars_deep()` for nested structures
- ✅ Check `ctx["stop"].is_set()` in long operations
- ✅ Set `exportable=True` only if compatible with `ScriptToPy`
- ✅ Provide clear `arg_schema` for UI editor
- ✅ Document behavior in `doc` field

**When adding Python extensions:**
- ✅ Define `main(*args)` function
- ✅ Return JSON-serializable values
- ✅ Handle frame payloads with `decode_frame()` pattern
- ✅ Keep processing time reasonable (script thread waits)
- ⚠️ No access to script engine context directly
- ⚠️ Runs in subprocess (isolated environment)

---

## Available Commands Reference

| Command | Purpose | Exportable | Parameters |
|---------|---------|------------|------------|
| `comment` | Documentation | Yes | `text` |
| `wait` | Delay execution | Yes | `ms` |
| `press` | Press buttons briefly | Yes | `buttons` (list), `ms` |
| `hold` | Hold buttons | Yes | `buttons` (list), `ms` |
| `mash` | Rapidly mash buttons | Yes | `buttons` (list), `duration_ms`, `hold_ms` (default 25), `wait_ms` (default 25) |
| `set` | Set variable | Yes | `var`, `value` |
| `add` | Add to variable | Yes | `var`, `value` |
| `contains` | Membership test (like Python `in`) | Yes | `needle`, `haystack`, `out` |
| `label` | Define jump target | No | `name` |
| `goto` | Jump to label | No | `label` |
| `if` | Conditional block start | Yes | `left`, `op`, `right` |
| `end_if` | Conditional block end | Yes | - |
| `while` | Loop block start | Yes | `left`, `op`, `right` |
| `end_while` | Loop block end | Yes | - |
| `find_color` | Pixel color detection | No | `x`, `y`, `rgb`, `tol`, `out` |
| `read_text` | OCR text from camera | No | `x`, `y`, `width`, `height`, `out` |
| `run_python` | Execute Python script | Yes | `file`, `args`, `out` |
| `tap_touch` | 3DS touchscreen tap | No | `x`, `y`, `down_time`, `settle` |
| `type_name` | Pokemon name entry | No | `name`, `confirm` |

**Operators for if/while**: `==`, `!=`, `<`, `<=`, `>`, `>=`

---

## Git Workflow

### Branch Strategy

- **Main branch**: Stable releases
- **Feature branches**: `claude/claude-md-mk4zq9gnf0f2taax-S0jbA` (current development branch)

### Commit Guidelines

1. **Descriptive messages**: Focus on "why" not "what"
2. **Logical grouping**: Related changes in single commit
3. **No secrets**: Don't commit `.env`, credentials, etc.


## Quick Reference for AI Assistants

### Before Making Changes

- [ ] Read relevant files completely
- [ ] Understand existing patterns and conventions
- [ ] Check for similar existing functionality
- [ ] Verify thread safety for concurrent access
- [ ] Consider impact on serialization/deserialization

### When Modifying Code

- [ ] Match existing naming conventions
- [ ] Use appropriate design patterns already in codebase
- [ ] Add to command registry if creating new commands
- [ ] Update arg_schema for UI compatibility
- [ ] Test with actual GUI if possible

### When Creating Scripts

- [ ] Use proper JSON structure with `cmd` key
- [ ] Close all if/while blocks with end_if/end_while
- [ ] Use `$varname` for variable references
- [ ] Validate button names against ALL_BUTTONS
- [ ] Check coordinate bounds for pixel operations

### When Writing Documentation

- [ ] Update README.md for user-facing changes
- [ ] Update CLAUDE.md for architectural changes
- [ ] Include examples for new features
- [ ] Document any new dependencies
- [ ] Note any breaking changes

---

## Additional Resources

- **README.md**: User-facing documentation and tutorials
- **py_scripts/example_*.py**: Example Python extension scripts
- **ScriptEngine.py**: Command registry and execution engine implementation
- **main.py**: GUI implementation and application structure

---

## Version Information

- **Python**: 3.10+ recommended
- **Tkinter**: Built-in (no version requirement)
- **FFmpeg**: Any recent build with DirectShow support
- **numpy**: Latest compatible version
- **pillow**: Latest compatible version
- **pyserial**: Latest compatible version

---

*Last Updated: 2026-01-12*
*This document is maintained for AI assistants working with the ControllerMacroRunner codebase.*
