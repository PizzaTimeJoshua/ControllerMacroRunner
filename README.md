# Controller Macro Runner

A Windows desktop application for **high-precision gaming automation** with sub-millisecond accuracy.

#### Key Capabilities:

- **High-precision timing system** - Sub-millisecond accuracy for frame-perfect inputs
- **Immediate serial transmission** - No delay for rapid button sequences (200+ presses/sec)
- **Visual scripting** - JSON-based macro scripts with variables, loops, and conditionals
- **Camera integration** - DirectShow camera feed with pixel sampling for vision-based automation
- **Extensible** - Custom Python tools and command system
- **Multiple backends** - USB serial transmitter or 3DS Input Redirection

---

## Performance & Capabilities

### Timing Precision
- **Sub-millisecond accuracy** - Supports fractional milliseconds (e.g., 3.5ms)
- **1ms minimum** - Button presses as short as 1 millisecond
- **Zero drift** - Maintains accuracy even in long sequences
- **200+ presses/second** - Ultra-fast mashing without missing inputs

### Serial Communication
- **Immediate transmission** - Button changes sent instantly (no 50ms delay)
- **1,000,000 baud** - High-speed USB serial communication
- **Frame-perfect combos** - Accurate for competitive gaming and TAS
- **Keep-alive backup** - Prevents receiver timeout with periodic updates

### Script Capabilities
- **15 built-in commands** - Press, mash, hold, wait, variables, control flow, image processing
- **Variables & expressions** - Math operations, conditionals, loops
- **Camera vision** - Pixel sampling and color detection
- **Python extensions** - Run custom Python scripts with full variable support
- **Export to Python** - Convert scripts to standalone Python programs

---

## Features

- Camera Preview
  - Lists cameras using FFmpeg DirectShow device enumeration

  - Streams video using FFmpeg rawvideo piping

  - Mouse hover shows pixel coordinates (x,y)

  - Click-to-copy coordinates for use in scripts

- Audio Streaming (Optional)
  - Lists available audio input and output devices

  - Real-time audio passthrough from input to output

  - Positioned below video output, hidden when camera panel is hidden

  - Requires PyAudio library (optional dependency)

- Controller / Serial
  - Uses a [USB Wireless TX (Transmitter) by insideGadgets](https://shop.insidegadgets.com/product/usb-wireless-tx-transmitter/)

  - Connect to a COM port at 1,000,000 baud

  - **Immediate packet transmission** - no delay for button changes

  - Sends keep-alive packets at ~20 Hz (50 ms) as backup

  - Pairing warm-up (neutral packets) on connect

  - Supports changing channels

  - Supports rapid button sequences and ultra-fast mashing

- Script Editor

  - Load/save JSON scripts in `./scripts`

  - Insert commands from a grouped “Insert Command” panel

  - Context menu, multi-select capable list behavior (TreeView selection)

  - Double-click to edit commands

  - Indented view for nested if/while

- Script Engine

  - Variables (`set`, `add`)

  - Control flow: `label`, `goto`, `if/end_if`, `while/end_while`

  - Image: `find_color`

  - Custom: `run_python` (calls `main(*args)` in `./py_scripts/*.py`)

    - Supports variable refs in args (e.g. `"$counter"`)

    - Supports passing current frame via `"$frame"` (PNG base64 payload)

  - **High-precision timing system** for accurate button inputs

    - Sub-millisecond accuracy using hybrid sleep approach

    - Supports fractional milliseconds (e.g. 3.5ms)

    - Eliminates timing drift for rapid sequences

#### Folder Layout
```
project/
  main.py                      # main file program
  ScriptEngine.py
  SerialController.py
  ThreeDSClasses.py
  scripts/                    # macro scripts (JSON)
    example.json
  py_scripts/                 # user python helpers for run_python
    my_tool.py
  bin/
    ffmpeg.exe                # Can also be on PATH
    *.dll                     # Dll files from ffmpeg
```
The app will create `scripts/` and `py_scripts/` if missing.

#### Requirements (development)

- Python 3.10+ recommended

- Dependencies:

  - numpy

  - pillow

  - pyserial

  - pyaudio (optional, for audio streaming)

Install:
```bat
py -m pip install numpy pillow pyserial pyaudio
```

Note: PyAudio is optional. If not installed, the audio feature will be disabled but the application will still function normally.


FFmpeg:

- Recommended: place `ffmpeg.exe` and `*.dll` files in the folder `bin/` (or ensure ffmpeg is on PATH)

#### Running
```bat
py main.py
```


1. Select a Camera and click Start Cam

2. Select COM port and click Connect

3. Load a script from the Script dropdown or click New

4. Run with Run, stop with Stop

#### Camera Coordinate Helper

- Hover your mouse over the video to display x,y

- Click the video to copy `x,y` to the clipboard

- Shift+Click copies `{"x":123,"y":45}`

Coordinates are pixel coordinates in the source frame.

#### Serial Packet Format

The app sends packets shaped like:

- Transmit packet:

  - Byte0: `0x54`

  - Byte1: high byte (L/R/X/Y bits)

  - Byte2: low byte (A/B/DPAD/Start/Select bits)

Buttons mapping:

High byte:

- L trigger = `0x01`

- R trigger = `0x02`

- X = `0x04`

- Y = `0x08`

- Low byte:

- A = `0x01`

- B = `0x02`

- Right = `0x04`

- Left = `0x08`

- Up = `0x10`

- Down = `0x20`

- Select = `0x40`

- Start = `0x80`

The receiver requires continuous updates or it times out after a few seconds.

#### Script Format

Scripts are JSON arrays of command objects:
```json
[
  {"cmd":"comment","text":"Example: press A every second"},
  {"cmd":"press","buttons":["A"],"ms":80},
  {"cmd":"wait","ms":920}
]
```

#### Controller Commands

- **Press** - Press and release buttons:
```json
{"cmd":"press","buttons":["A"],"ms":80}
```

- **Hold** - Hold buttons indefinitely (until another command changes them):
```json
{"cmd":"hold","buttons":["A","B"]}
```

- **Mash** - Rapidly mash buttons for a duration (default: ~20 presses/second):
```json
{"cmd":"mash","buttons":["A"],"duration_ms":1000,"hold_ms":25,"wait_ms":25}
```
  - `buttons`: List of buttons to mash
  - `duration_ms`: Total time to mash in milliseconds
  - `hold_ms`: How long to hold each press (default: 25ms)
  - `wait_ms`: Wait time between presses (default: 25ms)
  - Press rate = 1000 / (hold_ms + wait_ms) presses/second

#### Variables

- Set variable:
```json
{"cmd":"set","var":"counter","value":0}
```

- Add:
```json
{"cmd":"add","var":"counter","value":1}
```


To reference a variable inside a condition or args, use `$name` (string):

- Example condition: `"left":"$counter"`

#### Control Flow

- Labels / goto:
```json
{"cmd":"label","name":"start"}
{"cmd":"goto","label":"start"}
```


- If block:
```json
{"cmd":"if","left":"$flag","op":"==","right":true}
  {"cmd":"press","buttons":["A"],"ms":50}
{"cmd":"end_if"}
```


- While block:
```json
{"cmd":"while","left":"$counter","op":"<","right":10}
  {"cmd":"press","buttons":["A"],"ms":50}
  {"cmd":"add","var":"counter","value":1}
{"cmd":"end_while"}
```

#### Image: find_color

Samples the pixel at `(x,y)` from the latest camera frame and compares against an RGB target with a tolerance:
```json
{"cmd":"find_color","x":100,"y":200,"rgb":[255,0,0],"tol":20,"out":"match"}
```


Stores boolean result in `$match`.

#### Custom Python: run_python

Runs a python file from `./py_scripts` (or an absolute path). The file must define:
```python
def main(*args):
    ...
    return something_json_serializable
```

Example command:
```json
{"cmd":"run_python","file":"my_tool.py","args":[1,2,3],"out":"result"}
```

##### Variable references in args

Arguments can contain `$var` references (including nested lists/dicts):
```json
{"cmd":"run_python","file":"tool.py","args":["$counter", {"pt":["$x","$y"]}], "out":"r"}
```

##### Passing the current camera frame

Use `"$frame"` as an argument to pass the latest frame as a PNG base64 payload:
```json
{"cmd":"run_python","file":"vision.py","args":["$frame", 10, 20], "out":"ok"}
```


In `example_frame_input.py`, decode like this:
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
    # ... process ...
    return True
```

#### Export to Python

Scripts can be exported to standalone Python files for distribution or direct execution.

**Supported Commands:**
- ✅ `comment` - Comments in generated code
- ✅ `wait` - Time delays
- ✅ `press` - Button presses
- ✅ `hold` - Button holds
- ✅ `mash` - Button mashing (NEW)
- ✅ `set`, `add` - Variable operations
- ✅ `if/end_if` - Conditional blocks
- ✅ `while/end_while` - Loop blocks
- ✅ `run_python` - Python script execution

**Limitations:**
- ❌ `find_color` - Camera frame processing not included in export
- ❌ `label`, `goto` - Not compatible with structured Python export
- ❌ `$frame` references - Camera functionality excluded
- ❌ `tap_touch` - 3DS-specific command not exported

**Benefits of Python Export:**
- **No timing delays** - Runs natively without engine overhead
- **Standalone execution** - No need for the full application
- **Distribution** - Share scripts as single Python files
- **Performance** - Direct serial communication without intermediary
- **Customization** - Edit generated Python code as needed

**How to Export:**
1. Load your script in the editor
2. Click "Export to Python" button (if available in GUI)
3. Or use `ScriptToPy.export_script_to_python()`
4. Generated file includes all necessary serial communication code

**Example Generated Code:**
```python
# Variables
counter = 0  # from $counter

# Pairing warm-up (~3 seconds neutral)
send_buttons([])
wait_with_keepalive(3.0)

# Press A every second, 10 times
while counter < 10:
    press(['A'], 80.0)
    wait_with_keepalive(920.0/1000.0)
    counter += 1
```

---

#### Adding Custom Commands

Commands live in `ScriptEngine._build_default_registry()` as `CommandSpec` entries. To add one:

1. Create a formatter `fmt_mycmd(c)` (pretty display)

2. Create an executor `cmd_mycmd(ctx, c)` (runtime behavior)

3. Register it via `CommandSpec(...)` with:

   - `group` and `order` for the Insert panel

   - `arg_schema` so the editor knows how to edit it

See the existing `find_color` implementation for a camera-reading example.

#### Custom Command Tutorial

To add a new command, you generally do 5 things in `ScriptEngine.py`:

1. Decide the command’s JSON shape

    Every script line is a dict with at least:

    ```json
    { "cmd": "your_command_name", "...": "args" }
    ```


    Pick required fields and optional fields.

    Example target:
    ```json

    { "cmd": "beep", "freq": 880, "ms": 120 }
    ```

2. Add a pretty formatter (how it displays in the Script list)

    Inside `_build_default_registry()` (near the other `fmt_*` functions), add:
    ```python
    def fmt_beep(c):
        return f"Beep {c.get('freq', 440)} Hz for {c.get('ms', 100)} ms"
    ```

    _Optional_, but makes the command appear more clear in the command list.

3. Implement the runtime function `cmd_*`

    Inside `_build_default_registry()` (near the other `cmd_*` functions), add:
    ```python
    def cmd_beep(ctx, c):
        freq = int(resolve_value(ctx, c.get("freq", 440)))
        ms = int(resolve_value(ctx, c.get("ms", 100)))

        # Your behavior here...
        # For example: set a variable to prove it ran
        ctx["vars"]["last_beep"] = {"freq": freq, "ms": ms}

        # If you want delays, prefer the existing wait:
        # cmd_wait(ctx, {"ms": ms})
    ```

    Notes:

    - You get access to engine context via ctx:

        - `ctx["vars"] `– variable dict

        - `ctx["get_frame"]()` – latest camera frame (BGR numpy array) or None

        - `ctx["labels"]`, `ctx["if_map"]`, `ctx["while_to_end"]`, etc.

        - `ctx["stop"].is_set()` – stop flag

        - `ctx["ip"]` – current instruction pointer (for jumps)

    If you want your command to support `$var` references, run values through `resolve_value(ctx, ...)` or, for nested lists/dicts, `resolve_vars_deep(ctx, ...)`.

4. Register it with a `CommandSpec`

    Add a new `CommandSpec(...)` entry to the `specs = [...]` list:
    ```python
    CommandSpec(
        "beep",
        ["freq", "ms"],                 # required keys
        cmd_beep,                       # runtime function
        doc="Play a beep sound (example).",
        arg_schema=[
            {"key": "freq", "type": "int", "default": 440, "help": "Frequency in Hz"},
            {"key": "ms", "type": "int", "default": 100, "help": "Duration in milliseconds"},
        ],
        format_fn=fmt_beep,
        group="Custom",
        order=20
    ),
    ```

    What these fields do:

    - `name`: command name used in JSON cmd

    - `required_keys`: validation when loading

    - `fn`: runtime execution

    - `doc`: shown in Insert panel and editor

    - `arg_schema`: drives the Edit dialog UI defaults and field types

    - `format_fn`: nice display string in the script list

    - `group`/`order`: where it appears in the Insert Command panel

#### Packaging as an EXE (Windows)

Recommended: PyInstaller `--onedir` with bundled ffmpeg.

1. Ensure the code uses a bundled `ffmpeg.exe` if present (helper like `ffmpeg_path()`).

2. Build:
```bat
py -m pip install pyinstaller
pyinstaller --noconsole --onedir --clean --name ControllerMacroRunner --add-data "scripts;scripts" --add-data "py_scripts;py_scripts" main.py
```

Distribute the resulting `dist/ControllerMacroRunner/` folder as a zip.

---

## Test Scripts

The `scripts/` directory includes comprehensive test scripts to verify functionality:

### Timing & Performance Tests
- **test_timing_precision.json** - Sub-millisecond timing accuracy verification
- **test_rapid_buttons.json** - Immediate transmission for fast sequences
- **test_mash_speeds.json** - Various mashing speeds (5-50 presses/sec)

### Feature Tests
- **test_quick.json** - Fast sanity check (~3 seconds)
- **test_mash_basic.json** - Basic mash command functionality
- **test_all_buttons.json** - All button types and combinations
- **test_variables_mash.json** - Variable-controlled parameters
- **test_comprehensive.json** - Full integration test (~20 seconds)

See `scripts/README_TESTS.md` for detailed descriptions and usage instructions.

---

## Troubleshooting

- No cameras listed

    - Verify FFmpeg can list dshow devices:
      ```
      ffmpeg -list_devices true -f dshow -i dummy
      ```

    - If using bundled ffmpeg, ensure `ffmpeg.exe` and the `*.dll` files are in the the `./bin` folder.

- Video only shows top-left / decode errors

    - Use rawvideo piping approach (already implemented).

    - Ensure output is `bgr24` and frame size matches reshape.

- Script won’t run

  - Make sure serial is connected

  - Ensure `if/while` blocks are properly closed when running

  - Editor tolerates incomplete blocks, but Run is strict

---

## Recent Improvements

### High-Precision Timing System
- Hybrid sleep approach (sleep + busy-wait) for sub-millisecond accuracy
- Supports fractional milliseconds (e.g., 3.5ms, 7.25ms)
- Eliminates timing drift in long sequences
- Enables 200+ button presses per second

### Immediate Serial Transmission
- Button state changes send immediately (no 50ms keepalive delay)
- Enables rapid sequences faster than 50ms
- Critical for ultra-fast mashing and frame-perfect combos
- Keep-alive thread still runs as backup safety net

### Mash Command
- Rapidly mash buttons at configurable rates
- Default: 20 presses/second (customizable to 200+)
- Independent hold_ms and wait_ms parameters
- Precise timing for each press cycle

---

## Safety Notes

- `run_python` executes local code. Only run scripts you trust.
- Keep controller output neutral when not actively running actions.
- Test scripts carefully before using in production environments.

---

## Documentation

- **CLAUDE.md** - Comprehensive guide for AI assistants and developers
- **scripts/README_TESTS.md** - Test script documentation
- See inline code comments for implementation details
