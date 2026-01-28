# Controller Macro Runner

A Windows desktop application for **high-precision gaming automation** with sub-millisecond timing accuracy. Automate controller inputs for game consoles through macro scripting with support for vision-based automation.

---

## Table of Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [User Interface](#user-interface)
- [Script Commands Reference](#script-commands-reference)
- [Custom Python Scripts](#custom-python-scripts)
- [Export to Standalone Python](#export-to-standalone-python)
- [Keyboard Control Mode](#keyboard-control-mode)
- [Troubleshooting](#troubleshooting)
- [Advanced Topics](#advanced-topics)
- [Safety Notes](#safety-notes)

---

## Features

### Core Capabilities

- **High-Precision Timing** - Sub-millisecond accuracy for frame-perfect inputs
- **Immediate Serial Transmission** - No delay for rapid button sequences (200+ presses/sec)
- **Visual Scripting** - JSON-based macro scripts with variables, loops, and conditionals
- **Camera Integration** - DirectShow camera feed with pixel sampling for vision-based automation
- **Multiple Output Backends** - USB serial transmitter or Nintendo 3DS Input Redirection
- **Extensible** - Custom Python tools and command system
- **Custom Themes** - Create and preview a personalized color theme with live apply

### Camera Features

- Live video preview from DirectShow cameras via FFmpeg
- Mouse hover shows pixel coordinates (x,y)
- Double-click to pop out camera to a separate window
- Configurable aspect ratios (GBA, DS, 3DS, Standard)
- Region selector for OCR areas
- Color picker for find_color command

### Audio Features (Optional)

- Real-time audio passthrough from input to output device
- Requires PyAudio library (optional dependency)
- Useful for monitoring game audio while automating

### Script Engine

- **40+ Built-in Commands** - Controller, timing, variables, control flow, vision, audio, and user input
- **Variables & Expressions** - Math operations with `$variable` references
- **Control Flow** - Labels, goto, if/end_if, while/end_while
- **Vision Commands** - Pixel color detection (find_color) and OCR (read_text)
- **Python Extensions** - Run custom Python scripts with full variable support
- **Export to Python** - Convert scripts to standalone Python programs

---

## Requirements

### System Requirements

- **Operating System**: Windows 10/11 (required for DirectShow camera capture)
- **Python**: 3.10 or higher recommended

### Hardware (for controller output)

Choose one of the following output methods:

1. **USB Wireless TX Transmitter** by insideGadgets
   - Purchase: [insideGadgets Shop](https://shop.insidegadgets.com/product/usb-wireless-tx-transmitter/)
   - Connects via USB COM port at 1,000,000 baud
   - Works with: NES Classic, SNES Classic, GBA, and other compatible receivers

2. **Nintendo 3DS with Input Redirection**
   - Requires: 3DS with custom firmware and InputRedirectionNTR or similar
   - Connects via UDP over WiFi
   - Supports touchscreen input

---

## Installation

### Step 1: Install Python

Download and install Python 3.10 or higher from [python.org](https://www.python.org/downloads/).

During installation, make sure to check **"Add Python to PATH"**.

### Step 2: Install Required Dependencies

Open a command prompt and run:

```bat
py -m pip install numpy pillow pyserial
```

**Required packages:**
| Package | Purpose | Link |
|---------|---------|------|
| numpy | Image array processing | [PyPI](https://pypi.org/project/numpy/) |
| pillow | Image handling | [PyPI](https://pypi.org/project/Pillow/) |
| pyserial | Serial communication | [PyPI](https://pypi.org/project/pyserial/) |

### Step 3: Install Optional Dependencies

```bat
py -m pip install pyaudio pytesseract opencv-python
```

**Optional packages:**
| Package | Purpose | Link |
|---------|---------|------|
| pyaudio | Audio streaming/passthrough | [PyPI](https://pypi.org/project/PyAudio/) |
| pytesseract | OCR text recognition | [PyPI](https://pypi.org/project/pytesseract/) |
| opencv-python | Advanced image processing for OCR | [PyPI](https://pypi.org/project/opencv-python/) |

**Note for PyAudio on Windows:** If pip install fails, download the appropriate wheel from [Unofficial Windows Binaries](https://www.lfd.uci.edu/~gohlke/pythonlibs/#pyaudio).

**Note for pytesseract:** You also need to install Tesseract OCR:
- **Option A: Bundled (Recommended)** - Place `tesseract.exe` in the `bin/` folder (auto-detected)
- **Option B: System PATH** - Download from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki) and add to PATH

### Step 4: Install FFmpeg

FFmpeg is required for camera capture.

**Option A: Bundled (Recommended)**
1. Download FFmpeg from [ffmpeg.org](https://ffmpeg.org/download.html) or [gyan.dev builds](https://www.gyan.dev/ffmpeg/builds/)
2. Extract `ffmpeg.exe` and all `*.dll` files to the `bin/` folder in your project directory

**Option B: System PATH**
1. Download and extract FFmpeg
2. Add the `bin` folder to your system PATH
3. Verify with: `ffmpeg -version`

### Step 5: Download the Application

Clone or download this repository:

```bat
git clone https://github.com/yourusername/ControllerMacroRunner.git
cd ControllerMacroRunner
```

### Folder Structure

```
ControllerMacroRunner/
  main.py                 # Main application
  ScriptEngine.py         # Script execution engine
  SerialController.py     # USB serial communication
  InputRedirection.py       # 3DS Input Redirection backend
  ScriptToPy.py           # Script to Python exporter
  scripts/                # Macro scripts (JSON) - auto-created
  py_scripts/             # Custom Python helpers - auto-created
  bin/
    ffmpeg.exe            # FFmpeg binary
    tesseract.exe         # Tesseract OCR binary (optional)
    *.dll                 # FFmpeg/Tesseract dependencies
```

---

## Quick Start

### Running the Application

```bat
py main.py
```

### Basic Workflow

1. **Select a Camera** - Choose from the dropdown and click **Start Cam**
2. **Connect Controller** - Select COM port and click **Connect** (for USB Serial) or configure 3DS IP
3. **Load or Create Script** - Use the Script dropdown or click **New**
4. **Run the Script** - Click **Run** to execute, **Stop** to halt

### Your First Script

1. Click **New** and name your script (e.g., "test")
2. Click **Add** to insert commands
3. Select "press" from the command type dropdown
4. Choose button "A" and set duration to 100ms
5. Click **OK** to add the command
6. Click **Save** then **Run**

---

## User Interface

### Top Bar Controls

| Control | Description |
|---------|-------------|
| **Camera** | Select and start/stop camera feed |
| **Cam Ratio** | Adjust aspect ratio (GBA, DS, 3DS, etc.) |
| **Show/Hide Cam** | Toggle camera panel visibility |
| **COM** | Select serial port for USB transmitter |
| **Connect** | Connect/disconnect serial device |
| **Channel** | Set wireless channel (1-16) |
| **Output** | Choose backend: USB Serial or 3DS Input Redirection |
| **Script** | Load, save, and manage scripts |
| **Keyboard Control** | Enable manual keyboard input |

### Script Editor

- **Script Commands Panel** - Displays script with syntax highlighting
- **Variables Panel** - Shows current variable values during execution
- **Insert/Edit/Delete** - Modify script commands
- **Up/Down** - Reorder commands
- **Right-click** - Context menu with nested Add > Category > Command submenu

### Appearance and Themes

- **Theme Mode** - Switch between Auto, Light, Dark, or Custom
- **Custom Theme Editor** - Adjust per-color swatches with live previews and Apply
- **Save vs Apply** - Apply updates without closing Settings, save to persist

### Keyboard Shortcuts

- **Double-click** on video: Pop out camera window
- **Click** on video: Enable keyboard control focus
- **Double-click** on command: Edit command
- **Delete** on command: Remove selected command

---

## Script Commands Reference

Scripts are JSON arrays of command objects. Each command has a `cmd` field and command-specific parameters.

### Controller Commands

#### press
Press buttons for a duration, then release.

```json
{"cmd": "press", "buttons": ["A"], "ms": 80}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| buttons | list | required | Buttons to press: A, B, X, Y, Up, Down, Left, Right, Start, Select, L, R |
| ms | number | required | Hold duration in milliseconds |

#### hold
Hold buttons indefinitely until changed by another command.

```json
{"cmd": "hold", "buttons": ["Up", "A"]}
```

#### mash
Rapidly press buttons for a duration.

```json
{"cmd": "mash", "buttons": ["A"], "duration_ms": 1000, "hold_ms": 25, "wait_ms": 25}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| buttons | list | required | Buttons to mash |
| duration_ms | number | 1000 | Total mashing duration |
| hold_ms | number | 25 | Hold time per press |
| wait_ms | number | 25 | Wait time between presses |
| until_ms | number | 0 | If set, mash until this many ms elapsed since start_timing (overrides duration_ms) |

Press rate = 1000 / (hold_ms + wait_ms) presses/second. Default is 20 presses/sec.

**Timing reference mode:** Use `until_ms` with `start_timing` for frame-perfect mash sequences:
```json
{"cmd": "start_timing"}
{"cmd": "mash", "buttons": ["A"], "until_ms": 2000, "hold_ms": 30, "wait_ms": 30}
```
This mashes A until exactly 2000ms have elapsed since start_timing.

### Timing Commands

#### wait
Pause script execution.

```json
{"cmd": "wait", "ms": 500}
```

#### start_timing
Set a timing reference point for cumulative timing. Used with `wait_until` for frame-perfect sequences.

```json
{"cmd": "start_timing"}
```

#### wait_until
Wait until the specified milliseconds have elapsed since the last `start_timing`. Automatically compensates for execution overhead.

```json
{"cmd": "wait_until", "ms": 1000}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| ms | number | required | Target elapsed time since start_timing |

**Example: Frame-perfect input sequence**
```json
{"cmd": "start_timing"}
{"cmd": "press", "buttons": ["A"], "ms": 50}
{"cmd": "wait_until", "ms": 500}
{"cmd": "press", "buttons": ["B"], "ms": 50}
{"cmd": "wait_until", "ms": 1000}
```
This ensures the B press happens exactly 500ms after start_timing, regardless of how long the A press took.

#### get_elapsed
Get the milliseconds elapsed since the last `start_timing`.

```json
{"cmd": "get_elapsed", "out": "elapsed"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| out | string | "elapsed" | Variable to store elapsed milliseconds |

### Variable Commands

#### set
Set a variable value.

```json
{"cmd": "set", "var": "counter", "value": 0}
```
Use `$varname` to reference variables in other commands.

**Math expressions:** Prefix with `=` to evaluate expressions:
```json
{"cmd": "set", "var": "result", "value": "=$x + $y * 2"}
```

#### add
Add a value to an existing variable.

```json
{"cmd": "add", "var": "counter", "value": 1}
```

#### contains
Check if a value exists in another (like Python's `in` operator).

```json
{"cmd": "contains", "needle": "abc", "haystack": "abcdefgh", "out": "found"}
```
Works with strings (substring check) and lists (membership check).

#### random
Randomly select one value from a list of choices.

```json
{"cmd": "random", "choices": [1, 2, 3, 4, 5], "out": "random_value"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| choices | list | required | List of values to choose from (literal list or $var) |
| out | string | "random_value" | Variable name to store selected value (no $) |

**Examples:**

**Random number from a list:**
```json
{"cmd": "random", "choices": [100, 200, 300, 400, 500], "out": "wait_time"}
{"cmd": "wait", "ms": "$wait_time"}
```

**Random button selection:**
```json
{"cmd": "random", "choices": ["A", "B", "X", "Y"], "out": "button"}
{"cmd": "press", "buttons": ["$button"], "ms": 50}
```

**Random choice from a variable:**
```json
{"cmd": "set", "var": "options", "value": ["up", "down", "left", "right"]}
{"cmd": "random", "choices": "$options", "out": "direction"}
```

**Note:** The random number generator is seeded with the current time when the script engine starts, ensuring different random sequences on each run.

#### random_range
Generate a random number between min and max (inclusive).

```json
{"cmd": "random_range", "min": 0, "max": 100, "integer": false, "out": "random_value"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| min | number | 0 | Minimum value (inclusive, supports $var and =expr) |
| max | number | 100 | Maximum value (inclusive, supports $var and =expr) |
| integer | bool | false | If true, returns integer; if false, returns float |
| out | string | "random_value" | Variable name to store result (no $) |

**Examples:**

**Random integer for dice roll (1-6):**
```json
{"cmd": "random_range", "min": 1, "max": 6, "integer": true, "out": "dice"}
{"cmd": "wait", "ms": "=$dice * 100"}
```

**Random float for timing variation:**
```json
{"cmd": "random_range", "min": 0.5, "max": 2.0, "integer": false, "out": "multiplier"}
{"cmd": "set", "var": "wait_time", "value": "=100 * $multiplier"}
{"cmd": "wait", "ms": "$wait_time"}
```

**Random range with variables:**
```json
{"cmd": "set", "var": "min_wait", "value": 50}
{"cmd": "set", "var": "max_wait", "value": 150}
{"cmd": "random_range", "min": "$min_wait", "max": "$max_wait", "integer": true, "out": "delay"}
{"cmd": "wait", "ms": "$delay"}
```

#### random_value
Generate a random float between 0.0 and 1.0 (exclusive of 1.0).

```json
{"cmd": "random_value", "out": "random_value"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| out | string | "random_value" | Variable name to store result (no $) |

**Examples:**

**Random probability check (50% chance):**
```json
{"cmd": "random_value", "out": "chance"}
{"cmd": "if", "left": "$chance", "op": "<", "right": 0.5}
  {"cmd": "press", "buttons": ["A"], "ms": 50}
{"cmd": "end_if"}
```

**Random scaling for wait time:**
```json
{"cmd": "random_value", "out": "scale"}
{"cmd": "set", "var": "wait_time", "value": "=50 + $scale * 100"}
{"cmd": "wait", "ms": "$wait_time"}
```

**Random weighted decision (30% chance for action A, 70% for action B):**
```json
{"cmd": "random_value", "out": "roll"}
{"cmd": "if", "left": "$roll", "op": "<", "right": 0.3}
  {"cmd": "press", "buttons": ["A"], "ms": 50}
{"cmd": "end_if"}
{"cmd": "if", "left": "$roll", "op": ">=", "right": 0.3}
  {"cmd": "press", "buttons": ["B"], "ms": 50}
{"cmd": "end_if"}
```

#### export_json
Export script variables to a JSON file.

```json
{"cmd": "export_json", "file": "save_data.json", "vars": ["counter", "score", "player_name"]}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| file | string | required | Output filename (relative to scripts folder or absolute path) |
| vars | list | [] | Variable names to export (empty = export all non-internal variables) |

#### import_json
Import variables from a JSON file.

```json
{"cmd": "import_json", "file": "save_data.json"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| file | string | required | Input filename (relative to scripts folder or absolute path) |

**Example: Save and restore progress**
```json
{"cmd": "set", "var": "counter", "value": 100}
{"cmd": "set", "var": "score", "value": 5000}
{"cmd": "export_json", "file": "progress.json", "vars": ["counter", "score"]}

{"cmd": "import_json", "file": "progress.json"}
```

### Control Flow Commands

#### label / goto
Define jump targets and jump to them.

```json
{"cmd": "label", "name": "start"}
{"cmd": "goto", "label": "start"}
```

#### if / end_if
Conditional block execution.

```json
{"cmd": "if", "left": "$counter", "op": "<", "right": 10}
  {"cmd": "press", "buttons": ["A"], "ms": 50}
{"cmd": "end_if"}
```
Operators: `==`, `!=`, `<`, `<=`, `>`, `>=`

#### while / end_while
Loop while condition is true.

```json
{"cmd": "while", "left": "$counter", "op": "<", "right": 10}
  {"cmd": "press", "buttons": ["A"], "ms": 50}
  {"cmd": "add", "var": "counter", "value": 1}
{"cmd": "end_while"}
```

### Image/Vision Commands

#### find_color
Sample a pixel and compare against a target color using perceptual CIE76 Delta E.

```json
{"cmd": "find_color", "x": 100, "y": 200, "rgb": [255, 0, 0], "tol": 10, "out": "match"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x, y | int | required | Pixel coordinates |
| rgb | [R,G,B] | required | Target color (0-255 each) |
| tol | number | 10 | Delta E tolerance |
| out | string | "match" | Variable to store boolean result |

**Delta E interpretation:**
- 0-1: Imperceptible difference
- 1-2: Perceptible through close observation
- 2-10: Perceptible at a glance
- 10+: Obvious difference

#### find_area_color
Calculate average color in a region and compare against a target color using perceptual CIE76 Delta E.

```json
{"cmd": "find_area_color", "x": 100, "y": 200, "width": 50, "height": 50, "rgb": [255, 0, 0], "tol": 10, "out": "match"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x, y | int | 0 | Top-left corner coordinates |
| width, height | int | 10 | Region size |
| rgb | [R,G,B] | required | Target color (0-255 each) |
| tol | number | 10 | Delta E tolerance |
| out | string | "match" | Variable to store boolean result |

#### wait_for_color
Wait until a pixel matches or doesn't match a target color. Polls at regular intervals until the condition is met or timeout occurs.

```json
{"cmd": "wait_for_color", "x": 100, "y": 200, "rgb": [255, 0, 0], "tol": 10, "interval": 0.1, "timeout": 30, "wait_for": true, "out": "match"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x, y | int | required | Pixel coordinates |
| rgb | [R,G,B] | required | Target color (0-255 each) |
| tol | number | 10 | Delta E tolerance |
| interval | number | 0.1 | Check interval in seconds |
| timeout | number | 0 | Timeout in seconds (0 = no timeout) |
| wait_for | bool | true | true = wait for match, false = wait for no match |
| out | string | "match" | Variable to store boolean result (true if condition met, false if timeout) |

**Example use cases:**
- Wait for a UI element to appear: `"wait_for": true`
- Wait for a UI element to disappear: `"wait_for": false`
- Poll every 0.5 seconds: `"interval": 0.5`
- Timeout after 60 seconds: `"timeout": 60`

#### wait_for_color_area
Wait until the average color in a region matches or doesn't match a target color. Polls at regular intervals until the condition is met or timeout occurs.

```json
{"cmd": "wait_for_color_area", "x": 100, "y": 200, "width": 50, "height": 50, "rgb": [255, 0, 0], "tol": 10, "interval": 0.1, "timeout": 30, "wait_for": true, "out": "match"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x, y | int | 0 | Top-left corner coordinates |
| width, height | int | 10 | Region size |
| rgb | [R,G,B] | required | Target color (0-255 each) |
| tol | number | 10 | Delta E tolerance |
| interval | number | 0.1 | Check interval in seconds |
| timeout | number | 0 | Timeout in seconds (0 = no timeout) |
| wait_for | bool | true | true = wait for match, false = wait for no match |
| out | string | "match" | Variable to store boolean result (true if condition met, false if timeout) |

#### read_text
OCR a region of the camera frame (requires pytesseract).

```json
{"cmd": "read_text", "x": 50, "y": 100, "width": 200, "height": 30, "out": "text"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x, y | int | required | Top-left corner |
| width, height | int | required | Region size |
| out | string | "text" | Variable to store result |
| scale | int | 4 | Upscale factor (higher = better for small fonts) |
| threshold | int | 0 | Binary threshold (0 = auto OTSU) |
| invert | bool | false | Invert colors for light-on-dark text |
| psm | int | 7 | Tesseract page segmentation mode |
| whitelist | string | "" | Allowed characters (e.g., "0123456789") |

#### save_frame
Save the current camera frame to `./saved_images`.

```json
{"cmd": "save_frame", "filename": "frame.png", "out": "saved_path"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| filename | string | "" | Optional filename (saved under saved_images) |
| out | string | "" | Variable to store saved path (no $) |

### Custom Commands

#### discord_status
Send a Discord webhook status update (optional ping and image).

```json
{"cmd": "discord_status", "message": "Run complete", "ping": true, "image": "$frame"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| message | string | "" | Message text (supports $var) |
| ping | bool | false | Ping the configured user ID |
| image | string | "" | Image file path or $frame |

Configure the webhook URL and optional user ID in **Settings > Discord**.

#### run_python
Execute a Python script from `./py_scripts`.

```json
{"cmd": "run_python", "file": "my_tool.py", "args": [10, "$counter"], "out": "result"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| file | string | required | Filename in ./py_scripts or absolute path |
| args | list | [] | Arguments passed to main(*args) |
| out | string | "" | Variable to store return value |
| timeout_s | int | 10 | Timeout in seconds |

#### comment
Documentation that does nothing at runtime.

```json
{"cmd": "comment", "text": "This is a comment"}
```

### User Input Commands

#### prompt_input
Pause the script and prompt the user for text input.

```json
{"cmd": "prompt_input", "prompt": "Enter target ID:", "out": "target_id", "confirm": true}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| prompt | string | "Enter value:" | Message shown to the user |
| out | string | "input" | Variable to store the entered value |
| confirm | bool | false | If true, show confirmation dialog before continuing |

#### prompt_choice
Pause the script and prompt the user to select from a list of options.

```json
{"cmd": "prompt_choice", "prompt": "Select action:", "choices": ["Attack", "Defend", "Run"], "out": "action", "style": "buttons"}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| prompt | string | "Select:" | Message shown to the user |
| choices | list | required | List of options to choose from |
| out | string | "choice" | Variable to store the selected value |
| style | string | "dropdown" | Display style: "dropdown" or "buttons" |

**Example: Interactive decision point**
```json
{"cmd": "prompt_choice", "prompt": "Shiny found! What should we do?", "choices": ["Catch it", "Run away", "Take screenshot"], "out": "decision", "style": "buttons"}
{"cmd": "if", "left": "$decision", "op": "==", "right": "Catch it"}
  {"cmd": "press", "buttons": ["A"], "ms": 50}
{"cmd": "end_if"}
```

### Audio Commands

#### play_sound
Play a sound file from the `./bin/sounds/` folder.

```json
{"cmd": "play_sound", "file": "notification.wav", "volume": 50, "wait": false}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| file | string | required | Sound filename in ./bin/sounds/ |
| volume | int | 100 | Volume level (0-100) |
| wait | bool | false | If true, wait for sound to finish before continuing |

**Supported formats:** WAV, MP3, and other formats supported by ffplay.

**Example: Audio notification on event**
```json
{"cmd": "find_color", "x": 100, "y": 200, "rgb": [255, 215, 0], "tol": 10, "out": "found_shiny"}
{"cmd": "if", "left": "$found_shiny", "op": "==", "right": true}
  {"cmd": "play_sound", "file": "alert.wav", "volume": 80}
{"cmd": "end_if"}
```

### Stick Commands

#### set_left_stick
Set the left stick position (3DS or PABotBase).
Position stays active until another stick command or reset.

```json
{"cmd": "set_left_stick", "x": 1.0, "y": 0.0}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x | number | 0.0 | X axis (-1.0..1.0) |
| y | number | 0.0 | Y axis (-1.0..1.0) |

#### reset_left_stick
Reset left stick to center.

```json
{"cmd": "reset_left_stick"}
```

#### set_right_stick
Set the right stick position (3DS or PABotBase).
Position stays active until another stick command or reset.

```json
{"cmd": "set_right_stick", "x": 0.0, "y": -1.0}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x | number | 0.0 | X axis (-1.0..1.0) |
| y | number | 0.0 | Y axis (-1.0..1.0) |

#### reset_right_stick
Reset right stick to center.

```json
{"cmd": "reset_right_stick"}
```

### 3DS-Specific Commands

#### tap_touch
Tap the 3DS touchscreen (3DS backend only).

```json
{"cmd": "tap_touch", "x": 160, "y": 120, "down_time": 0.1, "settle": 0.1}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| x | int | required | X pixel (0-319) |
| y | int | required | Y pixel (0-239) |
| down_time | float | 0.1 | Seconds to hold touch |
| settle | float | 0.1 | Seconds to wait after release |

#### press_ir
Press ZL/ZR for a duration.

```json
{"cmd": "press_ir", "buttons": ["ZL"], "ms": 80}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| buttons | list | required | IR buttons: ZL, ZR |
| ms | number | required | Hold duration in milliseconds |

#### hold_ir
Hold ZL/ZR until changed by another command.

```json
{"cmd": "hold_ir", "buttons": ["ZL", "ZR"]}
```

#### press_interface
Press Home/Power buttons.

```json
{"cmd": "press_interface", "buttons": ["Home"], "ms": 80}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| buttons | list | required | Interface buttons: Home, Power, PowerLong |
| ms | number | required | Hold duration in milliseconds |

Note: PowerLong triggers the power-off dialog. Use with care.

#### hold_interface
Hold Home/Power buttons until changed.

```json
{"cmd": "hold_interface", "buttons": ["Home"]}
```

### Pokemon Commands

#### type_name
Type a name on Pokemon FRLG/RSE naming screens.

```json
{"cmd": "type_name", "name": "RED", "confirm": true}
```
| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| name | string | "Red" | Name to type |
| confirm | bool | true | Press A after Start to confirm |
| move_delay_ms | int | 200 | Delay after D-pad moves |
| select_delay_ms | int | 600 | Delay after page switch |
| press_delay_ms | int | 400 | Delay after letter selection |
| button_hold_ms | int | 50 | Button hold duration |

---

## Custom Python Scripts

Create Python scripts in `./py_scripts/` to extend functionality.

### Basic Structure

```python
def main(*args):
    # Process args
    # Return JSON-serializable value
    return {"result": "success"}
```

### Accessing Camera Frames

Use `"$frame"` as an argument to receive the current camera frame:

```json
{"cmd": "run_python", "file": "vision.py", "args": ["$frame", 100, 200], "out": "color"}
```

```python
# vision.py
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
    if img is None:
        return None
    pixel = img.getpixel((x, y))
    return {"r": pixel[0], "g": pixel[1], "b": pixel[2]}
```

### Using Variables in Args

Reference script variables with `$varname`:

```json
{"cmd": "run_python", "file": "tool.py", "args": ["$counter", {"coords": ["$x", "$y"]}], "out": "result"}
```

### Example: Math Operations

```python
# example_math.py
def main(a, b, operation="add"):
    if operation == "add":
        return a + b
    elif operation == "multiply":
        return a * b
    return None
```

---

## Export to Standalone Python

Convert scripts to standalone Python files for distribution or direct execution.

### Supported Commands

| Command | Exported | Notes |
|---------|----------|-------|
| comment | Yes | Becomes Python comment |
| wait | Yes | Direct time.sleep |
| press | Yes | Serial communication |
| hold | Yes | Serial communication |
| mash | Yes | Loop with timing |
| set, add | Yes | Python variables |
| contains | Yes | Python `in` operator |
| random | Yes | Python `random.choice()` |
| random_range | Yes | Python `random.randint()` or `random.uniform()` |
| random_value | Yes | Python `random.random()` |
| if/end_if | Yes | Python if statements |
| while/end_while | Yes | Python while loops |
| run_python | Yes | Subprocess call |
| start_timing | No | Requires timing reference system |
| wait_until | No | Requires timing reference system |
| get_elapsed | No | Requires timing reference system |
| export_json | No | Requires filesystem |
| import_json | No | Requires filesystem |
| prompt_input | No | Requires GUI |
| prompt_choice | No | Requires GUI |
| play_sound | No | Requires audio system |
| save_frame | No | Requires camera and filesystem |
| discord_status | No | Requires Discord webhook |
| find_color | No | Requires camera |
| find_area_color | No | Requires camera |
| wait_for_color | No | Requires camera |
| wait_for_color_area | No | Requires camera |
| read_text | No | Requires camera + pytesseract |
| label/goto | No | Not compatible with structured code |
| tap_touch | No | 3DS-specific |
| set_left_stick | No | Requires stick backend |
| reset_left_stick | No | Requires stick backend |
| set_right_stick | No | Requires stick backend |
| reset_right_stick | No | Requires stick backend |
| press_ir | No | 3DS-specific |
| hold_ir | No | 3DS-specific |
| press_interface | No | 3DS-specific |
| hold_interface | No | 3DS-specific |
| type_name | No | Complex navigation |

### How to Export

1. Load your script in the editor
2. Click **Export .py** button
3. Choose save location
4. Run the generated Python file directly

### Generated Code Features

- Standalone serial communication
- Built-in keep-alive mechanism
- All variable and control flow logic
- No external dependencies beyond pyserial

---

## Keyboard Control Mode

Control the game directly with your keyboard when no script is running.

### Enabling

1. Check **Keyboard Control** checkbox in the top bar
2. Ensure serial is connected
3. Script must not be running

### Default Key Bindings

| Key | Button |
|-----|--------|
| W | Up |
| A | Left |
| S | Down |
| D | Right |
| J | A |
| K | B |
| U | X |
| I | Y |
| Enter | Start |
| Space | Select |
| Q | L |
| E | R |
| Up Arrow | Left Stick Up |
| Down Arrow | Left Stick Down |
| Left Arrow | Left Stick Left |
| Right Arrow | Left Stick Right |

### Customizing Bindings

Click **Keybinds...** to open the keybind configuration window:
- Select a button and click **Rebind...**
- Press the desired key
- Press Escape to cancel
- Click **Restore defaults** to reset all bindings

---

## Troubleshooting

### Camera Issues

**No cameras listed:**
1. Verify FFmpeg can detect cameras:
   ```bat
   ffmpeg -list_devices true -f dshow -i dummy
   ```
2. Ensure `ffmpeg.exe` and DLLs are in `./bin/` folder
3. Check if camera is in use by another application

**Video only shows partial frame or decode errors:**
- Ensure the camera supports the selected resolution
- Try a different aspect ratio setting

### Serial Connection Issues

**COM port not listed:**
1. Check Device Manager for the correct COM port
2. Install drivers if needed (e.g., CH340 drivers for some USB-serial adapters)
3. Click **Refresh** to update the list

**Connection fails:**
- Ensure no other application is using the COM port
- Try disconnecting and reconnecting the USB device

### Script Issues

**Script won't run:**
- Check that serial is connected
- Ensure all if/while blocks have matching end_if/end_while
- Check the status bar for error messages

**Timing seems off:**
- The timing system is optimized for Windows; precision may vary on virtual machines
- For ultra-fast sequences, ensure no CPU throttling is active

### OCR Issues (read_text)

**pytesseract not found:**
1. Install pytesseract: `pip install pytesseract`
2. Install Tesseract OCR from [UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki)
3. Place `tesseract.exe` in the `bin/` folder (recommended), or add to your PATH

**Poor recognition:**
- Increase the `scale` parameter (4-8 for pixel fonts)
- Try different `threshold` values
- Use `invert: true` for light text on dark background
- Restrict characters with `whitelist`

---

## Advanced Topics

### Serial Packet Format

The USB transmitter uses 3-byte packets:

| Byte | Value | Description |
|------|-------|-------------|
| 0 | 0x54 | Header |
| 1 | High byte | L=0x01, R=0x02, X=0x04, Y=0x08 |
| 2 | Low byte | A=0x01, B=0x02, Right=0x04, Left=0x08, Up=0x10, Down=0x20, Select=0x40, Start=0x80 |

### 3DS Input Redirection

Uses 20-byte UDP packets containing:
- HID pad state (4 bytes)
- Touch state (4 bytes)
- Circle pad (4 bytes)
- C-stick/CPP state (4 bytes)
- Interface buttons (4 bytes)

Touch coordinates are mapped from 320x240 pixel space to 12-bit HID values.

### High-Precision Timing System

The engine uses a hybrid approach for sub-millisecond accuracy:
- **Durations < 2ms**: Pure busy-wait for maximum precision
- **Durations >= 2ms**: Sleep for most of the duration, busy-wait for final 2ms
- Supports fractional milliseconds (e.g., `"ms": 3.5`)
- Eliminates timing drift in long sequences

### Adding Custom Commands

Edit `ScriptEngine.py` in `_build_default_registry()`:

1. Create a formatter function:
   ```python
   def fmt_mycommand(c):
       return f"MyCommand param={c.get('param')}"
   ```

2. Create an executor function:
   ```python
   def cmd_mycommand(ctx, c):
       param = resolve_value(ctx, c.get("param"))
       # Your implementation
       ctx["vars"]["result"] = param * 2
   ```

3. Register with CommandSpec:
   ```python
   CommandSpec(
       "mycommand",
       ["param"],  # required keys
       cmd_mycommand,
       doc="Description of what this command does.",
       arg_schema=[
           {"key": "param", "type": "int", "default": 0, "help": "Parameter description"}
       ],
       format_fn=fmt_mycommand,
       group="Custom",
       order=10
   )
   ```

### Packaging as EXE

Use PyInstaller to create a standalone executable:

```bat
python -m pip install pyinstaller

python -m PyInstaller --noconsole --onedir --icon="bin/icon.ico" --clean --name ControllerMacroRunner main.py
```

Distribute the `dist/ControllerMacroRunner/` folder as a zip file. Include the `bin/` folder with FFmpeg.

---

## Safety Notes

- **run_python** executes local code. Only run scripts you trust.
- Keep controller output neutral when not actively running scripts.
- Test scripts carefully before using in production environments.
- The application has no network access except for 3DS Input Redirection (local UDP).

---

## Contributing

See [AGENTS.md](AGENTS.md) for developer documentation and contribution guidelines.

---

## License

This project is provided as-is for personal use. See individual component licenses for dependencies.

---

## Acknowledgments

- [insideGadgets](https://shop.insidegadgets.com/) for the USB Wireless TX hardware
- [FFmpeg](https://ffmpeg.org/) for video capture capabilities
- [Tesseract OCR](https://github.com/tesseract-ocr/tesseract) for text recognition
- [Pokemon Automation](https://pokemonautomation.github.io/index.html) for inspiration and controller compatibility