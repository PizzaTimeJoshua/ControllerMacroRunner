# Controller Macro Runner

#### A Windows desktop app for:

- Displaying a DirectShow camera feed (via FFmpeg)

- Sending controller button packets continuously over a USB serial transmitter

- Editing and running macro scripts (JSON) with `if/while`, labels, variables

- Extending the system with custom commands, including running custom Python tools

#### Features

- Camera Preview
  - Lists cameras using FFmpeg DirectShow device enumeration

  - Streams video using FFmpeg rawvideo piping

  - Mouse hover shows pixel coordinates (x,y)

  - Click-to-copy coordinates for use in scripts

- Controller / Serial
  - Uses a [USB Wireless TX (Transmitter) by insideGadgets](https://shop.insidegadgets.com/product/usb-wireless-tx-transmitter/)

  - Connect to a COM port at 1,000,000 baud

  - Sends keep-alive packets at ~20 Hz (50 ms)

  - Pairing warm-up (neutral packets) on connect

  - Supports changing channels

  - Supports button press/hold behavior

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

#### Folder Layout
```
project/
  app.py                      # the single-file program
  scripts/                    # macro scripts (JSON)
    example.json
  py_scripts/                 # user python helpers for run_python
    my_tool.py
  bin/
    ffmpeg.exe                  # optional (recommended for packaging)
    *.dll
```
The app will create `scripts/` and `py_scripts/` if missing.

#### Requirements (development)

- Python 3.10+ recommended

- Dependencies:

  - numpy

  - pillow

  - pyserial

Install:
```bat
py -m pip install numpy pillow pyserial
```


FFmpeg:

- Recommended: place ffmpeg.exe alongside the app (or ensure ffmpeg is on PATH)

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


In `vision.py`, decode like this:
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
#### Adding Custom Commands

Commands live in `ScriptEngine._build_default_registry()` as `CommandSpec` entries. To add one:

1. Create a formatter 1 (pretty display)

2. Create an executor cmd_mycmd(ctx, c) (runtime behavior)

3. Register it via CommandSpec(...) with:

   - `group` and `order` for the Insert panel

   - `arg_schema` so the editor knows how to edit it

See the existing `find_color` implementation for a camera-reading example.

#### Packaging as an EXE (Windows)

Recommended: PyInstaller `--onedir` with bundled ffmpeg.

1. Ensure the code uses a bundled `ffmpeg.exe` if present (helper like `ffmpeg_path()`).

2. Build:
```bat
py -m pip install pyinstaller
pyinstaller --noconsole --onedir --clean --name ControllerMacroRunner --add-data "scripts;scripts" --add-data "py_scripts;py_scripts" main.py
```

Distribute the resulting `dist/ControllerMacroRunner/` folder as a zip.

#### Troubleshooting

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

#### Safety Notes

- `run_python` executes local code. Only run scripts you trust.

- Keep controller output neutral when not actively running actions.
