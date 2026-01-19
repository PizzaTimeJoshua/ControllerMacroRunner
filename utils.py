"""
Utility functions for Controller Macro Runner.
"""
import os
import sys
import re


def resource_path(rel_path: str) -> str:
    """Get absolute path to resource, works for dev and PyInstaller."""
    base = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base, rel_path)


def exe_dir_path(rel_path: str) -> str:
    """Get absolute path relative to the executable's directory.

    When running as a PyInstaller bundle, this returns the path relative to
    where the .exe is located (not the temp extraction folder).
    When running from source, this returns the path relative to the script directory.
    """
    if getattr(sys, "frozen", False):
        # Running as PyInstaller bundle - use executable's directory
        base = os.path.dirname(sys.executable)
    else:
        # Running from source - use current working directory
        base = os.path.abspath(".")
    return os.path.join(base, rel_path)


def ffmpeg_path() -> str:
    """Get path to ffmpeg executable."""
    # Check 1: PyInstaller temp folder (if bundled with --add-data)
    bundled = resource_path("bin/ffmpeg/ffmpeg.exe")
    if os.path.exists(bundled):
        return bundled

    # Check 2: Directory next to the executable (for external bin folder)
    exe_local = exe_dir_path("bin/ffmpeg/ffmpeg.exe")
    if os.path.exists(exe_local):
        return exe_local

    return "ffmpeg"  # fallback to PATH


def tesseract_path() -> str:
    """Get path to tesseract executable."""
    # Check 1: PyInstaller temp folder (if bundled with --add-data)
    bundled = resource_path("bin/Tesseract-OCR/tesseract.exe")
    if os.path.exists(bundled):
        return bundled

    # Check 2: Directory next to the executable (for external bin folder)
    exe_local = exe_dir_path("bin/Tesseract-OCR/tesseract.exe")
    if os.path.exists(exe_local):
        return exe_local

    return "tesseract"  # fallback to PATH


# ----------------------------
# Embedded Python Support
# ----------------------------

PYTHON_EMBED_URL = "https://www.python.org/ftp/python/3.12.8/python-3.12.8-embed-amd64.zip"
PYTHON_EMBED_DIR = "bin/python-embedded"
PYTHON_EMBED_VERSION = "3.12.8"


def python_path() -> str:
    """Get path to Python executable, preferring embedded version."""
    # Check 1: PyInstaller temp folder (if bundled with --add-data)
    bundled = resource_path(f"{PYTHON_EMBED_DIR}/python.exe")
    if os.path.exists(bundled):
        return bundled

    # Check 2: Directory next to the executable (for external bin folder)
    exe_local = exe_dir_path(f"{PYTHON_EMBED_DIR}/python.exe")
    if os.path.exists(exe_local):
        return exe_local

    # Fallback to current interpreter (works in dev, may not work in frozen exe)
    return sys.executable


def is_embedded_python_available() -> bool:
    """Check if embedded Python is installed locally."""
    exe_local = exe_dir_path(f"{PYTHON_EMBED_DIR}/python.exe")
    return os.path.exists(exe_local)


def is_python_available() -> bool:
    """Check if any Python is available for running scripts."""
    # If we're not frozen (running from source), Python is always available
    if not getattr(sys, "frozen", False):
        return True
    # If frozen, check for embedded Python
    return is_embedded_python_available()


def get_embedded_python_dir() -> str:
    """Get the directory where embedded Python should be installed."""
    return exe_dir_path(PYTHON_EMBED_DIR)


def download_embedded_python(progress_callback=None) -> tuple[bool, str]:
    """
    Download and extract Python embeddable package.

    Args:
        progress_callback: Optional function(percent: int, status: str) called during download

    Returns:
        Tuple of (success: bool, message: str)
    """
    import urllib.request
    import zipfile
    import tempfile
    import shutil

    target_dir = get_embedded_python_dir()

    try:
        # Create target directory if needed
        os.makedirs(target_dir, exist_ok=True)

        # Download to temp file
        if progress_callback:
            progress_callback(0, "Connecting to python.org...")

        # Create a temp file for the download
        with tempfile.NamedTemporaryFile(suffix=".zip", delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            # Open URL and get content length
            req = urllib.request.Request(
                PYTHON_EMBED_URL,
                headers={"User-Agent": "ControllerMacroRunner/1.0"}
            )

            with urllib.request.urlopen(req, timeout=30) as response:
                total_size = int(response.headers.get("Content-Length", 0))
                downloaded = 0
                chunk_size = 65536  # 64KB chunks

                if progress_callback:
                    progress_callback(0, f"Downloading Python {PYTHON_EMBED_VERSION}...")

                with open(tmp_path, "wb") as out_file:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        out_file.write(chunk)
                        downloaded += len(chunk)

                        if total_size > 0 and progress_callback:
                            percent = int((downloaded / total_size) * 80)  # 0-80% for download
                            mb_done = downloaded / (1024 * 1024)
                            mb_total = total_size / (1024 * 1024)
                            progress_callback(percent, f"Downloading... {mb_done:.1f}/{mb_total:.1f} MB")

            # Extract zip
            if progress_callback:
                progress_callback(85, "Extracting Python...")

            # Remove existing installation if present
            if os.path.exists(target_dir):
                shutil.rmtree(target_dir)
            os.makedirs(target_dir, exist_ok=True)

            with zipfile.ZipFile(tmp_path, "r") as zip_ref:
                zip_ref.extractall(target_dir)

            if progress_callback:
                progress_callback(100, "Done!")

            # Verify installation
            python_exe = os.path.join(target_dir, "python.exe")
            if not os.path.exists(python_exe):
                return False, "Extraction failed: python.exe not found"

            return True, f"Python {PYTHON_EMBED_VERSION} installed successfully"

        finally:
            # Clean up temp file
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    except urllib.error.URLError as e:
        return False, f"Download failed: {e.reason}"
    except urllib.error.HTTPError as e:
        return False, f"Download failed: HTTP {e.code}"
    except zipfile.BadZipFile:
        return False, "Downloaded file is corrupted"
    except Exception as e:
        return False, f"Installation failed: {e}"


def safe_script_filename(name: str) -> str:
    """Sanitize a script filename for safe filesystem use."""
    name = name.strip()
    name = re.sub(r'[<>:"/\\|?*\x00-\x1F]', "_", name)  # Windows-illegal chars
    name = re.sub(r"\s+", " ", name).strip()
    if not name:
        return ""
    if not name.lower().endswith(".json"):
        name += ".json"
    return name


def list_python_files():
    """List Python files in the py_scripts folder."""
    folder = "py_scripts"
    if not os.path.isdir(folder):
        return []
    return sorted([n for n in os.listdir(folder) if n.lower().endswith(".py")])


def list_script_files():
    """List script JSON files in the scripts folder."""
    folder = "scripts"
    if not os.path.isdir(folder):
        return []
    return sorted([n for n in os.listdir(folder) if n.lower().endswith(".json")])


def list_com_ports():
    """List available COM ports."""
    from serial.tools import list_ports
    return [p.device for p in list_ports.comports()]


# ----------------------------
# Settings Management
# ----------------------------

SETTINGS_FILE = "settings.json"

DEFAULT_SETTINGS = {
    "keybindings": {
        "w": "Up",
        "a": "Left",
        "s": "Down",
        "d": "Right",
        "enter": "A",
        "shift_r": "B",
        "apostrophe": "X",
        "slash": "Y",
        "equal": "Start",
        "minus": "Select",
        "q": "L",
        "e": "R",
        "up": "Left Stick Up",
        "down": "Left Stick Down",
        "left": "Left Stick Left",
        "right": "Left Stick Right",
    },
    "threeds": {
        "ip": "192.168.1.1",
        "port": 4950,
    },
    "discord": {
        "webhook_url": "",
        "user_id": "",
    },
    "camera_ratio": "3:2 (GBA)",
    "default_camera_device": "",
    "default_audio_input_device": "",
    "default_audio_output_device": "",
    "default_com_port": "",
    "theme": "auto",
    "custom_theme": {
        "bg": "#0f1612",
        "panel": "#16201a",
        "text": "#d9e2dc",
        "muted": "#9fb0a4",
        "border": "#223128",
        "accent": "#7bd88f",
        "entry_bg": "#1b2620",
        "button_bg": "#1f2c25",
        "button_fg": "#d9e2dc",
        "select_bg": "#2a3a31",
        "select_fg": "#f5faf7",
        "tree_bg": "#141e18",
        "tree_fg": "#d9e2dc",
        "text_bg": "#0f1713",
        "text_fg": "#d9e2dc",
        "insert_fg": "#f5faf7",
        "text_sel_bg": "#2a3a31",
        "text_sel_fg": "#f5faf7",
        "pane_bg": "#0c120f",
        "ip_bg": "#27362e",
        "comment_fg": "#7bd88f",
        "variable_fg": "#8cc9a4",
        "math_fg": "#f2c374",
        "selected_bg": "#1f2b24",
    },
    "confirm_delete": True,
}


def load_settings() -> dict:
    """Load settings from settings.json, returning defaults if file doesn't exist."""
    import json
    import copy
    settings_path = exe_dir_path(SETTINGS_FILE)

    if not os.path.exists(settings_path):
        return copy.deepcopy(DEFAULT_SETTINGS)

    try:
        with open(settings_path, "r", encoding="utf-8") as f:
            loaded = json.load(f)

        # Merge with defaults to ensure all keys exist
        result = copy.deepcopy(DEFAULT_SETTINGS)
        if "keybindings" in loaded:
            result["keybindings"] = loaded["keybindings"]
        if "threeds" in loaded:
            result["threeds"] = {**DEFAULT_SETTINGS["threeds"], **loaded["threeds"]}
        if "discord" in loaded:
            result["discord"] = {**DEFAULT_SETTINGS["discord"], **loaded["discord"]}
        if isinstance(loaded.get("camera_ratio"), str):
            result["camera_ratio"] = loaded["camera_ratio"]
        if isinstance(loaded.get("default_camera_device"), str):
            result["default_camera_device"] = loaded["default_camera_device"]
        if isinstance(loaded.get("default_audio_input_device"), str):
            result["default_audio_input_device"] = loaded["default_audio_input_device"]
        if isinstance(loaded.get("default_audio_output_device"), str):
            result["default_audio_output_device"] = loaded["default_audio_output_device"]
        if isinstance(loaded.get("default_com_port"), str):
            result["default_com_port"] = loaded["default_com_port"]
        if isinstance(loaded.get("theme"), str):
            theme = normalize_theme_setting(loaded["theme"])
            result["theme"] = theme
        if isinstance(loaded.get("custom_theme"), dict):
            result["custom_theme"] = loaded["custom_theme"]
        if isinstance(loaded.get("confirm_delete"), bool):
            result["confirm_delete"] = loaded["confirm_delete"]
        return result
    except Exception:
        return copy.deepcopy(DEFAULT_SETTINGS)


def save_settings(settings: dict) -> bool:
    """Save settings to settings.json. Returns True on success."""
    import json
    settings_path = exe_dir_path(SETTINGS_FILE)

    try:
        with open(settings_path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2, ensure_ascii=False)
        return True
    except Exception:
        return False


def get_default_keybindings() -> dict:
    """Return a copy of the default keybindings."""
    return DEFAULT_SETTINGS["keybindings"].copy()


def normalize_theme_setting(value: str) -> str:
    """Normalize theme setting to 'auto', 'dark', or 'light'."""
    value = (value or "").strip().lower()
    if value in ("auto", "dark", "light", "custom"):
        return value
    return "auto"


def get_system_theme() -> str:
    """Return the current system theme as 'dark' or 'light'."""
    if sys.platform != "win32":
        return "light"
    try:
        import winreg
        key_path = r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
        return "light" if int(value) else "dark"
    except Exception:
        return "light"


def resolve_theme_mode(setting: str) -> str:
    """Resolve a theme setting into 'dark' or 'light'."""
    setting = normalize_theme_setting(setting)
    if setting == "auto":
        return get_system_theme()
    if setting == "custom":
        return "custom"
    return setting
