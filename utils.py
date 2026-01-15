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
        "j": "A",
        "k": "B",
        "u": "X",
        "i": "Y",
        "enter": "Start",
        "space": "Select",
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
    "theme": "auto",
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
        if isinstance(loaded.get("theme"), str):
            theme = normalize_theme_setting(loaded["theme"])
            result["theme"] = theme

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
    if value in ("auto", "dark", "light"):
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
    return setting
