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
# FFmpeg Download Support
# ----------------------------

# Using full-shared build which includes DLLs for codec support
# Note: .7z format requires 7-Zip for extraction
FFMPEG_URL = "https://www.gyan.dev/ffmpeg/builds/ffmpeg-release-full-shared.7z"
FFMPEG_DIR = "bin/ffmpeg"
FFMPEG_VERSION = "release-full-shared"


# ----------------------------
# Tesseract Download Support
# ----------------------------

# UB-Mannheim provides Windows installers - we extract with 7z to avoid elevation
TESSERACT_URL = "https://github.com/tesseract-ocr/tesseract/releases/download/5.5.0/tesseract-ocr-w64-setup-5.5.0.20241111.exe"
TESSERACT_DIR = "bin/Tesseract-OCR"
TESSERACT_VERSION = "5.5.0"


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


# ----------------------------
# 7-Zip Helper (needed for .7z and NSIS extraction)
# ----------------------------

def _get_7za_path(progress_callback=None) -> str | None:
    """Get path to 7z.exe, downloading if necessary.

    Downloads modern 7-Zip (v24.09) using a two-step process:
    1. Download 7zr.exe (standalone, can extract .7z)
    2. Use 7zr.exe to extract full 7z.exe + 7z.dll from the installer

    We need the full 7z.exe (not 7za.exe) for NSIS installer support.
    """
    import shutil
    import urllib.request
    import subprocess
    import tempfile

    # Check if 7z is on PATH
    seven_zip = shutil.which("7z") or shutil.which("7za")
    if seven_zip:
        return seven_zip

    # Check for bundled 7z.exe (full version with NSIS support)
    bin_dir = exe_dir_path("bin")
    bundled_7z = os.path.join(bin_dir, "7z.exe")
    if os.path.exists(bundled_7z):
        return bundled_7z

    # Download modern 7-Zip using two-step process
    try:
        os.makedirs(bin_dir, exist_ok=True)

        if progress_callback:
            progress_callback(0, "Downloading 7-Zip...")

        # Step 1: Download 7zr.exe (standalone reduced version, ~600KB)
        bundled_7zr = os.path.join(bin_dir, "7zr.exe")
        if not os.path.exists(bundled_7zr):
            url_7zr = "https://www.7-zip.org/a/7zr.exe"
            req = urllib.request.Request(url_7zr, headers={"User-Agent": "ControllerMacroRunner/1.0"})
            with urllib.request.urlopen(req, timeout=60) as response:
                with open(bundled_7zr, "wb") as out:
                    out.write(response.read())

        # Step 2: Download and extract full 7-Zip installer to get 7z.exe + 7z.dll
        # The x64 installer is a 7z archive that can be extracted with 7zr
        url_installer = "https://www.7-zip.org/a/7z2409-x64.exe"

        with tempfile.NamedTemporaryFile(suffix=".exe", delete=False) as tmp:
            tmp_path = tmp.name

        req = urllib.request.Request(url_installer, headers={"User-Agent": "ControllerMacroRunner/1.0"})
        with urllib.request.urlopen(req, timeout=60) as response:
            with open(tmp_path, "wb") as out:
                out.write(response.read())

        # Extract 7z.exe and 7z.dll from the installer
        with tempfile.TemporaryDirectory() as extract_dir:
            result = subprocess.run(
                [bundled_7zr, "x", "-y", f"-o{extract_dir}", tmp_path],
                capture_output=True,
                timeout=60
            )

            if result.returncode == 0:
                # Copy 7z.exe and 7z.dll to bin folder
                for filename in ["7z.exe", "7z.dll"]:
                    src = os.path.join(extract_dir, filename)
                    if os.path.exists(src):
                        shutil.copy2(src, os.path.join(bin_dir, filename))

        os.remove(tmp_path)

        if os.path.exists(bundled_7z):
            return bundled_7z

    except Exception:
        pass

    return None


# ----------------------------
# FFmpeg Availability and Download
# ----------------------------

def is_ffmpeg_available() -> bool:
    """Check if FFmpeg is installed locally or on PATH."""
    import shutil

    # Check local installation
    exe_local = exe_dir_path(f"{FFMPEG_DIR}/ffmpeg.exe")
    if os.path.exists(exe_local):
        return True

    # Check bundled (PyInstaller)
    bundled = resource_path(f"{FFMPEG_DIR}/ffmpeg.exe")
    if os.path.exists(bundled):
        return True

    # Check PATH
    return shutil.which("ffmpeg") is not None


def get_ffmpeg_dir() -> str:
    """Get the directory where FFmpeg should be installed."""
    return exe_dir_path(FFMPEG_DIR)


def download_ffmpeg(progress_callback=None) -> tuple[bool, str]:
    """
    Download and extract FFmpeg full-shared package.

    Args:
        progress_callback: Optional function(percent: int, status: str) called during download

    Returns:
        Tuple of (success: bool, message: str)
    """
    import urllib.request
    import tempfile
    import shutil
    import subprocess

    target_dir = get_ffmpeg_dir()

    try:
        if progress_callback:
            progress_callback(0, "Checking for 7-Zip...")

        # Get 7za for extraction (.7z files require 7-Zip)
        seven_zip = _get_7za_path(progress_callback)
        if not seven_zip:
            return False, "Could not find or download 7-Zip. Please install 7-Zip and try again."

        if progress_callback:
            progress_callback(5, "Connecting to gyan.dev...")

        with tempfile.NamedTemporaryFile(suffix=".7z", delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            req = urllib.request.Request(
                FFMPEG_URL,
                headers={"User-Agent": "ControllerMacroRunner/1.0"}
            )

            with urllib.request.urlopen(req, timeout=300) as response:
                total_size = int(response.headers.get("Content-Length", 0))
                downloaded = 0
                chunk_size = 65536

                if progress_callback:
                    progress_callback(5, "Downloading FFmpeg...")

                with open(tmp_path, "wb") as out_file:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        out_file.write(chunk)
                        downloaded += len(chunk)

                        if total_size > 0 and progress_callback:
                            percent = 5 + int((downloaded / total_size) * 75)
                            mb_done = downloaded / (1024 * 1024)
                            mb_total = total_size / (1024 * 1024)
                            progress_callback(percent, f"Downloading... {mb_done:.1f}/{mb_total:.1f} MB")

            if progress_callback:
                progress_callback(85, "Extracting FFmpeg...")

            # Remove existing installation
            if os.path.exists(target_dir):
                shutil.rmtree(target_dir)
            os.makedirs(target_dir, exist_ok=True)

            # Extract to temp directory first using 7z
            with tempfile.TemporaryDirectory() as extract_dir:
                result = subprocess.run(
                    [seven_zip, "x", "-y", f"-o{extract_dir}", tmp_path],
                    capture_output=True,
                    timeout=300
                )

                if result.returncode != 0:
                    stderr = result.stderr.decode('utf-8', errors='ignore')
                    return False, f"Extraction failed: {stderr}"

                # Find the bin folder containing ffmpeg.exe
                bin_folder = None
                for root, dirs, files in os.walk(extract_dir):
                    if "ffmpeg.exe" in files:
                        bin_folder = root
                        break

                if not bin_folder:
                    return False, "Extraction failed: ffmpeg.exe not found in archive"

                # Copy all files from bin folder to target directory
                for filename in os.listdir(bin_folder):
                    src_path = os.path.join(bin_folder, filename)
                    if os.path.isfile(src_path):
                        shutil.copy2(src_path, os.path.join(target_dir, filename))

            if progress_callback:
                progress_callback(100, "Done!")

            # Verify installation
            ffmpeg_exe = os.path.join(target_dir, "ffmpeg.exe")
            if not os.path.exists(ffmpeg_exe):
                return False, "Extraction failed: ffmpeg.exe not found"

            return True, "FFmpeg installed successfully"

        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    except urllib.error.URLError as e:
        return False, f"Download failed: {e.reason}"
    except urllib.error.HTTPError as e:
        return False, f"Download failed: HTTP {e.code}"
    except subprocess.TimeoutExpired:
        return False, "Extraction timed out"
    except Exception as e:
        return False, f"Installation failed: {e}"


# ----------------------------
# Tesseract Availability and Download
# ----------------------------

def is_tesseract_available() -> bool:
    """Check if Tesseract is installed locally or on PATH."""
    import shutil

    # Check local installation
    exe_local = exe_dir_path(f"{TESSERACT_DIR}/tesseract.exe")
    if os.path.exists(exe_local):
        return True

    # Check bundled (PyInstaller)
    bundled = resource_path(f"{TESSERACT_DIR}/tesseract.exe")
    if os.path.exists(bundled):
        return True

    # Check PATH
    return shutil.which("tesseract") is not None


def get_tesseract_dir() -> str:
    """Get the directory where Tesseract should be installed."""
    return exe_dir_path(TESSERACT_DIR)


def download_tesseract(progress_callback=None) -> tuple[bool, str]:
    """
    Download and install Tesseract OCR by extracting the NSIS installer.

    Args:
        progress_callback: Optional function(percent: int, status: str) called during download

    Returns:
        Tuple of (success: bool, message: str)
    """
    import urllib.request
    import tempfile
    import shutil
    import subprocess

    target_dir = get_tesseract_dir()

    try:
        os.makedirs(os.path.dirname(target_dir), exist_ok=True)

        if progress_callback:
            progress_callback(0, "Checking for 7-Zip...")

        # Get 7za for extraction (need full version with DLLs for NSIS support)
        seven_zip = _get_7za_path(progress_callback)
        if not seven_zip:
            return False, "Could not find or download 7-Zip. Please install 7-Zip and try again."

        if progress_callback:
            progress_callback(5, "Connecting to GitHub...")

        with tempfile.NamedTemporaryFile(suffix=".exe", delete=False) as tmp_file:
            tmp_path = tmp_file.name

        try:
            req = urllib.request.Request(
                TESSERACT_URL,
                headers={"User-Agent": "ControllerMacroRunner/1.0"}
            )

            with urllib.request.urlopen(req, timeout=120) as response:
                total_size = int(response.headers.get("Content-Length", 0))
                downloaded = 0
                chunk_size = 65536

                if progress_callback:
                    progress_callback(5, "Downloading Tesseract...")

                with open(tmp_path, "wb") as out_file:
                    while True:
                        chunk = response.read(chunk_size)
                        if not chunk:
                            break
                        out_file.write(chunk)
                        downloaded += len(chunk)

                        if total_size > 0 and progress_callback:
                            percent = 5 + int((downloaded / total_size) * 75)
                            mb_done = downloaded / (1024 * 1024)
                            mb_total = total_size / (1024 * 1024)
                            progress_callback(percent, f"Downloading... {mb_done:.1f}/{mb_total:.1f} MB")

            if progress_callback:
                progress_callback(85, "Extracting Tesseract...")

            # Remove existing installation
            if os.path.exists(target_dir):
                shutil.rmtree(target_dir)
            os.makedirs(target_dir, exist_ok=True)

            # Extract NSIS installer using 7z (requires full 7za with DLLs)
            result = subprocess.run(
                [seven_zip, "x", "-y", f"-o{target_dir}", tmp_path],
                capture_output=True,
                timeout=300
            )

            if result.returncode != 0:
                stdout = result.stdout.decode('utf-8', errors='ignore').strip()
                stderr = result.stderr.decode('utf-8', errors='ignore').strip()
                error_msg = stderr or stdout or f"Exit code {result.returncode}"
                return False, f"Extraction failed: {error_msg}"

            # Verify tesseract.exe exists
            tesseract_exe = os.path.join(target_dir, "tesseract.exe")
            if not os.path.exists(tesseract_exe):
                return False, "Extraction failed: tesseract.exe not found"

            # Clean up NSIS metadata files
            for cleanup_item in ["$PLUGINSDIR", "$TEMP", "Uninstall.exe", "uninstall.exe"]:
                cleanup_path = os.path.join(target_dir, cleanup_item)
                if os.path.isdir(cleanup_path):
                    shutil.rmtree(cleanup_path, ignore_errors=True)
                elif os.path.isfile(cleanup_path):
                    try:
                        os.remove(cleanup_path)
                    except Exception:
                        pass

            if progress_callback:
                progress_callback(100, "Done!")

            # Verify installation
            tesseract_exe = os.path.join(target_dir, "tesseract.exe")
            if not os.path.exists(tesseract_exe):
                return False, "Extraction failed: tesseract.exe not found"

            return True, f"Tesseract {TESSERACT_VERSION} installed successfully"

        finally:
            if os.path.exists(tmp_path):
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass

    except urllib.error.URLError as e:
        return False, f"Download failed: {e.reason}"
    except urllib.error.HTTPError as e:
        return False, f"Download failed: HTTP {e.code}"
    except subprocess.TimeoutExpired:
        return False, "Extraction timed out"
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
