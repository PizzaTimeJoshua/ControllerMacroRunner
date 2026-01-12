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
