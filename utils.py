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


def ffmpeg_path() -> str:
    """Get path to ffmpeg executable."""
    # Prefer bundled ffmpeg.exe
    bundled = resource_path("bin/ffmpeg.exe")
    if os.path.exists(bundled):
        return bundled
    return "ffmpeg"  # fallback to PATH


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
