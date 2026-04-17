import os
import shutil
import sys
from pathlib import Path
import time

def ensure_directory(path):
    """Ensure a directory exists, creating it if necessary."""
    os.makedirs(path, exist_ok=True)
    print(f"Ensured directory exists: {path}")

def remove_directory(path):
    """Remove a directory and its contents if it exists."""
    if os.path.exists(path):
        shutil.rmtree(path)
        print(f"Removed directory: {path}")

def get_default_base_folder(folder_path):
    """Return the provided folder path if valid, else return user's home directory."""
    if folder_path and os.path.isdir(folder_path):
        return folder_path
    print(f"Invalid or unset folder path: {folder_path}, using home directory")
    return str(Path.home())

def log_message(message):
    """Log a message with a timestamp."""
    timestamp = time.strftime('%H:%M:%S %Z on %Y-%m-%d', time.localtime())
    print(f"[{timestamp}] {message}")


def get_resource_path(relative_path: str) -> str:
    """Get absolute path to resource, works for dev and PyInstaller.

    Tries several sensible base paths so resources placed at project root
    (next to src/) or inside the extracted bundle are found.
    """
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))

    # Primary candidate: base_path + relative_path
    candidate = os.path.abspath(os.path.join(base_path, relative_path))
    if os.path.exists(candidate):
        return candidate

    # Try one level up (project root when utils.py is inside src/)
    parent_candidate = os.path.abspath(os.path.join(os.path.dirname(base_path), relative_path))
    if os.path.exists(parent_candidate):
        return parent_candidate

    # Last-resort: return the primary candidate (even if it doesn't exist yet)
    return candidate