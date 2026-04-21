import os
import shutil
import sys
from pathlib import Path
import time

def ensure_directory(path):
    os.makedirs(path, exist_ok=True)

def remove_directory(path):
    if os.path.exists(path):
        shutil.rmtree(path)

def get_default_base_folder(folder_path):
    if folder_path and os.path.isdir(folder_path):
        return folder_path
    return str(Path.home())

def log_message(message):
    timestamp = time.strftime('%H:%M:%S', time.localtime())
    print(f"[{timestamp}] {message}")

def get_resource_path(relative_path):
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)