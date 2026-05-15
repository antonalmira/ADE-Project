import os
import shutil
import sys
from pathlib import Path
import time
from PyQt5.QtWidgets import QMessageBox

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

def show_popup(parent, title, text, icon_type="info"):
    msg_box = QMessageBox(parent)
    msg_box.setStyleSheet("""a  
                          QMessageBox {
                              background-color: #f5f5f5;
                              }
                              QLabel {
                                  color: #000000;
                                  font-size: 13px;
                                  font-weight: bold;
                                  
                              }QPushButton {
                                  background-color: #0085ca;
                                  color: #ffffff;
                                  border-radius: 4px;
                                  padding: 6px 20px;
                                  font-weight: bold;
                                  min-width: 60px;
                              }
                              QPushButton:hover {
                                  background-color: #3c649f;
                              }
                          """)
    msg_box.setWindowTitle(title)
    msg_box.setText(text)
    
    if icon_type == "info":
        msg_box.setIcon(QMessageBox.Information)
    elif icon_type == "warning":
        msg_box.setIcon(QMessageBox.Warning)
    elif icon_type == "error":
        msg_box.setIcon(QMessageBox.Critical)
    
    msg_box.exec_()
    