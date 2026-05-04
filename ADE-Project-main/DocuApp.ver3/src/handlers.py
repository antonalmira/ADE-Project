import os
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QListWidgetItem, QFileDialog, QInputDialog
from PyQt5.QtGui import QFont
from list_updater import performancedata_testnames, waveform_testnames, save_performance_dict, save_waveform_dict

class AddItemDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, is_waveform=False):
        super().__init__(parent)
        self.setWindowTitle("Add Waveform Item" if is_waveform else "Add Performance Item")
        self.setMinimumWidth(450)
        
        self.layout = QtWidgets.QFormLayout(self)
        
        # Test Type Input
        self.test_type_input = QtWidgets.QLineEdit(self)
        self.test_type_input.setToolTip("Please enter the full name of your Test Type")
        self.test_type_input.setPlaceholderText("e.g., Start-up Condition")
        
        # File Name Prefix Input
        self.prefix_input = QtWidgets.QLineEdit(self)
        self.prefix_input.setToolTip("Please enter the word(s) before the first underscore '_' of your file name so that TARDIS can identify your files.")
        self.prefix_input.setPlaceholderText("e.g., output start-up")
        
        self.layout.addRow("Test Type:", self.test_type_input)
        self.layout.addRow("Test File Name:", self.prefix_input)
        
        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)
        
        self.setStyleSheet("""
            QDialog { background-color: #f5f5f5; }
            QLabel { color: #000000; font-weight: bold; }
            QLineEdit { border: 1px solid #b0b0b0; border-radius: 4px; padding: 6px; background: white; color: black; }
            QPushButton { background-color: #0085ca; color: #ffffff; border-radius: 4px; padding: 6px 20px; font-weight: bold; }
            QPushButton:hover { background-color: #3c649f; }
        """)

    def get_data(self):
        return self.test_type_input.text().strip(), self.prefix_input.text().strip().lower()


def select_template_file(app):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    default_path = os.path.join(current_dir, "templates")
    
    if not os.path.exists(default_path):
        default_path = os.path.join(os.path.expanduser("~"), "Desktop")

    file, _ = QFileDialog.getOpenFileName(
        app, 
        "Select Word Template", 
        default_path,
        "Word Documents (*.docx)"
    )
    
    if file:
        app.template_path_display.setText(file)
        app.template_path_display.setToolTip(file)
        print(f"Selected template: {file}")

def select_performance_folder(app):
    folder = QFileDialog.getExistingDirectory(app, "Select Performance Data Folder")
    if folder:
        app.performancedata_path.setText(folder)
        print(f"Selected performance folder: {folder}")

def select_waveform_folder(app):
    folder = QFileDialog.getExistingDirectory(app, "Select Waveforms Folder")
    if folder:
        app.waveforms_path.setText(folder)
        print(f"Selected waveform folder: {folder}")

def add_performance_item(app):
    dialog = AddItemDialog(app, is_waveform=False)
    if dialog.exec_():
        item_name, filename_prefix = dialog.get_data()
        
        if item_name and filename_prefix:
            performancedata_testnames[filename_prefix] = item_name
            save_performance_dict()
            
            item = QListWidgetItem(item_name)
            item_font = QFont()
            item_font.setPointSize(12)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.performancedata_list.addItem(item)
            print(f"Added performance item: {item_name} with prefix: {filename_prefix}")

def add_waveform_item(app):
    dialog = AddItemDialog(app, is_waveform=True)
    if dialog.exec_():
        item_name, filename_prefix = dialog.get_data()
        
        if item_name and filename_prefix:
            waveform_testnames[filename_prefix] = item_name
            save_waveform_dict()
            
            item = QListWidgetItem(item_name)
            item_font = QFont()
            item_font.setPointSize(12)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.waveforms_list.addItem(item)
            print(f"Added waveform item: {item_name} with prefix: {filename_prefix}")

def delete_performance_item(app):
    selected_items = app.performancedata_list.selectedItems()
    for item in selected_items:
        item_text = item.text()
        app.performancedata_list.takeItem(app.performancedata_list.row(item))
        for key, value in list(performancedata_testnames.items()):
            if value == item_text:
                del performancedata_testnames[key]
                break
        save_performance_dict()
        print(f"Permanently deleted performance item: {item_text}")

def delete_waveform_item(app):
    selected_items = app.waveforms_list.selectedItems()
    for item in selected_items:
        item_text = item.text()
        app.waveforms_list.takeItem(app.waveforms_list.row(item))
        for key, value in list(waveform_testnames.items()):
            if value == item_text:
                del waveform_testnames[key]
                break
        save_waveform_dict()
        print(f"Permanently deleted waveform item: {item_text}")

def toggle_maximize(app):
    if app.isMaximized():
        app.showNormal()
    else:
        app.showMaximized()