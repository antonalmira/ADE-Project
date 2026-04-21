import os
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QListWidgetItem, QFileDialog, QInputDialog
from PyQt5.QtGui import QFont
from list_updater import performancedata_testnames, waveform_testnames, save_performance_dict, save_waveform_dict

def select_template_file(app):
    """Opens a file dialog starting automatically in the project's template folder"""
    
    # 1. Get the directory where handlers.py is located
    current_dir = os.path.dirname(os.path.abspath(__file__))
    
    # 2. Point to the templates folder (assuming it's in the same folder as your scripts)
    # If templates is one level up, use: os.path.join(current_dir, "..", "templates")
    default_path = os.path.join(current_dir, "templates")
    
    # Ensure the path exists, otherwise default to Desktop
    if not os.path.exists(default_path):
        default_path = os.path.join(os.path.expanduser("~"), "Desktop")

    file, _ = QFileDialog.getOpenFileName(
        app, 
        "Select Word Template", 
        default_path, # This is the starting directory
        "Word Documents (*.docx)"
    )
    
    if file:
        app.template_path_display.setText(file)
        app.template_path_display.setToolTip(file)
        print(f"Selected template: {file}")

def select_performance_folder(app):
    """Opens a folder dialog for Performance Data"""
    folder = QFileDialog.getExistingDirectory(app, "Select Performance Data Folder")
    if folder:
        app.performancedata_path.setText(folder)
        print(f"Selected performance folder: {folder}")

def select_waveform_folder(app):
    """Opens a folder dialog for Waveforms"""
    folder = QFileDialog.getExistingDirectory(app, "Select Waveforms Folder")
    if folder:
        app.waveforms_path.setText(folder)
        print(f"Selected waveform folder: {folder}")

def add_performance_item(app):
    item_name, ok = QInputDialog.getText(app, "Add Performance Item", "Enter the test name:")
    if ok and item_name.strip():
        filename_prefix, ok = QInputDialog.getText(app, "Filename Prefix", f"Enter the first two words of the filename for '{item_name}' (e.g., 'efficiency test'):")
        if ok and filename_prefix.strip():
            # Add to dictionary permanently
            performancedata_testnames[filename_prefix.lower()] = item_name.strip()
            save_performance_dict()
            # Add to list with consistent formatting
            item = QListWidgetItem(item_name.strip())
            item_font = QFont()
            item_font.setPointSize(12)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.performancedata_list.addItem(item)
            print(f"Added performance item: {item_name}")

def add_waveform_item(app):
    item_name, ok = QInputDialog.getText(app, "Add Waveform Item", "Enter the test name:")
    if ok and item_name.strip():
        filename_prefix, ok = QInputDialog.getText(app, "Filename Prefix", f"Enter the first two words of the filename for '{item_name}' (e.g., 'cc load'):")
        if ok and filename_prefix.strip():
            # Add to dictionary permanently
            waveform_testnames[filename_prefix.lower()] = item_name.strip()
            save_waveform_dict()
            # Add to list with consistent formatting
            item = QListWidgetItem(item_name.strip())
            item_font = QFont()
            item_font.setPointSize(12)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.waveforms_list.addItem(item)
            print(f"Added waveform item: {item_name}")

def delete_performance_item(app):
    selected_items = app.performancedata_list.selectedItems()
    for item in selected_items:
        item_text = item.text()
        app.performancedata_list.takeItem(app.performancedata_list.row(item))
        # Remove from dictionary permanently
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
        # Remove from dictionary permanently
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