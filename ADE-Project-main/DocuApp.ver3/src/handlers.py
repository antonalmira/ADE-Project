import os
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QListWidgetItem, QFileDialog, QInputDialog
from PyQt5.QtGui import QFont
from list_updater import performancedata_testnames, waveform_testnames, save_performance_dict, save_waveform_dict

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
            item_font.setPointSize(16)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.performancedata_list.addItem(item)
            print(f"Added performance item: {item_name} with filename prefix: {filename_prefix}")
        else:
            print("No filename prefix entered; performance item not added")
    else:
        print("No test name entered; performance item not added")

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
            item_font.setPointSize(16)
            item.setFont(item_font)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked)
            app.waveforms_list.addItem(item)
            print(f"Added waveform item: {item_name} with filename prefix: {filename_prefix}")
        else:
            print("No filename prefix entered; waveform item not added")
    else:
        print("No test name entered; waveform item not added")

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

def select_generate_document_folder(app):
    folder = QFileDialog.getExistingDirectory(app, "Select Output Folder for Generated Document")
    if folder:
        app.generated_document_path.setText(folder)
    print(f"Selected output folder for generated document: {folder}")

def select_update_document(app):
    file, _ = QFileDialog.getOpenFileName(app, "Select Document to Update", "", "Word Documents (*.docx)")
    if file:
        app.update_document_path.setText(os.path.basename(file))
        app.update_document_path.setToolTip(file)
    print(f"Selected document to update: {file}")

def toggle_maximize(app):
    if app.isMaximized():
        app.showNormal()
    else:
        app.showMaximized()
    print(f"Toggled maximize state to {'Normal' if not app.isMaximized() else 'Maximized'}")