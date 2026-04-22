from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QListWidgetItem, QMessageBox
import os
import re
import json
from pathlib import Path
from utils import get_resource_path, ensure_directory, show_popup

# Default dictionaries
DEFAULT_PERFORMANCE = {
    "efficiency test": "Efficiency Test",
    "line regulation": "Line Regulation Test",
    "load regulation": "Load Regulation",
    "no load": "No load Input Power Test",
    "power factor": "Power Factor",
    "load transient": "Load Transient Response",
    "output voltage": "Output Voltage Distortion",
    "line input": "Line Input Harmonic Content",
    "led dimming": "LED Dimming Test",
    "thermal test": "Thermal Test",
    "standby input": "Standby Input Power"
}

DEFAULT_WAVEFORM = {
    "cc load": "CC Load Test",
    "output ripple": "Output Ripple Measurements",
    "load transient": "Load Transient Response",
    "switching waveforms": "Switching Waveforms",
    "recovery test": "Brown-In/Brown-Out Recovery Test",
    "start-up condition": "Start-up Condition",
    "steady state": "Steady State Condition",
    "transient load": "Transient Load Condition"
}

# Load from JSON if exists, else use defaults
# Load JSON mappings from a user-writable config location when possible.
# Default resource files are bundled with the app; user copies live under ~/.docuapp/.
config_dir = os.path.join(str(Path.home()), '.docuapp')
ensure_directory(config_dir)

# Performance dict: prefer user copy, fallback to bundled resource, otherwise defaults
user_perf_path = os.path.join(config_dir, 'performancedata_testnames.json')
bundled_perf_path = get_resource_path('performancedata_testnames.json')
performancedata_testnames = DEFAULT_PERFORMANCE.copy()
if os.path.exists(user_perf_path):
    try:
        with open(user_perf_path, 'r', encoding='utf-8') as f:
            performancedata_testnames = json.load(f)
        print("Loaded performance dict from user config")
    except Exception:
        print("Failed to load user performance dict, using defaults")
elif os.path.exists(bundled_perf_path):
    try:
        with open(bundled_perf_path, 'r', encoding='utf-8') as f:
            performancedata_testnames = json.load(f)
        print("Loaded performance dict from bundled resource")
    except Exception:
        print("Failed to load bundled performance dict, using defaults")

# Waveform dict: same policy
user_waveform_path = os.path.join(config_dir, 'waveform_testnames.json')
bundled_waveform_path = get_resource_path('waveform_testnames.json')
waveform_testnames = DEFAULT_WAVEFORM.copy()
if os.path.exists(user_waveform_path):
    try:
        with open(user_waveform_path, 'r', encoding='utf-8') as f:
            waveform_testnames = json.load(f)
        print("Loaded waveform dict from user config")
    except Exception:
        print("Failed to load user waveform dict, using defaults")
elif os.path.exists(bundled_waveform_path):
    try:
        with open(bundled_waveform_path, 'r', encoding='utf-8') as f:
            waveform_testnames = json.load(f)
        print("Loaded waveform dict from bundled resource")
    except Exception:
        print("Failed to load bundled waveform dict, using defaults")

def save_performance_dict():
    try:
        with open(user_perf_path, 'w', encoding='utf-8') as f:
            json.dump(performancedata_testnames, f, indent=2)
        print(f"Saved performance dict to {user_perf_path}")
    except Exception as e:
        print(f"Failed to save performance dict: {e}")

def save_waveform_dict():
    try:
        with open(user_waveform_path, 'w', encoding='utf-8') as f:
            json.dump(waveform_testnames, f, indent=2)
        print(f"Saved waveform dict to {user_waveform_path}")
    except Exception as e:
        print(f"Failed to save waveform dict: {e}")

def update_available_data_list(app):
    # Clear checklists and available lists
    app.performancedata_list.clear()
    app.waveforms_list.clear()
    app.available_data_list_performance.clear()
    app.available_data_list__waveforms.clear()

    # Helper function to get first two words from a filename
    def get_first_two_words(filename):
        words = re.split(r'\s+|-|_', filename.lower())
        return ' '.join(words[:2]).strip()

    # Fetch fresh performance data from folder, respecting dictionary
    performance_folder = app.performancedata_path.text()
    performance_items = set()  # Discovered items
    if performance_folder and os.path.isdir(performance_folder):
        for file in os.listdir(performance_folder):
            if file.lower().endswith(('.xlsx', '.xls')):
                first_two_words = get_first_two_words(file)
                if first_two_words in performancedata_testnames:
                    performance_items.add(performancedata_testnames[first_two_words])

    # Populate performance checklist with discovered items
    for item_name in sorted(performance_items):
        item = QListWidgetItem(item_name)
        item_font = QFont()
        item_font.setPointSize(16)
        item.setFont(item_font)
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked)
        app.performancedata_list.addItem(item)
        print(f"Added performance item from folder: {item_name}")

    # Get ordered items from checklist (all items, as just populated)
    performance_items_ordered = [
        app.performancedata_list.item(i).text()
        for i in range(app.performancedata_list.count())
    ]

    # Populate performance_files dict for discovered items
    performance_files = {}
    if performance_folder and os.path.isdir(performance_folder):
        for item_name in performance_items_ordered:
            performance_files[item_name] = []
            for file in sorted(os.listdir(performance_folder)):
                if file.lower().endswith(('.xlsx', '.xls')):
                    first_two_words = get_first_two_words(file)
                    if first_two_words in performancedata_testnames and performancedata_testnames[first_two_words] == item_name:
                        performance_files[item_name].append(file)

    # Populate available performance list in checklist order
    print("Populating performance data list:")
    for index, item_name in enumerate(performance_items_ordered):
        subheader_item = QListWidgetItem(item_name)
        subheader_item.setBackground(QBrush(QColor(34,35,39)))
        subheader_font = QFont()
        subheader_font.setBold(True)
        subheader_font.setPointSize(12)
        subheader_item.setFont(subheader_font)
        subheader_item.setFlags(Qt.NoItemFlags)
        app.available_data_list_performance.addItem(subheader_item)
        print(f"  Added subheader: {item_name}")

        for file_name in performance_files.get(item_name, []):
            file_item = QListWidgetItem(file_name)
            file_item.setFlags(file_item.flags() | Qt.ItemIsUserCheckable)
            file_item.setCheckState(Qt.Checked)
            app.available_data_list_performance.addItem(file_item)
            print(f"    Added file: {file_name}")

        if index < len(performance_items_ordered) - 1:
            spacer_item = QListWidgetItem("")
            spacer_item.setFlags(Qt.NoItemFlags)
            app.available_data_list_performance.addItem(spacer_item)
            print("    Added spacer")

    # Fetch fresh waveform data from folder, respecting dictionary
    waveform_folder = app.waveforms_path.text()
    waveform_items = set()
    if waveform_folder and os.path.isdir(waveform_folder):
        for file in os.listdir(waveform_folder):
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                first_two_words = get_first_two_words(file)
                if first_two_words in waveform_testnames:
                    waveform_items.add(waveform_testnames[first_two_words])

    # Populate waveform checklist with discovered items
    for item_name in sorted(waveform_items):
        item = QListWidgetItem(item_name)
        item_font = QFont()
        item_font.setPointSize(16)
        item.setFont(item_font)
        item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
        item.setCheckState(Qt.Checked)
        app.waveforms_list.addItem(item)
        print(f"Added waveform item from folder: {item_name}")

    # Get ordered items from checklist (all items, as just populated)
    waveform_items_ordered = [
        app.waveforms_list.item(i).text()
        for i in range(app.waveforms_list.count())
    ]

    # Populate waveform_files dict for discovered items
    waveform_files = {}
    if waveform_folder and os.path.isdir(waveform_folder):
        for item_name in waveform_items_ordered:
            waveform_files[item_name] = []
            for file in sorted(os.listdir(waveform_folder)):
                if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                    first_two_words = get_first_two_words(file)
                    if first_two_words in waveform_testnames and waveform_testnames[first_two_words] == item_name:
                        waveform_files[item_name].append(file)

    # Populate available waveform list in checklist order
    print("Populating waveform data list:")
    for index, item_name in enumerate(waveform_items_ordered):
        subheader_item = QListWidgetItem(item_name)
        subheader_item.setBackground(QBrush(QColor(34,35,39)))
        subheader_font = QFont()
        subheader_font.setBold(True)
        subheader_font.setPointSize(12)
        subheader_item.setFont(subheader_font)
        subheader_item.setFlags(Qt.NoItemFlags)
        app.available_data_list__waveforms.addItem(subheader_item)
        print(f"  Added subheader: {item_name}")

        for file_name in waveform_files.get(item_name, []):
            file_item = QListWidgetItem(file_name)
            file_item.setFlags(file_item.flags() | Qt.ItemIsUserCheckable)
            file_item.setCheckState(Qt.Checked)
            app.available_data_list__waveforms.addItem(file_item)
            print(f"    Added file: {file_name}")

        if index < len(waveform_items_ordered) - 1:
            spacer_item = QListWidgetItem("")
            spacer_item.setFlags(Qt.NoItemFlags)
            app.available_data_list__waveforms.addItem(spacer_item)
            print("    Added spacer")

    show_popup(app, "Data Loaded", "Successfully loaded data.", "info")
    print("Lists updated")

def refresh_data_lists(app):
    # Clear the available lists
    app.available_data_list_performance.clear()
    app.available_data_list__waveforms.clear()

    # Helper function to get first two words from a filename
    def get_first_two_words(filename):
        words = re.split(r'\s+|-|_', filename.lower())
        return ' '.join(words[:2]).strip()

    # Get checked items from checklists in their current order
    performance_items_ordered = [
        app.performancedata_list.item(i).text()
        for i in range(app.performancedata_list.count())
        if app.performancedata_list.item(i).checkState() == Qt.Checked
    ]
    waveform_items_ordered = [
        app.waveforms_list.item(i).text()
        for i in range(app.waveforms_list.count())
        if app.waveforms_list.item(i).checkState() == Qt.Checked
    ]

    # Fetch fresh performance files for checked items
    performance_folder = app.performancedata_path.text()
    performance_files = {item_name: [] for item_name in performance_items_ordered}
    if performance_folder and os.path.isdir(performance_folder):
        for file in os.listdir(performance_folder):
            if file.lower().endswith(('.xlsx', '.xls')):
                first_two_words = get_first_two_words(file)
                if first_two_words in performancedata_testnames:
                    value = performancedata_testnames[first_two_words]
                    if value in performance_files:
                        performance_files[value].append(file)

    # Fetch fresh waveform files for checked items
    waveform_folder = app.waveforms_path.text()
    waveform_files = {item_name: [] for item_name in waveform_items_ordered}
    if waveform_folder and os.path.isdir(waveform_folder):
        for file in os.listdir(waveform_folder):
            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                first_two_words = get_first_two_words(file)
                if first_two_words in waveform_testnames:
                    value = waveform_testnames[first_two_words]
                    if value in waveform_files:
                        waveform_files[value].append(file)

    # Populate performance available list (filtered by checks, in checklist order)
    print("Populating performance data list:")
    for index, item_name in enumerate(performance_items_ordered):
        if not performance_files.get(item_name):  # Skip if no files for this item
            continue
        subheader_item = QListWidgetItem(item_name)
        subheader_item.setBackground(QBrush(QColor(34,35,39)))
        subheader_font = QFont()
        subheader_font.setBold(True)
        subheader_font.setPointSize(12)
        subheader_item.setFont(subheader_font)
        subheader_item.setFlags(Qt.NoItemFlags)
        app.available_data_list_performance.addItem(subheader_item)
        print(f"  Added subheader: {item_name}")

        for file_name in sorted(performance_files.get(item_name, [])):
            file_item = QListWidgetItem(file_name)
            file_item.setFlags(file_item.flags() | Qt.ItemIsUserCheckable)
            file_item.setCheckState(Qt.Checked)
            app.available_data_list_performance.addItem(file_item)
            print(f"    Added file: {file_name}")

        if index < len(performance_items_ordered) - 1:
            spacer_item = QListWidgetItem("")
            spacer_item.setFlags(Qt.NoItemFlags)
            app.available_data_list_performance.addItem(spacer_item)
            print("    Added spacer")

    # Populate waveform available list (filtered by checks, in checklist order)
    print("Populating waveform data list:")
    for index, item_name in enumerate(waveform_items_ordered):
        if not waveform_files.get(item_name):  # Skip if no files for this item
            continue
        subheader_item = QListWidgetItem(item_name)
        subheader_item.setBackground(QBrush(QColor(34,35,39)))
        subheader_font = QFont()
        subheader_font.setBold(True)
        subheader_font.setPointSize(12)
        subheader_item.setFont(subheader_font)
        subheader_item.setFlags(Qt.NoItemFlags)
        app.available_data_list__waveforms.addItem(subheader_item)
        print(f"  Added subheader: {item_name}")

        for file_name in sorted(waveform_files.get(item_name, [])):
            file_item = QListWidgetItem(file_name)
            file_item.setFlags(file_item.flags() | Qt.ItemIsUserCheckable)
            file_item.setCheckState(Qt.Checked)
            app.available_data_list__waveforms.addItem(file_item)
            print(f"    Added file: {file_name}")

        if index < len(waveform_items_ordered) - 1:
            spacer_item = QListWidgetItem("")
            spacer_item.setFlags(Qt.NoItemFlags)
            app.available_data_list__waveforms.addItem(spacer_item)
            print("    Added spacer")

    print("Available lists refreshed based on checked checklist items")