import os
import sys
import time
from pathlib import Path
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QListWidgetItem

# Add project root to Python path to import resource_rc
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import resource_rc

from handlers import (
    add_performance_item, add_waveform_item, delete_performance_item, delete_waveform_item,
    select_performance_folder, select_waveform_folder,
    select_generate_document_folder, select_update_document, toggle_maximize
)
from preview import show_file_preview, crop_and_update_preview
from list_updater import update_available_data_list, refresh_data_lists, performancedata_testnames, waveform_testnames
from chart_extractor import save_chart_screenshots
from document_handler import generate_document, update_document
from document_generator import DocGenerator
from utils import get_resource_path

class DocuApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(DocuApp, self).__init__()
        # Load the UI file (works in dev and PyInstaller using get_resource_path)
        ui_path = get_resource_path('DocuApp_ver6.ui')
        uic.loadUi(ui_path, self)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.showMaximized()

        # Store crop values globally
        self.crop_upper = 0
        self.crop_lower = 0
        self.crop_left = 0
        self.crop_right = 0

        # Connect crop button to cropping logic
        self.crop_button.clicked.connect(lambda: crop_and_update_preview(self))

        # Connect buttons to handlers
        self.performancedata_add.clicked.connect(lambda: add_performance_item(self))
        self.waveforms_add.clicked.connect(lambda: add_waveform_item(self))
        self.performancedata_delete.clicked.connect(lambda: delete_performance_item(self))
        self.waveforms_delete.clicked.connect(lambda: delete_waveform_item(self))
        self.performancedata_sel.clicked.connect(lambda: select_performance_folder(self))
        self.waveforms_sel.clicked.connect(lambda: select_waveform_folder(self))
        self.exit_button.clicked.connect(self.close)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.maximize_button.clicked.connect(lambda: toggle_maximize(self))
        self.load_files_button.clicked.connect(lambda: update_available_data_list(self))
        self.refresh_button.clicked.connect(lambda: refresh_data_lists(self))
        self.generate_document_button.clicked.connect(lambda: generate_document(self))
        self.generated_document_sel_location.clicked.connect(lambda: select_generate_document_folder(self))
        # self.update_document_sel_location.clicked.connect(lambda: select_update_document(self))
        # self.update_document_button.clicked.connect(lambda: update_document(self))

        # Configure lists
        self.available_data_list_performance.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.available_data_list_performance.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.available_data_list_performance.itemSelectionChanged.connect(lambda: show_file_preview(self))

        self.available_data_list__waveforms.setSelectionMode(QtWidgets.QAbstractItemView.SingleSelection)
        self.available_data_list__waveforms.setSizePolicy(QtWidgets.QSizePolicy.Expanding, QtWidgets.QSizePolicy.Expanding)
        self.available_data_list__waveforms.itemSelectionChanged.connect(lambda: show_file_preview(self))

        # Enable drag-and-drop reordering for performance data list
        self.performancedata_list.setDragEnabled(True)
        self.performancedata_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        # Enable drag-and-drop reordering for waveforms list
        self.waveforms_list.setDragEnabled(True)
        self.waveforms_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
