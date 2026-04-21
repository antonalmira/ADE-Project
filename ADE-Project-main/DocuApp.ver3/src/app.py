import os
import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt
import resource_rc
from document_handler import generate_document, update_document
from utils import get_resource_path
from preview import show_file_preview, crop_and_update_preview
from list_updater import update_available_data_list, refresh_data_lists

class DocuApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(DocuApp, self).__init__()
        uic.loadUi(get_resource_path('DocuApp_ver6.ui'), self)
        self.setWindowFlags(Qt.FramelessWindowHint)
        
        # Crop logic
        self.crop_upper = 0
        self.crop_lower = 0
        self.crop_left = 0
        self.crop_right = 0

        # Connections
        self.generate_document_button.clicked.connect(lambda: generate_document(self))
        # self.update_document_button.clicked.connect(lambda: update_document(self))
        
        self.exit_button.clicked.connect(self.close)
        self.load_files_button.clicked.connect(lambda: update_available_data_list(self))
        self.refresh_button.clicked.connect(lambda: refresh_data_lists(self))
        self.crop_button.clicked.connect(lambda: crop_and_update_preview(self))
        
        self.available_data_list_performance.itemSelectionChanged.connect(lambda: show_file_preview(self))
        self.available_data_list__waveforms.itemSelectionChanged.connect(lambda: show_file_preview(self))