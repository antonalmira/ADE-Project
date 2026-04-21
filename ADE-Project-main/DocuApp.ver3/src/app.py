import os
import sys
from PyQt5 import QtWidgets, uic
from PyQt5.QtCore import Qt, QPoint
import resource_rc

# Import handlers (Removed select_template_file since we use a dropdown now)
from handlers import (
    select_performance_folder, 
    select_waveform_folder, 
    add_performance_item, 
    delete_performance_item, 
    add_waveform_item, 
    delete_waveform_item
)

from document_handler import generate_document
from utils import get_resource_path
from preview import show_file_preview, crop_and_update_preview
from list_updater import update_available_data_list, refresh_data_lists

class DocuApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(DocuApp, self).__init__()
        # Load the UI file 
        uic.loadUi(get_resource_path('DocuApp_ver6.ui'), self)
        
        # 1. WINDOW SETTINGS
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.old_pos = None

        # 2. CROP LOGIC INITIALIZATION
        self.crop_upper = 0
        self.crop_lower = 0
        self.crop_left = 0
        self.crop_right = 0

        # Set crop inputs to "0" by default in the UI
        for line_edit in [self.upper_input, self.lower_input, self.left_input, self.right_input]:
            line_edit.setText("0")

        # 3. FILE & FOLDER SELECTION CONNECTIONS
        self.performancedata_sel.clicked.connect(lambda: select_performance_folder(self))
        self.waveforms_sel.clicked.connect(lambda: select_waveform_folder(self))

        # 4. DATA LIST MANAGEMENT (Add/Delete)
        self.performancedata_add.clicked.connect(lambda: add_performance_item(self))
        self.performancedata_delete.clicked.connect(lambda: delete_performance_item(self))
        self.waveforms_add.clicked.connect(lambda: add_waveform_item(self))
        self.waveforms_delete.clicked.connect(lambda: delete_waveform_item(self))

        # 5. CORE ACTION CONNECTIONS
        self.exit_button.clicked.connect(self.close)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.maximize_button.clicked.connect(self.toggle_maximize)
        
        self.load_files_button.clicked.connect(lambda: update_available_data_list(self))
        self.refresh_button.clicked.connect(lambda: refresh_data_lists(self))
        
        # Crop button triggers the image update in preview.py
        self.crop_button.clicked.connect(lambda: crop_and_update_preview(self))
        
        # Generate button opens the Save dialog and starts processing
        self.generate_document_button.clicked.connect(lambda: generate_document(self))

        # 6. LIST SELECTION CONNECTIONS (Preview Updates)
        self.available_data_list_performance.itemSelectionChanged.connect(lambda: show_file_preview(self))
        self.available_data_list__waveforms.itemSelectionChanged.connect(lambda: show_file_preview(self))

        # 7. POPULATE TEMPLATE DROPDOWN
        self.populate_templates_dropdown()

    # --- NEW METHOD: POPULATE DROPDOWN ---
    def populate_templates_dropdown(self):
        # Path to the templates folder (assuming it's in the same folder as this script)
        templates_folder = os.path.join(os.path.dirname(__file__), "templates")
        
        self.template_dropdown.clear()
        
        if os.path.exists(templates_folder):
            # Find all .docx files in the templates folder
            template_files = [f for f in os.listdir(templates_folder) if f.endswith('.docx')]
            
            if template_files:
                self.template_dropdown.addItems(template_files)
            else:
                self.template_dropdown.addItem("No templates found")
                self.template_dropdown.setEnabled(False)
        else:
            self.template_dropdown.addItem("Templates folder missing!")
            self.template_dropdown.setEnabled(False)

    # --- HELPER METHODS ---

    def toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    # --- MOUSE EVENTS FOR DRAGGING FRAMELESS WINDOW (Restricted to Header) ---
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            # Only allow drag if clicking on the header frame
            if self.headerr.underMouse():
                self.old_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = QPoint(event.globalPos() - self.old_pos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.old_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.old_pos = None