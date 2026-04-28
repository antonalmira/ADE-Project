import os
import sys
from PyQt5 import QtWidgets, uic, QtCore
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

from document_handler import generate_document, update_document_prompt
from utils import get_resource_path
from preview import show_file_preview, crop_and_update_preview
from list_updater import update_available_data_list, refresh_data_lists

# --- NEW: CUSTOM DIALOG FOR MULTIPLE CAPTION VARIABLES ---
class CaptionDialog(QtWidgets.QDialog):
    def __init__(self, parent=None, current_data=None):
        super().__init__(parent)
        self.setWindowTitle("Custom Caption & Details")
        self.setMinimumWidth(450)
        
        self.layout = QtWidgets.QFormLayout(self)
        
        self.caption_input = QtWidgets.QLineEdit(self)
        self.caption_input.setPlaceholderText("e.g., 85 VAC 60 Hz.")
        
        self.ch_info_input = QtWidgets.QLineEdit(self)
        self.ch_info_input.setPlaceholderText("e.g., CH4: VRIPPLE, 20 mV / div., 20 ms / div.")
        
        self.zoom_info_input = QtWidgets.QLineEdit(self)
        self.zoom_info_input.setPlaceholderText("e.g., Zoom: 10 µs / div.")
        
        self.meas_info_input = QtWidgets.QLineEdit(self)
        self.meas_info_input.setPlaceholderText("e.g., Output Ripple = 67.2 mV")
        
        # Pre-fill if data already exists
        if isinstance(current_data, dict):
            self.caption_input.setText(current_data.get('caption', ''))
            self.ch_info_input.setText(current_data.get('ch_info', ''))
            self.zoom_info_input.setText(current_data.get('zoom_info', ''))
            self.meas_info_input.setText(current_data.get('meas_info', ''))
        elif isinstance(current_data, str):
            self.caption_input.setText(current_data)
            
        self.layout.addRow("Main Caption:", self.caption_input)
        self.layout.addRow("Channel Info:", self.ch_info_input)
        self.layout.addRow("Zoom Info:", self.zoom_info_input)
        self.layout.addRow("Measurement:", self.meas_info_input)
        
        self.buttons = QtWidgets.QDialogButtonBox(QtWidgets.QDialogButtonBox.Ok | QtWidgets.QDialogButtonBox.Cancel)
        self.buttons.accepted.connect(self.accept)
        self.buttons.rejected.connect(self.reject)
        self.layout.addRow(self.buttons)
        
        # Keep styling clean
        self.setStyleSheet("""
            QDialog { background-color: #f5f5f5; }
            QLabel { color: #000000; font-weight: bold; }
            QLineEdit { border: 1px solid #b0b0b0; border-radius: 4px; padding: 6px; background: white; color: black; }
            QPushButton { background-color: #0085ca; color: #ffffff; border-radius: 4px; padding: 6px 20px; font-weight: bold; }
            QPushButton:hover { background-color: #3c649f; }
        """)

    def get_data(self):
        return {
            'caption': self.caption_input.text().strip(),
            'ch_info': self.ch_info_input.text().strip(),
            'zoom_info': self.zoom_info_input.text().strip(),
            'meas_info': self.meas_info_input.text().strip()
        }


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

        # 6. LIST SELECTION CONNECTIONS (Preview Updates & Custom Functions)
        self.available_data_list_performance.itemSelectionChanged.connect(lambda: show_file_preview(self))
        self.available_data_list__waveforms.itemSelectionChanged.connect(lambda: show_file_preview(self))

        # ENABLE DRAG AND DROP REORDERING
        self.performancedata_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.waveforms_list.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.available_data_list_performance.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
        self.available_data_list__waveforms.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)

        # ENABLE DOUBLE CLICK FOR CUSTOM CAPTIONS
        self.available_data_list_performance.itemDoubleClicked.connect(self.set_custom_caption)
        self.available_data_list__waveforms.itemDoubleClicked.connect(self.set_custom_caption)

        # CONNECT THE NEW UPDATE DOCUMENT BUTTON
        if hasattr(self, 'update_document_button'):
            self.update_document_button.clicked.connect(lambda: update_document_prompt(self))

        # 7. POPULATE TEMPLATE DROPDOWN
        self.populate_templates_dropdown()

    def populate_templates_dropdown(self):
        # NEW: Check if we are running as an .exe or a Python script
        if getattr(sys, 'frozen', False):
            # Running as .exe: Look in the same folder as the .exe
            base_path = sys._MEIPASS
        else:
            # Running as script: Go up to the ADE-Project-main folder
            current_dir = os.path.dirname(os.path.abspath(__file__))
            base_path = os.path.abspath(os.path.join(current_dir, "..", ".."))

        templates_folder = os.path.join(base_path, "templates")
        
        self.template_dropdown.clear()
        if os.path.exists(templates_folder):
            template_files = [f for f in os.listdir(templates_folder) if f.endswith('.docx')]
            if template_files:
                self.template_dropdown.addItems(template_files)
            else:
                self.template_dropdown.addItem("No templates found")
                self.template_dropdown.setEnabled(False)
        else:
            self.template_dropdown.addItem("Templates folder missing!")
            self.template_dropdown.setEnabled(False)

    # --- CUSTOM CAPTIONS METHOD ---
    def set_custom_caption(self, item):
        if item.flags() & Qt.ItemIsUserCheckable:
            QtCore.QTimer.singleShot(0, lambda: self._prompt_custom_caption(item))

    def _prompt_custom_caption(self, item):
        current_data = item.data(Qt.UserRole)
        dialog = CaptionDialog(self, current_data)
        
        if dialog.exec_():
            new_data = dialog.get_data()
            # If at least one field has text, save it
            if any(new_data.values()):
                item.setData(Qt.UserRole, new_data)
                
                # Build preview for the Tooltip
                preview = []
                if new_data['caption']: preview.append(f"Caption: {new_data['caption']}")
                if new_data['ch_info']: preview.append(f"Ch: {new_data['ch_info']}")
                if new_data['zoom_info']: preview.append(f"Zoom: {new_data['zoom_info']}")
                if new_data['meas_info']: preview.append(f"Meas: {new_data['meas_info']}")
                
                item.setToolTip("\n".join(preview))
                item.setBackground(QtCore.Qt.darkBlue)
            else:
                item.setData(Qt.UserRole, None)
                item.setToolTip("")
                item.setBackground(QtCore.Qt.transparent)

    # --- HELPER METHODS ---
    def toggle_maximize(self):
        if self.isMaximized():
            self.showNormal()
        else:
            self.showMaximized()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.headerr.underMouse():
                self.old_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = QPoint(event.globalPos() - self.old_pos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.old_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.old_pos = None