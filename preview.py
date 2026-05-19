from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QFont
import os
from utils import ensure_directory
from image_utils import crop_and_save

def get_ui_crop_values(app):
    try:
        return {
            'left': int(app.left_input.text()) if app.left_input.text() else 0,
            'top': int(app.upper_input.text()) if app.upper_input.text() else 0,
            'right': int(app.right_input.text()) if app.right_input.text() else 0,
            'bottom': int(app.lower_input.text()) if app.lower_input.text() else 0
        }
    except ValueError:
        return {'left': 0, 'top': 0, 'right': 0, 'bottom': 0}

def _get_first_image_path(item):
    """Recursively search a folder item to find the first available image file."""
    for i in range(item.childCount()):
        child = item.child(i)
        if child.data(0, Qt.UserRole + 2) == "file":
            return child.data(0, Qt.UserRole + 1)
        elif child.data(0, Qt.UserRole + 2) == "folder":
            res = _get_first_image_path(child)
            if res: return res
    return None

def crop_and_update_preview(app):
    wave_items = app.waveform_tree.selectedItems()
    perf_items = app.performance_tree.selectedItems()
    
    item = None
    if wave_items: item = wave_items[0]
    elif perf_items: item = perf_items[0]

    file_path = None
    if item:
        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        if is_folder:
            file_path = _get_first_image_path(item)
        else:
            file_path = item.data(0, Qt.UserRole + 1)

    if not file_path or not os.path.exists(file_path):
        app.file_view.clear()
        return

    v = get_ui_crop_values(app)
    cropped_path = crop_and_save(file_path, v['left'], v['top'], v['right'], v['bottom'], "temp_preview")
    
    if cropped_path:
        pixmap = QPixmap(cropped_path)
        scaled = pixmap.scaled(app.file_view.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        app.file_view.setPixmap(scaled)
    else:
        app.file_view.clear()

def show_file_preview(app):
    sender = app.sender()
    crop = {'left': '0', 'top': '0', 'right': '0', 'bottom': '0'}
    
    if sender == app.performance_tree and app.performance_tree.selectedItems():
        app.waveform_tree.clearSelection()
        item = app.performance_tree.selectedItems()[0]
    elif sender == app.waveform_tree and app.waveform_tree.selectedItems():
        app.performance_tree.clearSelection()
        item = app.waveform_tree.selectedItems()[0]
    else:
        app.images_preview_text.setText("PREVIEW")
        return
        
    # --- Update Preview Label Dynamically with instructions ---
    clean_name = item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip()
    app.images_preview_text.setText(f"PREVIEW: {clean_name} ")

    current = item
    while current:
        saved_crop = current.data(0, Qt.UserRole + 3)
        if saved_crop:
            crop = saved_crop
            break
        current = current.parent()

    app.left_input.setText(str(crop['left']))
    app.upper_input.setText(str(crop['top']))
    app.right_input.setText(str(crop['right']))
    app.lower_input.setText(str(crop['bottom']))
        
    crop_and_update_preview(app)

# ==========================================
# EXPANDED PREVIEW DIALOG LOGIC
# ==========================================

class ExpandedPreviewDialog(QtWidgets.QDialog):
    def __init__(self, main_app, file_path, current_crop, item_name):
        super().__init__(main_app)
        self.main_app = main_app
        self.file_path = file_path
        self.current_pixmap = None
        
        self.setWindowTitle(f"Preview & Crop - {item_name}")
        self.resize(1000, 800)
        self.setStyleSheet("background-color: #16161a; color: white;")
        
        # Main Layout
        layout = QtWidgets.QVBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # Title Label
        self.title_label = QtWidgets.QLabel(f"Cropping: {item_name}")
        self.title_label.setAlignment(Qt.AlignCenter)
        font = QFont("Arial", 14, QFont.Bold)
        self.title_label.setFont(font)
        self.title_label.setStyleSheet("color: #0085ca; margin-bottom: 10px;")
        layout.addWidget(self.title_label)
        
        # Image Display
        self.image_label = QtWidgets.QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.image_label.setMinimumSize(50, 50)
        self.image_label.setStyleSheet("background: #0b0b0d; border: 2px solid #3a3a40; border-radius: 8px;")
        layout.addWidget(self.image_label, 1) # Expanding vertically
        
        # Bottom Controls
        ctrl_layout = QtWidgets.QHBoxLayout()
        ctrl_layout.setContentsMargins(0, 15, 0, 0)
        
        self.inp_top = QtWidgets.QLineEdit(str(current_crop.get('top', 0)))
        self.inp_bottom = QtWidgets.QLineEdit(str(current_crop.get('bottom', 0)))
        self.inp_left = QtWidgets.QLineEdit(str(current_crop.get('left', 0)))
        self.inp_right = QtWidgets.QLineEdit(str(current_crop.get('right', 0)))
        
        input_style = "background: #1a1a1e; color: #e0e0e0; border: 1px solid #3a3a40; border-radius: 4px; padding: 5px;"
        for inp in [self.inp_top, self.inp_bottom, self.inp_left, self.inp_right]:
            inp.setStyleSheet(input_style)
            inp.setMaximumWidth(70)
            inp.setAlignment(Qt.AlignCenter)
            
        ctrl_layout.addStretch()
        ctrl_layout.addWidget(QtWidgets.QLabel("Top:"))
        ctrl_layout.addWidget(self.inp_top)
        ctrl_layout.addSpacing(15)
        ctrl_layout.addWidget(QtWidgets.QLabel("Bottom:"))
        ctrl_layout.addWidget(self.inp_bottom)
        ctrl_layout.addSpacing(15)
        ctrl_layout.addWidget(QtWidgets.QLabel("Left:"))
        ctrl_layout.addWidget(self.inp_left)
        ctrl_layout.addSpacing(15)
        ctrl_layout.addWidget(QtWidgets.QLabel("Right:"))
        ctrl_layout.addWidget(self.inp_right)
        ctrl_layout.addSpacing(20)
        
        self.btn_crop = QtWidgets.QPushButton("APPLY CROP")
        self.btn_crop.setStyleSheet("background: #0085ca; color: white; border-radius: 4px; padding: 8px 25px; font-weight: bold;")
        self.btn_crop.setCursor(Qt.PointingHandCursor)
        self.btn_crop.clicked.connect(self.apply_crop)
        ctrl_layout.addWidget(self.btn_crop)
        ctrl_layout.addStretch()
        
        layout.addLayout(ctrl_layout)
        
        # Initial Render
        self.refresh_image(current_crop)
        
    def refresh_image(self, crop_dict):
        # We ensure inputs are integers before saving
        c_top = int(crop_dict.get('top') or 0)
        c_bottom = int(crop_dict.get('bottom') or 0)
        c_left = int(crop_dict.get('left') or 0)
        c_right = int(crop_dict.get('right') or 0)

        cropped_path = crop_and_save(self.file_path, c_left, c_top, c_right, c_bottom, "temp_preview")
        if cropped_path:
            self.current_pixmap = QPixmap(cropped_path)
            self.update_image_display()
            
    def update_image_display(self):
        """Scales the existing pixmap smoothly when the window resizes, without having to re-crop."""
        if self.current_pixmap:
            scaled = self.current_pixmap.scaled(self.image_label.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            self.image_label.setPixmap(scaled)
            
    def apply_crop(self):
        # 1. Sync values back to main window line edits
        self.main_app.upper_input.setText(self.inp_top.text() or '0')
        self.main_app.lower_input.setText(self.inp_bottom.text() or '0')
        self.main_app.left_input.setText(self.inp_left.text() or '0')
        self.main_app.right_input.setText(self.inp_right.text() or '0')
        
        # 2. Trigger the main app's save function (updates trees & main preview)
        self.main_app.save_crop_to_selected()
        
        # 3. Refresh our local high-res image
        crop_dict = {
            'top': self.inp_top.text(), 'bottom': self.inp_bottom.text(),
            'left': self.inp_left.text(), 'right': self.inp_right.text()
        }
        self.refresh_image(crop_dict)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        self.update_image_display()

def open_expanded_preview(app):
    wave_items = app.waveform_tree.selectedItems()
    perf_items = app.performance_tree.selectedItems()
    
    item = None
    if wave_items: item = wave_items[0]
    elif perf_items: item = perf_items[0]
    
    if not item: return
    
    file_path = None
    is_folder = item.data(0, Qt.UserRole + 2) == "folder"
    if is_folder:
        file_path = _get_first_image_path(item)
    else:
        file_path = item.data(0, Qt.UserRole + 1)
        
    if not file_path or not os.path.exists(file_path): return
    
    clean_name = item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip()
    crop = get_ui_crop_values(app)
    
    dialog = ExpandedPreviewDialog(app, file_path, crop, clean_name)
    dialog.exec_()