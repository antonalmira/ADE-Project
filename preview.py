from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
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
        
    # --- Update Preview Label Dynamically ---
    clean_name = item.text(0).replace(" CROPPED", "").strip()
    app.images_preview_text.setText(f"PREVIEW: {clean_name}")

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