from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
import os
from image_utils import crop_and_save

def get_ui_crop_values(app):
    """Safely parse crop values from the UI text boxes."""
    try:
        return {
            'left': int(app.left_input.text()) if app.left_input.text() else 0,
            'top': int(app.upper_input.text()) if app.upper_input.text() else 0,
            'right': int(app.right_input.text()) if app.right_input.text() else 0,
            'bottom': int(app.lower_input.text()) if app.lower_input.text() else 0
        }
    except ValueError:
        return {'left': 0, 'top': 0, 'right': 0, 'bottom': 0}

def crop_and_update_preview(app):
    """Triggered by the CROP button."""
    wave_items = app.available_data_list__waveforms.selectedItems()
    perf_items = app.available_data_list_performance.selectedItems()
    
    selected_item = wave_items[0] if wave_items else (perf_items[0] if perf_items else None)
    if not selected_item:
        app.file_view.setText("Select an image first")
        return

    source_folder = app.waveforms_path.text() if wave_items else app.performancedata_path.text()
    file_path = os.path.join(source_folder, selected_item.text())
    
    v = get_ui_crop_values(app)
    # Use a specific preview folder to avoid permission conflicts
    cropped_path = crop_and_save(file_path, v['left'], v['top'], v['right'], v['bottom'], "temp_preview")
    
    if cropped_path:
        pixmap = QPixmap(cropped_path)
        scaled = pixmap.scaled(app.file_view.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
        app.file_view.setPixmap(scaled)
    else:
        app.file_view.setText("Preview failed")

def show_file_preview(app):
    """Triggered when an item is clicked in the list."""
    # Synchronize selection between the two lists
    if app.available_data_list_performance.selectedItems():
        app.available_data_list__waveforms.clearSelection()
    
    # Run the same crop logic immediately so the user sees the current crop settings
    crop_and_update_preview(app)