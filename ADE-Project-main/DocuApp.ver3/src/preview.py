# src/preview.py
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap
import os

from image_utils import crop_and_save
def crop_and_update_preview(app):
    # Get selected waveform image
    waveform_items = app.available_data_list__waveforms.selectedItems()
    if not waveform_items:
        app.file_view.setText("No waveform selected")
        return
    selected_item = waveform_items[0]
    source_folder = app.waveforms_path.text()
    file_name = selected_item.text()
    file_path = os.path.join(source_folder, file_name)
    # Get crop values from QLineEdit
    try:
        crop_upper = int(app.upper_input.text()) if app.upper_input.text() else 0
        crop_lower = int(app.lower_input.text()) if app.lower_input.text() else 0
        crop_left = int(app.left_input.text()) if app.left_input.text() else 0
        crop_right = int(app.right_input.text()) if app.right_input.text() else 0
    except ValueError:
        app.file_view.setText("Invalid crop values")
        return
    app.crop_upper = crop_upper
    app.crop_lower = crop_lower
    app.crop_left = crop_left
    app.crop_right = crop_right
    # Crop image and update preview
    cropped_path = crop_and_save(file_path, crop_left, crop_upper, crop_right, crop_lower)
    if cropped_path and os.path.exists(cropped_path):
        pixmap = QPixmap(cropped_path)
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(app.file_view.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            app.file_view.setPixmap(scaled_pixmap)
            app.file_view.setStyleSheet("border-radius:10px")
            app.file_view.setAlignment(Qt.AlignCenter)
        else:
            app.file_view.setText("Invalid Cropped Image File")
            app.file_view.setAlignment(Qt.AlignCenter)
    else:
        app.file_view.setText("Cropping updated. Please Select Waveform.")
        app.file_view.setAlignment(Qt.AlignCenter)

def show_file_preview(app):
    # Clear previous preview
    app.file_view.setPixmap(QPixmap())
    app.file_view.setText("")
    app.file_view.setStyleSheet("")

    # Get selected items from both lists
    performance_items = app.available_data_list_performance.selectedItems()
    waveform_items = app.available_data_list__waveforms.selectedItems()

    # Clear selections in the opposite list to ensure single selection
    if performance_items:
        app.available_data_list__waveforms.clearSelection()
    elif waveform_items:
        app.available_data_list_performance.clearSelection()

    # Determine selected item and source folder
    selected_item = None
    source_folder = None
    is_excel_file = False
    if waveform_items:
        selected_item = waveform_items[0]
        source_folder = app.waveforms_path.text()
    elif performance_items:
        selected_item = performance_items[0]
        source_folder = app.performancedata_path.text()
        is_excel_file = selected_item.text().lower().endswith(('.xlsx', '.xls'))

    if not selected_item or not source_folder or not os.path.isdir(source_folder):
        app.file_view.setText("No Preview Available")
        app.file_view.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(26, 26, 30); border-radius:10px; border: 1px solid #7a7a81")
        app.file_view.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        print("No valid selection or folder for preview")
        return

    file_name = selected_item.text()
    file_path = os.path.join(source_folder, file_name)

    if is_excel_file:
        app.file_view.setText("No Preview Available")
        app.file_view.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(26, 26, 30);border-radius:10px; border: 1px solid #7a7a81")
        app.file_view.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        print(f"No preview available for Excel file: {file_name}")
    elif file_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')) and os.path.exists(file_path):
        try:
            crop_upper = int(app.upper_input.text()) if app.upper_input.text() else 0
            crop_lower = int(app.lower_input.text()) if app.lower_input.text() else 0
            crop_left = int(app.left_input.text()) if app.left_input.text() else 0
            crop_right = int(app.right_input.text()) if app.right_input.text() else 0
        except ValueError:
            app.file_view.setText("Invalid crop values")
            return
        from PIL import Image
        try:
            with Image.open(file_path) as img:
                width, height = img.size
                
                if crop_left > 0 or crop_right > 0 or crop_upper > 0 or crop_lower > 0:
                    # FIX: Safe bounds applied to preview as well
                    left = min(crop_left, width - 1)
                    top = min(crop_upper, height - 1)
                    right = max(left + 1, width - crop_right)
                    bottom = max(top + 1, height - crop_lower)
                    
                    cropped_img = img.crop((left, top, right, bottom))
                    from tempfile import gettempdir
                    temp_path = os.path.join(gettempdir(), f"preview_cropped_{os.path.basename(file_path)}")
                    cropped_img.save(temp_path)
                    pixmap = QPixmap(temp_path)
                else:
                    pixmap = QPixmap(file_path)
        except Exception as e:
            app.file_view.setText(f"Cropping failed: {str(e)}")
            return
        except Exception as e:
            app.file_view.setText(f"Cropping failed: {str(e)}")
            return
        if not pixmap.isNull():
            scaled_pixmap = pixmap.scaled(app.file_view.size(), Qt.KeepAspectRatio, Qt.SmoothTransformation)
            app.file_view.setPixmap(scaled_pixmap)
            app.file_view.setStyleSheet("border-radius:10px")
            app.file_view.setAlignment(Qt.AlignCenter)
            print(f"Showing preview for image: {file_name}")
        else:
            app.file_view.setText("Invalid Image File")
            app.file_view.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(26, 26, 30);border-radius:10px; border: 1px solid #7a7a81")
            app.file_view.setAlignment(Qt.AlignCenter | Qt.AlignTop)
            print(f"Failed to load image: {file_name}")
    else:
        app.file_view.setText("No Preview Available")
        app.file_view.setStyleSheet("color: rgb(255, 255, 255); background-color: rgb(26, 26, 30);border-radius:10px; border: 1px solid #7a7a81")
        app.file_view.setAlignment(Qt.AlignCenter | Qt.AlignTop)
        print(f"No preview available for: {file_name}")