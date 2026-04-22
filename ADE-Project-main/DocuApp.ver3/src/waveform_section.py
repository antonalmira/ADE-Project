import os
from PyQt5.QtWidgets import QFileDialog
from PyQt5.QtCore import Qt
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from utils import log_message
from word_utils import add_caption_field, format_value_units
from image_utils import crop_and_save
from list_updater import waveform_testnames
import re

class WaveformSection:
    def __init__(self, app, temp_dir):
        self.app = app
        self.temp_dir = temp_dir

    def get_first_two_words(self, filename):
        words = re.split(r'\s+|-|_', filename.lower())
        return ' '.join(words[:2]).strip()

    def get_images(self, waveform_items):
        # ORIGINAL LOGIC (restored):
        waveform_folder = self.app.waveforms_path.text()
        if not waveform_folder or not os.path.isdir(waveform_folder):
            log_message(f"Invalid waveform folder: {waveform_folder}")
            folder_path = QFileDialog.getExistingDirectory(self.app, "Select Waveform Folder", self.app.performancedata_path.text())
            if folder_path:
                waveform_folder = folder_path
                log_message(f"Selected waveform folder: {waveform_folder}")
            else:
                log_message("No waveform folder selected")
                return {}
        # Get the ordered list of files from available_data_list__waveforms
        ordered_files = []
        current_item = None
        for index in range(self.app.available_data_list__waveforms.count()):
            item = self.app.available_data_list__waveforms.item(index)
            if not item:
                continue
            if item.text() in waveform_items:
                current_item = item.text()
            elif current_item and item.checkState() == Qt.Checked and item.text().lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                ordered_files.append((current_item, item.text()))
        log_message(f"Ordered waveform files from UI: {ordered_files}")

        waveform_files = {}
        for item_name in waveform_items:
            waveform_files[item_name] = []
            prefix = next((key for key, value in waveform_testnames.items() if value == item_name), '')
            for ordered_item_name, file_name in ordered_files:
                if ordered_item_name != item_name:
                    continue
                if self.get_first_two_words(file_name) == prefix:
                    original_path = os.path.join(waveform_folder, file_name)
                    if os.path.exists(original_path):
                        cropped_path = crop_and_save(original_path, 0, 70, 0, 220, self.temp_dir)
                        if cropped_path:
                            waveform_files[item_name].append(cropped_path)
        return waveform_files

    def get_images_with_custom_crop(self, waveform_items):
          try:
            crop_upper = int(self.app.upper_input.text()) if self.app.upper_input.text() else 0
            crop_lower = int(self.app.lower_input.text()) if self.app.lower_input.text() else 0
            crop_left = int(self.app.left_input.text()) if self.app.left_input.text() else 0
            crop_right = int(self.app.right_input.text()) if self.app.right_input.text() else 0
        except ValueError:
            crop_upper = crop_lower = crop_left = crop_right = 0

        waveform_folder = self.app.waveforms_path.text()
        waveform_folder = self.app.waveforms_path.text()
        if not waveform_folder or not os.path.isdir(waveform_folder):
            log_message(f"Invalid waveform folder: {waveform_folder}")
            folder_path = QFileDialog.getExistingDirectory(self.app, "Select Waveform Folder", self.app.performancedata_path.text())
            if folder_path:
                waveform_folder = folder_path
                log_message(f"Selected waveform folder: {waveform_folder}")
            else:
                log_message("No waveform folder selected")
                return {}
        ordered_files = []
        current_item = None
        for index in range(self.app.available_data_list__waveforms.count()):
            item = self.app.available_data_list__waveforms.item(index)
            if not item:
                continue
            if item.text() in waveform_items:
                current_item = item.text()
            elif current_item and item.checkState() == Qt.Checked and item.text().lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')):
                ordered_files.append((current_item, item.text()))
        log_message(f"Ordered waveform files from UI: {ordered_files}")

        waveform_files = {}
        for item_name in waveform_items:
            waveform_files[item_name] = []
            prefix = next((key for key, value in waveform_testnames.items() if value == item_name), '')
            for ordered_item_name, file_name in ordered_files:
                if ordered_item_name != item_name:
                    continue
                if self.get_first_two_words(file_name) == prefix:
                    original_path = os.path.join(waveform_folder, file_name)
                    if os.path.exists(original_path):
                        cropped_path = crop_and_save(original_path, crop_left, crop_upper, crop_right, crop_lower, self.temp_dir)
                        if cropped_path:
                            waveform_files[item_name].append(cropped_path)
        return waveform_files

    def add_section(self, doc, last_element, waveform_items, waveform_files):
        for item in waveform_items:
            log_message(f"Adding waveform subheader: {item}")
            new_para = doc.add_paragraph(item, style='Heading 2')
            new_para.runs[0].font.size = Pt(14)
            last_element.getparent().insert(last_element.getparent().index(last_element) + 1, new_para._element)
            last_element = new_para._element
            if item in waveform_files and waveform_files[item]:
                table = doc.add_table(rows=1, cols=2)
                table.autofit = False
                table.columns[0].width = Inches(3.5)
                table.columns[1].width = Inches(3.5)
                table.alignment = WD_ALIGN_PARAGRAPH.CENTER
                current_row = 0
                current_col = 0
                for image_path in waveform_files[item]:
                    if image_path:
                        if current_col >= 2:
                            current_col = 0
                            current_row += 1
                            table.add_row()
                        cell = table.cell(current_row, current_col)
                        cell_para = cell.paragraphs[0]
                        cell_para.alignment = WD_ALIGN_PARAGRAPH.LEFT
                        run = cell_para.add_run()
                        run.add_picture(image_path, width=Inches(3.5))
                        caption_cell = cell.add_paragraph()
                        # Use the image filename (without extension) as the caption
                        cap = os.path.splitext(os.path.basename(image_path))[0]
                        caption_text = format_value_units(cap)
                        add_caption_field(caption_cell, caption_text, "Figure")
                        current_col += 1
                last_element.getparent().insert(last_element.getparent().index(last_element) + 1, table._element)
                last_element = table._element
        return last_element