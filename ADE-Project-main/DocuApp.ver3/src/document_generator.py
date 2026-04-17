import os
import shutil
from docx import Document
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
from pathlib import Path
import re
from utils import ensure_directory, remove_directory, get_default_base_folder, log_message, get_resource_path
from performance_section import PerformanceSection
from waveform_section import WaveformSection

class DocGenerator:
    def __init__(self, app, output_folder, update_doc_path=""):
        self.app = app
        # Prefer bundled resource path (works for dev and PyInstaller)
        self.template_path = get_resource_path(os.path.join('DER-template.docx'))
        log_message(f"Template path: {self.template_path}")
        if not os.path.exists(self.template_path):
            log_message("Template not found, prompting selection")
            self.template_path = self.select_template_file()
            if not self.template_path:
                raise FileNotFoundError("No template selected")
        self.output_folder = get_default_base_folder(output_folder)
        self.output_path = os.path.join(self.output_folder, "Generated_Document.docx")
        self.update_doc_path = update_doc_path
        self.temp_dir = "temp_cropped_images"
        ensure_directory(self.temp_dir)
        self.performance = PerformanceSection(app, self.temp_dir)
        self.waveform = WaveformSection(app, self.temp_dir)

    def extract_headers(self, doc):
        headers = [para.text.strip() for para in doc.paragraphs if para.style.name.startswith('Heading')]
        log_message(f"Extracted headers: {headers}")
        return headers

    def find_target_headers(self, headers):
        performance_headers = [h for h in headers if re.search(r"performance\s*data", h.lower())]
        waveform_headers = [h for h in headers if re.search(r"waveforms", h.lower())]
        log_message(f"Performance headers: {performance_headers}")
        log_message(f"Waveform headers: {waveform_headers}")
        return performance_headers, waveform_headers

    def get_checked_items(self):
        performance_items = [self.app.performancedata_list.item(index).text()
                             for index in range(self.app.performancedata_list.count())
                             if self.app.performancedata_list.item(index).checkState() == Qt.Checked]
        waveform_items = [self.app.waveforms_list.item(index).text()
                          for index in range(self.app.waveforms_list.count())
                          if self.app.waveforms_list.item(index).checkState() == Qt.Checked]
        log_message(f"Checked performance items: {performance_items}")
        log_message(f"Checked waveform items: {waveform_items}")
        return performance_items, waveform_items

    def generate(self):
        try:
            doc_path = self.update_doc_path if self.update_doc_path and os.path.exists(self.update_doc_path) else self.template_path
            doc = Document(doc_path)
            log_message(f"Loaded document: {doc_path}")
            headers = self.extract_headers(doc)
            performance_headers, waveform_headers = self.find_target_headers(headers)
            performance_items, waveform_items = self.get_checked_items()

            # Initialize progress dialog
            total_steps = len(performance_items) + len(waveform_items)
            progress_dialog = QtWidgets.QProgressDialog("Generating document...", "", 0, total_steps, self.app)
            progress_dialog.setWindowTitle("Processing")
            progress_dialog.setWindowModality(Qt.WindowModal)
            progress_dialog.setCancelButton(None)
            progress_dialog.setMinimumDuration(0)
            progress_value = 0
            log_message("Progress dialog initialized for document generation")

            # Remove content under target headers (keep headers)
            elements_to_remove = []
            i = 0
            while i < len(doc.element.body):
                elem = doc.element.body[i]
                if elem.tag.endswith('p'):
                    para_index = next(idx for idx, p in enumerate(doc.paragraphs) if p._element == elem)
                    para = doc.paragraphs[para_index]
                    para_text = para.text.strip()
                    if para_text in performance_headers + waveform_headers:
                        j = i + 1
                        while j < len(doc.element.body):
                            next_elem = doc.element.body[j]
                            if next_elem.tag.endswith('p'):
                                next_para_index = next(idx for idx, p in enumerate(doc.paragraphs) if p._element == next_elem)
                                next_para = doc.paragraphs[next_para_index]
                                if next_para.style.name.startswith('Heading 1'):
                                    break
                            elements_to_remove.append(next_elem)
                            j += 1
                        i = j
                        continue
                i += 1
            for elem in reversed(elements_to_remove):
                elem.getparent().remove(elem)
            log_message(f"Removed {len(elements_to_remove)} elements under headers")

            # Gather data
            performance_data = self.performance.get_data(performance_items)
            efficiency_table = self.performance.get_efficiency_table() if "Efficiency Test" in performance_items else None
            waveform_files = self.waveform.get_images_with_custom_crop(waveform_items)

            # Add content under existing headers
            header_elements = [(p.text.strip(), p._element) for p in doc.paragraphs if p.text.strip() in performance_headers + waveform_headers]
            for header_text, header_elem in header_elements:
                last_element = header_elem
                if header_text in performance_headers:
                    last_element = self.performance.add_section(doc, last_element, performance_items, performance_data, efficiency_table)
                    progress_value += len(performance_items)
                    progress_dialog.setValue(progress_value)
                    log_message(f"Updated progress: {progress_value}/{total_steps} after performance section")
                if header_text in waveform_headers:
                    last_element = self.waveform.add_section(doc, last_element, waveform_items, waveform_files)
                    progress_value += len(waveform_items)
                    progress_dialog.setValue(progress_value)
                    log_message(f"Updated progress: {progress_value}/{total_steps} after waveform section")

            # Append new sections if no matching headers found
            last_element = doc.element.body[-1]
            if not performance_headers and performance_items:
                log_message("Appending Performance Data header")
                new_para = doc.add_paragraph("Performance Data", style='Heading 1')
                last_element.getparent().insert(last_element.getparent().index(last_element) + 1, new_para._element)
                last_element = new_para._element
                self.performance.add_section(doc, last_element, performance_items, performance_data, efficiency_table)
                progress_value += len(performance_items)
                progress_dialog.setValue(progress_value)
                log_message(f"Updated progress: {progress_value}/{total_steps} after appending performance section")
            if not waveform_headers and waveform_items:
                log_message("Appending Waveforms header")
                last_element = doc.element.body[-1]
                new_para = doc.add_paragraph("Waveforms", style='Heading 1')
                last_element.getparent().insert(last_element.getparent().index(last_element) + 1, new_para._element)
                last_element = new_para._element
                self.waveform.add_section(doc, last_element, waveform_items, waveform_files)
                progress_value += len(waveform_items)
                progress_dialog.setValue(progress_value)
                log_message(f"Updated progress: {progress_value}/{total_steps} after appending waveform section")

            doc.save(self.output_path)
            log_message(f"Saved document: {self.output_path}")
            os.startfile(self.output_path)
            for file in os.listdir(self.temp_dir):
                os.remove(os.path.join(self.temp_dir, file))
            remove_directory(self.temp_dir)
            # No QMessageBox.information here; silently complete
            log_message(f"Document {'updated' if self.update_doc_path else 'generated'} and saved as {self.output_path}")

        except FileNotFoundError as e:
            log_message(f"File not found: {str(e)}")
            QtWidgets.QMessageBox.critical(self.app, "Error", str(e))
        except Exception as e:
            log_message(f"Error: {str(e)}")
            QtWidgets.QMessageBox.critical(self.app, "Error", f"Failed to generate/update document: {str(e)}")
        finally:
            progress_dialog.setValue(total_steps)
            log_message("Progress dialog completed")