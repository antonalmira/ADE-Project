import os
from docx import Document
from PyQt5.QtCore import Qt
from performance_section import PerformanceSection
from waveform_section import WaveformSection
from utils import ensure_directory, remove_directory, get_default_base_folder, get_resource_path

class DocGenerator:
    # Changed 'output_folder' to 'output_path' since the Save Dialog gives us the full file path
    def __init__(self, app, output_path, update_doc_path=""):
        self.app = app
        
        # 1. Check if the app has a selected template from the dropdown
        if hasattr(app, 'selected_template_path') and app.selected_template_path:
            self.template_path = app.selected_template_path
        else:
            # Fallback just in case something goes wrong
            self.template_path = get_resource_path('DER-template.docx')
            
        # 2. Set the exact save destination chosen by the user
        self.output_path = output_path
        
        self.update_doc_path = update_doc_path
        self.temp_dir = "temp_cropped_images"
        ensure_directory(self.temp_dir)
        
        self.performance = PerformanceSection(app, self.temp_dir)
        self.waveform = WaveformSection(app, self.temp_dir)

    def generate(self, progress_callback=None):
        # Choose between an update document or the chosen template
        doc_path = self.update_doc_path if self.update_doc_path and os.path.exists(self.update_doc_path) else self.template_path
        doc = Document(doc_path)
        
        # Get checked items from the UI
        performance_items = [self.app.performancedata_list.item(i).text() for i in range(self.app.performancedata_list.count()) if self.app.performancedata_list.item(i).checkState() == Qt.Checked]
        waveform_items = [self.app.waveforms_list.item(i).text() for i in range(self.app.waveforms_list.count()) if self.app.waveforms_list.item(i).checkState() == Qt.Checked]

        # Performance Section
        if performance_items:
            if progress_callback: progress_callback(70, "Writing Performance Data...")
            perf_data = self.performance.get_data(performance_items)
            self.performance.add_section(doc, doc.element.body[-1], performance_items, perf_data, None)

        # Waveform Section
        if waveform_items:
            if progress_callback: progress_callback(90, "Writing Waveforms...")
            wave_files = self.waveform.get_images_with_custom_crop(waveform_items)
            self.waveform.add_section(doc, doc.element.body[-1], waveform_items, wave_files)

        # Save and open the document
        doc.save(self.output_path)
        os.startfile(self.output_path)
        remove_directory(self.temp_dir)