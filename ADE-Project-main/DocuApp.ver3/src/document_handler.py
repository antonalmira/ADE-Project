import os
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QFileDialog, QProgressDialog, QMessageBox
from document_generator import DocGenerator
from chart_extractor import save_chart_screenshots
from utils import log_message

class DocumentWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, app, is_update=False):
        super().__init__()
        self.app = app
        self.is_update = is_update

    def run(self):
        pythoncom.CoInitialize()
        try:
            # 1. Start Extraction (0% to 50%)
            self.progress_signal.emit(5, "Starting chart extraction...")
            save_chart_screenshots(
                self.app, 
                headless=True, 
                progress_callback=self.progress_signal.emit
            )
            
            # 2. Setup Generator (50% to 60%)
            self.progress_signal.emit(55, "Opening Document...")
            output_path = getattr(self.app, 'final_save_destination', "Generated_Document.docx")
            update_path = getattr(self.app, 'update_document_path', "") if self.is_update else ""
            
            generator = DocGenerator(self.app, output_path, update_path)
            
            # 3. Generate Content (60% to 90% handled inside generator)
            self.progress_signal.emit(60, "Configuring Document Sections...")
            generator.generate(progress_callback=self.progress_signal.emit)
            
            # 4. Finalize (95% to 100%)
            self.progress_signal.emit(95, "Saving Document...")
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated and saved!")
        except Exception as e:
            log_message(f"Worker Error: {str(e)}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def run_document_job(app, is_update=False):
    # Setup the Progress Dialog
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.setAutoClose(False)
    app.progress_dialog.setAutoReset(False)
    app.progress_dialog.setMinimumDuration(0) 
    
    # Apply clean Light Mode stylesheet to fix invisible text
    app.progress_dialog.setStyleSheet("""
        QProgressDialog { background-color: #f5f5f5; }
        QLabel { color: #000000; font-size: 12px; font-weight: bold; }
        QProgressBar { border: 1px solid #b0b0b0; border-radius: 4px; text-align: center; color: #000000; background-color: #e0e0e0; }
        QProgressBar::chunk { background-color: #0085ca; width: 15px; }
    """)
    
    app.progress_dialog.setValue(0)
    app.progress_dialog.show()

    app.worker = DocumentWorker(app, is_update)
    app.worker.progress_signal.connect(lambda val, text: _update_ui(app, val, text))
    app.worker.finished_signal.connect(lambda success, msg: _finish_ui(app, success, msg))
    app.worker.start()

def _update_ui(app, value, text):
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.setValue(value)
        black_text = f'<span style="color: black;">{text}</span>'
        app.progress_dialog.setLabelText(black_text)

def _finish_ui(app, success, message):
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.close()
    
    msg_box = QMessageBox(app)
    msg_box.setStyleSheet("""
        QMessageBox { background-color: #f5f5f5; }
        QLabel { color: #000000; font-size: 13px; }
        QPushButton { background-color: #0085ca; color: #ffffff; border-radius: 4px; padding: 6px 20px; font-weight: bold; min-width: 60px; }
        QPushButton:hover { background-color: #3c649f; }
    """)
    
    if success:
        msg_box.setIcon(QMessageBox.Information)
        msg_box.setWindowTitle("Success")
        msg_box.setText("<b>Task Completed Successfully!</b>")
        msg_box.setInformativeText(message)
    else:
        msg_box.setIcon(QMessageBox.Critical)
        msg_box.setWindowTitle("Error")
        msg_box.setText("<b>An Error Occurred.</b>")
        msg_box.setInformativeText(message)
        
    msg_box.exec_()

def generate_document(app):
    """Handles the 'Generate Document' button flow."""
    # 1. Validate the Template selection from dropdown
    selected_template = app.template_dropdown.currentText()
    if not selected_template or selected_template in ["No templates found", "Templates folder missing!", ""]:
        QMessageBox.warning(app, "Missing Template", "Please choose a valid template (.docx) from the dropdown.")
        return

    # 2. Build the full absolute path to the selected template
    templates_folder = os.path.join(os.path.dirname(__file__), "templates")
    app.selected_template_path = os.path.join(templates_folder, selected_template)

    # 3. Prompt for BOM Excel File (Optional)
    bom_path, _ = QFileDialog.getOpenFileName(
        app, "Optional: Select BOM Spreadsheet (Cancel to skip)", "", "Excel Files (*.xlsx *.xls)"
    )
    app.bom_file_path = bom_path if bom_path else None

    # 4. Prompt for Save Destination (Asking only once here)
    save_path, _ = QFileDialog.getSaveFileName(
        app, 
        "Save Generated Document", 
        "Generated_Document.docx", 
        "Word Documents (*.docx)"
    )
    
    if save_path:
        app.final_save_destination = save_path
        run_document_job(app, is_update=False)

def update_document_prompt(app):
    """Handles the 'Update Existing Report' button flow."""
    # 1. Prompt for the existing document to update
    update_path, _ = QFileDialog.getOpenFileName(
        app, 
        "Select Existing Report to Update", 
        "", 
        "Word Documents (*.docx)"
    )
    
    if not update_path:
        return # User cancelled

    app.update_document_path = update_path

    # 2. Prompt where to save the newly updated document
    save_path, _ = QFileDialog.getSaveFileName(
        app, 
        "Save Updated Report As", 
        update_path, # Defaults to the existing filename
        "Word Documents (*.docx)"
    )

    if save_path:
        app.final_save_destination = save_path 
        run_document_job(app, is_update=True)