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
            save_chart_screenshots(
                self.app, 
                headless=True, 
                progress_callback=self.progress_signal.emit
            )
            
            self.progress_signal.emit(50, "Generating Word document...")
            output_path = getattr(self.app, 'final_save_destination', "Generated_Document.docx")
            update_path = getattr(self.app, 'update_document_path', "") if self.is_update else ""
            
            generator = DocGenerator(self.app, output_path, update_path)
            generator.generate(progress_callback=self.progress_signal.emit)
            
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated!")
        except Exception as e:
            log_message(f"Worker Error: {str(e)}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def run_document_job(app, is_update=False):
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.show()

    app.worker = DocumentWorker(app, is_update)
    app.worker.progress_signal.connect(lambda val, text: _update_ui(app, val, text))
    app.worker.finished_signal.connect(lambda success, msg: _finish_ui(app, success, msg))
    app.worker.start()

def _update_ui(app, value, text):
    app.progress_dialog.setValue(value)
    app.progress_dialog.setLabelText(text)

def _finish_ui(app, success, message):
    app.progress_dialog.close()
    if success:
        QMessageBox.information(app, "Success", message)
    else:
        QMessageBox.critical(app, "Error", f"Failed: {message}")

def generate_document(app):
    selected_template = app.template_dropdown.currentText()
    
    # Validate the selection
    if not selected_template or selected_template in ["No templates found", "Templates folder missing!", ""]:
        QMessageBox.warning(app, "Missing Template", "Please choose a valid template (.docx) from the dropdown.")
        return

    # Build the full absolute path to the selected template
    templates_folder = os.path.join(os.path.dirname(__file__), "templates")
    full_template_path = os.path.join(templates_folder, selected_template)
    
    app.selected_template_path = full_template_path

    save_path, _ = QFileDialog.getSaveFileName(
        app, 
        "Save Generated Document", 
        "Generated_Document.docx", 
        "Word Documents (*.docx)"
    )

    if save_path:
        # Save the destination to the app instance so the worker can access it
        app.final_save_destination = save_path 
        run_document_job(app, is_update=False)

def update_document(app):
    run_document_job(app, is_update=True)