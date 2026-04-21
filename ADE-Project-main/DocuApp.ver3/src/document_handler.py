import os
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QProgressDialog, QMessageBox
from document_generator import DocGenerator
from chart_extractor import save_chart_screenshots
from utils import log_message

class DocumentWorker(QThread):
    # Signals for safely communicating with the PyQt GUI thread
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, app, is_update=False):
        super().__init__()
        self.app = app
        self.is_update = is_update

    def run(self):
        try:
            # 1. Start Extraction
            self.progress_signal.emit(10, "Extracting chart screenshots...")
            
            # NOTE: save_chart_screenshots() must no longer create its own QProgressDialog or QMessageBox
            save_chart_screenshots(self.app, headless=True)
            
            # 2. Start Generation
            self.progress_signal.emit(50, "Generating Word document...")
            output_path = self.app.generated_document_path.text()
            update_path = self.app.update_document_path.toolTip() if self.is_update else ""
            
            generator = DocGenerator(self.app, output_path, update_path)
            generator.generate()
            
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated!")
        except Exception as e:
            log_message(f"Worker Thread Error: {str(e)}")
            self.finished_signal.emit(False, str(e))

def _run_document_job(app, is_update):
    """Initializes the worker and the main thread's progress dialog."""
    # 1. Create a GUI Progress Dialog in the Main Thread
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.show()

    # 2. Create and connect the worker
    app.worker = DocumentWorker(app, is_update)
    app.worker.progress_signal.connect(lambda val, text: _update_progress(app, val, text))
    app.worker.finished_signal.connect(lambda success, msg: _finish_processing(app, success, msg))
    
    # 3. Start the background execution
    app.worker.start()

def _update_progress(app, value, text):
    app.progress_dialog.setValue(value)
    app.progress_dialog.setLabelText(text)

def _finish_processing(app, success, message):
    app.progress_dialog.close()
    if success:
        QMessageBox.information(app, "Success", message)
    else:
        QMessageBox.critical(app, "Error", f"Failed: {message}")

# --- Exposed Handlers ---
def generate_document(app):
    log_message("Starting document generation via worker")
    _run_document_job(app, is_update=False)

def update_document(app):
    log_message("Starting document update via worker")
    _run_document_job(app, is_update=True)