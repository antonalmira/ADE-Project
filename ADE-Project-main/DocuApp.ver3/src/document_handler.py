import os
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QProgressDialog, QMessageBox
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
        # Mandatory for Excel COM to work in a thread
        pythoncom.CoInitialize()
        try:
            # 1. Start Extraction (Reports progress 0-50)
            save_chart_screenshots(
                self.app, 
                headless=True, 
                progress_callback=self.progress_signal.emit
            )
            
            # 2. Start Generation (Reports progress 50-100)
            self.progress_signal.emit(50, "Generating Word document...")
            output_path = self.app.generated_document_path.text()
            update_path = self.app.update_document_path.toolTip() if self.is_update else ""
            
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
    run_document_job(app, is_update=False)

def update_document(app):
    run_document_job(app, is_update=True)