import os
from PyQt5.QtCore import pyqtSignal, QThread
from PyQt5.QtWidgets import QProgressDialog, QMessageBox
from chart_extractor import save_chart_screenshots
from document_generator import DocGenerator
from utils import log_message

class DocumentWorker(QThread): 
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, app, is_update=False):
        super().__init__()
        self.app = app
        self.is_update = is_update

    def run(self):
        try:
            # 1. Extraction (takes 0% -> 50% of progress)
            save_chart_screenshots(
                self.app, 
                headless=True, 
                progress_callback=self.progress_signal.emit
            ) 
            
            # 2. Document Generation (takes 50% -> 90%)
            self.progress_signal.emit(60, "Building Word document...")
            output_path = self.app.generated_document_path.text()
            update_path = self.app.update_document_path.toolTip()
            
            generator = DocGenerator(self.app, output_path, update_path)
            generator.generate()
            
            # 3. Finalize (100%)
            self.progress_signal.emit(100, "Done!")
            self.finished_signal.emit(True, "Document generated successfully!")
            
        except Exception as e:
            log_message(f"Fatal Error: {str(e)}")
            self.finished_signal.emit(False, str(e))

# UI Helper Functions
def run_document_job(app, is_update=False):
    # Create the dialog
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app) 
    app.progress_dialog.setWindowTitle("Please Wait")
    app.progress_dialog.setWindowModality(2) # Modal
    app.progress_dialog.show()

    # Setup the background thread
    app.worker = DocumentWorker(app, is_update)
    
    # Connect signals to UI updates
    app.worker.progress_signal.connect(lambda val, text: update_progress(app, val, text))
    app.worker.finished_signal.connect(lambda success, msg: finish_processing(app, success, msg))
    
    # Launch thread
    app.worker.start()

def update_progress(app, value, text):
    app.progress_dialog.setValue(value)
    app.progress_dialog.setLabelText(text)

def finish_processing(app, success, msg):
    app.progress_dialog.close()
    if success:
        QMessageBox.information(app, "Success", msg)
    else:
        QMessageBox.critical(app, "Error", f"An error occurred:\n{msg}")