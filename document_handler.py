import os
import sys
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QFileDialog, QProgressDialog, QMessageBox
from document_generator import DocGenerator
from utils import log_message

class DocumentWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, app):
        super().__init__()
        self.app = app

    def run(self):
        # We must initialize COM on background threads for Excel/Word automation
        pythoncom.CoInitialize()
        try:
            self.progress_signal.emit(10, "Opening Document Template...")
            output_path = getattr(self.app, 'final_save_destination', "Generated_Document.docx")
            
            generator = DocGenerator(self.app, output_path)
            
            self.progress_signal.emit(30, "Processing Sections and Extracting Data...")
            generator.generate(progress_callback=self.progress_signal.emit)
            
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated!")
        except Exception as e:
            log_message(f"Worker Error: {str(e)}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def get_project_paths():
    if getattr(sys, 'frozen', False):
        project_root = sys._MEIPASS
        output_folder = os.path.join(os.path.dirname(sys.executable), "output")
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.abspath(os.path.join(current_dir, "..", ".."))
        output_folder = os.path.join(project_root, "output")
    
    paths = {
        "templates": os.path.join(project_root, "templates"),
        "output_dir": output_folder
    }
    return paths

def run_document_job(app):
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.setMinimumDuration(0) 
    app.progress_dialog.setStyleSheet("QProgressDialog { background-color: #f5f5f5; } QLabel { color: black; }")
    app.progress_dialog.setValue(0)
    app.progress_dialog.show()

    app.worker = DocumentWorker(app)
    app.worker.progress_signal.connect(lambda val, text: _update_ui(app, val, text))
    app.worker.finished_signal.connect(lambda success, msg: _finish_ui(app, success, msg))
    app.worker.start()

def _update_ui(app, value, text):
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.setValue(value)
        app.progress_dialog.setLabelText(f'<span style="color: black;">{text}</span>')

def _finish_ui(app, success, message):
    if hasattr(app, 'progress_dialog'): app.progress_dialog.close()
    msg = QMessageBox(app)
    msg.setStyleSheet("QMessageBox { background-color: #f5f5f5; } QLabel { color: black; }")
    msg.setIcon(QMessageBox.Information if success else QMessageBox.Critical)
    msg.setText(message)
    msg.exec_()

def generate_document(app):
    paths = get_project_paths()
    
    # 1. Template
    sel = app.template_dropdown.currentText()
    if not sel or "missing" in sel.lower():
        QMessageBox.warning(app, "Error", "Templates folder missing!")
        return
    app.selected_template_path = os.path.join(paths["templates"], sel)

    # 2. Extract Specific Files From UI
    app.bom_file_path = app.bom_path_display.text().strip() if app.bom_path_display.text() else None
    app.pixls_file_path = app.pix_path_display.text().strip() if app.pix_path_display.text() else None

    # 3. Output Path
    if not os.path.exists(paths["output_dir"]):
        os.makedirs(paths["output_dir"])

    save_path, _ = QFileDialog.getSaveFileName(
        app, "Save Report", 
        os.path.join(paths["output_dir"], "Generated_Report.docx"), 
        "Word (*.docx)"
    )
    
    if save_path:
        app.final_save_destination = save_path
        run_document_job(app)