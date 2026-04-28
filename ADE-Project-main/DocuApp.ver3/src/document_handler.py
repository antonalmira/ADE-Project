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
            self.progress_signal.emit(5, "Starting chart extraction...")
            save_chart_screenshots(self.app, headless=True, progress_callback=self.progress_signal.emit)
            
            self.progress_signal.emit(55, "Opening Document Template...")
            output_path = getattr(self.app, 'final_save_destination', "Generated_Document.docx")
            update_path = getattr(self.app, 'update_document_path', "") if self.is_update else ""
            
            generator = DocGenerator(self.app, output_path, update_path)
            
            self.progress_signal.emit(60, "Processing Sections and Cropping...")
            generator.generate(progress_callback=self.progress_signal.emit)
            
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated!")
        except Exception as e:
            log_message(f"Worker Error: {str(e)}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def get_project_paths():
    """
    Calculates paths relative to the script location.
    Works on any laptop as long as the folder structure is kept.
    """
    # Ver3/src/document_handler.py
    current_dir = os.path.dirname(os.path.abspath(__file__))
    # Ver3/
    ver_root = os.path.dirname(current_dir)
    # ADE-Project-main/
    project_root = os.path.dirname(ver_root)
    
    paths = {
        "bom": os.path.join(project_root, "resource", "BOM_PIXL.xlsx"),
        "output_dir": os.path.join(project_root, "output"),
        "templates": os.path.join(project_root, "templates")
    }
    return paths

def run_document_job(app, is_update=False):
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.setMinimumDuration(0) 
    app.progress_dialog.setStyleSheet("QProgressDialog { background-color: #f5f5f5; } QLabel { color: black; }")
    app.progress_dialog.setValue(0)
    app.progress_dialog.show()

    app.worker = DocumentWorker(app, is_update)
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

    # 2. Automatic BOM/PIX File
    app.bom_file_path = paths["bom"] if os.path.exists(paths["bom"]) else None

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
        run_document_job(app, is_update=False)

def update_document_prompt(app):
    paths = get_project_paths()
    update_path, _ = QFileDialog.getOpenFileName(app, "Select Report", paths["output_dir"], "Word (*.docx)")
    if not update_path: return

    app.update_document_path = update_path
    app.bom_file_path = paths["bom"] if os.path.exists(paths["bom"]) else None

    save_path, _ = QFileDialog.getSaveFileName(app, "Save As", update_path, "Word (*.docx)")
    if save_path:
        app.final_save_destination = save_path 
        run_document_job(app, is_update=True)