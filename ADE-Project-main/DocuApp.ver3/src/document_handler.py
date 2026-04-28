import os
import pythoncom
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QFileDialog, QProgressDialog, QMessageBox
from document_generator import DocGenerator
from chart_extractor import save_chart_screenshots
from utils import log_message

class DocumentWorker(QThread):
    """Handles the heavy lifting of chart extraction and document generation in a separate thread."""
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
            self.progress_signal.emit(55, "Opening Document Template...")
            output_path = getattr(self.app, 'final_save_destination', "Generated_Document.docx")
            update_path = getattr(self.app, 'update_document_path', "") if self.is_update else ""
            
            generator = DocGenerator(self.app, output_path, update_path)
            
            # 3. Generate Content (60% to 95%)
            self.progress_signal.emit(60, "Processing Sections and Cropping...")
            generator.generate(progress_callback=self.progress_signal.emit)
            
            # 4. Finalize
            self.progress_signal.emit(100, "Finalizing...")
            self.finished_signal.emit(True, "Document successfully generated and saved!")
        except Exception as e:
            log_message(f"Worker Error: {str(e)}")
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def get_automatic_bom_path():
    """
    Calculates the BOM path automatically.
    Moves from: ADE-Project-main/DocuApp.ver3/src
    To: ADE-Project-main/resource/BOM_PIXL.xlsx
    """
    try:
        current_dir = os.path.dirname(os.path.abspath(__file__))
        # Go up two levels: from 'src' to 'DocuApp.ver3' to 'ADE-Project-main'
        project_root = os.path.abspath(os.path.join(current_dir, "..", ".."))
        bom_path = os.path.join(project_root, "resource", "BOM_PIXL.xlsx")
        
        if os.path.exists(bom_path):
            return bom_path
        return None
    except:
        return None

def run_document_job(app, is_update=False):
    """Initializes the progress dialog and starts the worker thread."""
    app.progress_dialog = QProgressDialog("Initializing...", None, 0, 100, app)
    app.progress_dialog.setWindowTitle("Processing Document")
    app.progress_dialog.setModal(True)
    app.progress_dialog.setAutoClose(False)
    app.progress_dialog.setAutoReset(False)
    app.progress_dialog.setMinimumDuration(0) 
    
    # Apply clean Light Mode stylesheet to fix invisible text bugs
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
    """Updates the progress dialog text and bar."""
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.setValue(value)
        black_text = f'<span style="color: black;">{text}</span>'
        app.progress_dialog.setLabelText(black_text)

def _finish_ui(app, success, message):
    """Displays the final result popup."""
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
    """
    Handles the 'Generate Document' flow:
    1. Validation
    2. Auto-Detect BOM
    3. Save File Dialog (Once)
    4. Start Process
    """
    # 1. Validate Template selection
    selected_template = app.template_dropdown.currentText()
    if not selected_template or selected_template in ["No templates found", "Templates folder missing!", ""]:
        QMessageBox.warning(app, "Missing Template", "Please choose a valid template.")
        return

    # 2. NEW LOGIC: Point to the new template location in the project root
    current_dir = os.path.dirname(os.path.abspath(__file__))
    templates_folder = os.path.abspath(os.path.join(current_dir, "..", "..", "templates"))
    app.selected_template_path = os.path.join(templates_folder, selected_template)

    # 3. AUTOMATIC BOM FILE DETECTION
    app.bom_file_path = get_automatic_bom_path()
    if app.bom_file_path:
        log_message(f"BOM Auto-detected: {app.bom_file_path}")
    else:
        log_message("BOM file not found in resource folder. Skipping BOM section.")

    # 4. Save Dialog (ONLY CALLED ONCE)
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
    """Handles the 'Update Existing Report' flow."""
    # 1. Prompt for existing doc
    update_path, _ = QFileDialog.getOpenFileName(
        app, "Select Existing Report to Update", "", "Word Documents (*.docx)"
    )
    if not update_path:
        return 

    app.update_document_path = update_path

    # 2. Automatic BOM
    app.bom_file_path = get_automatic_bom_path()

    # 3. Save Dialog
    save_path, _ = QFileDialog.getSaveFileName(
        app, "Save Updated Report As", update_path, "Word Documents (*.docx)"
    )

    if save_path:
        app.final_save_destination = save_path 
        run_document_job(app, is_update=True)