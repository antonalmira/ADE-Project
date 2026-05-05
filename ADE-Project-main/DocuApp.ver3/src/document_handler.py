import os
import sys
from PyQt5.QtWidgets import QFileDialog, QMessageBox, QProgressDialog
from PyQt5.QtCore import QThread, pyqtSignal
import pythoncom

from document_generator import DocGenerator
from chart_extractor import save_chart_screenshots
from utils import show_popup, ensure_directory, remove_directory

class DocumentWorker(QThread):
    progress_signal = pyqtSignal(int, str)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, app, is_update=False):
        super().__init__()
        self.app = app
        self.is_update = is_update
        self.output_path = app.final_save_destination

    def run(self):
        pythoncom.CoInitialize()
        try:
            has_perf = any(self.app.performancedata_list.item(i).checkState() == 2 for i in range(self.app.performancedata_list.count()))
            if has_perf:
                self.progress_signal.emit(10, "Extracting Excel Charts...")
                try:
                    save_chart_screenshots(self.app, headless=True, progress_callback=self.progress_signal.emit)
                except Exception as e:
                    self.finished_signal.emit(False, f"Excel Chart Extraction Failed: {e}")
                    return

            self.progress_signal.emit(50, "Generating Word Document...")
            doc_gen = DocGenerator(self.app, self.output_path, self.app.update_document_path if self.is_update else "")
            doc_gen.generate(self.progress_signal.emit)

            self.progress_signal.emit(100, "Done!")
            self.finished_signal.emit(True, "Document generated successfully!")
        except Exception as e:
            self.finished_signal.emit(False, str(e))
        finally:
            pythoncom.CoUninitialize()

def get_project_paths():
    """ Calculates paths depending on whether running as a Python script or a compiled .exe """
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

def _update_ui(app, val, text):
    app.progress_dialog.setValue(val)
    app.progress_dialog.setLabelText(text)

def _finish_ui(app, success, msg):
    app.progress_dialog.close()
    if success:
        show_popup(app, "Success", msg, "info")
    else:
        show_popup(app, "Error", f"Failed to generate document:\n{msg}", "error")

def generate_document(app):
    paths = get_project_paths()
    
    # 1. Template
    sel = app.template_dropdown.currentText()
    if not sel or "missing" in sel.lower():
        QMessageBox.warning(app, "Error", "Templates folder missing!")
        return
    app.selected_template_path = os.path.join(paths["templates"], sel)

    # 2. Prompt for BOM and PIXL Files individually
    app.bom_file_path, _ = QFileDialog.getOpenFileName(app, "Select BOM File", paths["output_dir"], "Excel (*.xlsx *.xls)")
    app.pixl_file_path, _ = QFileDialog.getOpenFileName(app, "Select PIXLs / Design Spreadsheet File", paths["output_dir"], "Excel (*.xlsx *.xls)")

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
    update_path, _ = QFileDialog.getOpenFileName(app, "Select Report to Update", paths["output_dir"], "Word (*.docx)")
    if not update_path: return

    app.update_document_path = update_path
    
    # Prompt for BOM and PIXL Files
    app.bom_file_path, _ = QFileDialog.getOpenFileName(app, "Select BOM File", paths["output_dir"], "Excel (*.xlsx *.xls)")
    app.pixl_file_path, _ = QFileDialog.getOpenFileName(app, "Select PIXLs / Design Spreadsheet File", paths["output_dir"], "Excel (*.xlsx *.xls)")

    save_path, _ = QFileDialog.getSaveFileName(app, "Save As", update_path, "Word (*.docx)")
    if save_path:
        app.final_save_destination = save_path 
        run_document_job(app, is_update=True)