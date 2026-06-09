import os
import shutil
import json
import re
from PyQt5.QtCore import QThread, pyqtSignal
from PyQt5.QtWidgets import QFileDialog, QProgressDialog, QMessageBox
import win32com.client
import pythoncom
from excel_utils import peek_table_voltages
from list_updater import update_performance_tree, update_waveform_tree


def select_template_file(app):
    current_dir = os.path.dirname(os.path.abspath(__file__))
    default_path = os.path.join(current_dir, "templates")
    if not os.path.exists(default_path):
        default_path = os.path.join(os.path.expanduser("~"), "Desktop")

    file, _ = QFileDialog.getOpenFileName(
        app, "Select Word Template", default_path, "Word Documents (*.docx)"
    )
    if file:
        app.template_path_display.setText(file)
        app.template_path_display.setToolTip(file)


def add_waveform_folder(app):
    folder = QFileDialog.getExistingDirectory(app, "Select Waveform Folder", "")

    if folder:
        folder = os.path.normpath(folder)
        current = app.waveforms_path.text().strip()

        if current:
            existing_folders = [p.strip() for p in current.split(';') if p.strip()]
            if folder not in existing_folders:
                existing_folders.append(folder)
            app.waveforms_path.setText("; ".join(existing_folders))
        else:
            app.waveforms_path.setText(folder)

        update_waveform_tree(app)


def clear_waveform_folders(app):
    app.waveforms_path.clear()
    app.waveform_tree.clear()


def select_bom_file(app):
    file, _ = QFileDialog.getOpenFileName(
        app, "Select Bill of Materials (BOM)", "", "Excel Files (*.xlsx *.xls)"
    )
    if file:
        app.bom_path_display.setText(file)


def select_pixls_file(app):
    file, _ = QFileDialog.getOpenFileName(
        app, "Select Design Spreadsheet (PIXLs)", "", "Excel Files (*.xlsx *.xls)"
    )
    if file:
        app.pix_path_display.setText(file)


def clear_performance_folders(app):
    app.performancedata_path.clear()
    app.performance_tree.clear()
    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    if os.path.exists(perf_dir):
        shutil.rmtree(perf_dir)


# ---------------------------------------------------------------------------
# Performance data extraction worker
# ---------------------------------------------------------------------------

def _map_sheet_to_category(sheet_name):
    name_lower = sheet_name.lower().replace("-", " ").replace("_", " ")
    if "no load" in name_lower:
        return "No-Load Input Power"
    if "line eff" in name_lower:
        return "Full Load Efficiency vs. Line"
    if "load eff" in name_lower or "efficiency vs" in name_lower:
        return "Efficiency vs. Load"
    if "line reg" in name_lower:
        return "Line Regulation"
    if "load reg" in name_lower or "load ref" in name_lower:
        return "Load Regulation"
    if "eff" in name_lower and "table" in name_lower:
        return "Average and 10% Efficiency"
    if "average" in name_lower or "10%" in name_lower:
        return "Average and 10% Efficiency"
    return re.sub(r'\((Graph|Table|graph|table)\)', '', name_lower).strip().title()


class PerfImportWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, excel_paths, perf_dir):
        super().__init__()
        # Accept a single path or a list of paths
        if isinstance(excel_paths, str):
            self.excel_paths = [excel_paths]
        else:
            self.excel_paths = excel_paths
        self.perf_dir = perf_dir

    def run(self):
        pythoncom.CoInitialize()
        excel = None
        try:
            self.progress.emit(5, "Opening Excel...")
            excel = win32com.client.DispatchEx("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False

            os.makedirs(self.perf_dir, exist_ok=True)
            metadata = {}

            total_files = len(self.excel_paths)
            if total_files == 0:
                raise ValueError("No files selected.")

            for file_idx, excel_path in enumerate(self.excel_paths):
                base_pct = 5 + (file_idx / total_files) * 90
                next_pct = 5 + ((file_idx + 1) / total_files) * 90

                filename = os.path.basename(excel_path)
                self.progress.emit(int(base_pct), f"Processing {filename}...")

                wb = excel.Workbooks.Open(os.path.abspath(excel_path), ReadOnly=True)
                try:
                    sheets = [
                        s for s in wb.Sheets
                        if re.search(r'\((Graph|Table)\)', s.Name, re.IGNORECASE)
                    ]

                    if not sheets:
                        continue

                    for i, sheet in enumerate(sheets):
                        pct = int(base_pct + (i / len(sheets)) * (next_pct - base_pct))
                        self.progress.emit(pct, f"Extracting {sheet.Name} from {filename}...")

                        category = _map_sheet_to_category(sheet.Name)
                        cat_dir = os.path.join(self.perf_dir, category)
                        os.makedirs(cat_dir, exist_ok=True)

                        is_table = "(Table)" in sheet.Name

                        png_name = f"{sheet.Name}.png"
                        if total_files > 1:
                            base_filename = os.path.splitext(filename)[0]
                            png_name = f"{base_filename} - {sheet.Name}.png"

                        # Store path relative to perf_dir to avoid embedding
                        # absolute machine-specific paths in metadata.json
                        rel_png = os.path.join(
                            os.path.relpath(cat_dir, self.perf_dir), png_name
                        )
                        png_path = os.path.abspath(os.path.join(cat_dir, png_name))

                        try:
                            if is_table:
                                rng = sheet.UsedRange
                                rng.CopyPicture(Appearance=1, Format=2)
                                chart_obj = sheet.ChartObjects().Add(0, 0, rng.Width, rng.Height)
                                chart_obj.Chart.Paste()
                                chart_obj.Chart.Export(png_path, "PNG")
                                chart_obj.Delete()
                            else:
                                if sheet.Type in [-4169, 3]:
                                    sheet.Export(png_path, "PNG")
                                elif sheet.ChartObjects().Count > 0:
                                    sheet.ChartObjects(1).Chart.Export(png_path, "PNG")
                        except Exception as e:
                            print(f"Failed to export png for {sheet.Name}: {e}")

                        if not os.path.exists(png_path):
                            continue

                        if is_table:
                            voltages = peek_table_voltages(excel_path, sheet.Name)
                            if voltages:
                                for volt in voltages:
                                    split_name = f"{sheet.Name} - {volt} VAC"
                                    if total_files > 1:
                                        base_filename = os.path.splitext(filename)[0]
                                        split_name = f"{base_filename} - {split_name}"

                                    split_png_abs = os.path.abspath(
                                        os.path.join(cat_dir, f"{split_name}.png")
                                    )
                                    split_rel = os.path.join(
                                        os.path.relpath(cat_dir, self.perf_dir),
                                        f"{split_name}.png"
                                    )
                                    shutil.copy(png_path, split_png_abs)
                                    metadata[split_rel] = {
                                        "type": "table",
                                        "excel_path": excel_path,
                                        "sheet_name": sheet.Name,
                                        "voltage": volt,
                                        "original_name": f"{split_name}.png"
                                    }
                                os.remove(png_path)
                                continue

                        metadata[rel_png] = {
                            "type": "table" if is_table else "graph",
                            "excel_path": excel_path,
                            "sheet_name": sheet.Name,
                            "voltage": None,
                            "original_name": png_name
                        }
                finally:
                    try:
                        wb.Close(SaveChanges=False)
                    except Exception:
                        pass

            with open(os.path.join(self.perf_dir, "metadata.json"), "w") as f:
                json.dump(metadata, f)

            self.progress.emit(100, "Done!")
            self.finished.emit(True, "Success")

        except Exception as e:
            self.finished.emit(False, str(e))
        finally:
            if excel is not None:
                try:
                    excel.Quit()
                except Exception:
                    pass
            pythoncom.CoUninitialize()


def update_perf_progress(app, val, text):
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.setValue(val)
        app.progress_dialog.setLabelText(text)


def select_performance_file(app):
    files, _ = QFileDialog.getOpenFileNames(
        app, "Select Performance Data File(s)", "",
        "Excel Files (*.xlsx *.xls *.xlsm)"
    )
    if not files:
        return

    # Guard against launching a second worker while one is already running
    if hasattr(app, 'perf_worker') and app.perf_worker is not None:
        if app.perf_worker.isRunning():
            QMessageBox.warning(
                app, "Busy",
                "An import is already in progress. Please wait for it to finish."
            )
            return

    app.performancedata_path.setText("; ".join(files))

    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    if os.path.exists(perf_dir):
        shutil.rmtree(perf_dir)

    app.progress_dialog = QProgressDialog(
        "Importing Excel and Generating Structure...", "Cancel", 0, 100, app
    )
    app.progress_dialog.setWindowTitle("Processing")
    app.progress_dialog.setModal(True)
    app.progress_dialog.setMinimumDuration(0)
    app.progress_dialog.show()

    app.perf_worker = PerfImportWorker(files, perf_dir)
    app.perf_worker.progress.connect(lambda val, text: update_perf_progress(app, val, text))
    app.perf_worker.finished.connect(lambda s, m: on_perf_import_finished(app, s, m))
    app.perf_worker.start()


def on_perf_import_finished(app, success, msg):
    if hasattr(app, 'progress_dialog'):
        app.progress_dialog.close()
    if success:
        update_performance_tree(app)
    else:
        QMessageBox.warning(app, "Error", f"Import failed: {msg}")