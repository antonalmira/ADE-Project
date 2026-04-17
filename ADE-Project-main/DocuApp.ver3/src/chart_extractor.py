# src/chart_extractor.py
import os
import shutil
from pathlib import Path
from PyQt5 import QtWidgets
from PyQt5.QtCore import Qt
import win32com.client
from PIL import Image
import pyautogui
import time

def save_chart_screenshots(app):
    try:
        import win32com.client
    except ImportError:
        QtWidgets.QMessageBox.critical(app, "Error", "The 'pywin32' package is not installed. Please install it using 'pip install pywin32'.")
        print("Error: pywin32 not installed")
        return

    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        print("Excel application initialized")
    except Exception as e:
        QtWidgets.QMessageBox.critical(app, "Error", f"Failed to open Excel: {str(e)}. Ensure Microsoft Excel is installed.")
        print(f"Error initializing Excel: {str(e)}")
        return

    try:
        base_folder = app.performancedata_path.text() if app.performancedata_path.text() and os.path.isdir(app.performancedata_path.text()) else str(Path.home())
        charts_base_dir = os.path.join(base_folder, "Performance Data Charts")
        if os.path.exists(charts_base_dir):
            shutil.rmtree(charts_base_dir)
        os.makedirs(charts_base_dir, exist_ok=True)
        print(f"Created base folder: {charts_base_dir}")

        performance_items = []
        for index in range(app.performancedata_list.count()):
            item = app.performancedata_list.item(index)
            if item.checkState() == Qt.Checked:
                performance_items.append(item.text())

        selected_files = {}
        current_item = None
        for index in range(app.available_data_list_performance.count()):
            item = app.available_data_list_performance.item(index)
            if not item:
                continue
            if item.text() in performance_items:
                current_item = item.text()
                selected_files[current_item] = []
            elif current_item and item.checkState() == Qt.Checked:
                if item.text().lower().endswith(('.xlsx', '.xls')):
                    selected_files[current_item].append(item.text())
        print(f"Selected files: {selected_files}")

        if not any(selected_files.values()):
            QtWidgets.QMessageBox.warning(app, "Warning", "No Excel files selected.")
            excel.Quit()
            print("Warning: No Excel files selected")
            return

        performance_folder = app.performancedata_path.text()
        if not performance_folder or not os.path.isdir(performance_folder):
            QtWidgets.QMessageBox.critical(app, "Error", "Performance data folder is not set or invalid.")
            excel.Quit()
            print("Error: Invalid or unset performance data folder")
            return

        total_files = sum(len(files) for files in selected_files.values())
        progress_dialog = QtWidgets.QProgressDialog("Extracting charts...", "", 0, total_files, app)
        progress_dialog.setWindowTitle("Processing")
        progress_dialog.setWindowModality(Qt.WindowModal)
        progress_dialog.setCancelButton(None)
        progress_dialog.setMinimumDuration(0)
        progress_value = 0
        print("Progress dialog initialized")

        first_file_processed = False
        for item_name, files in selected_files.items():
            item_folder = os.path.join(charts_base_dir, f"{item_name} Charts")
            if os.path.exists(item_folder):
                shutil.rmtree(item_folder)
            os.makedirs(item_folder, exist_ok=True)
            print(f"Created item folder: {item_folder}")

            for file_name in files:
                file_path = os.path.join(performance_folder, file_name)
                file_subfolder = os.path.join(item_folder, os.path.splitext(file_name)[0])
                if os.path.exists(file_subfolder):
                    shutil.rmtree(file_subfolder)
                os.makedirs(file_subfolder, exist_ok=True)
                print(f"Created file subfolder: {file_subfolder}")

                chart_sheets = []
                try:
                    excel_wb = excel.Workbooks.Open(os.path.abspath(file_path))
                    chart_sheets = [sheet.Name for sheet in excel_wb.Sheets if sheet.Type in [-4169, 3]]
                    print(f"Found chart sheets with COM: {chart_sheets}")

                    if not chart_sheets:
                        excel_wb.Close(SaveChanges=False)
                        QtWidgets.QMessageBox.warning(app, "Warning", f"No chart sheets found in {file_name}.")
                        print(f"No chart sheets in {file_name}")
                        progress_value += 1
                        progress_dialog.setValue(progress_value)
                        print(f"Progress updated: {progress_value}/{total_files}")
                        continue

                    first_chart = True
                    for sheet_name in chart_sheets:
                        temp_image = os.path.join(file_subfolder, f"{sheet_name}.png")
                        success = False

                        if not first_file_processed and first_chart:
                            try:
                                chart_sheet = excel_wb.Sheets(sheet_name)
                                chart_sheet.Activate()
                                excel.Visible = True
                                time.sleep(3)
                                temp_ws = excel_wb.Sheets.Add()
                                chart = chart_sheet.Chart
                                if chart:
                                    chart.Copy()
                                    temp_ws.Paste()
                                    temp_ws.Export(temp_image, "PNG")
                                    temp_ws.Delete()
                                    with Image.open(temp_image) as img:
                                        img.save(temp_image, "PNG", quality=95)
                                    print(f"Exported first chart {sheet_name} with Method 2 to {temp_image}")
                                    success = True
                                excel.Visible = False
                            except Exception as e:
                                print(f"Method 2 failed for first chart {sheet_name}: {str(e)}")
                            first_file_processed = True

                        if not success:
                            try:
                                chart_sheet = excel_wb.Sheets(sheet_name)
                                chart_sheet.Activate()
                                for attempt in range(3):
                                    try:
                                        chart_sheet.Export(temp_image, "PNG")
                                        with Image.open(temp_image) as img:
                                            img.save(temp_image, "PNG", quality=95)
                                        print(f"Exported chart {sheet_name} with Method 1 to {temp_image}")
                                        success = True
                                        break
                                    except Exception as e:
                                        if attempt < 2:
                                            time.sleep(1)
                                            print(f"Retry {attempt + 1} for Method 1 ({sheet_name}): {str(e)}")
                                        else:
                                            raise
                            except Exception as e:
                                print(f"Method 1 failed for {sheet_name}: {str(e)}")

                            if not success:
                                try:
                                    chart_sheet = excel_wb.Sheets(sheet_name)
                                    chart_sheet.Activate()
                                    excel.Visible = True
                                    time.sleep(3)
                                    windows = pyautogui.getAllWindows()
                                    excel_window = next((w for w in windows if "Microsoft Excel" in w.title), None)
                                    if excel_window:
                                        left, top = excel_window.left, excel_window.top
                                        width, height = excel_window.width, excel_window.height
                                        screenshot = pyautogui.screenshot(region=(left, top, width, height))
                                        screenshot.save(temp_image)
                                        excel.Visible = False
                                        print(f"Exported chart {sheet_name} with Method 3 to {temp_image}")
                                        success = True
                                except Exception as e:
                                    print(f"Method 3 failed for {sheet_name}: {str(e)}")
                                    excel.Visible = False

                        if not success:
                            QtWidgets.QMessageBox.warning(app, "Warning", f"Failed to export chart {sheet_name} in {file_name} with all methods. Skipping this chart.")
                            print(f"All methods failed for {sheet_name}. Skipping.")

                        first_chart = False

                    excel_wb.Close(SaveChanges=False)
                except Exception as e:
                    QtWidgets.QMessageBox.critical(app, "Error", f"Failed to process {file_name}: {str(e)}")
                    print(f"COM processing error for {file_name}: {str(e)}")
                    continue

                progress_value += 1
                progress_dialog.setValue(progress_value)
                print(f"Progress updated: {progress_value}/{total_files}")

        QtWidgets.QMessageBox.information(app, "Success", "Charts successfully extracted.")
        print("Chart screenshots saved successfully")

    except Exception as e:
        QtWidgets.QMessageBox.critical(app, "Error", f"An error occurred: {str(e)}")
        print(f"General error: {str(e)}")

    finally:
        excel.Quit()
        print("Excel application closed")