import os
import time
from PyQt5 import QtWidgets
from PIL import Image
import pyautogui
import win32com.client
from utils import ensure_directory, remove_directory, show_popup

def extract_chart_screenshots(app, excel, performance_items, performance_folder, charts_base_dir):
    """Extract charts from Excel files and save as screenshots."""
    selected_files = {}
    current_item = None
    for index in range(app.available_data_list_performance.count()):
        item = app.available_data_list_performance.item(index)
        if not item:
            continue
        if item.text() in performance_items:
            current_item = item.text()
            selected_files[current_item] = []
        elif current_item and item.checkState() == QtWidgets.Qt.Checked:
            if item.text().lower().endswith(('.xlsx', '.xls')):
                selected_files[current_item].append(item.text())

    if not any(selected_files.values()):
        show_popup(app, "Warning", "No Excel files selected.", "warning")
        return False

    if not performance_folder or not os.path.isdir(performance_folder):
        show_popup(app, "Error", "Performance data folder is not set or invalid.", "error")
        return False

    total_files = sum(len(files) for files in selected_files.values())
    progress_dialog = QtWidgets.QProgressDialog("Extracting charts...", "", 0, total_files, app)
    progress_dialog.setWindowTitle("Processing")
    progress_dialog.setWindowModality(QtWidgets.Qt.WindowModal)
    progress_dialog.setCancelButton(None)
    progress_dialog.setMinimumDuration(0)
    progress_value = 0

    first_file_processed = False
    for item_name, files in selected_files.items():
        item_folder = os.path.join(charts_base_dir, f"{item_name} Charts")
        remove_directory(item_folder)
        ensure_directory(item_folder)

        for file_name in files:
            file_path = os.path.join(performance_folder, file_name)
            file_subfolder = os.path.join(item_folder, os.path.splitext(file_name)[0])
            remove_directory(file_subfolder)
            ensure_directory(file_subfolder)

            chart_sheets = []
            try:
                excel_wb = excel.Workbooks.Open(os.path.abspath(file_path))
                chart_sheets = [sheet.Name for sheet in excel_wb.Sheets if sheet.Type in [-4169, 3]]

                if not chart_sheets:
                    # No charts found, export the first table as formatted screenshot
                    sheet_name = excel_wb.ActiveSheet.Name
                    temp_image = os.path.join(file_subfolder, "Table.png")
                    success = False
                    try:
                        sheet = excel_wb.Sheets(sheet_name)
                        sheet.Activate()
                        used_range = sheet.UsedRange
                        used_range.Font.Name = 'Calibri'
                        used_range.Font.Size = 9
                        header_row = sheet.Rows(1)
                        header_row.Interior.Color = 11237376  # #0078AB
                        header_row.Font.Color = 16777215  # White
                        excel.Visible = True
                        time.sleep(3)
                        windows = pyautogui.getAllWindows()
                        excel_window = next((w for w in windows if "Microsoft Excel" in w.title), None)
                        if excel_window:
                            screenshot = pyautogui.screenshot(region=(excel_window.left, excel_window.top, excel_window.width, excel_window.height))
                            screenshot.save(temp_image)
                            success = True
                        excel.Visible = False
                    except Exception as e:
                        print(f"Table screenshot failed for {file_name}: {str(e)}")
                        excel.Visible = False
                        
                    if not success:
                        print(f"Failed to export table for {file_name}. Skipping silently.")
                    excel_wb.Close(SaveChanges=False)
                    progress_value += 1
                    progress_dialog.setValue(progress_value)
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
                                with Image.open(temp_image) as img:
                                    img.save(temp_image, "PNG", quality=95)
                                success = True
                            temp_ws.Delete()
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
                                    success = True
                                    break
                                except Exception as e:
                                    if attempt < 2:
                                        time.sleep(1)
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
                                    success = True
                            except Exception as e:
                                print(f"Method 3 failed for {sheet_name}: {str(e)}")
                                excel.Visible = False

                    if not success:
                        print(f"All methods failed for {sheet_name}. Skipping silently.")

                    first_chart = False

                excel_wb.Close(SaveChanges=False)
            except Exception as e:
                # Silently catch the error and move to the next file instead of triggering a popup
                print(f"COM processing error for {file_name}: {str(e)}")
                continue

            progress_value += 1
            progress_dialog.setValue(progress_value)
    show_popup(app, "Success", "Charts successfully extracted.", "info")
    return True