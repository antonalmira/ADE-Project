import os
import shutil
import time
from pathlib import Path
from PyQt5.QtCore import Qt
import win32com.client
from PIL import Image

def save_chart_screenshots(app, headless=True, progress_callback=None):
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = not headless
        excel.DisplayAlerts = False

        base_folder = app.performancedata_path.text()
        charts_base_dir = os.path.join(base_folder, "Performance Data Charts")
        if os.path.exists(charts_base_dir):
            shutil.rmtree(charts_base_dir)
        os.makedirs(charts_base_dir, exist_ok=True)

        performance_items = [app.performancedata_list.item(i).text() 
                            for i in range(app.performancedata_list.count()) 
                            if app.performancedata_list.item(i).checkState() == Qt.Checked]

        # Filter selected files
        selected_files = {}
        current_item = None
        for index in range(app.available_data_list_performance.count()):
            item = app.available_data_list_performance.item(index)
            if item.text() in performance_items:
                current_item = item.text()
                selected_files[current_item] = []
            elif current_item and item.checkState() == Qt.Checked:
                if item.text().lower().endswith(('.xlsx', '.xls')):
                    selected_files[current_item].append(item.text())

        total_files = sum(len(files) for files in selected_files.values())
        if total_files == 0: return
        
        processed_count = 0
        for item_name, files in selected_files.items():
            item_folder = os.path.join(charts_base_dir, f"{item_name} Charts")
            os.makedirs(item_folder, exist_ok=True)

            for file_name in files:
                processed_count += 1
                if progress_callback:
                    progress_callback(int((processed_count/total_files)*50), f"Excel: {file_name}")

                file_path = os.path.abspath(os.path.join(base_folder, file_name))
                file_subfolder = os.path.join(item_folder, os.path.splitext(file_name)[0])
                os.makedirs(file_subfolder, exist_ok=True)

                wb = excel.Workbooks.Open(file_path, ReadOnly=True)
                chart_sheets = [s for s in wb.Sheets if s.Type in [-4169, 3]]
                
                for sheet in chart_sheets:
                    temp_image = os.path.join(file_subfolder, f"{sheet.Name}.png")
                    try:
                        sheet.Export(temp_image, "PNG")
                    except:
                        continue
                wb.Close(SaveChanges=False)
    finally:
        if excel:
            excel.Quit()