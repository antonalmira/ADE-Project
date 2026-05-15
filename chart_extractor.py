import os
import shutil
import win32com.client
from PyQt5.QtCore import Qt

def save_chart_screenshots(app, headless=True, progress_callback=None):
    excel = None
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = not headless
        excel.DisplayAlerts = False

        base_folder = app.performancedata_path.text()
        if not base_folder or not os.path.exists(base_folder): 
            return

        charts_base_dir = os.path.join(base_folder, "Performance Data Charts")
        if os.path.exists(charts_base_dir):
            shutil.rmtree(charts_base_dir)
        os.makedirs(charts_base_dir, exist_ok=True)

        selected_files = {}
        
        for i in range(app.performance_tree.topLevelItemCount()):
            cat_node = app.performance_tree.topLevelItem(i)
            if cat_node.checkState(0) == Qt.Unchecked: 
                continue
            
            cat_name = cat_node.text(0).replace(" [CROP SET]", "")
            selected_files[cat_name] = []
            
            for j in range(cat_node.childCount()):
                file_node = cat_node.child(j)
                if file_node.checkState(0) != Qt.Unchecked:
                    file_name = file_node.data(0, Qt.UserRole + 4) 
                    # Ignore (Table) sheets during bulk image export
                    if file_name and "(Table)" not in file_name and file_name.lower().endswith(('.xlsx', '.xls')):
                        selected_files[cat_name].append(file_name)

        total_files = sum(len(files) for files in selected_files.values())
        if total_files == 0: 
            return
        
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