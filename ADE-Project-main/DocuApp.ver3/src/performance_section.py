import os
import openpyxl
import re
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PyQt5.QtCore import Qt
from utils import log_message
from excel_utils import extract_excel_table
from word_utils import add_styled_table, add_caption_field, format_value_units
from image_utils import crop_and_save
from list_updater import performancedata_testnames

class PerformanceSection:
    def __init__(self, app, temp_dir):
        self.app = app
        self.temp_dir = temp_dir

    def get_first_two_words(self, filename):
        words = re.split(r'\s+|-|_', filename.lower())
        return ' '.join(words[:2]).strip()

    def get_data(self, performance_items):
        performance_folder = self.app.performancedata_path.text()
        charts_base_dir = os.path.join(performance_folder, "Performance Data Charts")
        performance_data = {}
        
        # Get the ordered list of files from available_data_list_performance
        ordered_files = []
        current_item = None
        for index in range(self.app.available_data_list_performance.count()):
            item = self.app.available_data_list_performance.item(index)
            if not item:
                continue
            if item.text() in performance_items:
                current_item = item.text()
            elif current_item and item.checkState() == Qt.Checked and item.text().lower().endswith(('.xlsx', '.xls')):
                custom_cap = item.data(Qt.UserRole)
                ordered_files.append((current_item, item.text(), custom_cap))
        log_message(f"Ordered performance files from UI: {ordered_files}")

        for item_name in performance_items:
            performance_data[item_name] = {'charts': [], 'tables': []}
            prefix = next((key for key, value in performancedata_testnames.items() if value == item_name), '')
            item_folder = os.path.join(charts_base_dir, f"{item_name} Charts")
            
            # Process charts in the order of ordered_files
            if os.path.isdir(item_folder):
                for ordered_item_name, file_name, custom_cap in ordered_files:
                    if ordered_item_name != item_name:
                        continue
                    subfolder = os.path.splitext(file_name)[0]
                    subfolder_path = os.path.join(item_folder, subfolder)
                    if os.path.isdir(subfolder_path):
                        # Check for Table.png first
                        table_image = os.path.join(subfolder_path, "Table.png")
                        if os.path.exists(table_image):
                            performance_data[item_name]['charts'].append({'path': table_image, 'custom_cap': custom_cap})
                        # Then collect other chart images (e.g., chart sheets)
                        for file in sorted(os.listdir(subfolder_path)):
                            if file.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp')) and file != "Table.png":
                                performance_data[item_name]['charts'].append({'path': os.path.join(subfolder_path, file), 'custom_cap': custom_cap})

            # Process tables in the order of ordered_files
            for ordered_item_name, file_name, custom_cap in ordered_files:
                if ordered_item_name != item_name:
                    continue
                if self.get_first_two_words(file_name) == prefix:
                    file_path = os.path.join(performance_folder, file_name)
                    try:
                        wb = openpyxl.load_workbook(file_path, data_only=True)
                        ws = wb.active
                        table_data, merged_cells = extract_excel_table(ws)
                        if table_data:
                            performance_data[item_name]['tables'].append({
                                'file_name': file_name,
                                'data': table_data,
                                'merged_cells': merged_cells,
                                'custom_cap': custom_cap
                            })
                        wb.close()
                    except Exception as e:
                        log_message(f"Error extracting table from {file_name}: {str(e)}")
            log_message(f"Performance data for {item_name}: {performance_data[item_name]}")
        return performance_data

    def get_efficiency_table(self):
        return None

    def add_section(self, doc, last_element, performance_items, performance_data, efficiency_table):
        for item in performance_items:
            log_message(f"Adding performance subheader: {item}")
            new_para = doc.add_paragraph(item, style='Heading 2')
            new_para.runs[0].font.size = Pt(12)
            last_element.getparent().insert(last_element.getparent().index(last_element) + 1, new_para._element)
            last_element = new_para._element

            # Add charts (including Table.png)
            for chart_data in performance_data.get(item, {}).get('charts', []):
                image_path = chart_data['path']
                custom_cap = chart_data.get('custom_cap')
                cropped_path = crop_and_save(image_path, 2, 2, 2, 2, self.temp_dir)
                
                if cropped_path:
                    img_para = doc.add_paragraph()
                    img_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run = img_para.add_run()
                    run.add_picture(cropped_path, width=Inches(6.5))
                    last_element.getparent().insert(last_element.getparent().index(last_element) + 1, img_para._element)
                    last_element = img_para._element
                    file_name = os.path.basename(image_path).lower()
                    
                    if custom_cap:
                        caption_text = custom_cap
                    else:
                        if 'table' in file_name:
                            caption_text = f"{item} Table"
                        else:
                            cap = f"{item} Chart"
                            caption_text = format_value_units(cap)
                            
                            edited_name = file_name[:-4].upper()
                            categories = edited_name.split("_")
                            current = categories[-1] if categories else ""
                            volts = categories[-2] if len(categories) > 1 else ""
                            if 'lnveff' in file_name:
                                cap = f"Line vs Efficiency, {volts}, {current}"
                                caption_text = format_value_units(cap)
                            elif 'ldveff' in file_name:
                                cap = f"Efficiency vs. Load, {volts}, {current}"
                                caption_text = format_value_units(cap)
                            elif 'loadvripple' in file_name:
                                caption_text = "Load vs. Ripple"
                            elif 'linereg' in file_name:
                                cap = f"Full Load Line Regulation, {volts}, {current}"
                                caption_text = format_value_units(cap)
                            elif 'loadreg' in file_name:
                                cap = f"Load Regulation, {volts}, {current}"
                                caption_text = format_value_units(cap)
                            else:
                                caption_text = "No Load Input Power"
                                
                    caption_para = doc.add_paragraph()
                    add_caption_field(caption_para, caption_text, "Figure")
                    last_element.getparent().insert(last_element.getparent().index(last_element) + 1, caption_para._element)
                    last_element = caption_para._element

            # Add tables for all items
            for table_info in performance_data.get(item, {}).get('tables', []):
                table_data = table_info['data']
                merged_cells = table_info['merged_cells']
                custom_cap = table_info.get('custom_cap')
                
                rows = len(table_data)
                cols = max(len(row) for row in table_data) if rows > 0 else 1
                table = add_styled_table(doc, rows, cols, table_data, merged_cells, header_color='#0078AB', font_name='Calibri', font_size=9)
                last_element.getparent().insert(last_element.getparent().index(last_element) + 1, table._element)
                last_element = table._element
                
                caption_para = doc.add_paragraph()
                if custom_cap:
                    caption_text = custom_cap
                else:
                    caption_stripped = str(table_info['file_name']).strip(".xlsx")
                    cap = f"Data from {caption_stripped}"
                    caption_text = format_value_units(cap)
                    
                add_caption_field(caption_para, caption_text, "Table")
                last_element.getparent().insert(last_element.getparent().index(last_element) + 1, caption_para._element)
                last_element = caption_para._element

        return last_element