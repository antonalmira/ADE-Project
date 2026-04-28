import os
import re
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.shared import OxmlElement, qn
from PyQt5.QtCore import Qt
from performance_section import PerformanceSection
from waveform_section import WaveformSection
from utils import ensure_directory, remove_directory

def set_cell_background(cell, hex_color):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_table_inner_borders(table, hex_color):
    """BOM: No Outer Border, Inner Border Only. 1/2 pt (sz=4)."""
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:color'), hex_color)
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
def set_table_all_borders(table, hex_color):
    """PIXls: All Borders. 1/2 pt (sz=4)."""
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:color'), hex_color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

def format_text_specs(text):
    """Applies the 'Other' specifications: spacing and capitalization."""
    if not isinstance(text, str):
        return text
    # Space between numeric value and symbol/unit (e.g., 10uF -> 10 uF)
    text = re.sub(r'(?<=\d)(?=[a-zA-ZµΩ°])', ' ', text)
    # Capitalize specific terms
    for term in ['vac', 'vdc', 'vor', 'kp']:
        # Use regex boundary \b to only replace the exact word, ignoring case
        text = re.sub(fr'\b{term}\b', term.upper(), text, flags=re.IGNORECASE)
    return text

def apply_column_widths(table, width_inches_list):
    """Forces exact column widths across the entire table."""
    for row in table.rows:
        for idx, width in enumerate(width_inches_list):
            if idx < len(row.cells):
                row.cells[idx].width = Inches(width)

class DocGenerator:
    def __init__(self, app, output_path, update_doc_path=""):
        self.app = app
        self.template_path = getattr(app, 'selected_template_path', '')
        self.output_path = output_path
        self.update_doc_path = update_doc_path
        self.temp_dir = "temp_cropped_images"
        ensure_directory(self.temp_dir)
        
        self.performance = PerformanceSection(app, self.temp_dir)
        self.waveform = WaveformSection(app, self.temp_dir)

    def generate(self, progress_callback=None):
        doc_path = self.update_doc_path if self.update_doc_path and os.path.exists(self.update_doc_path) else self.template_path
        doc = Document(doc_path)
        
        perf_checked = [self.app.performancedata_list.item(i).text() for i in range(self.app.performancedata_list.count()) if self.app.performancedata_list.item(i).checkState() == Qt.Checked]
        wave_checked = [self.app.waveforms_list.item(i).text() for i in range(self.app.waveforms_list.count()) if self.app.waveforms_list.item(i).checkState() == Qt.Checked]

        if perf_checked:
            if progress_callback: progress_callback(70, "Writing Performance Data...")
            perf_data = self.performance.get_data(perf_checked)
            self.performance.add_section(doc, doc.element.body[-1], perf_checked, perf_data, None)

        if wave_checked:
            if progress_callback: progress_callback(85, "Writing Waveforms...")
            wave_files = self.waveform.get_images_with_custom_crop(wave_checked)
            self.waveform.add_section(doc, doc.element.body[-1], wave_checked, wave_files)

        if hasattr(self.app, 'bom_file_path') and self.app.bom_file_path:
            if progress_callback: progress_callback(95, "Appending Bill of Materials and PIXls...")
            add_pixls_designer_table(doc, self.app.bom_file_path)
            add_bom_table(doc, self.app.bom_file_path)

        doc.save(self.output_path)
        os.startfile(self.output_path)
        remove_directory(self.temp_dir)

def add_bom_table(document, excel_path):
    try:
        df = pd.read_excel(excel_path, sheet_name='BOM', skiprows=2)
        bom_columns = ['Item', 'Quantity', 'Designator', 'Value', 'Description', 'Manufacturer Part Number', 'Manufacturer']
        
        rename_map = {}
        for col in df.columns:
            if 'part number' in str(col).lower() and 'man' in str(col).lower(): rename_map[col] = 'Manufacturer Part Number'
            elif 'man' in str(col).lower() and 'part' not in str(col).lower(): rename_map[col] = 'Manufacturer'
        df = df.rename(columns=rename_map)
        
        existing_cols = [col for col in bom_columns if col in df.columns]
        df = df[existing_cols].dropna(how='all')
        
        document.add_paragraph("Bill of Materials", style='Heading 2')
        table = document.add_table(rows=1, cols=len(df.columns))
        set_table_inner_borders(table, 'C0C0C0')

        # Format Headers
        for i, col_name in enumerate(df.columns):
            cell = table.rows[0].cells[i]
            cell.text = str(col_name)
            set_cell_background(cell, '0085CA') # POWI Blue
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.name = 'Calibri'
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Insert Rows
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                cell_val = "" if pd.isna(value) else str(value)
                col_name = df.columns[i]
                
                # Apply specific formatting rules
                if col_name == 'Designator':
                    # Designators are line-break-separated instead of comma-separated
                    cell_val = cell_val.replace(',', '\n').replace(' ', '')
                else:
                    cell_val = format_text_specs(cell_val)

                cell = row_cells[i]
                cell.text = cell_val
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                p = cell.paragraphs[0]
                # Description left aligned, rest centered
                if col_name == 'Description':
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(8)
                    run.font.bold = False
                    run.font.color.rgb = RGBColor(0, 0, 0)

        # Apply exact Column Widths (Scaled down by 10 to fit 6.5" page)
        # Spec ratios: 4", 6", 7.2", 21.6", 16.5", 10.1". (Added 6" for 'Value' column)
        apply_column_widths(table, [0.4, 0.6, 0.72, 0.6, 2.16, 1.65, 1.01])

    except Exception as e:
        print(f"BOM Error: {e}")

def add_pixls_designer_table(document, excel_path):
    try:
        xl = pd.ExcelFile(excel_path)
        sheet_name = next((s for s in xl.sheet_names if 'PIX' in s or 'Design' in s), None)
        if not sheet_name: return

        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        header_idx = 0
        for i, row in df_raw.iterrows():
            row_str = [str(x).upper() for x in row.values]
            if 'INPUT' in row_str and 'OUTPUT' in row_str:
                header_idx = i
                break

        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=header_idx).iloc[:, :6]
        df.columns = ['Parameter', 'INPUT', 'INFO', 'OUTPUT', 'UNIT', 'Description']
        
        document.add_paragraph("Design Spreadsheet", style='Heading 2')
        table = document.add_table(rows=1, cols=len(df.columns))
        set_table_all_borders(table, 'C0C0C0') # POWI Gray

        # Format Headers
        for i, col_name in enumerate(df.columns):
            cell = table.rows[0].cells[i]
            cell.text = str(col_name)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.runs[0]
            run.font.name = 'Calibri'
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)

        # Insert Rows
        for _, row in df.iterrows():
            if pd.isna(row['Parameter']) and pd.isna(row['INPUT']): continue
            row_cells = table.add_row().cells
            is_subheader = pd.isna(row['INPUT']) and pd.isna(row['OUTPUT']) and pd.isna(row['UNIT'])

            if is_subheader:
                cell = row_cells[0]
                cell.merge(row_cells[-1])
                cell.text = format_text_specs(str(row['Parameter']))
                set_cell_background(cell, 'D9D9D9') # Light Gray
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in p.runs:
                    run.font.name = 'Calibri'
                    run.font.size = Pt(8)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255) # White Font
            else:
                for i, col_name in enumerate(df.columns):
                    val = "" if pd.isna(row[col_name]) else format_text_specs(str(row[col_name]))
                    cell = row_cells[i]
                    cell.text = val
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
                    p = cell.paragraphs[0]
                    if col_name in ['Parameter', 'Description']:
                        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                    for run in p.runs:
                        run.font.name = 'Calibri'
                        run.font.size = Pt(8)
                        run.font.bold = False
                        run.font.color.rgb = RGBColor(0, 0, 0)

        # Apply exact Column Widths (Scaled down by 10 to fit 6.5" page)
        # Spec ratios: 13.7", 7", 7", 7", 5", 25.2"
        apply_column_widths(table, [1.37, 0.7, 0.7, 0.7, 0.5, 2.52])

    except Exception as e:
        print(f"PIXls Error: {e}")