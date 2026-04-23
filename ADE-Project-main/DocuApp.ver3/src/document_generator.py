import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.shared import OxmlElement, qn
from PyQt5.QtCore import Qt
from performance_section import PerformanceSection
from waveform_section import WaveformSection
from utils import ensure_directory, remove_directory, get_default_base_folder, get_resource_path

def set_cell_background(cell, hex_color):
    """Applies a background hex color to a table cell."""
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)

def set_table_inner_borders(table, hex_color):
    """Applies inner borders to the table according to specifications."""
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    
    for border_name in ['insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), hex_color)
        tblBorders.append(border)
        
    tblPr.append(tblBorders)
    
def set_table_all_borders(table, hex_color):
    """Applies outer and inner borders to the table (used for PIXls Designer)."""
    tblPr = table._element.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4') 
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), hex_color)
        tblBorders.append(border)
    tblPr.append(tblBorders)

class DocGenerator:
    # Changed 'output_folder' to 'output_path' since the Save Dialog gives us the full file path
    def __init__(self, app, output_path, update_doc_path=""):
        self.app = app
        
        # 1. Check if the app has a selected template from the dropdown
        if hasattr(app, 'selected_template_path') and app.selected_template_path:
            self.template_path = app.selected_template_path
        else:
            # Fallback just in case something goes wrong
            self.template_path = get_resource_path('DER-template.docx')
            
        # 2. Set the exact save destination chosen by the user
        self.output_path = output_path
        
        self.update_doc_path = update_doc_path
        self.temp_dir = "temp_cropped_images"
        ensure_directory(self.temp_dir)
        
        self.performance = PerformanceSection(app, self.temp_dir)
        self.waveform = WaveformSection(app, self.temp_dir)

    def generate(self, progress_callback=None):
        # Choose between an update document or the chosen template
        doc_path = self.update_doc_path if self.update_doc_path and os.path.exists(self.update_doc_path) else self.template_path
        doc = Document(doc_path)
        
        # Get checked items from the UI
        performance_items = [self.app.performancedata_list.item(i).text() for i in range(self.app.performancedata_list.count()) if self.app.performancedata_list.item(i).checkState() == Qt.Checked]
        waveform_items = [self.app.waveforms_list.item(i).text() for i in range(self.app.waveforms_list.count()) if self.app.waveforms_list.item(i).checkState() == Qt.Checked]

        # Performance Section
        if performance_items:
            if progress_callback: progress_callback(70, "Writing Performance Data...")
            perf_data = self.performance.get_data(performance_items)
            self.performance.add_section(doc, doc.element.body[-1], performance_items, perf_data, None)

        # Waveform Section
        if waveform_items:
            if progress_callback: progress_callback(90, "Writing Waveforms...")
            wave_files = self.waveform.get_images_with_custom_crop(waveform_items)
            self.waveform.add_section(doc, doc.element.body[-1], waveform_items, wave_files)

        # --- BOM SECTION (Added after waveforms) ---
        if hasattr(self.app, 'bom_file_path') and self.app.bom_file_path:
            if progress_callback: progress_callback(95, "Appending Bill of Materials...")
            add_bom_table(doc, self.app.bom_file_path)

        # Save and open the document
        doc.save(self.output_path)
        os.startfile(self.output_path)
        remove_directory(self.temp_dir)

def add_bom_table(document, excel_path):
    """Reads the BOM from Excel and inserts a formatted table into the Word document."""
    try:
        df = pd.read_excel(excel_path, sheet_name='BOM', skiprows=2)
        bom_columns = [
            'Item', 'Quantity', 'Designator', 'Value', 
            'Description', 'Manufacturer Part Number', 'Manufacturer'
        ]
        
        # Handle variations/typos in the Manufacturer column names from raw files
        rename_map = {}
        for col in df.columns:
            if 'part number' in col.lower() and 'man' in col.lower():
                rename_map[col] = 'Manufacturer Part Number'
            elif 'man' in col.lower() and 'part' not in col.lower():
                rename_map[col] = 'Manufacturer'
                
        df = df.rename(columns=rename_map)
        
        # Filter only the columns that exist
        existing_cols = [col for col in bom_columns if col in df.columns]
        df = df[existing_cols].dropna(how='all')
        
        document.add_paragraph("Bill of Materials", style='Heading 2')
        table = document.add_table(rows=1, cols=len(df.columns))
        table.autofit = True
        
        set_table_inner_borders(table, 'C0C0C0')

        # Format Headers
        hdr_cells = table.rows[0].cells
        for i, col_name in enumerate(df.columns):
            cell = hdr_cells[i]
            cell.text = str(col_name)
            set_cell_background(cell, '0085CA')
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = paragraph.runs[0]
            run.font.name = 'Calibri'
            run.font.size = Pt(8)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Insert Rows
        for _, row in df.iterrows():
            row_cells = table.add_row().cells
            for i, value in enumerate(row):
                cell_val = "" if pd.isna(value) else str(value)
                cell = row_cells[i]
                cell.text = cell_val
                cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                paragraph = cell.paragraphs[0]
                if df.columns[i] == 'Description':
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                else:
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    
                if paragraph.runs:
                    run = paragraph.runs[0]
                    run.font.name = 'Calibri'
                    run.font.size = Pt(8)
                    run.font.bold = False
                    run.font.color.rgb = RGBColor(0, 0, 0)
    except Exception as e:
        print(f"Failed to load BOM from {excel_path}: {e}")
        
def add_pixls_designer_table(document, excel_path):
    """Reads the PIXls Designer spreadsheet and inserts a formatted table."""
    try:
        # Dynamically find the right sheet name (could be PIXls Designer, PIXI Designer, etc.)
        xl = pd.ExcelFile(excel_path)
        sheet_name = next((s for s in xl.sheet_names if 'PIX' in s or 'Design' in s), None)
        
        if not sheet_name:
            print("No PIXls Designer sheet found.")
            return

        # Find the actual header row by scanning for 'INPUT' and 'OUTPUT'
        df_raw = pd.read_excel(excel_path, sheet_name=sheet_name, header=None)
        header_idx = 0
        for i, row in df_raw.iterrows():
            row_str = [str(x).upper() for x in row.values]
            if 'INPUT' in row_str and 'OUTPUT' in row_str:
                header_idx = i
                break

        df = pd.read_excel(excel_path, sheet_name=sheet_name, skiprows=header_idx)
        
        # Isolate the core parameters (first 6 columns)
        df = df.iloc[:, :6]
        df.columns = ['Parameter', 'INPUT', 'INFO', 'OUTPUT', 'UNIT', 'Description']
        
        document.add_paragraph("Design Spreadsheet", style='Heading 2')
        table = document.add_table(rows=1, cols=len(df.columns))
        table.autofit = True
        
        # PIXls table uses standard all-around borders
        set_table_all_borders(table, 'C0C0C0')

        # Format Headers
        hdr_cells = table.rows[0].cells
        for i, col_name in enumerate(df.columns):
            cell = hdr_cells[i]
            cell.text = str(col_name)
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            
            paragraph = cell.paragraphs[0]
            paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            if paragraph.runs:
                run = paragraph.runs[0]
                run.font.name = 'Calibri'
                run.font.size = Pt(8)
                run.font.bold = True

        # Insert Rows
        for _, row in df.iterrows():
            # Skip completely empty rows
            if pd.isna(row['Parameter']) and pd.isna(row['INPUT']):
                continue

            row_cells = table.add_row().cells
            
            # Detect Sub-Headers (Parameter exists, but Input/Output/Unit are completely blank)
            is_subheader = pd.isna(row['INPUT']) and pd.isna(row['OUTPUT']) and pd.isna(row['UNIT'])

            if is_subheader:
                # Merge the whole row for the sub-header
                a, b = row_cells[0], row_cells[-1]
                a.merge(b)
                a.text = str(row['Parameter'])
                set_cell_background(a, 'D9D9D9') # #D9D9D9 Light Gray background
                a.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                
                paragraph = a.paragraphs[0]
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                if paragraph.runs:
                    run = paragraph.runs[0]
                    run.font.name = 'Calibri'
                    run.font.size = Pt(8)
                    run.font.bold = True
                    run.font.color.rgb = RGBColor(255, 255, 255) # White font as requested
            else:
                # Regular data rows
                for i, col in enumerate(df.columns):
                    val = "" if pd.isna(row[col]) else str(row[col])
                    cell = row_cells[i]
                    cell.text = val
                    cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    
                    paragraph = cell.paragraphs[0]
                    # Left align Parameter and Description, Center the rest
                    if col in ['Parameter', 'Description']:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    else:
                        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        
                    if paragraph.runs:
                        run = paragraph.runs[0]
                        run.font.name = 'Calibri'
                        run.font.size = Pt(8)
                        run.font.bold = False

    except Exception as e:
        print(f"Failed to load PIXls Designer from {excel_path}: {e}")