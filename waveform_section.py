import os
import re
from PyQt5.QtCore import Qt
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph
from utils import log_message
from word_utils import add_caption_field, format_value_units
from image_utils import crop_and_save

class WaveformSection:
    def __init__(self, app, temp_dir):
        self.app = app
        self.temp_dir = temp_dir
        self.waveforms_start_idx = 0

    def _strip_chapter_numbering(self, text):
        """
        Safely strips section numbers like "7.1 ", "1.2.3 ", "1. ", or "1 - " 
        so Word's auto-numbering doesn't double them up. 
        It safely ignores technical values like "85 VAC", "1.5V", or "100%".
        """
        pattern = r'^(?:\d+\.(?:\d+\.?)*\s+|\d+(?:\.\d+)*\s*[-_]+\s+)(?![vV][aA]?[cCdD]?\b|[aA]\b|[mM][aA]\b|[wW]\b|[hH][zZ]\b|%)'
        return re.sub(pattern, '', text).strip()

    def add_section(self, doc, last_element):
        self.waveforms_start_idx, main_wave_anchor = self._get_main_waveforms_anchor(doc)
        
        if main_wave_anchor:
            last_element = main_wave_anchor._element
            self._wipe_template_waveforms_section(doc, last_element)
        else:
            new_para = doc.add_paragraph("Waveforms", style='Heading 1')
            last_element.getparent().insert(last_element.getparent().index(last_element) + 1, new_para._element)
            last_element = new_para._element
            self.waveforms_start_idx = len(doc.paragraphs) - 1

        root = self.app.waveform_tree.invisibleRootItem()
        return self._process_node(root, doc, last_element, current_heading_level=2)

    def _get_main_waveforms_anchor(self, doc):
        """Finds the main 'Waveforms' chapter heading."""
        for idx, p in enumerate(doc.paragraphs):
            if p.style and p.style.name.startswith('Heading'):
                clean_p_text = re.sub(r'^\d+(\.\d+)*\s*[-\._]*\s*', '', p.text.lower()).strip()
                if "waveform" in clean_p_text:
                    return idx, p
        return 0, None

    def _wipe_template_waveforms_section(self, doc, start_element):
        """
        Deletes all paragraphs and tables occurring after the 'Waveforms' heading
        until it hits the next Chapter (Heading 1) or the end of the document.
        """
        parent = start_element.getparent()
        found_start = False
        elements_to_remove = []
        
        for child in parent:
            if not found_start:
                if child == start_element:
                    found_start = True
                continue
            
            # If it's a paragraph, check if it's the start of the next chapter
            if child.tag.endswith('p'):
                p = Paragraph(child, doc)
                if p.style and p.style.name == 'Heading 1':
                    break # Stop wiping, we reached the next section!
            
            elements_to_remove.append(child)
            
        for el in elements_to_remove:
            parent.remove(el)
            
        log_message(f"Wiped {len(elements_to_remove)} legacy elements from the Waveforms template section.")

    def _process_node(self, node, doc, last_element, current_heading_level):
        is_root = (node == self.app.waveform_tree.invisibleRootItem())

        if not is_root and node.checkState(0) == Qt.Unchecked:
            return last_element

        current_anchor = last_element

        # 1. PROCESS FOLDER HEADINGS
        if not is_root and node.data(0, Qt.UserRole + 2) == "folder":
            clean_name = node.text(0).replace(" CROPPED", "")
            
            # Strip physical numbering (7.1, 7.2) so Word's auto-numbering doesn't double it
            clean_name = self._strip_chapter_numbering(clean_name)
            
            level = min(current_heading_level, 9)
            
            # Write JUST the clean text. Word's native style adds the "7.1" automatically.
            new_para = doc.add_paragraph(clean_name, style=f'Heading {level}')
            
            # Apply slight visual styling for root headings if needed
            if level == 2 and new_para.runs:
                new_para.runs[0].font.size = Pt(14)
            elif level == 3 and new_para.runs:
                new_para.runs[0].font.size = Pt(12)
            
            current_anchor.getparent().insert(current_anchor.getparent().index(current_anchor) + 1, new_para._element)
            current_anchor = new_para._element 
            current_heading_level += 1

        # 2. GATHER FILES IN CURRENT FOLDER
        files_at_this_level = []
        for i in range(node.childCount()):
            child = node.child(i)
            if child.checkState(0) != Qt.Unchecked and child.data(0, Qt.UserRole + 2) == "file":
                files_at_this_level.append(child)

        # 3. RENDER FILES INTO TABLE
        if files_at_this_level:
            current_anchor = self._render_image_table(files_at_this_level, doc, current_anchor)

        # 4. RECURSE DEEPER INTO SUBFOLDERS
        for i in range(node.childCount()):
            child = node.child(i)
            if child.checkState(0) != Qt.Unchecked and child.data(0, Qt.UserRole + 2) == "folder":
                current_anchor = self._process_node(child, doc, current_anchor, current_heading_level)

        return current_anchor

    def _get_crop_for_node(self, node):
        current = node
        while current:
            data = current.data(0, Qt.UserRole + 3)
            if data: return data
            current = current.parent()
        return {'left': '0', 'top': '0', 'right': '0', 'bottom': '0'}

    def _render_image_table(self, file_nodes, doc, last_element):
        table = doc.add_table(rows=1, cols=2)
        table.autofit = False
        table.columns[0].width = Inches(3.5)
        table.columns[1].width = Inches(3.5)
        table.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        current_row = 0
        current_col = 0
        
        for node in file_nodes:
            original_path = node.data(0, Qt.UserRole + 1)
            
            crop_settings = self._get_crop_for_node(node)
            c_left = int(crop_settings.get('left', 0))
            c_top = int(crop_settings.get('top', 0))
            c_right = int(crop_settings.get('right', 0))
            c_bottom = int(crop_settings.get('bottom', 0))
            
            cropped_path = crop_and_save(original_path, c_left, c_top, c_right, c_bottom, self.temp_dir)
            if not cropped_path: continue

            if current_col >= 2:
                current_col = 0
                current_row += 1
                table.add_row()
            
            cell = table.cell(current_row, current_col)
            
            cell_para = cell.paragraphs[0]
            run = cell_para.add_run()
            run.add_picture(cropped_path, width=Inches(3.4))
            
            # LINK THE UI RENAME TO THE WORD CAPTION
            clean_base_name = node.text(0).replace(" CROPPED", "")
            
            # Strip physical numbering so image captions aren't prefixed with "7.2.1 "
            clean_base_name = self._strip_chapter_numbering(clean_base_name)
            
            main_cap_text = format_value_units(clean_base_name)

            caption_cell = cell.add_paragraph()
            add_caption_field(caption_cell, main_cap_text, "Figure")
            caption_cell.alignment = WD_ALIGN_PARAGRAPH.LEFT

            current_col += 1
        
        last_element.getparent().insert(last_element.getparent().index(last_element) + 1, table._element)
        return table._element