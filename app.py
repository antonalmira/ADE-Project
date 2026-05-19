import os
import sys
import shutil
import json
from PyQt5 import QtWidgets, uic, QtCore, QtGui
from PyQt5.QtCore import Qt, QPoint
import resource_rc

from handlers import (
    select_bom_file, select_pixls_file, select_performance_file, 
    add_waveform_folder, clear_waveform_folders, clear_performance_folders
)
from document_handler import generate_document
from utils import get_resource_path
from preview import show_file_preview, crop_and_update_preview
from list_updater import (
    update_performance_tree, update_waveform_tree,
    capture_wave_state, capture_perf_state,
    sync_state_on_rename, sync_metadata_on_rename, sync_metadata_on_delete
)

class DocuApp(QtWidgets.QMainWindow):
    def __init__(self):
        super(DocuApp, self).__init__()
        uic.loadUi(get_resource_path('DocuApp_ver6.ui'), self)
        
        icon_path = get_resource_path(os.path.join('resources', 'icons', 'tardis_icon.ico'))
        self.setWindowIcon(QtGui.QIcon(icon_path))
        
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)
        self.old_pos = None

        for line_edit in [self.upper_input, self.lower_input, self.left_input, self.right_input]:
            line_edit.setText("0")

        self.tab_group = QtWidgets.QButtonGroup(self)
        self.tab_group.addButton(self.btn_tab_perf, 0)
        self.tab_group.addButton(self.btn_tab_wave, 1)
        self.btn_tab_perf.clicked.connect(lambda: self.switch_tab(0))
        self.btn_tab_wave.clicked.connect(lambda: self.switch_tab(1))

        self.performancedata_sel.clicked.connect(lambda: select_performance_file(self))
        self.waveforms_add_folder.clicked.connect(lambda: add_waveform_folder(self))
        
        self.refresh_button_perf.clicked.connect(lambda: update_performance_tree(self))
        self.clear_perf_button.clicked.connect(lambda: clear_performance_folders(self))
        self.refresh_button_wave.clicked.connect(lambda: update_waveform_tree(self))
        self.waveforms_clear_folders.clicked.connect(lambda: clear_waveform_folders(self))

        self.exit_button.clicked.connect(self.close)
        self.minimize_button.clicked.connect(self.showMinimized)
        self.maximize_button.clicked.connect(self.toggle_maximize)
        
        self.crop_button.clicked.connect(self.save_crop_to_selected)
        self.generate_document_button.clicked.connect(lambda: generate_document(self))

        # --- UNIFIED TREE SETUP ---
        for tree in [self.performance_tree, self.waveform_tree]:
            tree.itemSelectionChanged.connect(lambda: show_file_preview(self))
            tree.setDragDropMode(QtWidgets.QAbstractItemView.InternalMove)
            tree.setDragEnabled(True)
            tree.setAcceptDrops(True)
            tree.setDropIndicatorShown(True)
            tree.setSelectionMode(QtWidgets.QAbstractItemView.ExtendedSelection)
            tree.setEditTriggers(QtWidgets.QAbstractItemView.DoubleClicked | QtWidgets.QAbstractItemView.EditKeyPressed)
            
            # Allow smooth scrolling when dragging an item near the top/bottom edges
            tree.setAutoScroll(True)
            tree.setAutoScrollMargin(35)
            tree.setVerticalScrollMode(QtWidgets.QAbstractItemView.ScrollPerPixel)
            
            tree.itemChanged.connect(self.on_tree_item_changed)
            tree.setContextMenuPolicy(Qt.CustomContextMenu)
            tree.customContextMenuRequested.connect(self.show_context_menu)

        # Override Drag and Drop 
        self.original_waveform_drop_event = self.waveform_tree.dropEvent
        self.waveform_tree.dropEvent = lambda event: self.custom_drop_event(event, self.waveform_tree, self.original_waveform_drop_event)
        
        self.original_performance_drop_event = self.performance_tree.dropEvent
        self.performance_tree.dropEvent = lambda event: self.custom_drop_event(event, self.performance_tree, self.original_performance_drop_event)

        if hasattr(self, 'select_bom_button'):
            self.select_bom_button.clicked.connect(lambda: select_bom_file(self))
            self.select_pixls_button.clicked.connect(lambda: select_pixls_file(self))

        self.waveforms_path.editingFinished.connect(lambda: update_waveform_tree(self))
        self.populate_templates_dropdown()

    # --- PHYSICAL DRAG AND DROP (RELOCATION) ---
    def custom_drop_event(self, event, tree_widget, original_event_call):
        dragged_items = tree_widget.selectedItems()
        if not dragged_items:
            return original_event_call(event)

        target_item = tree_widget.itemAt(event.pos())
        indicator = tree_widget.dropIndicatorPosition()
        
        target_dir = None
        
        if target_item:
            is_folder = target_item.data(0, Qt.UserRole + 2) == "folder"
            if indicator == QtWidgets.QAbstractItemView.OnItem:
                if is_folder:
                    target_dir = target_item.data(0, Qt.UserRole + 1)
                else:
                    parent = target_item.parent()
                    if parent: target_dir = parent.data(0, Qt.UserRole + 1)
            elif indicator in [QtWidgets.QAbstractItemView.AboveItem, QtWidgets.QAbstractItemView.BelowItem]:
                parent = target_item.parent()
                if parent:
                    target_dir = parent.data(0, Qt.UserRole + 1)
        
        if not target_dir:
            if tree_widget == self.waveform_tree:
                valid_paths = [p.strip() for p in self.waveforms_path.text().split(';') if p.strip() and os.path.isdir(p.strip())]
                if valid_paths: target_dir = valid_paths[0]
            else:
                target_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")

        if not target_dir or not os.path.isdir(target_dir):
            return original_event_call(event)

        tree_type = "wave" if tree_widget == self.waveform_tree else "perf"
        if tree_type == "wave": capture_wave_state(self)
        else: capture_perf_state(self)

        physically_moved = False
        
        for item in dragged_items:
            old_path = item.data(0, Qt.UserRole + 1)
            if not old_path or not os.path.exists(old_path): continue
            
            if os.path.dirname(old_path) != target_dir:
                new_path = os.path.join(target_dir, os.path.basename(old_path))
                if os.path.isdir(old_path) and new_path.startswith(old_path):
                    continue 

                try:
                    old_name = os.path.basename(old_path)
                    new_name = os.path.basename(new_path)
                    parent_path = os.path.dirname(old_path)
                    
                    shutil.move(old_path, new_path)
                    sync_state_on_rename(tree_type, old_path, new_path, old_name, new_name, parent_path)
                    if tree_type == "perf":
                        sync_metadata_on_rename(old_path, new_path)
                    physically_moved = True
                except Exception as e:
                    print(f"Error moving {old_path}: {e}")

        original_event_call(event)
        
        if physically_moved:
            if tree_widget == self.waveform_tree: QtCore.QTimer.singleShot(50, lambda: update_waveform_tree(self, capture=False))
            else: QtCore.QTimer.singleShot(50, lambda: update_performance_tree(self, capture=False))

    # --- PHYSICAL RENAME ---
    def on_tree_item_changed(self, item, column):
        if column != 0: return
        tree_widget = item.treeWidget()
        old_path = item.data(0, Qt.UserRole + 1)
        if not old_path or not os.path.exists(old_path): return
        
        original_ui_name = item.data(0, Qt.UserRole + 6)
        
        # Safely strip any crop tag
        new_display_name = item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip()
        
        if new_display_name == original_ui_name: return

        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        old_display_name = original_ui_name
        forbidden = ['<', '>', ':', '"', '/', '\\', '|', '?', '*']
        if any(c in new_display_name for c in forbidden):
            QtWidgets.QMessageBox.warning(self, "Invalid Name", "File name contains invalid characters.")
            self.revert_item_text(tree_widget, item, old_display_name)
            return
        
        ext = os.path.splitext(old_path)[1] if not is_folder else ""
        new_filename = new_display_name + ext
        new_path = os.path.join(os.path.dirname(old_path), new_filename)
        
        if os.path.exists(new_path) and new_path.lower() != old_path.lower():
            QtWidgets.QMessageBox.warning(self, "Exists", "A file or folder with this name already exists.")
            self.revert_item_text(tree_widget, item, old_display_name)
            return

        try:
            tree_type = "wave" if tree_widget == self.waveform_tree else "perf"
            if tree_type == "wave": capture_wave_state(self)
            else: capture_perf_state(self)
            
            old_filename = os.path.basename(old_path)
            parent_path = os.path.dirname(old_path)
            
            os.rename(old_path, new_path)
            
            sync_state_on_rename(tree_type, old_path, new_path, old_filename, new_filename, parent_path)
            if tree_type == "perf":
                sync_metadata_on_rename(old_path, new_path)
                
            item.setData(0, Qt.UserRole + 1, new_path)
            item.setData(0, Qt.UserRole + 4, new_filename)
            item.setData(0, Qt.UserRole + 6, new_display_name) 
            
            if is_folder:
                if tree_type == "wave": QtCore.QTimer.singleShot(50, lambda: update_waveform_tree(self, capture=False))
                else: QtCore.QTimer.singleShot(50, lambda: update_performance_tree(self, capture=False))
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "Rename Error", f"Could not rename:\n{str(e)}")
            self.revert_item_text(tree_widget, item, old_display_name)
            
    def revert_item_text(self, tree_widget, item, old_display_name):
        revert_text = old_display_name
        if item.data(0, Qt.UserRole + 3): 
            is_folder = item.data(0, Qt.UserRole + 2) == "folder"
            revert_text += " [FOLDER CROPPED]" if is_folder else " [IMAGE CROPPED]"
            
        tree_widget.blockSignals(True)
        item.setText(0, revert_text)
        tree_widget.blockSignals(False)

    # --- CONTEXT MENU LOGIC ---
    def show_context_menu(self, position):
        tree_widget = self.sender()
        item = tree_widget.itemAt(position)
        if not item: return
        
        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        item_path = item.data(0, Qt.UserRole + 1)

        menu = QtWidgets.QMenu()
        menu.setStyleSheet("""
            QMenu { background-color: #202025; color: white; border: 1px solid #3a3a40; }
            QMenu::item { padding: 5px 25px; }
            QMenu::item:selected { background-color: #0085ca; }
            QMenu::separator { background-color: #3a3a40; height: 1px; margin: 4px 0px; }
        """)
        
        add_folder_action = None
        add_file_action = None
        toggle_setup_action = None
        
        if is_folder:
            if tree_widget == self.performance_tree:
                include_setup = item.data(0, Qt.UserRole + 11)
                if include_setup is None: include_setup = False
                
                selected_folders = [i for i in tree_widget.selectedItems() if i.data(0, Qt.UserRole + 2) == "folder"]
                suffix = " (Selected)" if len(selected_folders) > 1 and item in selected_folders else ""
                
                toggle_setup_action = menu.addAction(f"Remove Test Set-up Table{suffix}" if include_setup else f"Add Test Set-up Table{suffix}")
                menu.addSeparator()

            add_folder_action = menu.addAction("Add New Folder")
            add_file_action = menu.addAction("Add New File(s)")
            menu.addSeparator()
            
        rename_action = menu.addAction("Rename")
        delete_action = menu.addAction("Delete")
        
        action = menu.exec_(tree_widget.viewport().mapToGlobal(position))
        if not action: return
        
        if action == add_folder_action: self.create_physical_folder(tree_widget, item_path)
        elif action == add_file_action: self.add_physical_files(tree_widget, item_path)
        elif action == rename_action: tree_widget.editItem(item, 0)
        elif action == delete_action: self.delete_physical_item(tree_widget, item)
        elif action == toggle_setup_action: self.toggle_test_setup(tree_widget, item)

    def toggle_test_setup(self, tree_widget, clicked_item):
        current = clicked_item.data(0, Qt.UserRole + 11)
        new_val = False if current in [True] else True
        
        selected_items = tree_widget.selectedItems()
        if clicked_item not in selected_items:
            targets = [clicked_item]
        else:
            targets = [i for i in selected_items if i.data(0, Qt.UserRole + 2) == "folder"]
            
        perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
        meta_path = os.path.join(perf_dir, "metadata.json")
        
        meta = {}
        if os.path.exists(meta_path):
            try:
                with open(meta_path, 'r') as f: meta = json.load(f)
            except: pass
            
        for t_item in targets:
            t_item.setData(0, Qt.UserRole + 11, new_val)
            folder_path = t_item.data(0, Qt.UserRole + 1)
            if folder_path not in meta: meta[folder_path] = {}
            meta[folder_path]["include_setup_table"] = new_val
            
        with open(meta_path, 'w') as f: json.dump(meta, f)

    def delete_physical_item(self, tree_widget, item):
        selected_items = tree_widget.selectedItems()
        if not selected_items or item not in selected_items:
            selected_items = [item]

        msg = f"Are you sure you want to permanently delete {len(selected_items)} item(s)?\n\nThis cannot be undone!"
        reply = QtWidgets.QMessageBox.question(self, "Confirm Delete", msg, QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        
        if reply == QtWidgets.QMessageBox.Yes:
            paths_to_delete = []
            for sel_item in selected_items:
                path = sel_item.data(0, Qt.UserRole + 1)
                if path and os.path.exists(path):
                    paths_to_delete.append(path)
                    is_folder = sel_item.data(0, Qt.UserRole + 2) == "folder"
                    try:
                        if is_folder: shutil.rmtree(path)
                        else: os.remove(path)
                    except Exception as e:
                        print(f"Error deleting {path}: {e}")

            tree_type = "wave" if tree_widget == self.waveform_tree else "perf"
            if tree_type == "perf" and paths_to_delete:
                sync_metadata_on_delete(paths_to_delete)
                
            if tree_type == "wave": update_waveform_tree(self)
            else: update_performance_tree(self)

    def create_physical_folder(self, tree_widget, parent_path):
        if not parent_path or not os.path.exists(parent_path): return
        new_path = os.path.join(parent_path, "New Folder")
        counter = 1
        while os.path.exists(new_path):
            new_path = os.path.join(parent_path, f"New Folder ({counter})")
            counter += 1
        try:
            os.makedirs(new_path)
            if tree_widget == self.waveform_tree: update_waveform_tree(self)
            else: update_performance_tree(self)
        except Exception as e: print(f"Error creating folder: {e}")

    def add_physical_files(self, tree_widget, parent_path):
        if not parent_path or not os.path.exists(parent_path): return
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(self, "Select Images", "", "Images (*.png *.jpg *.jpeg *.bmp)")
        if not files: return
        for f in files:
            try: shutil.copy(f, parent_path)
            except Exception as e: print(f"Error copying {f}: {e}")
        if tree_widget == self.waveform_tree: update_waveform_tree(self)
        else: update_performance_tree(self)

    def switch_tab(self, index):
        if self.stackedWidget.currentIndex() == index: return
        current_widget = self.stackedWidget.currentWidget()
        next_widget = self.stackedWidget.widget(index)
        width = self.stackedWidget.width()
        height = self.stackedWidget.height()
        offset_x = width if self.stackedWidget.currentIndex() < index else -width

        next_widget.setGeometry(0, 0, width, height)
        next_widget.move(offset_x, 0)
        next_widget.show()
        next_widget.raise_()

        self.anim_group = QtCore.QParallelAnimationGroup()
        anim_out = QtCore.QPropertyAnimation(current_widget, b"pos")
        anim_out.setDuration(300); anim_out.setEasingCurve(QtCore.QEasingCurve.InOutQuart)
        anim_out.setStartValue(QtCore.QPoint(0, 0)); anim_out.setEndValue(QtCore.QPoint(-offset_x, 0))
        anim_in = QtCore.QPropertyAnimation(next_widget, b"pos")
        anim_in.setDuration(300); anim_in.setEasingCurve(QtCore.QEasingCurve.InOutQuart)
        anim_in.setStartValue(QtCore.QPoint(offset_x, 0)); anim_in.setEndValue(QtCore.QPoint(0, 0))
        self.anim_group.addAnimation(anim_out); self.anim_group.addAnimation(anim_in)
        self.anim_group.finished.connect(lambda: self.stackedWidget.setCurrentIndex(index))
        self.anim_group.start()

    def save_crop_to_selected(self):
        wave_items = self.waveform_tree.selectedItems()
        perf_items = self.performance_tree.selectedItems()
        crop_data = {
            'left': self.left_input.text() or '0',
            'top': self.upper_input.text() or '0',
            'right': self.right_input.text() or '0',
            'bottom': self.lower_input.text() or '0'
        }
        items = wave_items + perf_items
        if items:
            for item in items:
                item.setData(0, Qt.UserRole + 3, crop_data) 
                
                # Apply the correct tag based on type
                base_text = item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip()
                is_folder = item.data(0, Qt.UserRole + 2) == "folder"
                tag = " [FOLDER CROPPED]" if is_folder else " [IMAGE CROPPED]"
                item.setText(0, f"{base_text}{tag}")
                
            crop_and_update_preview(self)
        else:
            QtWidgets.QMessageBox.warning(self, "Selection", "Select a folder or file to apply crop.")

    def populate_templates_dropdown(self):
        if getattr(sys, 'frozen', False): base_path = sys._MEIPASS
        else: base_path = os.path.abspath(os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
        templates_folder = os.path.join(base_path, "templates")
        self.template_dropdown.clear()
        if os.path.exists(templates_folder):
            template_files = [f for f in os.listdir(templates_folder) if f.endswith('.docx')]
            if template_files: self.template_dropdown.addItems(template_files)
            else: self.template_dropdown.addItem("No templates found"); self.template_dropdown.setEnabled(False)
        else:
            self.template_dropdown.addItem("Templates folder missing!"); self.template_dropdown.setEnabled(False)

    def toggle_maximize(self):
        if self.isMaximized(): self.showNormal()
        else: self.showMaximized()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.headerr.underMouse(): self.old_pos = event.globalPos()

    def mouseMoveEvent(self, event):
        if self.old_pos:
            delta = QPoint(event.globalPos() - self.old_pos)
            self.move(self.x() + delta.x(), self.y() + delta.y())
            self.old_pos = event.globalPos()

    def mouseReleaseEvent(self, event):
        self.old_pos = None