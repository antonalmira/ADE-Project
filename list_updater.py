from PyQt5.QtCore import Qt
from PyQt5.QtGui import QBrush, QColor, QFont
from PyQt5.QtWidgets import QListWidgetItem, QTreeWidgetItem
import os
import re
import json

global_perf_state = {}

def capture_perf_state(app):
    global global_perf_state
    def process_item(item):
        item_path = item.data(0, Qt.UserRole + 1)
        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        if item_path:
            global_perf_state[item_path] = {
                "check_state": item.checkState(0), 
                "crop": item.data(0, Qt.UserRole + 3), 
                "caption": item.data(0, Qt.UserRole),
                "custom_name": item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip(),
                "expanded": item.isExpanded()
            }
            if is_folder:
                order = [item.child(i).data(0, Qt.UserRole + 4) for i in range(item.childCount()) if item.child(i).data(0, Qt.UserRole + 4)]
                global_perf_state[item_path]["order"] = order
                for i in range(item.childCount()): process_item(item.child(i))

    root_paths = {os.path.dirname(app.performance_tree.topLevelItem(i).data(0, Qt.UserRole + 1)) 
                  for i in range(app.performance_tree.topLevelItemCount()) if app.performance_tree.topLevelItem(i).data(0, Qt.UserRole + 1)}
            
    for rp in root_paths:
        if rp not in global_perf_state: global_perf_state[rp] = {}
        global_perf_state[rp]["order"] = [app.performance_tree.topLevelItem(i).data(0, Qt.UserRole + 4) 
                                          for i in range(app.performance_tree.topLevelItemCount()) 
                                          if app.performance_tree.topLevelItem(i).data(0, Qt.UserRole + 1) and os.path.dirname(app.performance_tree.topLevelItem(i).data(0, Qt.UserRole + 1)) == rp]
    
    for i in range(app.performance_tree.topLevelItemCount()): process_item(app.performance_tree.topLevelItem(i))

def update_performance_tree(app, capture=True):
    if capture: capture_perf_state(app)
    
    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    if not os.path.exists(perf_dir): return
    
    meta_path = os.path.join(perf_dir, "metadata.json")
    metadata = {}
    if os.path.exists(meta_path):
        with open(meta_path, 'r') as f: metadata = json.load(f)

    app.performance_tree.blockSignals(True)
    app.performance_tree.clear()
    
    build_perf_tree(app.performance_tree, perf_dir, metadata)
    app.performance_tree.blockSignals(False)

def build_perf_tree(parent_widget, current_path, metadata):
    try: items = os.listdir(current_path)
    except PermissionError: return False

    saved_order = global_perf_state.get(current_path, {}).get("order", [])
    items.sort(key=lambda x: saved_order.index(x) if x in saved_order else 999999)

    for item_name in items:
        if item_name == "metadata.json": continue
        
        item_path = os.path.join(current_path, item_name)
        item_state = global_perf_state.get(item_path, {}) 
        is_dir = os.path.isdir(item_path)
        is_img = item_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))
        
        if not is_dir and not is_img: continue

        display_name = item_state.get("custom_name")
        if not display_name:
            if is_dir: display_name = item_name
            else:
                # Removed .rstrip('.') to preserve user's intended trailing periods (e.g. Hz.)
                display_name = os.path.splitext(item_name)[0]
                if not display_name: display_name = item_name

        new_item = QTreeWidgetItem([display_name])
        
        base_flags = new_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled | Qt.ItemIsEditable
        if is_dir: new_item.setFlags(base_flags | Qt.ItemIsAutoTristate | Qt.ItemIsDropEnabled)
        else: new_item.setFlags(base_flags & ~Qt.ItemIsDropEnabled)
        
        new_item.setCheckState(0, item_state.get("check_state", Qt.Checked))
            
        if item_state.get("crop"):
            new_item.setData(0, Qt.UserRole + 3, item_state["crop"])
            tag = " [FOLDER CROPPED]" if is_dir else " [IMAGE CROPPED]"
            new_item.setText(0, f"{display_name}{tag}") 
            
        cap = item_state.get("caption")
        if cap:
            new_item.setData(0, Qt.UserRole, cap)
            new_item.setBackground(0, Qt.darkBlue)

        new_item.setData(0, Qt.UserRole + 1, item_path)
        new_item.setData(0, Qt.UserRole + 2, "folder" if is_dir else "file")
        new_item.setData(0, Qt.UserRole + 4, item_name) 
        new_item.setData(0, Qt.UserRole + 6, display_name) 

        if item_path in metadata:
            meta = metadata[item_path]
            new_item.setData(0, Qt.UserRole + 7, meta.get("type"))
            new_item.setData(0, Qt.UserRole + 8, meta.get("excel_path"))
            new_item.setData(0, Qt.UserRole + 9, meta.get("sheet_name"))
            new_item.setData(0, Qt.UserRole + 10, meta.get("voltage"))
            if is_dir: new_item.setData(0, Qt.UserRole + 11, meta.get("include_setup_table", False))
        else:
            if is_dir: new_item.setData(0, Qt.UserRole + 11, False)

        if isinstance(parent_widget, QTreeWidgetItem): parent_widget.addChild(new_item)
        else: parent_widget.addTopLevelItem(new_item)

        if is_dir:
            new_item.setExpanded(item_state.get("expanded", True))
            build_perf_tree(new_item, item_path, metadata)
            
    return True

global_wave_state = {}

def capture_wave_state(app):
    global global_wave_state
    def process_item(item):
        item_path = item.data(0, Qt.UserRole + 1)
        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        if item_path:
            global_wave_state[item_path] = {
                "check_state": item.checkState(0), 
                "crop": item.data(0, Qt.UserRole + 3), 
                "caption": item.data(0, Qt.UserRole),
                "custom_name": item.text(0).replace(" [FOLDER CROPPED]", "").replace(" [IMAGE CROPPED]", "").strip(),
                "expanded": item.isExpanded()
            }
            if is_folder:
                order = [item.child(i).data(0, Qt.UserRole + 4) for i in range(item.childCount()) if item.child(i).data(0, Qt.UserRole + 4)]
                global_wave_state[item_path]["order"] = order
                for i in range(item.childCount()): process_item(item.child(i))

    root_paths = {os.path.dirname(app.waveform_tree.topLevelItem(i).data(0, Qt.UserRole + 1)) 
                  for i in range(app.waveform_tree.topLevelItemCount()) if app.waveform_tree.topLevelItem(i).data(0, Qt.UserRole + 1)}
            
    for rp in root_paths:
        if rp not in global_wave_state: global_wave_state[rp] = {}
        global_wave_state[rp]["order"] = [app.waveform_tree.topLevelItem(i).data(0, Qt.UserRole + 4) 
                                          for i in range(app.waveform_tree.topLevelItemCount()) 
                                          if app.waveform_tree.topLevelItem(i).data(0, Qt.UserRole + 1) and os.path.dirname(app.waveform_tree.topLevelItem(i).data(0, Qt.UserRole + 1)) == rp]
    
    for i in range(app.waveform_tree.topLevelItemCount()): process_item(app.waveform_tree.topLevelItem(i))

def update_waveform_tree(app, capture=True):
    if capture: capture_wave_state(app)
    app.waveform_tree.blockSignals(True)
    app.waveform_tree.clear()
    valid_paths = [p.strip() for p in app.waveforms_path.text().split(';') if p.strip() and os.path.isdir(p.strip())]
    for path in valid_paths: build_tree(app.waveform_tree, path)
    app.waveform_tree.blockSignals(False)

def build_tree(parent_widget, current_path):
    try: items = os.listdir(current_path)
    except PermissionError: return False

    saved_order = global_wave_state.get(current_path, {}).get("order", [])
    items.sort(key=lambda x: saved_order.index(x) if x in saved_order else 999999)

    for item_name in items:
        item_path = os.path.join(current_path, item_name)
        item_state = global_wave_state.get(item_path, {}) 
        is_dir = os.path.isdir(item_path)
        is_img = item_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))
        
        if not is_dir and not is_img: continue

        display_name = item_state.get("custom_name")
        if not display_name:
            if is_dir: display_name = item_name
            else:
                # Removed .rstrip('.') to preserve user's intended trailing periods (e.g. Hz.)
                display_name = os.path.splitext(item_name)[0]
                if not display_name: display_name = item_name

        new_item = QTreeWidgetItem([display_name])
        
        base_flags = new_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled | Qt.ItemIsEditable
        if is_dir: new_item.setFlags(base_flags | Qt.ItemIsAutoTristate | Qt.ItemIsDropEnabled)
        else: new_item.setFlags(base_flags & ~Qt.ItemIsDropEnabled)
        
        new_item.setCheckState(0, item_state.get("check_state", Qt.Checked))
            
        if item_state.get("crop"):
            new_item.setData(0, Qt.UserRole + 3, item_state["crop"])
            tag = " [FOLDER CROPPED]" if is_dir else " [IMAGE CROPPED]"
            new_item.setText(0, f"{display_name}{tag}") 
            
        cap = item_state.get("caption")
        if cap:
            new_item.setData(0, Qt.UserRole, cap)
            new_item.setBackground(0, Qt.darkBlue)

        new_item.setData(0, Qt.UserRole + 1, item_path)
        new_item.setData(0, Qt.UserRole + 2, "folder" if is_dir else "file")
        new_item.setData(0, Qt.UserRole + 4, item_name) 
        new_item.setData(0, Qt.UserRole + 6, display_name) 

        if isinstance(parent_widget, QTreeWidgetItem): parent_widget.addChild(new_item)
        else: parent_widget.addTopLevelItem(new_item)

        if is_dir:
            new_item.setExpanded(item_state.get("expanded", True))
            build_tree(new_item, item_path)

    return True

# --- STATE AND METADATA SYNC HELPERS ---

def sync_state_on_rename(tree_type, old_path, new_path, old_name, new_name, parent_path):
    state = global_perf_state if tree_type == "perf" else global_wave_state
    
    if parent_path in state and "order" in state[parent_path]:
        order_list = state[parent_path]["order"]
        if old_name in order_list:
            order_list[order_list.index(old_name)] = new_name
            
    new_state = {}
    for k, v in state.items():
        if k == old_path or k.startswith(old_path + os.sep):
            new_k = k.replace(old_path, new_path, 1)
            new_state[new_k] = v
        else:
            new_state[k] = v
    state.clear()
    state.update(new_state)

def sync_metadata_on_rename(old_path, new_path):
    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    meta_path = os.path.join(perf_dir, "metadata.json")
    if not os.path.exists(meta_path): return
    
    with open(meta_path, 'r') as f: meta = json.load(f)
    new_meta = {}
    changed = False
    
    for k, v in meta.items():
        if k == old_path or k.startswith(old_path + os.sep):
            new_k = k.replace(old_path, new_path, 1)
            new_meta[new_k] = v
            changed = True
        else:
            new_meta[k] = v
            
    if changed:
        with open(meta_path, 'w') as f: json.dump(new_meta, f)

def sync_metadata_on_delete(paths_to_delete):
    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    meta_path = os.path.join(perf_dir, "metadata.json")
    if not os.path.exists(meta_path): return
    
    with open(meta_path, 'r') as f: meta = json.load(f)
    new_meta = {}
    changed = False
    
    for k, v in meta.items():
        if any(k == p or k.startswith(p + os.sep) for p in paths_to_delete):
            changed = True
        else:
            new_meta[k] = v
            
    if changed:
        with open(meta_path, 'w') as f: json.dump(new_meta, f)