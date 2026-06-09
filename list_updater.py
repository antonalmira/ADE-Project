from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QTreeWidgetItem
import os
import json

# NOTE: These module-level dicts are a known limitation — they persist for the
# entire process lifetime and are shared by all DocuApp instances. For a
# single-window application this is acceptable, but if multiple windows are
# ever needed these should be moved into the DocuApp instance itself.
global_perf_state = {}
global_wave_state = {}


def capture_tree_state(tree_widget, state_dict):
    def process_item(item):
        item_path = item.data(0, Qt.UserRole + 1)
        is_folder = item.data(0, Qt.UserRole + 2) == "folder"
        if item_path:
            state_dict[item_path] = {
                "check_state": item.checkState(0),
                "crop": item.data(0, Qt.UserRole + 3),
                "caption": item.data(0, Qt.UserRole),
                "custom_name": (
                    item.text(0)
                    .replace(" [FOLDER CROPPED]", "")
                    .replace(" [IMAGE CROPPED]", "")
                    .strip()
                ),
                "expanded": item.isExpanded(),
                "setup_table": item.data(0, Qt.UserRole + 11)
            }
            if is_folder:
                order = [
                    item.child(i).data(0, Qt.UserRole + 4)
                    for i in range(item.childCount())
                    if item.child(i).data(0, Qt.UserRole + 4)
                ]
                state_dict[item_path]["order"] = order
                for i in range(item.childCount()):
                    process_item(item.child(i))

    root_paths = {
        os.path.dirname(tree_widget.topLevelItem(i).data(0, Qt.UserRole + 1))
        for i in range(tree_widget.topLevelItemCount())
        if tree_widget.topLevelItem(i).data(0, Qt.UserRole + 1)
    }

    for rp in root_paths:
        if rp not in state_dict:
            state_dict[rp] = {}
        state_dict[rp]["order"] = [
            tree_widget.topLevelItem(i).data(0, Qt.UserRole + 4)
            for i in range(tree_widget.topLevelItemCount())
            if (tree_widget.topLevelItem(i).data(0, Qt.UserRole + 1)
                and os.path.dirname(tree_widget.topLevelItem(i).data(0, Qt.UserRole + 1)) == rp)
        ]

    for i in range(tree_widget.topLevelItemCount()):
        process_item(tree_widget.topLevelItem(i))


def capture_perf_state(app):
    capture_tree_state(app.performance_tree, global_perf_state)


def capture_wave_state(app):
    capture_tree_state(app.waveform_tree, global_wave_state)


def build_tree(parent_widget, current_path, state_dict, metadata=None, perf_dir=None):
    try:
        items = os.listdir(current_path)
    except PermissionError:
        return False

    saved_order = state_dict.get(current_path, {}).get("order", [])
    items.sort(key=lambda x: saved_order.index(x) if x in saved_order else 999999)

    for item_name in items:
        if item_name == "metadata.json":
            continue

        item_path = os.path.join(current_path, item_name)
        item_state = state_dict.get(item_path, {})
        is_dir = os.path.isdir(item_path)
        is_img = item_name.lower().endswith(('.png', '.jpg', '.jpeg', '.bmp'))

        if not is_dir and not is_img:
            continue

        display_name = item_state.get("custom_name")
        if not display_name:
            if is_dir:
                display_name = item_name
            else:
                display_name = os.path.splitext(item_name)[0]
                if not display_name:
                    display_name = item_name

        new_item = QTreeWidgetItem([display_name])

        base_flags = new_item.flags() | Qt.ItemIsUserCheckable | Qt.ItemIsDragEnabled | Qt.ItemIsEditable
        if is_dir:
            new_item.setFlags(base_flags | Qt.ItemIsAutoTristate | Qt.ItemIsDropEnabled)
        else:
            new_item.setFlags(base_flags & ~Qt.ItemIsDropEnabled)

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

        # Resolve metadata using relative paths when a perf_dir base is provided.
        # This avoids hard-coded absolute paths from another machine breaking lookups.
        if metadata is not None:
            rel_key = (
                os.path.relpath(item_path, perf_dir)
                if perf_dir and os.path.isabs(item_path)
                else item_path
            )
            # Try relative key first, then fall back to absolute for backward compat
            meta = metadata.get(rel_key) or metadata.get(item_path)
            if meta:
                new_item.setData(0, Qt.UserRole + 7, meta.get("type"))
                new_item.setData(0, Qt.UserRole + 8, meta.get("excel_path"))
                new_item.setData(0, Qt.UserRole + 9, meta.get("sheet_name"))
                new_item.setData(0, Qt.UserRole + 10, meta.get("voltage"))
                if is_dir:
                    new_item.setData(0, Qt.UserRole + 11, meta.get("include_setup_table", False))
            else:
                if is_dir:
                    new_item.setData(0, Qt.UserRole + 11, item_state.get("setup_table", False))
        else:
            if is_dir:
                setup_val = item_state.get("setup_table", False)
                if setup_val is None:
                    setup_val = False
                new_item.setData(0, Qt.UserRole + 11, setup_val)

        if isinstance(parent_widget, QTreeWidgetItem):
            parent_widget.addChild(new_item)
        else:
            parent_widget.addTopLevelItem(new_item)

        if is_dir:
            build_tree(new_item, item_path, state_dict, metadata, perf_dir)
            new_item.setExpanded(item_state.get("expanded", False))

    return True


def update_performance_tree(app, capture=True):
    if capture:
        capture_perf_state(app)

    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    if not os.path.exists(perf_dir):
        return

    meta_path = os.path.join(perf_dir, "metadata.json")
    metadata = {}
    if os.path.exists(meta_path):
        with open(meta_path, 'r') as f:
            metadata = json.load(f)

    app.performance_tree.blockSignals(True)
    app.performance_tree.clear()
    build_tree(app.performance_tree, perf_dir, global_perf_state, metadata, perf_dir)
    app.performance_tree.blockSignals(False)


def update_waveform_tree(app, capture=True):
    if capture:
        capture_wave_state(app)
    app.waveform_tree.blockSignals(True)
    app.waveform_tree.clear()
    valid_paths = [
        p.strip()
        for p in app.waveforms_path.text().split(';')
        if p.strip() and os.path.isdir(p.strip())
    ]
    for path in valid_paths:
        build_tree(app.waveform_tree, path, global_wave_state)
    app.waveform_tree.blockSignals(False)


# ---------------------------------------------------------------------------
# State and metadata sync helpers
# ---------------------------------------------------------------------------

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
    if not os.path.exists(meta_path):
        return

    with open(meta_path, 'r') as f:
        meta = json.load(f)

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
        with open(meta_path, 'w') as f:
            json.dump(new_meta, f)


def sync_metadata_on_delete(paths_to_delete):
    perf_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Performance Data")
    meta_path = os.path.join(perf_dir, "metadata.json")
    if not os.path.exists(meta_path):
        return

    with open(meta_path, 'r') as f:
        meta = json.load(f)

    new_meta = {}
    changed = False

    for k, v in meta.items():
        if any(k == p or k.startswith(p + os.sep) for p in paths_to_delete):
            changed = True
        else:
            new_meta[k] = v

    if changed:
        with open(meta_path, 'w') as f:
            json.dump(new_meta, f)