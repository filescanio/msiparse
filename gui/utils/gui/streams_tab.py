"""
Streams tab functionality for the MSI Parser GUI
"""

import json
import tempfile
import shutil
from PyQt5.QtWidgets import QTreeWidgetItem, QMenu, QAction
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
from PyQt5.QtWidgets import QApplication

from threads.identifystreams import IdentifyStreamsThread
from utils.gui.preview import show_hex_view, show_text_preview, show_image_preview, show_archive_preview
from utils.gui.extraction import extract_single_stream
from utils.common import format_file_size

def list_streams(parent):
    """List MSI streams"""
    if not parent.msi_file_path:
        return
        
    command = [parent.msiparse_path, "list_streams", parent.msi_file_path]
    output = parent.run_command_safe(command)
    if output:
        display_streams(parent, output)
        
def display_streams(parent, output):
    """Display streams in the streams tab"""
    try:
        streams = json.loads(output)
        parent.streams_data = streams
        
        parent.streams_filter.clear()
        parent.streams_tree.setSortingEnabled(False)
        parent.streams_tree.clear()
        
        for stream in streams:
            item = QTreeWidgetItem([stream, "", "", "", ""])
            item.setIcon(0, parent.group_icons['unknown'])
            parent.streams_tree.addTopLevelItem(item)
            
        parent.streams_tree.setSortingEnabled(True)
        parent.original_order = True
        resize_streams_columns(parent)
        parent.statusBar().showMessage(f"Found {len(streams)} streams")
    except json.JSONDecodeError:
        parent.handle_error("Parse Error", "Error parsing streams output", show_dialog=True)
        
def identify_streams(parent):
    """Identify the file types of all streams using Magika"""
    if not parent.msi_file_path or not parent.streams_data:
        return
        
    parent.statusBar().showMessage("Identifying stream file types...")
    parent.identify_streams_button.setEnabled(False)
    
    temp_dir = tempfile.mkdtemp()
    thread = IdentifyStreamsThread(parent, parent.streams_data, temp_dir)
    parent.active_threads.append(thread)
    
    thread.progress_updated.connect(lambda current, total: update_identify_progress(parent, current, total))
    thread.stream_identified.connect(lambda name, group, mime, size, hash: update_stream_file_type(parent, name, group, mime, size, hash))
    thread.finished.connect(lambda: identify_streams_finished(parent, thread, temp_dir))
    thread.error_occurred.connect(lambda msg: parent.handle_error("Identification Error", msg))
    
    thread.start()

def update_identify_progress(parent, current, total):
    """Update the progress during stream identification"""
    percentage = (current / total) * 100
    parent.statusBar().showMessage(f"Identifying stream types: {current}/{total} ({percentage:.1f}%)")
    QApplication.processEvents()

def update_stream_file_type(parent, stream_name, group, mime_type, file_size, sha1_hash):
    """Update the group, MIME type, size, and SHA1 hash for a stream in the tree"""
    try:
        was_sorting_enabled = parent.streams_tree.isSortingEnabled()
        parent.streams_tree.setSortingEnabled(False)
        current_filter = parent.streams_filter.text().lower()
        
        for i in range(parent.streams_tree.topLevelItemCount()):
            item = parent.streams_tree.topLevelItem(i)
            if item.text(0) == stream_name:
                item.setText(1, group)
                item.setText(2, mime_type)
                item.setText(3, file_size)
                item.setText(4, sha1_hash)
                
                set_icon_for_group(parent, item, group)
                
                if file_size != "Unknown":
                    try:
                        # Convert the formatted size back to bytes for sorting
                        size_parts = file_size.split()
                        if len(size_parts) == 2:
                            value = float(size_parts[0])
                            unit = size_parts[1]
                            multiplier = {
                                'B': 1,
                                'KB': 1024,
                                'MB': 1024 * 1024,
                                'GB': 1024 * 1024 * 1024,
                                'TB': 1024 * 1024 * 1024 * 1024
                            }.get(unit, 0)
                            size_value = value * multiplier
                            item.setData(3, Qt.UserRole, size_value)
                    except (ValueError, IndexError):
                        pass
                        
                if current_filter:
                    match_found = any(current_filter in item.text(col).lower() 
                                    for col in range(parent.streams_tree.columnCount()))
                    item.setHidden(not match_found)
                break
        
        parent.streams_tree.setSortingEnabled(was_sorting_enabled)
        QApplication.processEvents()
        
    except Exception as e:
        print(f"Error updating stream type: {str(e)}")

def set_icon_for_group(parent, item, group):
    """Set an appropriate icon based on the file group"""
    if not group or group == "undefined":
        item.setIcon(0, parent.group_icons['unknown'])
        return
        
    item.setIcon(0, parent.group_icons.get(group, parent.group_icons['unknown']))
        
def identify_streams_finished(parent, thread, temp_dir):
    """Called when stream identification is complete"""
    try:
        if thread in parent.active_threads:
            parent.active_threads.remove(thread)
        thread.cleanup()
        
        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception as e:
            parent.handle_error("Cleanup Error", f"Failed to clean up temporary directory: {str(e)}")
        
        parent.identify_streams_button.setEnabled(True)
        current_filter = parent.streams_filter.text().lower()
        
        if current_filter:
            visible_count = sum(1 for i in range(parent.streams_tree.topLevelItemCount())
                              if not parent.streams_tree.topLevelItem(i).isHidden())
            parent.statusBar().showMessage(f"Stream identification completed. Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams")
        else:
            parent.statusBar().showMessage("Stream identification completed")
        
        resize_streams_columns(parent)
        
    except Exception as e:
        parent.handle_error("Cleanup Error", f"Error during cleanup: {str(e)}")
        parent.identify_streams_button.setEnabled(True)
    
def resize_streams_columns(parent):
    """Resize the columns in the streams tree to fit content but not more than 33% of width"""
    if parent.streams_tree.topLevelItemCount() == 0:
        return
        
    for col in range(5):  # 5 columns total
        parent.streams_tree.resizeColumnToContents(col)
        max_width = parent.streams_tree.width() // 5
        if parent.streams_tree.columnWidth(col) > max_width:
            parent.streams_tree.setColumnWidth(col, max_width)

def on_sort_indicator_changed(parent, logical_index, order):
    """Handle changes to the sort indicator"""
    column_name = parent.streams_tree.headerItem().text(logical_index)
    order_str = "ascending" if order == Qt.AscendingOrder else "descending"
    parent.statusBar().showMessage(f"Sorted by {column_name} ({order_str})")
    parent.original_order = False
    
def reset_to_original_order(parent):
    """Reset the tree to the original order"""
    parent.streams_tree.setSortingEnabled(False)
    
    selected_streams = [item.text(0) for item in parent.streams_tree.selectedItems()]
    stream_data = {item.text(0): (item.text(1), item.text(2), item.text(3), item.text(4), item.data(3, Qt.UserRole))
                  for i in range(parent.streams_tree.topLevelItemCount())
                  for item in [parent.streams_tree.topLevelItem(i)]}
    
    current_filter = parent.streams_filter.text().lower()
    parent.streams_tree.clear()
    
    mono_font = QFont("Courier")
    mono_font.setStyleHint(QFont.Monospace)
    mono_font.setFixedPitch(True)
    
    visible_count = 0
    
    for stream in parent.streams_data:
        group, mime_type, file_size, sha1_hash, size_value = stream_data.get(stream, ("", "", "", "", None))
        
        item = QTreeWidgetItem([stream, group, mime_type, file_size, sha1_hash])
        
        if sha1_hash and sha1_hash not in ("Error calculating hash", ""):
            item.setFont(4, mono_font)
        
        if size_value is not None:
            item.setData(3, Qt.UserRole, size_value)
            
        set_icon_for_group(parent, item, group)
        parent.streams_tree.addTopLevelItem(item)
        
        if stream in selected_streams:
            item.setSelected(True)
            
        if current_filter:
            match_found = any(current_filter in item.text(col).lower() 
                            for col in range(parent.streams_tree.columnCount()))
            item.setHidden(not match_found)
            if match_found:
                visible_count += 1
    
    parent.streams_tree.setSortingEnabled(True)
    parent.original_order = True
    resize_streams_columns(parent)
    
    if current_filter:
        parent.statusBar().showMessage(f"Restored original order. Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams")
    else:
        parent.statusBar().showMessage("Restored original order")

def show_streams_context_menu(parent, position):
    """Show context menu for the streams tree"""
    item = parent.streams_tree.itemAt(position)
    if not item:
        return
        
    stream_name = item.text(0)
    group = item.text(1)
    mime_type = item.text(2)
    sha1_hash = item.text(4)
    
    context_menu = QMenu(parent)
    
    # Add Hex View action (always available)
    hex_view_action = QAction("Hex View", parent)
    hex_view_action.triggered.connect(lambda: show_hex_view(parent, stream_name))
    context_menu.addAction(hex_view_action)
    
    # Add group-specific actions
    if group == "image":
        preview_image_action = QAction("Preview Image", parent)
        preview_image_action.triggered.connect(lambda: show_image_preview(parent, stream_name))
        context_menu.addAction(preview_image_action)
        
    elif group in ("text", "code", "document"):
        preview_text_action = QAction("Preview Text", parent)
        preview_text_action.triggered.connect(lambda: show_text_preview(parent, stream_name))
        context_menu.addAction(preview_text_action)
        
    elif group == "archive" and parent.archive_support:
        preview_archive_action = QAction("Preview Archive", parent)
        preview_archive_action.triggered.connect(lambda: show_archive_preview(parent, stream_name))
        context_menu.addAction(preview_archive_action)
    
    context_menu.addSeparator()
    
    # Add Extract File option
    extract_action = QAction("Extract File...", parent)
    extract_action.triggered.connect(lambda: extract_single_stream(parent, stream_name))
    context_menu.addAction(extract_action)
    
    # Add Hash Lookup option if hash is available
    if sha1_hash and sha1_hash not in ("Error calculating hash", ""):
        context_menu.addSeparator()
        
        hash_menu = QMenu("Lookup Hash", parent)
        
        for service, name in [("filescan", "FileScan.io"), 
                            ("metadefender", "MetaDefender Cloud"),
                            ("virustotal", "VirusTotal")]:
            action = QAction(name, parent)
            # Use a regular function instead of lambda to ensure proper closure
            def create_hash_lookup_handler(s):
                def handler():
                    parent.open_hash_lookup(sha1_hash, s)
                return handler
            action.triggered.connect(create_hash_lookup_handler(service))
            hash_menu.addAction(action)
        
        context_menu.addMenu(hash_menu)
    
    context_menu.addSeparator()
    
    # Add Copy submenu
    copy_menu = QMenu("Copy", parent)
    
    copy_name_action = QAction("Stream Name", parent)
    copy_name_action.triggered.connect(lambda: parent.copy_to_clipboard(stream_name))
    copy_menu.addAction(copy_name_action)
    
    if mime_type:
        copy_type_action = QAction("MIME Type", parent)
        copy_type_action.triggered.connect(lambda: parent.copy_to_clipboard(mime_type))
        copy_menu.addAction(copy_type_action)
        
    if sha1_hash and sha1_hash not in ("Error calculating hash", ""):
        copy_hash_action = QAction("SHA1 Hash", parent)
        copy_hash_action.triggered.connect(lambda: parent.copy_to_clipboard(sha1_hash))
        copy_menu.addAction(copy_hash_action)
    
    # Add separator and "Copy Line" option that copies all columns
    copy_menu.addSeparator()
    copy_line_action = QAction("Copy Line", parent)
    copy_line_action.triggered.connect(lambda: parent.copy_to_clipboard("\t".join([
        stream_name,
        group if group else "",
        mime_type if mime_type else "",
        item.text(3) if item.text(3) else "",  # File Size
        sha1_hash if sha1_hash not in ("Error calculating hash", "") else ""
    ])))
    copy_menu.addAction(copy_line_action)
    
    context_menu.addMenu(copy_menu)
    
    context_menu.exec_(parent.streams_tree.mapToGlobal(position))

def filter_streams(parent, filter_text):
    """Filter the streams tree based on the input text"""
    if not filter_text:
        for i in range(parent.streams_tree.topLevelItemCount()):
            parent.streams_tree.topLevelItem(i).setHidden(False)
        parent.statusBar().showMessage(f"Showing all {parent.streams_tree.topLevelItemCount()} streams")
        return
        
    filter_text = filter_text.lower()
    visible_count = 0
    
    for i in range(parent.streams_tree.topLevelItemCount()):
        item = parent.streams_tree.topLevelItem(i)
        match_found = any(filter_text in item.text(col).lower() 
                         for col in range(parent.streams_tree.columnCount()))
        item.setHidden(not match_found)
        if match_found:
            visible_count += 1
            
    parent.statusBar().showMessage(f"Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams") 