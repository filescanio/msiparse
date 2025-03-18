"""
Streams tab functionality for the MSI Parser GUI
"""

import json
import tempfile
import shutil
from PyQt5.QtWidgets import QTreeWidgetItem, QMenu, QAction
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

from threads.identifystreams import IdentifyStreamsThread
from utils.gui.preview import show_hex_view, show_text_preview, show_image_preview, show_archive_preview
from utils.gui.extraction import extract_single_stream

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
        parent.streams_data = streams  # Store for later use
        
        # Clear any existing filter
        parent.streams_filter.clear()
        
        # Disable sorting while populating
        parent.streams_tree.setSortingEnabled(False)
        parent.streams_tree.clear()
        
        for stream in streams:
            # Create item with four columns: Stream Name, Group, MIME Type, and File Size (empty for now)
            item = QTreeWidgetItem([stream, "", "", "", ""])
            # Set default icon
            item.setIcon(0, parent.group_icons['unknown'])
            parent.streams_tree.addTopLevelItem(item)
            
        # Re-enable sorting
        parent.streams_tree.setSortingEnabled(True)
        
        # Mark as original order
        parent.original_order = True
            
        # Resize columns to fit content
        resize_streams_columns(parent)
            
        parent.statusBar().showMessage(f"Found {len(streams)} streams")
    except json.JSONDecodeError:
        parent.handle_error("Parse Error", "Error parsing streams output", show_dialog=True)
        
def identify_streams(parent):
    """Identify the file types of all streams using Magika"""
    if not parent.msi_file_path or not parent.streams_data:
        return
        
    # Set progress bar to show progress
    parent.progress_bar.setRange(0, len(parent.streams_data))
    parent.progress_bar.setValue(0)
    parent.progress_bar.setVisible(True)
    parent.statusBar().showMessage("Identifying stream file types...")
    
    # Disable identify button while running
    parent.identify_streams_button.setEnabled(False)
    
    # Create a temporary directory for stream extraction
    temp_dir = tempfile.mkdtemp()
    
    # Create and start the identification thread
    parent.identify_thread = IdentifyStreamsThread(
        parent,
        parent.streams_data,
        temp_dir
    )
    parent.active_threads.append(parent.identify_thread)
    
    # Connect signals
    parent.identify_thread.progress_updated.connect(parent.update_identify_progress)
    parent.identify_thread.stream_identified.connect(parent.update_stream_file_type)
    parent.identify_thread.finished.connect(parent.identify_streams_finished)
    parent.identify_thread.error_occurred.connect(lambda msg: parent.handle_error("Identification Error", msg))
    
    # Start the thread
    parent.identify_thread.start()
    
def update_identify_progress(parent, current, total):
    """Update the progress bar during stream identification"""
    parent.progress_bar.setValue(current)
    parent.statusBar().showMessage(f"Identifying stream types: {current}/{total}")
    
def update_stream_file_type(parent, stream_name, group, mime_type, file_size, sha1_hash):
    """Update the group, MIME type, size, and SHA1 hash for a stream in the tree"""
    # Temporarily disable sorting while updating
    was_sorting_enabled = parent.streams_tree.isSortingEnabled()
    parent.streams_tree.setSortingEnabled(False)
    
    # Get current filter text
    current_filter = parent.streams_filter.text().lower()
    
    # Find the item for this stream
    for i in range(parent.streams_tree.topLevelItemCount()):
        item = parent.streams_tree.topLevelItem(i)
        if item.text(0) == stream_name:
            item.setText(1, group)
            item.setText(2, mime_type)
            item.setText(3, file_size)
            item.setText(4, sha1_hash)
            
            # Set monospaced font for the hash column only
            if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
                mono_font = QFont("Courier New", 10)
                mono_font.setFixedPitch(True)
                item.setFont(4, mono_font)
            
            # Set icon based on group
            set_icon_for_group(parent, item, group)
            
            # Set data for proper sorting of file sizes
            if file_size != "Unknown":
                try:
                    # Extract the numeric value for sorting
                    if file_size.endswith(" B"):
                        size_value = float(file_size.split(" ")[0])
                    elif file_size.endswith(" KB"):
                        size_value = float(file_size.split(" ")[0]) * 1024
                    elif file_size.endswith(" MB"):
                        size_value = float(file_size.split(" ")[0]) * 1024 * 1024
                    elif file_size.endswith(" GB"):
                        size_value = float(file_size.split(" ")[0]) * 1024 * 1024 * 1024
                    else:
                        size_value = 0
                    item.setData(3, Qt.UserRole, size_value)
                except (ValueError, IndexError):
                    pass
                    
            # Apply current filter if any
            if current_filter:
                # Check if any column contains the filter text
                match_found = False
                for col in range(parent.streams_tree.columnCount()):
                    if current_filter in item.text(col).lower():
                        match_found = True
                        break
                        
                # Show or hide the item based on the match
                item.setHidden(not match_found)
            break
    
    # Restore sorting state
    parent.streams_tree.setSortingEnabled(was_sorting_enabled)
    
def set_icon_for_group(parent, item, group):
    """Set an appropriate icon based on the file group"""
    if not group or group == "undefined":
        item.setIcon(0, parent.group_icons['unknown'])
        return
        
    # Set the icon based on the group
    if group in parent.group_icons:
        item.setIcon(0, parent.group_icons[group])
    else:
        item.setIcon(0, parent.group_icons['unknown'])
        
def identify_streams_finished(parent):
    """Called when stream identification is complete"""
    parent.progress_bar.setVisible(False)
    parent.identify_streams_button.setEnabled(True)
    
    # Clean up the temporary directory
    if hasattr(parent.identify_thread, 'temp_dir'):
        try:
            shutil.rmtree(parent.identify_thread.temp_dir, ignore_errors=True)
        except Exception as e:
            parent.handle_error("Cleanup Error", f"Failed to clean up temporary directory: {str(e)}")
    
    # Get current filter text
    current_filter = parent.streams_filter.text().lower()
    
    # Count visible items if there's a filter
    if current_filter:
        visible_count = 0
        for i in range(parent.streams_tree.topLevelItemCount()):
            if not parent.streams_tree.topLevelItem(i).isHidden():
                visible_count += 1
        parent.statusBar().showMessage(f"Stream identification completed. Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams")
    else:
        parent.statusBar().showMessage("Stream identification completed")
    
    # Resize columns to fit content
    resize_streams_columns(parent)
    
def resize_streams_columns(parent):
    """Resize the columns in the streams tree to fit content but not more than 33% of width"""
    if parent.streams_tree.topLevelItemCount() == 0:
        return
        
    # Resize columns to fit content
    parent.streams_tree.resizeColumnToContents(0)  # Stream Name column
    parent.streams_tree.resizeColumnToContents(1)  # Group column
    parent.streams_tree.resizeColumnToContents(2)  # MIME Type column
    parent.streams_tree.resizeColumnToContents(3)  # File Size column
    parent.streams_tree.resizeColumnToContents(4)  # SHA1 Hash column
    
    # Get total width
    total_width = parent.streams_tree.width()
    
    # Limit each column to 33% of total width
    max_width = total_width // 5
    
    if parent.streams_tree.columnWidth(0) > max_width:
        parent.streams_tree.setColumnWidth(0, max_width)
        
    if parent.streams_tree.columnWidth(1) > max_width:
        parent.streams_tree.setColumnWidth(1, max_width)
        
    if parent.streams_tree.columnWidth(2) > max_width:
        parent.streams_tree.setColumnWidth(2, max_width)
        
    if parent.streams_tree.columnWidth(3) > max_width:
        parent.streams_tree.setColumnWidth(3, max_width)
        
    if parent.streams_tree.columnWidth(4) > max_width:
        parent.streams_tree.setColumnWidth(4, max_width)

def on_sort_indicator_changed(parent, logical_index, order):
    """Handle changes to the sort indicator"""
    # Update status message
    column_name = parent.streams_tree.headerItem().text(logical_index)
    order_str = "ascending" if order == Qt.AscendingOrder else "descending"
    parent.statusBar().showMessage(f"Sorted by {column_name} ({order_str})")
    
    # We're no longer in original order
    parent.original_order = False
    
def reset_to_original_order(parent):
    """Reset the tree to the original order"""
    # Disable sorting temporarily
    parent.streams_tree.setSortingEnabled(False)
    
    # Store current selection
    selected_streams = []
    for item in parent.streams_tree.selectedItems():
        selected_streams.append(item.text(0))
    
    # Store current group, MIME type and sizes
    stream_data = {}
    for i in range(parent.streams_tree.topLevelItemCount()):
        item = parent.streams_tree.topLevelItem(i)
        stream_name = item.text(0)
        group = item.text(1)
        mime_type = item.text(2)
        file_size = item.text(3)
        sha1_hash = item.text(4)
        size_value = item.data(3, Qt.UserRole)
        stream_data[stream_name] = (group, mime_type, file_size, sha1_hash, size_value)
    
    # Get current filter text
    current_filter = parent.streams_filter.text().lower()
    
    # Clear and repopulate the tree
    parent.streams_tree.clear()
    
    # Create monospaced font for hash columns
    mono_font = QFont("Courier New", 10)
    mono_font.setFixedPitch(True)
    
    # Count visible items for status message
    visible_count = 0
    
    for stream in parent.streams_data:
        # Get stored data if available
        group = ""
        mime_type = ""
        file_size = ""
        sha1_hash = ""
        size_value = None
        
        if stream in stream_data:
            group, mime_type, file_size, sha1_hash, size_value = stream_data[stream]
        
        # Create item with five columns
        item = QTreeWidgetItem([stream, group, mime_type, file_size, sha1_hash])
        
        # Set monospaced font for the hash column only
        if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
            item.setFont(4, mono_font)
        
        # Set the size value for proper sorting if available
        if size_value is not None:
            item.setData(3, Qt.UserRole, size_value)
            
        # Set icon based on group
        set_icon_for_group(parent, item, group)
            
        parent.streams_tree.addTopLevelItem(item)
        
        # Restore selection
        if stream in selected_streams:
            item.setSelected(True)
            
        # Apply current filter if any
        if current_filter:
            # Check if any column contains the filter text
            match_found = False
            for col in range(parent.streams_tree.columnCount()):
                if current_filter in item.text(col).lower():
                    match_found = True
                    break
                    
            # Show or hide the item based on the match
            item.setHidden(not match_found)
            
            # Count visible items
            if match_found:
                visible_count += 1
    
    # Re-enable sorting
    parent.streams_tree.setSortingEnabled(True)
    
    # Mark as original order
    parent.original_order = True
    
    # Resize columns to fit content
    resize_streams_columns(parent)
    
    # Update status
    if current_filter:
        parent.statusBar().showMessage(f"Restored original order. Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams")
    else:
        parent.statusBar().showMessage("Restored original order")

def show_streams_context_menu(parent, position):
    """Show context menu for the streams tree"""
    # Get the item at the position
    item = parent.streams_tree.itemAt(position)
    if not item:
        return
        
    # Get the stream name, group and MIME type
    stream_name = item.text(0)
    group = item.text(1)
    mime_type = item.text(2)
    sha1_hash = item.text(4)
    
    # Create context menu
    context_menu = QMenu(parent)
    
    # Add Hex View action (always available)
    hex_view_action = QAction("Hex View", parent)
    hex_view_action.triggered.connect(lambda: show_hex_view(parent, stream_name))
    context_menu.addAction(hex_view_action)
    
    # Add group-specific actions
    if group == "image":
        # Image preview action
        preview_image_action = QAction("Preview Image", parent)
        preview_image_action.triggered.connect(lambda: show_image_preview(parent, stream_name))
        context_menu.addAction(preview_image_action)
        
    elif group == "text" or group == "code" or group == "document":
        # Text preview action
        preview_text_action = QAction("Preview Text", parent)
        preview_text_action.triggered.connect(lambda: show_text_preview(parent, stream_name))
        context_menu.addAction(preview_text_action)
        
    elif group == "archive" and parent.archive_support:
        # Archive preview action
        preview_archive_action = QAction("Preview Archive", parent)
        preview_archive_action.triggered.connect(lambda: show_archive_preview(parent, stream_name))
        context_menu.addAction(preview_archive_action)
    
    # Add separator
    context_menu.addSeparator()
    
    # Add Extract File option
    extract_action = QAction("Extract File...", parent)
    extract_action.triggered.connect(lambda: extract_single_stream(parent, stream_name))
    context_menu.addAction(extract_action)
    
    # Add Hash Lookup option if hash is available
    if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
        # Add separator
        context_menu.addSeparator()
        
        # Add Hash Lookup submenu
        hash_menu = QMenu("Lookup Hash", parent)
        
        # Add FileScan.io option (first)
        fs_action = QAction("FileScan.io", parent)
        fs_action.triggered.connect(lambda: parent.open_hash_lookup(sha1_hash, "filescan"))
        hash_menu.addAction(fs_action)
        
        # Add MetaDefender Cloud option (second)
        md_action = QAction("MetaDefender Cloud", parent)
        md_action.triggered.connect(lambda: parent.open_hash_lookup(sha1_hash, "metadefender"))
        hash_menu.addAction(md_action)
        
        # Add VirusTotal option (third)
        vt_action = QAction("VirusTotal", parent)
        vt_action.triggered.connect(lambda: parent.open_hash_lookup(sha1_hash, "virustotal"))
        hash_menu.addAction(vt_action)
        
        context_menu.addMenu(hash_menu)
    
    # Add separator before copy options
    context_menu.addSeparator()
    
    # Add Copy submenu at the bottom
    copy_menu = QMenu("Copy", parent)
    
    # Add Copy options to the submenu
    copy_name_action = QAction("Stream Name", parent)
    copy_name_action.triggered.connect(lambda: parent.copy_to_clipboard(stream_name))
    copy_menu.addAction(copy_name_action)
    
    if mime_type:  # Only add if mime_type is not empty
        copy_type_action = QAction("MIME Type", parent)
        copy_type_action.triggered.connect(lambda: parent.copy_to_clipboard(mime_type))
        copy_menu.addAction(copy_type_action)
        
    if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
        copy_hash_action = QAction("SHA1 Hash", parent)
        copy_hash_action.triggered.connect(lambda: parent.copy_to_clipboard(sha1_hash))
        copy_menu.addAction(copy_hash_action)
    
    context_menu.addMenu(copy_menu)
    
    # Show the context menu
    context_menu.exec_(parent.streams_tree.mapToGlobal(position))

def filter_streams(parent, filter_text):
    """Filter the streams tree based on the input text"""
    # If no filter text, show all items
    if not filter_text:
        for i in range(parent.streams_tree.topLevelItemCount()):
            parent.streams_tree.topLevelItem(i).setHidden(False)
        parent.statusBar().showMessage(f"Showing all {parent.streams_tree.topLevelItemCount()} streams")
        return
        
    # Convert filter text to lowercase for case-insensitive matching
    filter_text = filter_text.lower()
    
    # Count visible items for status message
    visible_count = 0
    
    # Check each item against the filter
    for i in range(parent.streams_tree.topLevelItemCount()):
        item = parent.streams_tree.topLevelItem(i)
        
        # Check if any column contains the filter text
        match_found = False
        for col in range(parent.streams_tree.columnCount()):
            if filter_text in item.text(col).lower():
                match_found = True
                break
                
        # Show or hide the item based on the match
        item.setHidden(not match_found)
        
        # Count visible items
        if match_found:
            visible_count += 1
            
    # Update status message
    parent.statusBar().showMessage(f"Showing {visible_count} of {parent.streams_tree.topLevelItemCount()} streams") 