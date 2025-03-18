import os
import tempfile
import shutil
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                           QTreeWidget, QMessageBox, QProgressBar,
                           QMenu, QAction, QFileDialog, QApplication, QLineEdit, QShortcut)
from PyQt5.QtCore import Qt, QThread
from PyQt5.QtGui import QKeySequence, QFont
import magika
from concurrent.futures import ThreadPoolExecutor, as_completed
from queue import Queue, Empty
from threading import Lock

# Import common utilities
from utils.common import (format_file_size, calculate_sha1, TableHelper, TreeHelper, 
                         FileIdentificationHelper)
from utils.preview import (show_hex_view_dialog, show_text_preview_dialog, 
                          show_image_preview_dialog)

# Try to import libarchive for archive handling
try:
    import libarchive
    LIBARCHIVE_AVAILABLE = True
except (ImportError, TypeError):
    LIBARCHIVE_AVAILABLE = False

class ArchivePreviewDialog(QDialog):
    """Dialog for displaying archive contents"""
    def __init__(self, parent, archive_name, archive_path, group_icons, auto_identify=False):
        super().__init__(parent)
        self.parent = parent
        self.archive_name = archive_name
        self.archive_path = archive_path
        self.group_icons = group_icons
        self.archive_entries = []
        self.temp_dir = None
        self.extracted_files = set()  # Track extracted files for cleanup
        self.auto_identify = auto_identify  # Whether to auto-identify files
        
        # Initialize magika
        self.magika_client = magika.Magika()
        
        # Initialize parallel processing components
        self.result_queue = Queue()
        self.ui_lock = Lock()
        self.max_workers = min(32, os.cpu_count() * 4)  # Adjust based on system
            
        self.init_ui()
        self.load_archive_contents()
        
    def init_ui(self):
        self.setWindowTitle(f"Archive Preview: {self.archive_name}")
        self.setGeometry(100, 100, 900, 600)
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Status label at the top
        self.status_label = QLabel("Loading archive contents...")
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        layout.addWidget(self.progress_bar)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        # Add Identify Files button
        self.identify_button = QPushButton("Identify Files")
        self.identify_button.clicked.connect(self.auto_identify_files)
        button_layout.addWidget(self.identify_button)
        
        # Add spacer to push buttons to the left
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # Add filter input for archive contents
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filter:")
        self.contents_filter = QLineEdit()
        self.contents_filter.setPlaceholderText("Type to filter contents... (Ctrl+F)")
        self.contents_filter.textChanged.connect(self.filter_contents)
        self.contents_filter.setClearButtonEnabled(True)  # Add clear button inside the field
        
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.contents_filter)
        layout.addLayout(filter_layout)
        
        # Create a tree widget for the archive contents
        self.contents_tree = QTreeWidget()
        self.contents_tree.setHeaderLabels(["Name", "Group", "MIME Type", "Size", "SHA1 Hash"])
        self.contents_tree.setSelectionMode(QTreeWidget.SingleSelection)
        
        # Enable context menu for contents tree
        self.contents_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.contents_tree.customContextMenuRequested.connect(self.show_context_menu)
        
        layout.addWidget(self.contents_tree)
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close_and_cleanup)
        layout.addWidget(close_button)
        
        # Set up keyboard shortcuts
        self.setup_shortcuts()
        
    def setup_shortcuts(self):
        """Set up keyboard shortcuts for the dialog"""
        # Ctrl+F to focus on filter
        self.filter_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.filter_shortcut.activated.connect(lambda: self.contents_filter.setFocus())
        
        # Escape key to clear filter when it has focus
        self.filter_escape = QShortcut(QKeySequence("Escape"), self.contents_filter)
        self.filter_escape.activated.connect(self.contents_filter.clear)
        
    def filter_contents(self, filter_text):
        """Filter the contents tree based on the input text"""
        # If no filter text, show all items
        if not filter_text:
            self.show_all_items(self.contents_tree.invisibleRootItem())
            self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files)")
            return
            
        # Convert filter text to lowercase for case-insensitive matching
        filter_text = filter_text.lower()
        
        # Apply filter recursively and count visible items
        visible_count = self.apply_filter_recursive(self.contents_tree.invisibleRootItem(), filter_text)
        
        # Update status message
        self.status_label.setText(f"Archive: {self.archive_name} (showing {visible_count} of {len(self.archive_entries)} files)")
        
    def show_all_items(self, parent_item):
        """Show all items in the tree recursively"""
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            child.setHidden(False)
            self.show_all_items(child)
            
    def apply_filter_recursive(self, parent_item, filter_text):
        """Apply filter recursively to all items and return count of visible items"""
        visible_count = 0
        
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # Check if this is a file (has path data)
            is_file = child.data(0, Qt.UserRole) is not None
            
            if is_file:
                # Check if any column contains the filter text
                match_found = any(filter_text in child.text(col).lower() 
                                for col in range(self.contents_tree.columnCount()))
                
                # Show or hide the item based on the match
                child.setHidden(not match_found)
                
                # Count visible items
                if match_found:
                    visible_count += 1
            else:
                # For directories, check children recursively
                child_visible_count = self.apply_filter_recursive(child, filter_text)
                
                # Show directory if any children are visible
                child.setHidden(child_visible_count == 0)
                
                # Add to visible count
                visible_count += child_visible_count
                
        return visible_count
        
    def close_and_cleanup(self):
        """Close the dialog and clean up temporary files"""
        # Clean up temporary directory
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir)
            except Exception:
                pass
                
        # Close the dialog
        self.accept()
        
    def load_archive_contents(self):
        """Load the contents of the archive"""
        if not LIBARCHIVE_AVAILABLE:
            self.status_label.setText("Error: libarchive is not available")
            self.progress_bar.setVisible(False)
            QMessageBox.critical(self, "Error", "libarchive is not available. Cannot preview archive.")
            return
            
        try:
            # Create a temporary directory for extracted files
            self.temp_dir = tempfile.mkdtemp()
            
            # Show progress bar (indeterminate) during loading
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            self.progress_bar.setVisible(True)
            self.status_label.setText(f"Loading archive: {self.archive_name}...")
            QApplication.processEvents()  # Ensure UI updates
            
            # List archive entries
            with libarchive.file_reader(self.archive_path) as archive:
                # Count entries for progress
                self.archive_entries = list(archive)
                
            # Update status
            self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files)")
            
            # Populate the tree
            self.populate_tree()
            
            # Hide progress bar when done
            self.progress_bar.setVisible(False)
            
            # Auto-resize columns
            TableHelper.auto_resize_columns(self.contents_tree)
            
            # Auto-identify files only if enabled
            if self.auto_identify:
                self.auto_identify_files()
            
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            self.progress_bar.setVisible(False)
            QMessageBox.critical(self, "Error", f"Failed to read archive: {str(e)}")
        
    def process_item(self, item):
        """Process a single item in a worker thread"""
        try:
            file_path = self.extract_file(item)
            if file_path:
                file_path_obj = Path(file_path)
                if file_path_obj.exists() and file_path_obj.is_file():
                    group, mime_type = FileIdentificationHelper.identify_file_with_magika(file_path_obj, self.magika_client)
                    size = format_file_size(os.path.getsize(file_path))
                    sha1 = calculate_sha1(file_path)
                    return (item, group, mime_type, size, sha1)
        except Exception:
            pass
        return (item, "unknown", "application/octet-stream", "Unknown", "Error calculating hash")

    def update_ui_from_queue(self):
        """Update UI with results from the queue"""
        try:
            while True:
                item, group, mime_type, size, sha1 = self.result_queue.get_nowait()
                with self.ui_lock:
                    self.update_item_with_file_info(item, group, mime_type, None, size, sha1)
                self.result_queue.task_done()
        except Empty:
            pass

    def collect_items(self, parent_item):
        """Collect all items that need processing"""
        items_to_process = []
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            if child.data(0, Qt.UserRole):
                items_to_process.append(child)
            items_to_process.extend(self.collect_items(child))
        return items_to_process

    def auto_identify_files(self):
        self.identify_button.setEnabled(False)
        
        try:
            # Collect all items to process
            items_to_process = self.collect_items(self.contents_tree.invisibleRootItem())
            
            total_items = len(items_to_process)
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, total_items)
            self.status_label.setText(f"Identifying file types (0/{total_items})...")
            
            # Start processing items in parallel
            with ThreadPoolExecutor(max_workers=self.max_workers) as executor:
                future_to_item = {executor.submit(self.process_item, item): item 
                                for item in items_to_process}
                
                completed = 0
                for future in as_completed(future_to_item):
                    completed += 1
                    self.progress_bar.setValue(completed)
                    self.status_label.setText(f"Identifying file types ({completed}/{total_items})...")
                    
                    # Process results and update UI
                    result = future.result()
                    self.result_queue.put(result)
                    self.update_ui_from_queue()
                    
                    # Process Qt events
                    QApplication.processEvents()
            
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files, {completed} identified)")
            TableHelper.auto_resize_columns(self.contents_tree)
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error during identification: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error during identification: {str(e)}")
        finally:
            self.identify_button.setEnabled(True)
        
    def populate_tree(self):
        self.contents_tree.clear()
        tree_structure = {}
        
        for entry in self.archive_entries:
            if entry.isdir:
                continue
                
            path = entry.pathname
            path_parts = path.split('/')
            current_dict = tree_structure
            
            for i, part in enumerate(path_parts):
                if i == len(path_parts) - 1:  # Last part (file)
                    if '' not in current_dict:
                        current_dict[''] = []
                    current_dict[''].append((part, entry.size, path))
                else:  # Directory
                    if part not in current_dict:
                        current_dict[part] = {}
                    current_dict = current_dict[part]
        
        TreeHelper.populate_tree_from_structure(self.contents_tree, tree_structure, self.group_icons)
        TableHelper.auto_resize_columns(self.contents_tree)
            
    def show_context_menu(self, position):
        item = self.contents_tree.itemAt(position)
        if not item or not item.data(0, Qt.UserRole):
            return
            
        file_name = item.text(0)
        group = item.text(1)
        mime_type = item.text(2)
        sha1_hash = item.text(4)
        
        context_menu = QMenu(self)
        
        hex_view_action = QAction("Hex View", self)
        hex_view_action.triggered.connect(lambda: self.show_hex_view(item))
        context_menu.addAction(hex_view_action)
        
        if group == "image":
            preview_image_action = QAction("Preview Image", self)
            preview_image_action.triggered.connect(lambda: self.show_image_preview(item))
            context_menu.addAction(preview_image_action)
        elif group in ["text", "code", "document"]:
            preview_text_action = QAction("Preview Text", self)
            preview_text_action.triggered.connect(lambda: self.show_text_preview(item))
            context_menu.addAction(preview_text_action)
        elif group == "archive":
            preview_archive_action = QAction("Preview Archive", self)
            preview_archive_action.triggered.connect(lambda: self.show_nested_archive_preview(item))
            context_menu.addAction(preview_archive_action)
        
        context_menu.addSeparator()
        
        extract_action = QAction("Extract File...", self)
        extract_action.triggered.connect(lambda: self.extract_file_to_user_location(item))
        context_menu.addAction(extract_action)
        
        if sha1_hash and sha1_hash != "Error calculating hash":
            context_menu.addSeparator()
            hash_lookup_menu = QMenu("Lookup Hash", self)
            
            for service, name in [("filescan", "FileScan.io"), 
                                ("metadefender", "MetaDefender Cloud"),
                                ("virustotal", "VirusTotal")]:
                action = QAction(name, self)
                # Use a regular function instead of lambda to ensure proper closure
                def create_hash_lookup_handler(s):
                    def handler():
                        self.open_hash_lookup(sha1_hash, s)
                    return handler
                action.triggered.connect(create_hash_lookup_handler(service))
                hash_lookup_menu.addAction(action)
            
            context_menu.addMenu(hash_lookup_menu)
        
        context_menu.addSeparator()
        copy_menu = QMenu("Copy", self)
        
        copy_name_action = QAction("File Name", self)
        copy_name_action.triggered.connect(lambda: self.copy_to_clipboard(file_name))
        copy_menu.addAction(copy_name_action)
        
        if mime_type:
            copy_type_action = QAction("MIME Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_to_clipboard(mime_type))
            copy_menu.addAction(copy_type_action)
        
        if sha1_hash and sha1_hash != "Error calculating hash":
            copy_hash_action = QAction("SHA1 Hash", self)
            copy_hash_action.triggered.connect(lambda: self.copy_to_clipboard(sha1_hash))
            copy_menu.addAction(copy_hash_action)
        
        context_menu.addMenu(copy_menu)
        context_menu.exec_(self.contents_tree.mapToGlobal(position))
        
    def show_hex_view(self, item):
        file_path = self.extract_file(item)
        if file_path:
            show_hex_view_dialog(self, item.text(0), file_path)
        
    def show_image_preview(self, item):
        file_path = self.extract_file(item)
        if file_path:
            show_image_preview_dialog(self, item.text(0), file_path)
        
    def show_text_preview(self, item):
        file_path = self.extract_file(item)
        if file_path:
            show_text_preview_dialog(self, item.text(0), file_path)
            
    def show_nested_archive_preview(self, item):
        self.progress_bar.setRange(0, 0)
        self.progress_bar.setVisible(True)
        self.status_label.setText(f"Extracting nested archive: {item.text(0)}...")
            
        file_path = self.extract_file(item)
        if not file_path:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Failed to extract: {item.text(0)}")
            return
            
        try:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Opening nested archive: {item.text(0)}")
            
            from utils.preview import show_archive_preview_dialog
            show_archive_preview_dialog(self.parent, item.text(0), file_path, self.group_icons, self.auto_identify)
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error showing archive preview: {str(e)}")

    def identify_streams_finished(self):
        """Called when stream identification is complete"""
        self.progress_bar.setVisible(False)
        self.identify_button.setEnabled(True)
        
        # Get current filter text
        current_filter = self.contents_filter.text().lower()
        
        # Count visible items if there's a filter
        if current_filter:
            visible_count = self.count_visible_items_recursive(self.contents_tree.invisibleRootItem())
            self.status_label.setText(f"File identification completed. Showing {visible_count} of {len(self.archive_entries)} files")
        else:
            self.status_label.setText("File identification completed")
        
        # Resize columns to fit content
        TableHelper.auto_resize_columns(self.contents_tree)
        
    def count_visible_items_recursive(self, parent_item):
        """Count visible items in the tree recursively"""
        visible_count = 0
        
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If item is not hidden and is a file (has path data)
            if not child.isHidden() and child.data(0, Qt.UserRole) is not None:
                visible_count += 1
                
            # Add count from children
            visible_count += self.count_visible_items_recursive(child)
                
        return visible_count

    def extract_file(self, item):
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            return None
            
        try:
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp()
                
            normalized_paths = [
                full_path,
                full_path.replace('\\', '/'),
                full_path.replace('\\', '/').lstrip('/'),
                os.path.basename(full_path)
            ]
            
            safe_filename = os.path.basename(full_path.replace('\\', '/'))
            output_path = os.path.join(self.temp_dir, safe_filename)
            output_path_obj = Path(output_path)
            
            if output_path in self.extracted_files and output_path_obj.exists():
                return str(output_path_obj)
                
            for normalized_path in normalized_paths:
                for method in [self._extract_with_file_reader, self._extract_with_archive_class]:
                    try:
                        result = method(normalized_path, str(output_path_obj))
                        if result:
                            return result
                    except Exception:
                        continue
            
            try:
                output_path_obj.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(self.archive_path, str(output_path_obj))
                if output_path_obj.exists():
                    self.extracted_files.add(str(output_path_obj))
                    return str(output_path_obj)
            except Exception:
                pass
            
            return os.path.join(self.temp_dir, "dummy_" + safe_filename)
            
        except Exception:
            return os.path.join(self.temp_dir, "dummy_" + os.path.basename(full_path))
            
    def _extract_with_file_reader(self, normalized_path, output_path):
        with libarchive.file_reader(self.archive_path) as archive:
            for entry in archive:
                entry_path = entry.pathname.replace('\\', '/').lstrip('/')
                if entry_path == normalized_path:
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                    with open(output_path, 'wb') as output_file:
                        for block in entry.get_blocks():
                            output_file.write(block)
                    self.extracted_files.add(output_path)
                    return output_path
        return None
        
    def _extract_with_archive_class(self, normalized_path, output_path):
        with open(self.archive_path, 'rb') as archive_file:
            with libarchive.Archive(archive_file) as archive:
                for entry in archive:
                    entry_path = entry.pathname.replace('\\', '/').lstrip('/')
                    if entry_path == normalized_path:
                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                        with open(output_path, 'wb') as output_file:
                            for block in entry.get_blocks():
                                output_file.write(block)
                        self.extracted_files.add(output_path)
                        return output_path
        return None
        
    def extract_file_to_user_location(self, item):
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            QMessageBox.warning(self, "Error", "No file path found for this item")
            return
            
        file_name = item.text(0)
        save_path, _ = QFileDialog.getSaveFileName(self, "Save File", file_name, "All Files (*)")
        
        if not save_path:
            return
            
        try:
            self.progress_bar.setRange(0, 0)
            self.progress_bar.setVisible(True)
            self.status_label.setText(f"Extracting: {file_name}...")
            
            temp_path = self.extract_file(item)
            if not temp_path:
                self.progress_bar.setVisible(False)
                self.status_label.setText(f"Failed to extract: {file_name}")
                return
                
            shutil.copy2(temp_path, save_path)
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Extracted: {file_name} to {save_path}")
            
            QMessageBox.information(self, "Extraction Complete", f"File extracted to:\n{save_path}")
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
            
    def copy_to_clipboard(self, text):
        QApplication.clipboard().setText(text)
        self.status_label.setText(f"Copied to clipboard: {text}")
        
    def open_hash_lookup(self, hash_value, service):
        urls = {
            "filescan": f"https://www.filescan.io/search-result?query={hash_value}",
            "metadefender": f"https://metadefender.com/results/hash/{hash_value}",
            "virustotal": f"https://www.virustotal.com/gui/file/{hash_value}"
        }
        
        if service in urls:
            try:
                webbrowser.open(urls[service])
                self.status_label.setText(f"Opened {service} lookup for {hash_value}")
            except Exception as e:
                QMessageBox.warning(self, "Error", f"Failed to open browser: {str(e)}")

    def update_item_with_file_info(self, item, group, mime_type, file_path, size_text=None, hash_text=None):
        item.setText(1, group)
        item.setText(2, mime_type)
        
        if size_text:
            item.setText(3, size_text)
        elif file_path:
            try:
                item.setText(3, format_file_size(os.path.getsize(file_path)))
            except Exception:
                item.setText(3, "Unknown")
        else:
            item.setText(3, "Unknown")
            
        if hash_text:
            item.setText(4, hash_text)
        elif file_path:
            try:
                sha1 = calculate_sha1(file_path)
                item.setText(4, sha1)
                mono_font = QFont("Courier New", 10)
                mono_font.setFixedPitch(True)
                item.setFont(4, mono_font)
            except Exception:
                item.setText(4, "Error calculating hash")
        else:
            item.setText(4, "")
            
        if self.group_icons and group in self.group_icons:
            item.setIcon(0, self.group_icons[group])
        elif self.group_icons and "unknown" in self.group_icons:
            item.setIcon(0, self.group_icons["unknown"])
            
        current_filter = self.contents_filter.text().lower()
        if current_filter:
            match_found = any(current_filter in item.text(col).lower() 
                            for col in range(self.contents_tree.columnCount()))
            item.setHidden(not match_found)
            
            if match_found:
                parent = item.parent()
                while parent and parent != self.contents_tree.invisibleRootItem():
                    parent.setHidden(False)
                    parent = parent.parent()
