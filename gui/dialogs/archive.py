import os
import tempfile
import shutil
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                           QTreeWidget, QMessageBox, QProgressBar,
                           QMenu, QAction, QFileDialog, QApplication, QLineEdit, QShortcut)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence, QFont
import magika

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
                match_found = False
                for col in range(self.contents_tree.columnCount()):
                    if filter_text in child.text(col).lower():
                        match_found = True
                        break
                        
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
        
    def auto_identify_files(self):
        """Identify file types for all files in the archive"""
        # Disable identify button while running
        self.identify_button.setEnabled(False)
        
        try:
            # Count total items
            total_items = self.count_all_items()
            
            # Show progress bar for identification
            self.progress_bar.setVisible(True)
            self.progress_bar.setRange(0, total_items)
            self.status_label.setText(f"Identifying file types (0/{total_items})...")
            QApplication.processEvents()  # Ensure UI updates
            
            # Process all top-level items
            progress_count = 0
            progress_count = self.identify_items_recursive(self.contents_tree.invisibleRootItem(), progress_count)
            
            # Hide progress bar when done
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files, {progress_count} identified)")
            
            # Auto-resize columns after identification
            TableHelper.auto_resize_columns(self.contents_tree)
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error during identification: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error during identification: {str(e)}")
        finally:
            # Re-enable identify button
            self.identify_button.setEnabled(True)
        
    def count_all_items(self):
        """Count all file items in the tree recursively"""
        return TreeHelper.count_items_recursive(self.contents_tree.invisibleRootItem())
        
    def identify_items_recursive(self, parent_item, progress_count):
        """Recursively identify file types for all items"""
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If it's a file (has path data)
            full_path = child.data(0, Qt.UserRole)
            if full_path:
                # Update progress
                self.progress_bar.setValue(progress_count)
                self.status_label.setText(f"Identifying file types ({progress_count+1}/{self.progress_bar.maximum()}): {child.text(0)}")
                QApplication.processEvents()  # Keep UI responsive
                
                # Identify the file
                self.identify_file(child, show_message=False)
                
                progress_count += 1
            
            # Recursively process child items (for directories)
            progress_count = self.identify_items_recursive(child, progress_count)
            
        return progress_count
        
    def populate_tree(self):
        """Populate the tree with archive entries"""
        # Clear the tree
        self.contents_tree.clear()
        
        # Create a dictionary to store the tree structure
        tree_structure = {}
        
        # Process each entry
        for entry in self.archive_entries:
            path = entry.pathname
            is_dir = entry.isdir
            size = entry.size if not is_dir else 0
            
            # Skip directories for now
            if is_dir:
                continue
                
            # Split the path into components
            path_parts = path.split('/')
            
            # Create tree items for each part of the path
            current_dict = tree_structure
            for i, part in enumerate(path_parts):
                if i == len(path_parts) - 1:  # Last part (file)
                    if '' not in current_dict:
                        current_dict[''] = []
                    current_dict[''].append((part, size, path))
                else:  # Directory
                    if part not in current_dict:
                        current_dict[part] = {}
                    current_dict = current_dict[part]
        
        # Build the tree from the structure
        TreeHelper.populate_tree_from_structure(self.contents_tree, tree_structure, self.group_icons)
        
        # Auto-resize columns
        TableHelper.auto_resize_columns(self.contents_tree)
            
    def show_context_menu(self, position):
        """Show context menu for the contents tree"""
        # Get the item at the position
        item = self.contents_tree.itemAt(position)
        if not item:
            return
            
        # Check if it's a file (has path data)
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            return  # It's a directory
            
        # Get file name and group/mime type if available
        file_name = item.text(0)
        group = item.text(1)
        sha1_hash = item.text(4)
        
        # Create context menu
        context_menu = QMenu(self)
        
        # Add Hex View action (always available)
        hex_view_action = QAction("Hex View", self)
        hex_view_action.triggered.connect(lambda: self.show_hex_view(item))
        context_menu.addAction(hex_view_action)
        
        # Add group-specific actions
        if group == "image":
            # Image preview action
            preview_image_action = QAction("Preview Image", self)
            preview_image_action.triggered.connect(lambda: self.show_image_preview(item))
            context_menu.addAction(preview_image_action)
            
        elif group == "text" or group == "code" or group == "document":
            # Text preview action
            preview_text_action = QAction("Preview Text", self)
            preview_text_action.triggered.connect(lambda: self.show_text_preview(item))
            context_menu.addAction(preview_text_action)
            
        elif group == "archive":
            # Nested archive preview action
            preview_archive_action = QAction("Preview Archive", self)
            preview_archive_action.triggered.connect(lambda: self.show_nested_archive_preview(item))
            context_menu.addAction(preview_archive_action)
        
        # Add separator
        context_menu.addSeparator()
        
        # Add Extract action
        extract_action = QAction("Extract File...", self)
        extract_action.triggered.connect(lambda: self.extract_file_to_user_location(item))
        context_menu.addAction(extract_action)
        
        # Add Copy Hash action if hash is available
        if sha1_hash and sha1_hash != "Error calculating hash":
            context_menu.addSeparator()
            
            copy_hash_action = QAction("Copy SHA1 Hash", self)
            copy_hash_action.triggered.connect(lambda: self.copy_to_clipboard(sha1_hash))
            context_menu.addAction(copy_hash_action)
            
            # Add online hash lookup submenu
            hash_lookup_menu = QMenu("Lookup Hash Online", self)
            
            virustotal_action = QAction("VirusTotal", self)
            virustotal_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "virustotal"))
            hash_lookup_menu.addAction(virustotal_action)
            
            hybrid_action = QAction("Hybrid Analysis", self)
            hybrid_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "hybrid"))
            hash_lookup_menu.addAction(hybrid_action)
            
            context_menu.addMenu(hash_lookup_menu)
        
        # Show the context menu
        context_menu.exec_(self.contents_tree.mapToGlobal(position))
        
    def identify_file(self, item, show_message=True):
        """Identify a file using Magika"""
        # Get the full path from the item data
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            if show_message:
                self.status_label.setText("Error: No file path found for item")
            return
            
        try:
            # Update status to show which file we're processing
            if show_message:
                self.status_label.setText(f"Identifying: {item.text(0)}...")
                QApplication.processEvents()  # Ensure UI updates
            
            # Extract the file to a temporary location
            file_path = self.extract_file(item)
            
            # If extraction failed completely (should not happen with our new fallback)
            if not file_path:
                # Try to identify by extension as fallback
                group, mime_type = FileIdentificationHelper.identify_by_extension(item.text(0))
                
                # Update the item with file information
                if group and mime_type:
                    self.update_item_with_file_info(
                        item, group, mime_type, None, 
                        size_text="Unknown", hash_text="Extraction failed"
                    )
                    if show_message:
                        self.status_label.setText(f"Identified by extension: {item.text(0)} as {mime_type}")
                else:
                    if show_message:
                        self.status_label.setText(f"Failed to identify: {item.text(0)}")
                return
                
            # Convert to Path object for better file existence check
            file_path_obj = Path(file_path)
            file_exists = file_path_obj.exists() and file_path_obj.is_file()
            
            # Identify the file
            if file_exists:
                try:
                    # Use Magika for identification if file exists
                    group, mime_type = FileIdentificationHelper.identify_file_with_magika(file_path_obj, self.magika_client)
                except Exception:
                    # Fallback to extension-based identification
                    group, mime_type = FileIdentificationHelper.identify_by_extension(item.text(0))
            else:
                # Fallback to extension-based identification
                group, mime_type = FileIdentificationHelper.identify_by_extension(item.text(0))
            
            # Update the item with file information
            if file_exists:
                self.update_item_with_file_info(
                    item, group, mime_type, str(file_path_obj)
                )
            else:
                # Update with limited information if file doesn't exist
                self.update_item_with_file_info(
                    item, group, mime_type, None,
                    size_text="Unknown", hash_text="Extraction failed"
                )
            
            # Update status if requested
            if show_message:
                self.status_label.setText(f"Identified: {item.text(0)} as {mime_type}")
                
        except Exception as e:
            # Try to identify by extension as fallback
            try:
                group, mime_type = FileIdentificationHelper.identify_by_extension(item.text(0))
                if group and mime_type:
                    self.update_item_with_file_info(
                        item, group, mime_type, None,
                        size_text="Unknown", hash_text="Error: " + str(e)[:100]
                    )
                    if show_message:
                        self.status_label.setText(f"Identified by extension: {item.text(0)} as {mime_type}")
                    return
            except Exception:
                pass
                
            if show_message:
                self.status_label.setText(f"Error identifying file: {str(e)}")
                
            # Update item with error information
            item.setText(1, "unknown")  # Group
            item.setText(2, "Error: " + str(e)[:50])  # MIME type (truncated error)
            if self.group_icons and "unknown" in self.group_icons:
                item.setIcon(0, self.group_icons["unknown"])
                
    def update_item_with_file_info(self, item, group, mime_type, file_path, size_text=None, hash_text=None):
        """Update a tree item with file information, respecting the current filter"""
        # Get current filter text
        current_filter = self.contents_filter.text().lower()
        
        # Update the item with file information
        item.setText(1, group)
        item.setText(2, mime_type)
        
        # Set size if provided, otherwise calculate it
        if size_text:
            item.setText(3, size_text)
        elif file_path:
            try:
                file_size = os.path.getsize(file_path)
                item.setText(3, format_file_size(file_size))
            except Exception:
                item.setText(3, "Unknown")
        else:
            item.setText(3, "Unknown")
            
        # Set hash if provided, otherwise calculate it
        if hash_text:
            item.setText(4, hash_text)
        elif file_path:
            try:
                sha1 = calculate_sha1(file_path)
                item.setText(4, sha1)
                
                # Set monospaced font for the hash column
                mono_font = QFont("Courier New", 10)
                mono_font.setFixedPitch(True)
                item.setFont(4, mono_font)
            except Exception:
                item.setText(4, "Error calculating hash")
        else:
            item.setText(4, "")
            
        # Set icon based on group
        if self.group_icons and group in self.group_icons:
            item.setIcon(0, self.group_icons[group])
        elif self.group_icons and "unknown" in self.group_icons:
            item.setIcon(0, self.group_icons["unknown"])
            
        # Apply current filter if any
        if current_filter:
            # Check if any column contains the filter text
            match_found = False
            for col in range(self.contents_tree.columnCount()):
                if current_filter in item.text(col).lower():
                    match_found = True
                    break
                    
            # Show or hide the item based on the match
            item.setHidden(not match_found)
            
            # If this is a file in a directory, we need to make sure the parent directories are visible
            parent = item.parent()
            while parent and parent != self.contents_tree.invisibleRootItem():
                if match_found:  # If this item matches, make sure all parent directories are visible
                    parent.setHidden(False)
                parent = parent.parent()
        
    def extract_file(self, item):
        """Extract a file from the archive to a temporary directory"""
        # Get the full path from the item data
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            QMessageBox.warning(self, "Error", "No file path found for this item")
            return None
            
        try:
            # Create a temporary directory if needed
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp()
                
            # Normalize the path for extraction (try different normalizations)
            normalized_paths = [
                full_path,  # Original path
                full_path.replace('\\', '/'),  # Replace backslashes
                full_path.replace('\\', '/').lstrip('/'),  # Remove leading slash
                os.path.basename(full_path)  # Just the filename
            ]
            
            # Create a safe output path
            safe_filename = os.path.basename(full_path.replace('\\', '/'))
            output_path = os.path.join(self.temp_dir, safe_filename)
            
            # Convert to pathlib.Path for better compatibility
            output_path_obj = Path(output_path)
            
            # Check if the file has already been extracted
            if output_path in self.extracted_files and output_path_obj.exists():
                return str(output_path_obj)
                
            # Try different extraction methods with different path normalizations
            for normalized_path in normalized_paths:
                # Try different extraction methods
                extraction_methods = [
                    self._extract_with_file_reader,
                    self._extract_with_archive_class
                ]
                
                for method in extraction_methods:
                    try:
                        result = method(normalized_path, str(output_path_obj))
                        if result:
                            return result
                    except Exception:
                        continue
            
            # If all methods failed, try a direct file copy as a last resort
            try:
                # This is a fallback for when the archive is actually just a renamed file
                # Ensure parent directory exists
                output_path_obj.parent.mkdir(parents=True, exist_ok=True)
                # Copy the file
                shutil.copy2(self.archive_path, str(output_path_obj))
                if output_path_obj.exists():
                    self.extracted_files.add(str(output_path_obj))
                    return str(output_path_obj)
            except Exception:
                pass
            
            # If we get here, all extraction methods failed
            
            # Instead of showing an error message, return a dummy path for identification
            # This allows identification by extension to work even if extraction fails
            dummy_path = os.path.join(self.temp_dir, "dummy_" + safe_filename)
            return dummy_path
            
        except Exception:
            # Return a dummy path instead of None to allow identification by extension
            dummy_path = os.path.join(self.temp_dir, "dummy_" + os.path.basename(full_path))
            return dummy_path
            
    def _extract_with_file_reader(self, normalized_path, output_path):
        """Extract a file using libarchive.file_reader"""
        with libarchive.file_reader(self.archive_path) as archive:
            for entry in archive:
                entry_path = entry.pathname.replace('\\', '/').lstrip('/')
                if entry_path == normalized_path:
                    # Create parent directories if needed
                    os.makedirs(os.path.dirname(output_path), exist_ok=True)
                    
                    # Extract the file
                    with open(output_path, 'wb') as output_file:
                        for block in entry.get_blocks():
                            output_file.write(block)
                            
                    # Add to extracted files set
                    self.extracted_files.add(output_path)
                    return output_path
                    
        return None
        
    def _extract_with_archive_class(self, normalized_path, output_path):
        """Extract a file using libarchive.Archive class"""
        with open(self.archive_path, 'rb') as archive_file:
            with libarchive.Archive(archive_file) as archive:
                for entry in archive:
                    entry_path = entry.pathname.replace('\\', '/').lstrip('/')
                    if entry_path == normalized_path:
                        # Create parent directories if needed
                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                        
                        # Extract the file
                        with open(output_path, 'wb') as output_file:
                            for block in entry.get_blocks():
                                output_file.write(block)
                                
                        # Add to extracted files set
                        self.extracted_files.add(output_path)
                        return output_path
                        
        return None
        
    def extract_file_to_user_location(self, item):
        """Extract a file to a user-specified location"""
        # Get the full path from the item data
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            QMessageBox.warning(self, "Error", "No file path found for this item")
            return
            
        # Get the file name
        file_name = item.text(0)
        
        # Ask the user for a save location
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            file_name,
            "All Files (*)"
        )
        
        if not save_path:
            return  # User cancelled
            
        try:
            # Show progress during extraction
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            self.progress_bar.setVisible(True)
            self.status_label.setText(f"Extracting: {file_name}...")
            QApplication.processEvents()  # Ensure UI updates
            
            # Extract the file
            temp_path = self.extract_file(item)
            if not temp_path:
                self.progress_bar.setVisible(False)
                self.status_label.setText(f"Failed to extract: {file_name}")
                return
                
            # Copy to the user's chosen location
            shutil.copy2(temp_path, save_path)
            
            # Hide progress bar
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Extracted: {file_name} to {save_path}")
            
            # Show success message
            QMessageBox.information(
                self,
                "Extraction Complete",
                f"File extracted to:\n{save_path}"
            )
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
            
    def copy_to_clipboard(self, text):
        """Copy text to clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.status_label.setText(f"Copied to clipboard: {text}")
        
    def open_hash_lookup(self, hash_value, service):
        """Open a hash lookup service in the default browser"""
        if service == "virustotal":
            url = f"https://www.virustotal.com/gui/file/{hash_value}"
        elif service == "hybrid":
            url = f"https://www.hybrid-analysis.com/search?query={hash_value}"
        else:
            return
            
        try:
            webbrowser.open(url)
            self.status_label.setText(f"Opened {service} lookup for {hash_value}")
        except Exception as e:
            QMessageBox.warning(self, "Error", f"Failed to open browser: {str(e)}")
            
    def show_hex_view(self, item):
        """Show hex view of the selected file"""
        file_path = self.extract_file(item)
        if file_path:
            show_hex_view_dialog(self, item.text(0), file_path)
        
    def show_image_preview(self, item):
        """Show image preview of the selected file"""
        file_path = self.extract_file(item)
        if file_path:
            show_image_preview_dialog(self, item.text(0), file_path)
        
    def show_text_preview(self, item):
        """Show text preview of the selected file"""
        file_path = self.extract_file(item)
        if file_path:
            show_text_preview_dialog(self, item.text(0), file_path)
            
    def show_nested_archive_preview(self, item):
        """Show preview for a nested archive file"""
        # Show progress bar during extraction
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setVisible(True)
        self.status_label.setText(f"Extracting nested archive: {item.text(0)}...")
        QApplication.processEvents()  # Ensure UI updates
            
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Failed to extract: {item.text(0)}")
            return
            
        try:
            # Hide progress bar before showing dialog
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Opening nested archive: {item.text(0)}")
            
            # Show archive preview dialog
            from utils.preview import show_archive_preview_dialog
            # Use the same auto_identify setting as this dialog
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
