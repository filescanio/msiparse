import os
import tempfile
import shutil
import webbrowser
from pathlib import Path
from PyQt5.QtWidgets import (QDialog, QVBoxLayout, QHBoxLayout, QPushButton, QLabel, 
                           QTreeWidget, QTreeWidgetItem, QMessageBox, QProgressBar,
                           QMenu, QAction, QFileDialog, QApplication)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont
import magika

# Import common utilities
from utils.common import format_file_size, calculate_sha1
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
    def __init__(self, parent, archive_name, archive_path, group_icons):
        super().__init__(parent)
        self.parent = parent
        self.archive_name = archive_name
        self.archive_path = archive_path
        self.group_icons = group_icons
        self.archive_entries = []
        self.temp_dir = None
        self.extracted_files = set()  # Track extracted files for cleanup
        
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
        
        # Create a tree widget for the archive contents
        self.contents_tree = QTreeWidget()
        self.contents_tree.setHeaderLabels(["Name", "Group", "MIME Type", "Size", "SHA1 Hash"])
        self.contents_tree.setSelectionMode(QTreeWidget.SingleSelection)
        
        # Remove monospaced font for the entire tree
        # mono_font = QFont("Courier New", 10)
        # mono_font.setFixedPitch(True)
        # self.contents_tree.setFont(mono_font)
        
        # Enable context menu for contents tree
        self.contents_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.contents_tree.customContextMenuRequested.connect(self.show_context_menu)
        
        layout.addWidget(self.contents_tree)
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close_and_cleanup)
        layout.addWidget(close_button)
        
    def close_and_cleanup(self):
        """Close the dialog and clean up temporary files"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                # Clean up all extracted files first
                for file_path in self.extracted_files:
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                        except:
                            pass
                
                # Then remove the temp directory
                shutil.rmtree(self.temp_dir)
            except:
                pass
        self.accept()
        
    def closeEvent(self, event):
        """Handle dialog close event"""
        self.close_and_cleanup()
        event.accept()
        
    def load_archive_contents(self):
        """Load the contents of the archive file"""
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
            self.auto_resize_columns()
            
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            self.progress_bar.setVisible(False)
            QMessageBox.critical(self, "Error", f"Failed to read archive: {str(e)}")
        
    def auto_identify_files(self):
        """Identify file types for all files in the archive"""
        # Disable identify button while running
        self.identify_button.setEnabled(False)
            
        # Show progress bar for identification
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, self.count_all_items())
        self.status_label.setText("Identifying file types...")
        QApplication.processEvents()  # Ensure UI updates
        
        # Process all top-level items
        self.identify_items_recursive(self.contents_tree.invisibleRootItem(), 0)
        
        # Hide progress bar when done
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files)")
        
        # Auto-resize columns after identification
        self.auto_resize_columns()
        
        # Re-enable identify button
        self.identify_button.setEnabled(True)
        
    def count_all_items(self):
        """Count all file items in the tree recursively"""
        return self.count_items_recursive(self.contents_tree.invisibleRootItem())
        
    def count_items_recursive(self, parent_item):
        """Recursively count all file items under the parent item"""
        count = 0
        
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If it's a file (has path data)
            if child.data(0, Qt.UserRole):
                count += 1
                
            # Recursively count child items (for directories)
            count += self.count_items_recursive(child)
            
        return count
        
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
                progress_count += 1
                QApplication.processEvents()  # Keep UI responsive
                
                # Identify the file
                self.identify_file(child, show_message=False)
            
            # Recursively process child items (for directories)
            progress_count = self.identify_items_recursive(child, progress_count)
            
        return progress_count
        
    def auto_resize_columns(self):
        """Automatically resize columns to fit content"""
        for i in range(5):
            self.contents_tree.resizeColumnToContents(i)
            
        # Ensure columns don't get too wide
        total_width = self.contents_tree.width()
        max_width = total_width // 5
        
        for i in range(5):
            if self.contents_tree.columnWidth(i) > max_width:
                self.contents_tree.setColumnWidth(i, max_width)
            
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
        self.build_tree(tree_structure, self.contents_tree)
        
        # Auto-resize columns
        self.auto_resize_columns()
            
    def build_tree(self, structure, parent_item):
        """Recursively build the tree from the structure dictionary"""
        # Add directories first
        for key, value in structure.items():
            if key == '':  # Files at this level
                continue
                
            # Create a directory item
            dir_item = QTreeWidgetItem(parent_item if isinstance(parent_item, QTreeWidget) else parent_item)
            dir_item.setText(0, key)
            dir_item.setIcon(0, self.group_icons['inode'])
            
            # Recursively add children
            self.build_tree(value, dir_item)
            
        # Add files
        if '' in structure:
            for file_info in structure['']:
                file_name, size, full_path = file_info
                
                # Create a file item
                file_item = QTreeWidgetItem(parent_item if isinstance(parent_item, QTreeWidget) else parent_item)
                file_item.setText(0, file_name)
                file_item.setText(3, format_file_size(size))
                # Initialize empty group, MIME type, and SHA1 hash
                file_item.setText(1, "")  # Group
                file_item.setText(2, "")  # MIME type
                file_item.setText(4, "")  # SHA1 hash
                
                # Store the full path for later use
                file_item.setData(0, Qt.UserRole, full_path)
                
                # Set a default icon
                file_item.setIcon(0, self.group_icons['unknown'])
                
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
        mime_type = item.text(2)
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
        
        # Add Extract File option
        extract_action = QAction("Extract File...", self)
        extract_action.triggered.connect(lambda: self.extract_file_to_location(item))
        context_menu.addAction(extract_action)
        
        # Extract and identify the file if needed
        if not group or not mime_type:
            identify_action = QAction("Identify File Type", self)
            identify_action.triggered.connect(lambda: self.identify_file(item))
            context_menu.addAction(identify_action)
        
        # Add Hash Lookup option if hash is available
        if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
            # Add separator
            context_menu.addSeparator()
            
            # Add Hash Lookup submenu
            hash_menu = QMenu("Lookup Hash", self)
            
            # Add FileScan.io option (first)
            fs_action = QAction("FileScan.io", self)
            fs_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "filescan"))
            hash_menu.addAction(fs_action)
            
            # Add MetaDefender Cloud option (second)
            md_action = QAction("MetaDefender Cloud", self)
            md_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "metadefender"))
            hash_menu.addAction(md_action)
            
            # Add VirusTotal option (third)
            vt_action = QAction("VirusTotal", self)
            vt_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "virustotal"))
            hash_menu.addAction(vt_action)
            
            context_menu.addMenu(hash_menu)
        
        # Add separator before copy options
        context_menu.addSeparator()
        
        # Add Copy submenu at the bottom
        copy_menu = QMenu("Copy", self)
        
        # Add Copy options to the submenu
        copy_name_action = QAction("File Name", self)
        copy_name_action.triggered.connect(lambda: self.copy_to_clipboard(file_name))
        copy_menu.addAction(copy_name_action)
        
        if mime_type:  # Only add if mime_type is not empty
            copy_type_action = QAction("MIME Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_to_clipboard(mime_type))
            copy_menu.addAction(copy_type_action)
        
        if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
            copy_hash_action = QAction("SHA1 Hash", self)
            copy_hash_action.triggered.connect(lambda: self.copy_to_clipboard(sha1_hash))
            copy_menu.addAction(copy_hash_action)
        
        context_menu.addMenu(copy_menu)
        
        # Show the context menu
        context_menu.exec_(self.contents_tree.mapToGlobal(position))
        
    def extract_file_to_location(self, item):
        """Extract a file to a user-specified location"""
        file_name = item.text(0)
        full_path = item.data(0, Qt.UserRole)
        
        if not full_path:
            return
            
        # Ask user for save location
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            file_name,
            "All Files (*)"
        )
        
        if not save_path:
            return  # User cancelled
            
        try:
            # Extract the file first
            temp_path = self.extract_file(item)
            
            if not temp_path or not os.path.exists(temp_path):
                QMessageBox.warning(self, "Error", "Failed to extract file")
                return
                
            # Copy to the destination
            shutil.copy2(temp_path, save_path)
            
            self.status_label.setText(f"Extracted: {file_name} to {save_path}")
            QMessageBox.information(
                self,
                "File Extracted",
                f"File has been extracted to:\n{save_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
        
    def copy_to_clipboard(self, text):
        """Copy the given text to the clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.status_label.setText(f"Copied to clipboard: {text[:30]}{'...' if len(text) > 30 else ''}")
        
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
                
            # Extract the file
            output_path = os.path.join(self.temp_dir, os.path.basename(full_path))
            
            # Check if parent has extract_stream_unified method
            if hasattr(self.parent, 'extract_stream_unified'):
                # Use parent's unified extraction method
                return self.parent.extract_stream_unified(
                    full_path,  # stream name
                    self.temp_dir,  # output directory
                    temp=False,
                    show_messages=False
                )
            else:
                # Fallback to direct extraction using libarchive
                with open(self.archive_path, 'rb') as archive_file:
                    with libarchive.Archive(archive_file) as archive:
                        for entry in archive:
                            if entry.pathname == full_path:
                                # Create parent directories if needed
                                os.makedirs(os.path.dirname(output_path), exist_ok=True)
                                
                                # Extract the file
                                with open(output_path, 'wb') as output_file:
                                    for block in entry.get_blocks():
                                        output_file.write(block)
                                        
                                # Add to extracted files set
                                self.extracted_files.add(output_path)
                                return output_path
                                
            # File not found in archive
            QMessageBox.warning(self, "Error", f"File not found in archive: {full_path}")
            return None
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
            return None
            
    def identify_file(self, item, show_message=True):
        """Identify the type of a file in the archive"""
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            return
            
        try:
            # Identify file type using magika with Path object
            result = self.magika_client.identify_path(Path(file_path))
            group = result.output.group
            mime_type = result.output.mime_type
            
            # Calculate SHA1 hash
            sha1_hash = calculate_sha1(file_path)
            
            # Update the item
            item.setText(1, group)
            item.setText(2, mime_type)
            item.setText(4, sha1_hash)
            
            # Set monospaced font for the hash column only
            if sha1_hash and sha1_hash != "Error calculating hash":
                mono_font = QFont("Courier New", 10)
                mono_font.setFixedPitch(True)
                item.setFont(4, mono_font)
            
            # Set icon based on group
            if group in self.group_icons:
                item.setIcon(0, self.group_icons[group])
            else:
                item.setIcon(0, self.group_icons['unknown'])
                
            if show_message:
                self.status_label.setText(f"Identified: {item.text(0)} as {mime_type}")
            
        except Exception as e:
            if show_message:
                QMessageBox.critical(self, "Error", f"Failed to identify file: {str(e)}")
            
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
            show_archive_preview_dialog(self.parent, item.text(0), file_path, self.group_icons)
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error showing archive preview: {str(e)}")

    def open_hash_lookup(self, hash_value, service):
        """Open the hash lookup in the specified service"""
        if service == "virustotal":
            url = f"https://www.virustotal.com/gui/file/{hash_value}"
            self.status_label.setText(f"Opening hash in VirusTotal: {hash_value}")
        elif service == "metadefender":
            url = f"https://metadefender.com/results/hash/{hash_value}"
            self.status_label.setText(f"Opening hash in MetaDefender Cloud: {hash_value}")
        elif service == "filescan":
            url = f"https://www.filescan.io/search-result?query={hash_value}"
            self.status_label.setText(f"Opening hash in FileScan.io: {hash_value}")
        else:
            return
            
        try:
            webbrowser.open(url)
        except Exception as e:
            QMessageBox.critical(self, "Browser Error", f"Failed to open browser: {str(e)}")
