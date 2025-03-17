# ruff: noqa: E722
"""
Common utility functions used across the application.
"""
import hashlib
import tempfile
import shutil
import contextlib
import os
from PyQt5.QtWidgets import (QTableWidget, QTableWidgetItem, QTreeWidget, QTreeWidgetItem, QApplication)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont

def format_file_size(size_bytes):
    """Format file size in a human-readable format"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"

def read_text_file_with_fallback(file_path):
    """Read a text file with UTF-8 encoding, falling back to Latin-1 if needed.
    Returns the content or None if reading fails."""
    content = None
    try:
        # First try to read as UTF-8
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # If UTF-8 fails, try with Latin-1 (which should always work)
        try:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()
        except Exception:
            # If all fails, return None
            pass
    return content

def calculate_sha1(file_path):
    """Calculate SHA1 hash for a file"""
    try:
        with open(file_path, 'rb') as f:
            return hashlib.sha1(f.read()).hexdigest()
    except Exception:
        return "Error calculating hash"

@contextlib.contextmanager
def temp_directory():
    """Context manager for creating and cleaning up a temporary directory"""
    temp_dir = tempfile.mkdtemp()
    try:
        yield temp_dir
    finally:
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

class TableHelper:
    """Helper class for common table functionality used across the application."""
    
    @staticmethod
    def auto_resize_columns(table_widget):
        """Automatically resize columns to fit content for both QTableWidget and QTreeWidget"""
        if isinstance(table_widget, (QTableWidget, QTreeWidget)):
            # Get the number of columns
            if isinstance(table_widget, QTableWidget):
                num_columns = table_widget.columnCount()
            else:  # QTreeWidget
                num_columns = table_widget.columnCount()
                
            # Resize each column
            for i in range(num_columns):
                table_widget.resizeColumnToContents(i)
                
            # Ensure columns don't get too wide
            total_width = table_widget.width()
            max_width = total_width // num_columns
            
            for i in range(num_columns):
                if table_widget.columnWidth(i) > max_width:
                    table_widget.setColumnWidth(i, max_width)
    
    @staticmethod
    def create_mono_font():
        """Create a monospaced font for hash values"""
        mono_font = QFont("Courier New", 10)
        mono_font.setFixedPitch(True)
        return mono_font
    
    @staticmethod
    def is_hash_value(text):
        """Check if a text value looks like a hash (hex string of sufficient length)"""
        return (text and len(text) >= 32 and 
                all(c in "0123456789abcdefABCDEF" for c in text))
    
    @staticmethod
    def apply_hash_font(item, text, mono_font=None):
        """Apply monospaced font to an item if the text looks like a hash"""
        if mono_font is None:
            mono_font = TableHelper.create_mono_font()
            
        if TableHelper.is_hash_value(text):
            item.setFont(mono_font)
        
        return item
    
    @staticmethod
    def populate_table(table_widget, columns, rows, apply_mono_font=True):
        """Populate a QTableWidget with data
        
        Args:
            table_widget: The QTableWidget to populate
            columns: List of column names
            rows: List of rows, where each row is a list of cell values
            apply_mono_font: Whether to apply monospaced font to hash values
        """
        if not isinstance(table_widget, QTableWidget):
            return
            
        # Clear the table
        table_widget.clear()
        table_widget.setRowCount(len(rows))
        table_widget.setColumnCount(len(columns))
        table_widget.setHorizontalHeaderLabels(columns)
        
        # Create monospaced font for hash columns if needed
        mono_font = TableHelper.create_mono_font() if apply_mono_font else None
        
        # Fill the table with data
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(columns):  # Safety check
                    item = QTableWidgetItem(str(cell_data))
                    
                    # Apply monospaced font to hash values if requested
                    if apply_mono_font:
                        # Check if column name contains "hash" or if the data looks like a hash
                        column_name = columns[col_idx].lower()
                        if "hash" in column_name or TableHelper.is_hash_value(cell_data):
                            item.setFont(mono_font)
                            
                    table_widget.setItem(row_idx, col_idx, item)
        
        # Auto-resize columns
        TableHelper.auto_resize_columns(table_widget)
        
        return len(rows)

class TreeHelper:
    """Helper class for tree widget functionality used across the application."""
    
    @staticmethod
    def populate_tree_from_structure(tree_widget, structure, group_icons, parent_item=None):
        """Recursively build a tree from a nested dictionary structure
        
        Args:
            tree_widget: The QTreeWidget to populate
            structure: A nested dictionary representing the tree structure
            group_icons: Dictionary of icons for different file groups
            parent_item: The parent item to add children to (None for root)
        """
        if parent_item is None:
            # If no parent item is provided, use the tree widget's invisible root
            parent_item = tree_widget.invisibleRootItem()
            
        # Add directories first
        for key, value in structure.items():
            if key == '':  # Files at this level
                continue
                
            # Create a directory item
            dir_item = QTreeWidgetItem(parent_item)
            dir_item.setText(0, key)
            dir_item.setIcon(0, group_icons.get('inode', QApplication.style().standardIcon(QApplication.style().SP_DirIcon)))
            
            # Recursively add children
            TreeHelper.populate_tree_from_structure(tree_widget, value, group_icons, dir_item)
            
        # Add files
        if '' in structure:
            for file_info in structure['']:
                file_name, size, full_path = file_info
                
                # Create a file item
                file_item = QTreeWidgetItem(parent_item)
                file_item.setText(0, file_name)
                file_item.setText(3, format_file_size(size))
                # Initialize empty group, MIME type, and SHA1 hash
                file_item.setText(1, "")  # Group
                file_item.setText(2, "")  # MIME type
                file_item.setText(4, "")  # SHA1 hash
                
                # Store the full path for later use
                file_item.setData(0, Qt.UserRole, full_path)
                
                # Set a default icon
                file_item.setIcon(0, group_icons.get('unknown', QApplication.style().standardIcon(QApplication.style().SP_FileIcon)))
    
    @staticmethod
    def count_items_recursive(parent_item):
        """Recursively count all file items under the parent item"""
        count = 0
        
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If it's a file (has path data)
            if child.data(0, Qt.UserRole):
                count += 1
                
            # Recursively count child items (for directories)
            count += TreeHelper.count_items_recursive(child)
            
        return count
    
    @staticmethod
    def apply_hash_font_to_tree_item(item, column_index, mono_font=None):
        """Apply monospaced font to a tree item's hash column"""
        try:
            if mono_font is None:
                mono_font = TableHelper.create_mono_font()
                
            text = item.text(column_index)
            
            if TableHelper.is_hash_value(text):
                item.setFont(column_index, mono_font)
                return True
            return False
        except Exception:
            return False
    
    @staticmethod
    def set_icon_for_group(item, group, group_icons):
        """Set the appropriate icon for a file based on its group"""
        try:
            if group in group_icons:
                item.setIcon(0, group_icons[group])
            else:
                # Default to unknown icon
                item.setIcon(0, group_icons.get('unknown', QApplication.style().standardIcon(QApplication.style().SP_FileIcon)))
            return True
        except Exception:
            return False

class FileIdentificationHelper:
    """Helper class for file identification functionality used across the application."""
    
    # Common file type groups and their associated MIME types
    FILE_TYPE_GROUPS = {
        'image': ['image/'],
        'video': ['video/'],
        'audio': ['audio/'],
        'text': ['text/plain', 'text/csv', 'text/tab-separated-values'],
        'code': ['text/x-', 'text/html', 'application/json', 'application/xml', 'application/javascript'],
        'document': ['application/pdf', 'application/msword', 'application/vnd.openxmlformats-officedocument',
                    'application/vnd.ms-', 'application/rtf'],
        'archive': ['application/zip', 'application/x-rar-compressed', 'application/x-tar', 
                   'application/x-gzip', 'application/x-bzip2', 'application/x-7z-compressed',
                   'application/java-archive', 'application/x-archive'],
        'executable': ['application/x-executable', 'application/x-msi', 'application/x-ms-dos-executable',
                      'application/x-msdownload', 'application/x-ms-shortcut'],
        'font': ['font/', 'application/font-'],
    }
    
    @staticmethod
    def determine_file_group(mime_type):
        """Determine the file group based on MIME type"""
        if not mime_type:
            return 'unknown'
            
        mime_type = mime_type.lower()
        
        # Check each group
        for group, mime_patterns in FileIdentificationHelper.FILE_TYPE_GROUPS.items():
            for pattern in mime_patterns:
                if pattern in mime_type:
                    return group
                    
        # Default to application group for other application/ types
        if mime_type.startswith('application/'):
            return 'application'
            
        return 'unknown'
    
    @staticmethod
    def identify_file_with_magika(file_path, magika_client):
        """Identify a file using Magika
        
        Args:
            file_path: Path to the file to identify (str or pathlib.Path)
            magika_client: Initialized Magika client
            
        Returns:
            Tuple of (group, mime_type)
        """
        try:
            # Convert to pathlib.Path if it's a string
            from pathlib import Path
            if isinstance(file_path, str):
                path_obj = Path(file_path)
            else:
                path_obj = file_path
                
            # Check if file exists and is readable
            if not path_obj.exists() or not path_obj.is_file():
                return FileIdentificationHelper.identify_by_extension(str(path_obj))
                
            # Check if file is readable (convert to string for os.access)
            if not os.access(str(path_obj), os.R_OK):
                return FileIdentificationHelper.identify_by_extension(str(path_obj))
                
            # Get file size
            file_size = path_obj.stat().st_size
            
            # Skip empty files
            if file_size == 0:
                return ('unknown', 'Empty file')
                
            # Identify the file - Magika expects a Path object
            try:
                result = magika_client.identify_path(path_obj)
                mime_type = result.output.mime_type
            except Exception:
                # Some versions of Magika might expect string paths
                result = magika_client.identify_path(str(path_obj))
                mime_type = result.output.mime_type
            
            # Determine the group
            group = FileIdentificationHelper.determine_file_group(mime_type)
            
            return (group, mime_type)
        except Exception:
            # Fallback to extension-based identification
            return FileIdentificationHelper.identify_by_extension(str(file_path))
            
    @staticmethod
    def identify_by_extension(file_path):
        """Identify a file by its extension
        
        Args:
            file_path: Path to the file
            
        Returns:
            Tuple of (group, mime_type)
        """
        # Get the file extension
        _, ext = os.path.splitext(file_path)
        ext = ext.lower()
        
        # Common extensions and their MIME types
        extension_map = {
            '.exe': ('executable', 'application/x-msdownload'),
            '.dll': ('executable', 'application/x-msdownload'),
            '.sys': ('executable', 'application/x-msdownload'),
            '.txt': ('text', 'text/plain'),
            '.log': ('text', 'text/plain'),
            '.ini': ('text', 'text/plain'),
            '.cfg': ('text', 'text/plain'),
            '.conf': ('text', 'text/plain'),
            '.html': ('code', 'text/html'),
            '.htm': ('code', 'text/html'),
            '.xml': ('code', 'application/xml'),
            '.json': ('code', 'application/json'),
            '.js': ('code', 'application/javascript'),
            '.css': ('code', 'text/css'),
            '.py': ('code', 'text/x-python'),
            '.c': ('code', 'text/x-c'),
            '.cpp': ('code', 'text/x-c++'),
            '.h': ('code', 'text/x-c'),
            '.hpp': ('code', 'text/x-c++'),
            '.java': ('code', 'text/x-java'),
            '.cs': ('code', 'text/x-csharp'),
            '.php': ('code', 'text/x-php'),
            '.rb': ('code', 'text/x-ruby'),
            '.go': ('code', 'text/x-go'),
            '.rs': ('code', 'text/x-rust'),
            '.jpg': ('image', 'image/jpeg'),
            '.jpeg': ('image', 'image/jpeg'),
            '.png': ('image', 'image/png'),
            '.gif': ('image', 'image/gif'),
            '.bmp': ('image', 'image/bmp'),
            '.ico': ('image', 'image/x-icon'),
            '.svg': ('image', 'image/svg+xml'),
            '.mp3': ('audio', 'audio/mpeg'),
            '.wav': ('audio', 'audio/wav'),
            '.ogg': ('audio', 'audio/ogg'),
            '.mp4': ('video', 'video/mp4'),
            '.avi': ('video', 'video/x-msvideo'),
            '.mov': ('video', 'video/quicktime'),
            '.wmv': ('video', 'video/x-ms-wmv'),
            '.pdf': ('document', 'application/pdf'),
            '.doc': ('document', 'application/msword'),
            '.docx': ('document', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'),
            '.xls': ('document', 'application/vnd.ms-excel'),
            '.xlsx': ('document', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'),
            '.ppt': ('document', 'application/vnd.ms-powerpoint'),
            '.pptx': ('document', 'application/vnd.openxmlformats-officedocument.presentationml.presentation'),
            '.zip': ('archive', 'application/zip'),
            '.rar': ('archive', 'application/x-rar-compressed'),
            '.7z': ('archive', 'application/x-7z-compressed'),
            '.tar': ('archive', 'application/x-tar'),
            '.gz': ('archive', 'application/gzip'),
            '.bz2': ('archive', 'application/x-bzip2'),
            '.cab': ('archive', 'application/vnd.ms-cab-compressed'),
            '.msi': ('executable', 'application/x-msi'),
            '.ttf': ('font', 'font/ttf'),
            '.otf': ('font', 'font/otf'),
            '.woff': ('font', 'font/woff'),
            '.woff2': ('font', 'font/woff2'),
            '.sfx': ('audio', 'audio/x-sfx'),
        }
        
        # Check if the extension is in our map
        if ext in extension_map:
            group, mime_type = extension_map[ext]
            return (group, mime_type)
            
        # Default to unknown
        return ('unknown', f'application/octet-stream (unknown extension: {ext})')
    
    @staticmethod
    def update_tree_item_with_file_info(item, group, mime_type, file_path, group_icons, size_text=None, hash_text=None):
        """Update a tree item with file identification information
        
        Args:
            item: The QTreeWidgetItem to update
            group: The file group
            mime_type: The MIME type
            file_path: Path to the file (for hash calculation)
            group_icons: Dictionary of icons for different file groups
            size_text: Optional custom text for size column
            hash_text: Optional custom text for hash column
        """
        try:
            # Update the item with group and MIME type
            item.setText(1, group)
            item.setText(2, mime_type)
            
            # Set the appropriate icon
            TreeHelper.set_icon_for_group(item, group, group_icons)
            
            # Set custom size text if provided
            if size_text is not None:
                item.setText(3, size_text)
                
            # Set custom hash text if provided
            if hash_text is not None:
                item.setText(4, hash_text)
                # Apply monospaced font to hash
                TreeHelper.apply_hash_font_to_tree_item(item, 4)
            # Calculate and set SHA1 hash if file exists and hash is empty and no custom hash text
            elif item.text(4) == "" and file_path and os.path.isfile(file_path):
                sha1_hash = calculate_sha1(file_path)
                item.setText(4, sha1_hash)
                
                # Apply monospaced font to hash
                TreeHelper.apply_hash_font_to_tree_item(item, 4)
            
            # Force the tree widget to update
            if item.treeWidget():
                item.treeWidget().update()
                QApplication.processEvents()  # Ensure UI updates
                
            return True
        except Exception:
            return False 