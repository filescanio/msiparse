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
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"

def read_text_file_with_fallback(file_path):
    """Read a text file with UTF-8 encoding, falling back to Latin-1 if needed."""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return f.read()
    except UnicodeDecodeError:
        with open(file_path, 'r', encoding='latin-1') as f:
            return f.read()
    except Exception:
        return None

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
    """Helper class for common table functionality."""
    
    @staticmethod
    def auto_resize_columns(table_widget):
        """Automatically resize columns to fit content"""
        if not isinstance(table_widget, (QTableWidget, QTreeWidget)):
            return
            
        num_columns = table_widget.columnCount()
        for i in range(num_columns):
            table_widget.resizeColumnToContents(i)
            
        # Ensure columns don't get too wide
        max_width = table_widget.width() // num_columns
        for i in range(num_columns):
            if table_widget.columnWidth(i) > max_width:
                table_widget.setColumnWidth(i, max_width)
    
    @staticmethod
    def is_hash_value(text):
        """Check if a text value looks like a hash"""
        return (text and len(text) >= 32 and 
                all(c in "0123456789abcdefABCDEF" for c in text))
    
    @staticmethod
    def populate_table(table_widget, columns, rows):
        """Populate a QTableWidget with data"""
        if not isinstance(table_widget, QTableWidget):
            return
            
        table_widget.clear()
        table_widget.setRowCount(len(rows))
        table_widget.setColumnCount(len(columns))
        table_widget.setHorizontalHeaderLabels(columns)
        
        mono_font = QFont("Courier New", 10)
        mono_font.setFixedPitch(True)
        
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(columns):
                    item = QTableWidgetItem(str(cell_data))
                    # Apply monospace font to hash columns or hash-like values
                    if "hash" in columns[col_idx].lower() and cell_data and cell_data not in ("Error calculating hash", ""):
                        item.setFont(mono_font)
                    table_widget.setItem(row_idx, col_idx, item)
        
        TableHelper.auto_resize_columns(table_widget)
        return len(rows)

class TreeHelper:
    """Helper class for tree widget functionality."""
    
    @staticmethod
    def populate_tree_from_structure(tree_widget, structure, group_icons, parent_item=None):
        """Build a tree from a nested dictionary structure"""
        if parent_item is None:
            parent_item = tree_widget.invisibleRootItem()
            
        # Add directories first
        for key, value in structure.items():
            if key == '':  # Files at this level
                continue
                
            dir_item = QTreeWidgetItem(parent_item)
            dir_item.setText(0, key)
            dir_item.setIcon(0, group_icons.get('inode', QApplication.style().standardIcon(QApplication.style().SP_DirIcon)))
            TreeHelper.populate_tree_from_structure(tree_widget, value, group_icons, dir_item)
            
        # Add files
        if '' in structure:
            for file_name, size, full_path in structure['']:
                file_item = QTreeWidgetItem(parent_item)
                file_item.setText(0, file_name)
                file_item.setText(3, format_file_size(size))
                file_item.setText(1, "")  # Group
                file_item.setText(2, "")  # MIME type
                file_item.setText(4, "")  # SHA1 hash
                file_item.setData(0, Qt.UserRole, full_path)
                file_item.setIcon(0, group_icons.get('unknown', QApplication.style().standardIcon(QApplication.style().SP_FileIcon)))
    
    @staticmethod
    def count_items_recursive(parent_item):
        """Count all file items under the parent item"""
        count = 0
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            if child.data(0, Qt.UserRole):  # If it's a file
                count += 1
            count += TreeHelper.count_items_recursive(child)
        return count
    
    @staticmethod
    def apply_hash_font_to_tree_item(item, column_index, mono_font=None):
        """Apply monospaced font to a tree item's hash column"""
        if mono_font is None:
            mono_font = QFont("Courier")  # Use Courier instead of Courier New
            mono_font.setStyleHint(QFont.Monospace)  # Ensure monospace fallback
            mono_font.setFixedPitch(True)
        
        # Apply monospace font to all hash values, even if they're not valid hashes
        hash_text = item.text(column_index)
        if hash_text and hash_text not in ("Error calculating hash", ""):
            item.setFont(column_index, mono_font)
            return True
        return False
    
    @staticmethod
    def set_icon_for_group(item, group, group_icons):
        """Set the appropriate icon for a file based on its group"""
        item.setIcon(0, group_icons.get(group, group_icons.get('unknown', QApplication.style().standardIcon(QApplication.style().SP_FileIcon))))
        return True

class FileIdentificationHelper:
    """Helper class for file identification functionality."""
    
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
        for group, mime_patterns in FileIdentificationHelper.FILE_TYPE_GROUPS.items():
            if any(pattern in mime_type for pattern in mime_patterns):
                return group
        return 'application' if mime_type.startswith('application/') else 'unknown'
    
    @staticmethod
    def identify_file_with_magika(file_path, magika_client):
        """Identify a file using Magika"""
        try:
            from pathlib import Path
            path_obj = Path(file_path) if isinstance(file_path, str) else file_path
            
            if not path_obj.exists() or not path_obj.is_file() or not os.access(str(path_obj), os.R_OK):
                return FileIdentificationHelper.identify_by_extension(str(path_obj))
                
            if path_obj.stat().st_size == 0:
                return ('', 'Empty file')
                
            try:
                result = magika_client.identify_path(path_obj)
            except Exception:
                result = magika_client.identify_path(str(path_obj))
                
            mime_type = result.output.mime_type
            # Get group directly from Magika instead of determining it ourselves
            group = result.output.group
            return (group, mime_type)
        except Exception:
            return FileIdentificationHelper.identify_by_extension(str(file_path))
            
    @staticmethod
    def identify_by_extension(file_path):
        """Identify a file by its extension"""
        ext = os.path.splitext(file_path)[1].lower()
        
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
        
        return extension_map.get(ext, ('', f'application/octet-stream (unknown extension: {ext})'))
    
    @staticmethod
    def update_tree_item_with_file_info(item, group, mime_type, file_path, group_icons, size_text=None, hash_text=None):
        """Update a tree item with file identification information"""
        try:
            item.setText(1, group)
            item.setText(2, mime_type)
            TreeHelper.set_icon_for_group(item, group, group_icons)
            
            if size_text is not None:
                item.setText(3, size_text)
                
            if hash_text is not None:
                item.setText(4, hash_text)
                TreeHelper.apply_hash_font_to_tree_item(item, 4)
            elif item.text(4) == "" and file_path and os.path.isfile(file_path):
                sha1_hash = calculate_sha1(file_path)
                item.setText(4, sha1_hash)
                TreeHelper.apply_hash_font_to_tree_item(item, 4)
            
            if item.treeWidget():
                item.treeWidget().update()
                QApplication.processEvents()
                
            return True
        except Exception:
            return False 