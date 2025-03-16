# ruff: noqa: E722
import sys
import tempfile
import shutil
import contextlib
from PyQt5.QtWidgets import (QApplication, QMessageBox)

# Try to import libarchive for archive handling
try:
    import libarchive
    LIBARCHIVE_AVAILABLE = True
except (ImportError, TypeError):
    LIBARCHIVE_AVAILABLE = False

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

def show_text_preview_dialog(parent, file_name, file_path):
    """Show a text preview dialog for the given file path"""
    from dialogs.text import TextPreviewDialog
    
    content = read_text_file_with_fallback(file_path)
    if content is None:
        if hasattr(parent, 'show_warning'):
            parent.show_warning("Error", "Failed to read text file")
        else:
            QMessageBox.warning(parent, "Error", "Failed to read text file")
        return False
        
    text_dialog = TextPreviewDialog(parent, file_name, content)
    text_dialog.exec_()
    return True

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
    from dialogs.hex import HexViewDialog
    
    try:
        with open(file_path, 'rb') as f:
            content = f.read()
        
        hex_dialog = HexViewDialog(parent, file_name, content)
        hex_dialog.exec_()
        return True
    except Exception as e:
        if hasattr(parent, 'show_error'):
            parent.show_error("Hex View Error", e)
        else:
            QMessageBox.critical(parent, "Error", f"Error showing hex view: {str(e)}")
        return False

def show_image_preview_dialog(parent, file_name, file_path):
    """Show an image preview dialog for the given file path"""
    from dialogs.image import ImagePreviewDialog
    
    try:
        image_dialog = ImagePreviewDialog(parent, file_name, file_path)
        image_dialog.exec_()
        return True
    except Exception as e:
        if hasattr(parent, 'show_error'):
            parent.show_error("Image Preview Error", e)
        else:
            QMessageBox.critical(parent, "Error", f"Error showing image preview: {str(e)}")
        return False

def check_libarchive_support():
    """Check if libarchive is available and show a warning if not"""
    if not LIBARCHIVE_AVAILABLE:
        QMessageBox.warning(
            None,
            "LibArchive Not Available",
            "The libarchive-c library is not installed or the required DLL/SO files are missing.\n"
            "Archive preview functionality will be disabled.\n\n"
            "To enable archive support, install libarchive-c with:\n"
            "sudo apt install libarchive libarchive-dev"
        )
    return LIBARCHIVE_AVAILABLE

def main():
    # Import here to avoid circular imports
    from utils.gui import MSIParseGUI
    
    # Check libarchive support once at startup
    has_archive_support = check_libarchive_support()
    
    app = QApplication(sys.argv)
    window = MSIParseGUI(archive_support=has_archive_support)
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()