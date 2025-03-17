"""
Common preview functionality used across the application.
"""
from PyQt5.QtWidgets import QMessageBox
from utils.common import read_text_file_with_fallback

def show_text_preview_dialog(parent, file_name, file_path):
    """Show a text preview dialog for the given file path"""
    from dialogs.text import TextPreviewDialog
    
    # Read the file content
    content = read_text_file_with_fallback(file_path)
    
    if content is None:
        if hasattr(parent, 'show_error'):
            parent.show_error("Preview Error", f"Failed to read file: {file_name}")
        else:
            QMessageBox.critical(parent, "Error", f"Failed to read file: {file_name}")
        return False
        
    try:
        # Create and show the dialog
        text_dialog = TextPreviewDialog(parent, file_name, content)
        text_dialog.exec_()
        return True
    except Exception as e:
        if hasattr(parent, 'show_error'):
            parent.show_error("Preview Error", e)
        else:
            QMessageBox.critical(parent, "Error", f"Error showing text preview: {str(e)}")
        return False

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
    from dialogs.hex import HexViewDialog
    
    try:
        hex_dialog = HexViewDialog(parent, file_name, file_path)
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

def show_archive_preview_dialog(parent, archive_name, archive_path, group_icons, auto_identify=False):
    """Show an archive preview dialog for the given archive path"""
    from dialogs.archive import ArchivePreviewDialog
    
    try:
        archive_dialog = ArchivePreviewDialog(parent, archive_name, archive_path, group_icons, auto_identify)
        archive_dialog.exec_()
        return True
    except Exception as e:
        if hasattr(parent, 'show_error'):
            parent.show_error("Archive Preview Error", e)
        else:
            QMessageBox.critical(parent, "Error", f"Error showing archive preview: {str(e)}")
        return False 