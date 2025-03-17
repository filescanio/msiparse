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

class PreviewHelper:
    """Helper class for preview functionality used across the application."""
    
    @staticmethod
    def show_preview_for_file(parent, file_name, file_path, file_group, group_icons=None, auto_identify=False):
        """Show an appropriate preview dialog based on the file group
        
        Args:
            parent: The parent widget
            file_name: The name of the file
            file_path: The path to the file
            file_group: The group of the file (image, text, code, document, archive)
            group_icons: Dictionary of icons for different file groups (needed for archive preview)
            auto_identify: Whether to auto-identify files in archives (default: False)
            
        Returns:
            True if preview was shown successfully, False otherwise
        """
        if file_group == "image":
            return show_image_preview_dialog(parent, file_name, file_path)
        elif file_group in ["text", "code", "document"]:
            return show_text_preview_dialog(parent, file_name, file_path)
        elif file_group == "archive" and group_icons is not None:
            return show_archive_preview_dialog(parent, file_name, file_path, group_icons, auto_identify)
        else:
            # Default to hex view for other file types
            return show_hex_view_dialog(parent, file_name, file_path) 