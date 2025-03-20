"""
Common preview functionality used across the application.
"""
from PyQt5.QtWidgets import QMessageBox
from utils.common import read_text_file_with_fallback

def _handle_error(parent, title, error):
    """Centralized error handling for preview dialogs"""
    if hasattr(parent, 'show_error'):
        parent.show_error(title, error)
    else:
        QMessageBox.critical(parent, "Error", f"{title}: {str(error)}")

def show_text_preview_dialog(parent, file_name, file_path):
    """Show a text preview dialog for the given file path"""
    from dialogs.text import TextPreviewDialog
    
    content = read_text_file_with_fallback(file_path)
    if content is None:
        _handle_error(parent, "Preview Error", f"Failed to read file: {file_name}")
        return False
        
    try:
        TextPreviewDialog(parent, file_name, content).exec_()
        return True
    except Exception as e:
        _handle_error(parent, "Preview Error", e)
        return False

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
    from dialogs.hex import HexViewDialog
    
    try:
        HexViewDialog(parent, file_name, file_path).exec_()
        return True
    except Exception as e:
        _handle_error(parent, "Hex View Error", e)
        return False

def show_image_preview_dialog(parent, file_name, file_path):
    """Show an image preview dialog for the given file path"""
    from dialogs.image import ImagePreviewDialog
    
    try:
        ImagePreviewDialog(parent, file_name, file_path).exec_()
        return True
    except Exception as e:
        _handle_error(parent, "Image Preview Error", e)
        return False

def show_archive_preview_dialog(parent, archive_name, archive_path, group_icons):
    """Show an archive preview dialog for the given archive path"""
    from dialogs.archive import ArchivePreviewDialog
    
    try:
        ArchivePreviewDialog(parent, archive_name, archive_path, group_icons).exec_()
        return True
    except Exception as e:
        _handle_error(parent, "Archive Preview Error", e)
        return False

def show_pdf_preview_dialog(parent, file_name, file_path):
    """Show a PDF preview dialog for the given PDF file path"""
    try:
        # Import the dialog class
        from dialogs.pdf import PDFPreviewDialog
        PDFPreviewDialog(parent, file_name, file_path).exec_()
        return True
    except Exception as e:
        _handle_error(parent, "PDF Preview Error", 
                     f"Could not preview PDF: {str(e)}\n\n"
                     f"To enable PDF preview, install PyMuPDF:\n"
                     f"pip install PyMuPDF")
        return False 