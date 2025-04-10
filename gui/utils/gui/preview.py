"""
Preview functionality for the MSI Parser GUI
"""

import os
from PyQt5.QtWidgets import QMessageBox, QApplication, QFileDialog
from PyQt5.QtCore import Qt
from utils.common import temp_directory
from utils.gui.extraction import extract_file_to_temp
from utils.gui.helpers import show_hex_view_dialog, show_text_preview_dialog, show_image_preview_dialog

def show_preview(parent, stream_name, preview_func):
    """Show a preview of a stream using the specified preview function"""
    if not parent.msi_file_path:
        return
    
    # Get mime_type from streams tree if available
    mime_type = None
    for i in range(parent.streams_tree.topLevelItemCount()):
        item = parent.streams_tree.topLevelItem(i)
        if item.text(0) == stream_name:
            mime_type = item.text(2)
            break
            
    with temp_directory() as temp_dir:
        try:
            if file_path := extract_file_to_temp(parent, stream_name, temp_dir):
                preview_func(parent, stream_name, file_path, mime_type)
        except Exception as e:
            parent.handle_error("Preview Error", e)

def show_hex_view(parent, stream_name):
    """Show hexadecimal view of a stream"""
    show_preview(parent, stream_name, lambda p, name, path, mime: show_hex_view_dialog(p, name, path))
    
def show_text_preview(parent, stream_name):
    """Show text preview of a stream with syntax highlighting"""
    show_preview(parent, stream_name, lambda p, name, path, mime: show_text_preview_dialog(p, name, path, mime))
    
def show_image_preview(parent, stream_name):
    """Show image preview of a stream"""
    show_preview(parent, stream_name, lambda p, name, path, mime: show_image_preview_dialog(p, name, path))
    
def show_archive_preview(parent, stream_name):
    """Show an archive preview dialog for the given stream"""
    if not parent.archive_support:
        QMessageBox.information(parent, "Preview Not Available", 
                              "Archive preview is not available. Archive support is disabled.")
        return
        
    # Extract the file to a temporary location first
    file_path = None
    try:
        file_path = parent.extract_file_safe(stream_name, temp=True)
        if not file_path:
            parent.show_error("Preview Error", f"Failed to extract archive: {stream_name}", status_only=True)
            return
            
        # Call the archive preview dialog
        from utils.preview import show_archive_preview_dialog
        show_archive_preview_dialog(parent, stream_name, file_path, parent.group_icons)
        
    except Exception as e:
        parent.show_error("Preview Error", e, status_only=True)
        
def show_pdf_preview(parent, stream_name):
    """Show a PDF preview dialog for the given stream"""
    # Extract the file to a temporary location first
    file_path = None
    try:
        file_path = parent.extract_file_safe(stream_name, temp=True)
        if not file_path:
            parent.show_error("Preview Error", f"Failed to extract PDF: {stream_name}", status_only=True)
            return
            
        # Call the PDF preview dialog
        from utils.preview import show_pdf_preview_dialog
        result = show_pdf_preview_dialog(parent, stream_name, file_path)
        
        if not result:
            # If preview failed, try to open with default program
            parent.show_warning("Preview Error", 
                               "PDF preview failed. Would you like to extract and open with default viewer?",
                               status_only=True)
            if QMessageBox.question(parent, "Open with Default Viewer?", 
                                  "Would you like to extract the PDF and open with your default PDF viewer?",
                                  QMessageBox.Yes | QMessageBox.No) == QMessageBox.Yes:
                # Extract to user location and open
                output_dir = parent.get_output_directory()
                if output_dir:
                    extracted_path = parent.extract_file_safe(stream_name, output_dir, temp=False)
                    if extracted_path:
                        from PyQt5.QtGui import QDesktopServices
                        from PyQt5.QtCore import QUrl
                        QDesktopServices.openUrl(QUrl.fromLocalFile(extracted_path))
        
    except Exception as e:
        parent.show_error("Preview Error", e, status_only=True) 