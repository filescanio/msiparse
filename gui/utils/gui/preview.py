"""
Preview functionality for the MSI Parser GUI
"""

from PyQt5.QtWidgets import QMessageBox
from utils.common import temp_directory
from utils.gui.extraction import extract_file_to_temp
from utils.gui.helpers import show_hex_view_dialog, show_text_preview_dialog, show_image_preview_dialog

def show_preview(parent, stream_name, preview_func):
    """Base method for showing previews with common functionality"""
    if not parent.msi_file_path:
        return
            
    # Use context manager for temporary directory
    with temp_directory() as temp_dir:
        try:
            # Extract the file
            file_path = extract_file_to_temp(parent, stream_name, temp_dir)
            
            if file_path:
                # Call the specific preview function
                preview_func(parent, stream_name, file_path)
                
        except Exception as e:
            parent.handle_error("Preview Error", e)

def show_hex_view(parent, stream_name):
    """Show hex view of the stream content"""
    show_preview(parent, stream_name, show_hex_view_dialog)
    
def show_text_preview(parent, stream_name):
    """Show text preview for text files"""
    show_preview(parent, stream_name, show_text_preview_dialog)
    
def show_image_preview(parent, stream_name):
    """Show image preview for supported image formats"""
    show_preview(parent, stream_name, show_image_preview_dialog)
    
def show_archive_preview(parent, stream_name):
    """Show archive preview for archive files"""
    # Check if archive support is enabled
    if not parent.archive_support:
        QMessageBox.warning(
            parent,
            "Archive Support Disabled",
            "Archive preview functionality is disabled because libarchive-c is not available.\n\n"
            "To enable archive support, install libarchive-c with:\n"
            "pip install libarchive-c"
        )
        return
        
    with parent.status_progress(f"Extracting archive: {stream_name}..."):
        def show_archive(parent, name, path):
            from dialogs.archive import ArchivePreviewDialog
            
            parent.show_status(f"Opening archive preview: {name}")
            archive_dialog = ArchivePreviewDialog(parent, name, path, parent.group_icons)
            archive_dialog.exec_()
            
        show_preview(
            parent,
            stream_name,
            show_archive
        ) 