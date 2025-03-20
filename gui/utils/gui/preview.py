"""
Preview functionality for the MSI Parser GUI
"""

from PyQt5.QtWidgets import QMessageBox
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
    """Show archive preview of a stream"""
    if not parent.archive_support:
        QMessageBox.warning(
            parent,
            "Archive Support Disabled",
            "Archive preview functionality is disabled because the 7z command-line tool is not available.\n\n"
            "To enable archive support, install 7-Zip and ensure it is in your system path."
        )
        return
        
    with parent.status_progress(f"Extracting archive: {stream_name}..."):
        def show_archive(parent, name, path, mime_type):
            from dialogs.archive import ArchivePreviewDialog
            parent.show_status(f"Opening archive preview: {name}")
            ArchivePreviewDialog(parent, name, path, parent.group_icons).exec_()
            
        show_preview(parent, stream_name, show_archive) 