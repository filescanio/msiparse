"""
Utility modules for the MSI Parser GUI application.
"""
from utils.common import (
    format_file_size,
    read_text_file_with_fallback,
    calculate_sha1,
    temp_directory
)

# Import preview functions only when needed to avoid PyQt5 import issues
from utils.preview import (
    show_text_preview_dialog,
    show_hex_view_dialog,
    show_image_preview_dialog,
    show_archive_preview_dialog
)
