# ruff: noqa: E722
import sys
from PyQt5.QtWidgets import (QApplication, QMessageBox)

# Try to import libarchive for archive handling
try:
    import libarchive
    LIBARCHIVE_AVAILABLE = True
except (ImportError, TypeError):
    LIBARCHIVE_AVAILABLE = False

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