# ruff: noqa: E722
"""
MSI Parser GUI - Main Entry Point

This file contains the main entry point and setup logic for the MSI Parser GUI application.
It initializes the application, checks for dependencies, and launches the main window.
"""

import sys
import os
from PyQt5.QtWidgets import QApplication, QMessageBox
from PyQt5.QtGui import QIcon
from utils.common import get_bundle_path

# Try to import our custom 7z-based archive handler
try:
    # Ensure utils package is found (might need adjustment if structure changes)
    from utils import archive7z
    ARCHIVE_SUPPORT_AVAILABLE = archive7z.is_available()
except ImportError:
    ARCHIVE_SUPPORT_AVAILABLE = False

def check_archive_support():
    """Check if 7z archive support is available"""
    return ARCHIVE_SUPPORT_AVAILABLE

def show_archive_support_warning(parent=None):
    """Show a warning if 7z is not available"""
    if not ARCHIVE_SUPPORT_AVAILABLE:
        QMessageBox.warning(
            parent,
            "7-Zip Not Available",
            "The 7z command-line tool is not installed or not in the system path.\n"
            "Archive preview functionality will be disabled.\n\n"
            "To enable archive support, install 7-Zip and ensure it's in your system path."
        )

def main():
    """Main function to start the application."""
    # Always create QApplication first, before importing any modules that might create QWidgets
    app = QApplication(sys.argv)
    
    # Now it's safe to import modules that might create QWidgets
    # Assuming MSIParseGUI is the main window class in utils.gui
    try:
        from utils.gui.main_window import MSIParseGUI 
    except ImportError:
        print("Error: Could not import MSIParseGUI from utils.gui.main_window")
        print("Please ensure the utils/gui structure and main_window.py exist.")
        # Optionally, show a message box if QApplication is running
        if QApplication.instance():
             QMessageBox.critical(None, "Import Error", 
                                  "Could not import the main GUI class. Application cannot start.")
        sys.exit(1)

    # Check archive support once at startup
    has_archive_support = check_archive_support()
    
    # Create the main window
    window = MSIParseGUI(archive_support=has_archive_support)

    # Set the window icon
    icon_path = get_bundle_path("resources/icon.ico")
    if os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))
    else:
        print(f"Warning: Window icon not found at {icon_path}")
    
    # Now it's safe to show warnings
    if not has_archive_support:
        show_archive_support_warning(window)
    
    # Show the window
    window.show()
    
    # Start the event loop
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 