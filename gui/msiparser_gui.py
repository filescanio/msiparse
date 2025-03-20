#!/usr/bin/env python3
"""
MSI Parser GUI - Main Implementation

This file contains the main implementation for the MSI Parser GUI application.
It can be run directly as a script:
    python msiparser_gui.py

Or as a module (recommended, using the __main__.py entry point):
    python -m msiparser_gui
"""

# ruff: noqa: E722
import sys
import os
import pathlib
from PyQt5.QtWidgets import (QApplication, QMessageBox)

# Try to import our custom 7z-based archive handler
try:
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
    from utils.gui import MSIParseGUI
    
    # Check archive support once at startup
    has_archive_support = check_archive_support()
    
    # Create the main window
    window = MSIParseGUI(archive_support=has_archive_support)
    
    # Now it's safe to show warnings
    if not has_archive_support:
        show_archive_support_warning(window)
    
    # Show the window
    window.show()
    
    # Start the event loop
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 