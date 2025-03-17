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

# Check for archive.dll in the current directory before importing libarchive
archive_dll_path = pathlib.Path(__file__).parent / "archive.dll"
if archive_dll_path.exists():
    # Set LIBARCHIVE environment variable to the absolute path of archive.dll
    os.environ["LIBARCHIVE"] = str(archive_dll_path.absolute())
    print(f"Found archive.dll, setting LIBARCHIVE={os.environ['LIBARCHIVE']}")

# Try to import libarchive for archive handling
try:
    import libarchive
    LIBARCHIVE_AVAILABLE = True
except (ImportError, TypeError):
    LIBARCHIVE_AVAILABLE = False

def check_libarchive_support():
    """Check if libarchive is available"""
    return LIBARCHIVE_AVAILABLE

def show_libarchive_warning(parent=None):
    """Show a warning if libarchive is not available"""
    if not LIBARCHIVE_AVAILABLE:
        QMessageBox.warning(
            parent,
            "LibArchive Not Available",
            "The libarchive-c library is not installed or the required DLL/SO files are missing.\n"
            "Archive preview functionality will be disabled.\n\n"
            "To enable archive support, install libarchive-c with:\n"
            "pip install libarchive-c"
        )

def main():
    """Main function to start the application."""
    # Always create QApplication first, before importing any modules that might create QWidgets
    app = QApplication(sys.argv)
    
    # Now it's safe to import modules that might create QWidgets
    from utils.gui import MSIParseGUI
    
    # Check libarchive support once at startup
    has_archive_support = check_libarchive_support()
    
    # Create the main window
    window = MSIParseGUI(archive_support=has_archive_support)
    
    # Now it's safe to show warnings
    if not has_archive_support:
        show_libarchive_warning(window)
    
    # Show the window
    window.show()
    
    # Start the event loop
    sys.exit(app.exec_())

if __name__ == "__main__":
    main() 