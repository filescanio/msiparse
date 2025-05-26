"""
Helper functions for the MSI Parser GUI
"""

import os
import contextlib
from PyQt5.QtWidgets import QMessageBox
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import QApplication
from utils.subprocess_utils import run_subprocess
from PyQt5.QtWidgets import QWidget, QTreeWidget, QHeaderView
from PyQt5.QtGui import QFont

def show_text_preview_dialog(parent, file_name, file_path, mime_type=None):
    """Show a text preview dialog for the given file path with optional MIME type
    
    Args:
        parent: Parent widget
        file_name: Name of the file to display
        file_path: Path to the file to read
        mime_type: Optional MIME type from Magika identification
    """
    from dialogs.text import TextPreviewDialog
    
    content = None
    try:
        # First try to read as UTF-8
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
    except UnicodeDecodeError:
        # If UTF-8 fails, try with Latin-1 (which should always work)
        try:
            with open(file_path, 'r', encoding='latin-1') as f:
                content = f.read()
        except Exception:
            # If all fails, return None
            pass
            
    if content is None:
        if hasattr(parent, 'show_warning'):
            parent.show_warning("Error", "Failed to read text file")
        else:
            QMessageBox.warning(parent, "Error", "Failed to read text file")
        return False
        
    text_dialog = TextPreviewDialog(parent, file_name, content, mime_type)
    text_dialog.exec_()
    return True

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
    from dialogs.hex import HexViewDialog
    
    try:
        with open(file_path, 'rb') as f:
            content = f.read()
        
        hex_dialog = HexViewDialog(parent, file_name, content)
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

def get_group_icons():
    """Initialize icons for different file groups"""
    group_icons = {
        'video': QIcon.fromTheme('video-x-generic'),
        'unknown': QIcon.fromTheme('unknown', QIcon.fromTheme('dialog-question')),
        'archive': QIcon.fromTheme('package-x-generic', QIcon.fromTheme('application-x-archive')),
        'image': QIcon.fromTheme('image-x-generic'),
        'code': QIcon.fromTheme('text-x-script', QIcon.fromTheme('text-x-source')),
        'document': QIcon.fromTheme('x-office-document', QIcon.fromTheme('application-pdf')),
        'text': QIcon.fromTheme('text-x-generic', QIcon.fromTheme('text-plain')),
        'audio': QIcon.fromTheme('audio-x-generic'),
        'executable': QIcon.fromTheme('application-x-executable', QIcon.fromTheme('system-run')),
        'application': QIcon.fromTheme('application-x-object', QIcon.fromTheme('applications-other')),
        'inode': QIcon.fromTheme('inode-directory', QIcon.fromTheme('folder')),
        'font': QIcon.fromTheme('font-x-generic'),
        'undefined': QIcon.fromTheme('dialog-question', QIcon.fromTheme('unknown'))
    }
    
    # Fallback icons if theme icons are not available
    if not has_theme_icons():
        # Use system standard icons as fallbacks with more variety
        group_icons = {
            'video': QApplication.style().standardIcon(QApplication.style().SP_MediaPlay),
            'unknown': QApplication.style().standardIcon(QApplication.style().SP_MessageBoxQuestion),
            'archive': QApplication.style().standardIcon(QApplication.style().SP_DirClosedIcon),
            'image': QApplication.style().standardIcon(QApplication.style().SP_DesktopIcon),
            'code': QApplication.style().standardIcon(QApplication.style().SP_FileDialogDetailedView),
            'document': QApplication.style().standardIcon(QApplication.style().SP_FileDialogInfoView),
            'text': QApplication.style().standardIcon(QApplication.style().SP_FileIcon),
            'audio': QApplication.style().standardIcon(QApplication.style().SP_MediaVolume),
            'executable': QApplication.style().standardIcon(QApplication.style().SP_ComputerIcon),
            'application': QApplication.style().standardIcon(QApplication.style().SP_DriveFDIcon),
            'inode': QApplication.style().standardIcon(QApplication.style().SP_DirIcon),
            'font': QApplication.style().standardIcon(QApplication.style().SP_DirLinkIcon),
            'undefined': QApplication.style().standardIcon(QApplication.style().SP_MessageBoxQuestion)
        }
    
    return group_icons

def has_theme_icons():
    """Check if theme icons are available"""
    # Try to get a common theme icon
    test_icon = QIcon.fromTheme('document-new')
    return not test_icon.isNull()

def find_msiparse_executable():
    """Find the msiparse executable in common locations"""
    # Try current directory first
    if os.path.exists("./msiparse") and os.access("./msiparse", os.X_OK):
        return "./msiparse"
    elif os.path.exists("./msiparse.exe") and os.access("./msiparse.exe", os.X_OK):
        return "./msiparse.exe"
    
    # Try target directory
    if os.path.exists("./target/release/msiparse") and os.access("./target/release/msiparse", os.X_OK):
        return "./target/release/msiparse"
    elif os.path.exists("./target/release/msiparse.exe") and os.access("./target/release/msiparse.exe", os.X_OK):
        return "./target/release/msiparse.exe"
        
    # Return just the name and hope it's in PATH
    return "msiparse"

def get_msi_tables_descriptions():
    """Return descriptions for common MSI tables"""
    return {
        "_Streams": "Contains embedded data streams, often used for storing binary blobs such as custom actions, DLLs, and other resources.",
        "_Storages": "Holds different storage sections inside the MSI file, which can contain nested data.",
        "_StringData": "Stores string values used in the MSI database, including paths, registry keys, and commands.",
        "_StringPool": "A pool of string values used across different tables in the MSI database.",
        "_Tables": "Defines the structure of the MSI database, listing all available tables.",
        "_SummaryInformation": "Contains metadata about the MSI file, such as the author, timestamps, and security details.",
        "Binary": "Stores embedded executables, DLLs, VBScript, JScript, or other payloads used for custom actions.",
        "CustomAction": "Defines actions that can execute DLLs, scripts, or commands during installation.",
        "Property": "Stores global MSI properties, including installation paths, feature names, and custom parameters.",
        "FeatureComponents": "Maps features to their associated components, helping to reconstruct the installation structure.",
        "File": "Lists all files included in the MSI, along with their names, locations, and sizes.",
        "Registry": "Contains registry modifications that the MSI will apply upon installation.",
        "Shortcut": "Lists shortcuts created during installation, useful for tracking post-install behavior.",
        "MsiPatchSequence": "Defines patch dependencies and upgrade behaviors for MSI patches.",
        "InstallExecuteSequence": "Defines the order of execution for installation steps in silent or full installs.",
        "InstallUISequence": "Defines the order of execution for UI-related installation steps.",
        "ActionText": "Contains user-facing messages displayed during installation.",
        "Component": "Represents an atomic unit of installation, mapping files, registry keys, and other resources.",
        "Feature": "Represents logical features that users can select during installation.",
        "Media": "Lists media sources such as CAB files or external files used during installation.",
        "Directory": "Defines directory structures and where files will be installed.",
        "Upgrade": "Contains upgrade information for managing version control of installed applications.",
        "ServiceInstall": "Defines Windows services that will be installed and their configurations.",
        "ServiceControl": "Specifies service start, stop, delete, or other control actions during installation.",
        "Environment": "Modifies environment variables upon installation.",
        "Error": "Defines error messages displayed by the MSI installer.",
        "Condition": "Specifies conditions that determine whether certain components or actions will execute.",
        "IniFile": "Handles modifications to INI configuration files.",
        "ODBCDataSource": "Configures ODBC data sources for database connectivity.",
        "ProgId": "Registers COM ProgIDs (Programmatic Identifiers) in the Windows registry.",
        "TextStyle": "Defines fonts and text styles used in the installation UI.",
        "UIText": "Stores text strings used in the UI, such as button labels and prompts.",
        "User": "Defines user accounts to be created or modified during installation.",
        "ISSetupFiles": "Used in InstallShield-created MSIs to store additional setup files.",
        "ComponentQualifier": "Links components with additional qualification data for conditional installs.",
    }

def get_status_messages():
    """Return common status messages"""
    return {
        'ready': "Ready",
        'extract_complete': "Extraction completed",
        'command_complete': "Command completed successfully",
        'running_command': "Running command...",
    }

@contextlib.contextmanager
def status_progress(parent, message, show_progress=True, indeterminate=True):
    """Context manager for showing status message with optional progress bar"""
    parent.statusBar().showMessage(message)
    QApplication.processEvents()  # Keep UI responsive
    try:
        yield  # Allows code inside `with` to run
    finally:
        parent.statusBar().showMessage(parent.STATUS_MESSAGES['ready'])

def run_command_safe(parent, command, success_message=None):
    """Run a command with proper error handling and progress indication"""
    try:
        with status_progress(parent, parent.STATUS_MESSAGES['running_command']):
            result = run_subprocess(command, capture_output=True, text=True, check=True)
            if success_message:
                parent.show_status(success_message)
            return result.stdout
    except Exception as e:
        parent.show_error("Command Error", e)
        return None

def apply_scaling_to_dialog(dialog_widget, scale_factor, original_fonts_dict):
    """
    Apply font scaling to all relevant widgets within a dialog.

    Args:
        dialog_widget: The dialog (QWidget) whose children will be scaled.
        scale_factor: The factor by which to scale the fonts (e.g., 1.0, 1.1).
        original_fonts_dict: A dictionary to store/retrieve original QFont objects for widgets.
    """
    if not hasattr(dialog_widget, 'findChildren'):
        return # Not a QWidget or similar

    # Ensure all widgets have their original fonts stored
    for widget in dialog_widget.findChildren(QWidget):
        widget_id = id(widget)
        if widget_id not in original_fonts_dict:
            original_fonts_dict[widget_id] = widget.font()

    all_widgets = dialog_widget.findChildren(QWidget)
    for widget in all_widgets:
        widget_id = id(widget)
        # Retrieve the original font; if not found, use current and store it (should not happen if pre-populated)
        original_font = original_fonts_dict.get(widget_id, widget.font())
        if widget_id not in original_fonts_dict : 
             original_fonts_dict[widget_id] = original_font


        scaled_font = QFont(original_font) # Create a new font object from the original
        
        original_point_size = original_font.pointSize()
        if original_point_size <= 0: 
            # If point size isn't set (e.g., -1), use application's base font size
            app_font_size = QApplication.font().pointSize()
            original_point_size = app_font_size if app_font_size > 0 else 10 # Default to 10 if app_font_size is also invalid

        new_size = int(original_point_size * scale_factor)
        if new_size <= 0:
            new_size = 1 # Minimum font size
        
        scaled_font.setPointSize(new_size)
        widget.setFont(scaled_font)

        # Special handling for QTreeWidget headers
        if isinstance(widget, QTreeWidget):
            header = widget.header()
            if header:
                header_id = id(header)
                original_header_font = original_fonts_dict.get(header_id, header.font())
                if header_id not in original_fonts_dict:
                     original_fonts_dict[header_id] = original_header_font
                
                scaled_header_font = QFont(original_header_font)
                original_header_point_size = original_header_font.pointSize()

                if original_header_point_size <= 0:
                    app_font_size = QApplication.font().pointSize()
                    original_header_point_size = app_font_size if app_font_size > 0 else 10

                new_header_size = int(original_header_point_size * scale_factor)
                if new_header_size <= 0:
                    new_header_size = 1
                
                scaled_header_font.setPointSize(new_header_size)
                header.setFont(scaled_header_font)
                # Optionally, resize sections if needed, though this might be disruptive
                # header.resizeSections(QHeaderView.ResizeToContents)

    # Request layout recalculation and repaint for the dialog
    dialog_widget.updateGeometry()
    dialog_widget.update() 