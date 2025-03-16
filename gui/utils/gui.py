import os
import json
import subprocess
import tempfile
import shutil
import contextlib
import webbrowser

from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QTabWidget, QTextEdit, 
                            QTreeWidget, QTreeWidgetItem, QMessageBox, QProgressBar,
                            QSplitter, QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
                            QToolButton, QListWidgetItem, QMenu, QAction, QApplication)
from PyQt5.QtCore import Qt, QPoint
from PyQt5.QtGui import QIcon

# Import custom modules
from threads.command import CommandThread
from threads.identifystreams import IdentifyStreamsThread
from dialogs.archive import ArchivePreviewDialog

# Import helper functions
from dialogs.text import TextPreviewDialog
from dialogs.hex import HexViewDialog
from dialogs.image import ImagePreviewDialog

# Helper functions that were previously imported from main
@contextlib.contextmanager
def temp_directory():
    """Context manager for creating and cleaning up a temporary directory"""
    temp_dir = tempfile.mkdtemp()
    try:
        yield temp_dir
    finally:
        try:
            shutil.rmtree(temp_dir)
        except:
            pass

def format_file_size(size_bytes):
    """Format file size in a human-readable format"""
    if size_bytes < 1024:
        return f"{size_bytes} B"
    elif size_bytes < 1024 * 1024:
        return f"{size_bytes / 1024:.1f} KB"
    elif size_bytes < 1024 * 1024 * 1024:
        return f"{size_bytes / (1024 * 1024):.1f} MB"
    else:
        return f"{size_bytes / (1024 * 1024 * 1024):.1f} GB"

def show_text_preview_dialog(parent, file_name, file_path):
    """Show a text preview dialog for the given file path"""
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
        
    text_dialog = TextPreviewDialog(parent, file_name, content)
    text_dialog.exec_()
    return True

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
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

class MSIParseGUI(QMainWindow):
    def __init__(self, archive_support=True):
        super().__init__()
        self.msi_file_path = None
        self.output_dir = None  # Will be set when needed
        self.last_output_dir = None  # Remember the last chosen directory
        self.msiparse_path = self.find_msiparse_executable()
        self.active_threads = []  # Keep track of active threads
        self.tables_data = None  # Store the tables data
        self.streams_data = []  # Store stream names
        self.archive_support = archive_support  # Store archive support status
        
        # Initialize icons for different file groups
        self.group_icons = {
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
        if self.has_theme_icons():
            # Theme icons are available, no need for fallbacks
            pass
        else:
            # Use system standard icons as fallbacks with more variety
            self.group_icons = {
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
        
        # Table descriptions for common MSI tables
        self.msi_tables = {
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
        
        # Status message constants
        self.STATUS_MESSAGES = {
            'ready': "Ready",
            'extract_complete': "Extraction completed",
            'command_complete': "Command completed successfully",
            'running_command': "Running command...",
        }
        
        self.init_ui()
        
    def closeEvent(self, event):
        """Clean up threads when the application is closing"""
        # Stop all active threads
        for thread in self.active_threads:
            thread.stop()
            thread.wait(100)  # Wait a bit for thread to finish
            
        # Accept the close event
        event.accept()
        
    def find_msiparse_executable(self):
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
        
    def init_ui(self):
        self.setWindowTitle("MSI Parser GUI")
        self.setGeometry(100, 100, 800, 600)
        
        # Main layout
        main_widget = QWidget()
        main_layout = QVBoxLayout()
        main_widget.setLayout(main_layout)
        self.setCentralWidget(main_widget)
        
        # File selection area
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No MSI file selected")
        self.browse_button = QPushButton("Browse MSI File")
        self.browse_button.clicked.connect(self.browse_msi_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_button)
        main_layout.addLayout(file_layout)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setVisible(False)
        main_layout.addWidget(self.progress_bar)
        
        # Tab widget for different functions
        self.tabs = QTabWidget()
        main_layout.addWidget(self.tabs)
        
        # Metadata tab
        self.metadata_tab = QWidget()
        metadata_layout = QVBoxLayout()
        self.metadata_tab.setLayout(metadata_layout)
        
        self.metadata_text = QTextEdit()
        self.metadata_text.setReadOnly(True)
        metadata_layout.addWidget(self.metadata_text)
        
        self.tabs.addTab(self.metadata_tab, "Metadata")
        
        # Streams tab
        self.streams_tab = QWidget()
        streams_layout = QVBoxLayout()
        self.streams_tab.setLayout(streams_layout)
        
        streams_button_layout = QHBoxLayout()
        
        # Add Identify Streams button
        self.identify_streams_button = QPushButton("Identify Streams")
        self.identify_streams_button.clicked.connect(self.identify_streams)
        streams_button_layout.addWidget(self.identify_streams_button)
        
        self.extract_all_button = QPushButton("Extract All Streams")
        self.extract_all_button.clicked.connect(self.extract_all_streams)
        streams_button_layout.addWidget(self.extract_all_button)
        
        # Move Extract Stream button to top row
        self.extract_stream_button = QPushButton("Extract Selected Streams")
        self.extract_stream_button.clicked.connect(self.extract_stream)
        streams_button_layout.addWidget(self.extract_stream_button)
        
        streams_layout.addLayout(streams_button_layout)
        
        # Update streams tree to have four columns
        self.streams_tree = QTreeWidget()
        self.streams_tree.setHeaderLabels(["Stream Name", "Group", "MIME Type", "File Size", "SHA1 Hash"])
        self.streams_tree.itemClicked.connect(self.stream_selected)
        self.streams_tree.setSelectionMode(QTreeWidget.ExtendedSelection)  # Allow multiple selection
        
        # Enable context menu for streams tree
        self.streams_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.streams_tree.customContextMenuRequested.connect(self.show_streams_context_menu)
        
        # Enable sorting but keep original order initially
        self.streams_tree.setSortingEnabled(True)  # Enable sorting
        self.streams_tree.header().setSortIndicator(-1, Qt.AscendingOrder)  # No initial sorting
        
        # Store original order
        self.original_order = True
        
        # Connect to sort indicator changed signal
        self.streams_tree.header().sortIndicatorChanged.connect(self.on_sort_indicator_changed)
        
        streams_layout.addWidget(self.streams_tree)
        
        self.tabs.addTab(self.streams_tab, "Streams")
        
        # Tables tab - redesigned with split view
        self.tables_tab = QWidget()
        tables_layout = QVBoxLayout()
        self.tables_tab.setLayout(tables_layout)
        
        # Add buttons for table export
        tables_button_layout = QHBoxLayout()
        
        self.export_selected_table_button = QPushButton("Export Selected Table")
        self.export_selected_table_button.clicked.connect(self.export_selected_table)
        self.export_selected_table_button.setEnabled(False)  # Disabled until a table is selected
        tables_button_layout.addWidget(self.export_selected_table_button)
        
        self.export_all_tables_button = QPushButton("Export All Tables")
        self.export_all_tables_button.clicked.connect(self.export_all_tables)
        self.export_all_tables_button.setEnabled(False)  # Disabled until tables are loaded
        tables_button_layout.addWidget(self.export_all_tables_button)
        
        tables_layout.addLayout(tables_button_layout)
        
        # Create a splitter for the tables view
        tables_splitter = QSplitter(Qt.Horizontal)
        tables_layout.addWidget(tables_splitter)
        
        # Left side - Table list with info buttons
        table_list_widget = QWidget()
        table_list_layout = QVBoxLayout()
        table_list_widget.setLayout(table_list_layout)
        
        self.table_list = QListWidget()
        self.table_list.setMaximumWidth(200)  # Limit width of the table list
        self.table_list.currentItemChanged.connect(self.table_selected)
        table_list_layout.addWidget(self.table_list)
        
        tables_splitter.addWidget(table_list_widget)
        
        # Right side - Table content
        self.table_content = QTableWidget()
        self.table_content.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tables_splitter.addWidget(self.table_content)
        
        # Set initial splitter sizes
        tables_splitter.setSizes([200, 600])
        
        self.tabs.addTab(self.tables_tab, "Tables")
        
        # Status bar
        self.statusBar().showMessage("Ready")
        
        # Disable buttons until file is selected
        self.update_button_states()
        
    def update_button_states(self):
        has_file = self.msi_file_path is not None
        has_selected_streams = len(self.streams_tree.selectedItems()) > 0
        has_tables = self.tables_data is not None and len(self.tables_data) > 0
        has_selected_table = self.table_list.currentItem() is not None
        
        # Check if any selected stream has a hash
        has_hash = False
        for item in self.streams_tree.selectedItems():
            if item.text(4) and item.text(4) != "Error calculating hash" and item.text(4) != "":
                has_hash = True
                break
        
        self.identify_streams_button.setEnabled(has_file)
        self.extract_all_button.setEnabled(has_file)
        self.extract_stream_button.setEnabled(has_file and has_selected_streams)
        self.export_selected_table_button.setEnabled(has_file and has_selected_table)
        self.export_all_tables_button.setEnabled(has_file and has_tables)
        
    def browse_msi_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select MSI File", "", "MSI Files (*.msi);;All Files (*)"
        )
        if file_path:
            self.msi_file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.update_button_states()
            self.statusBar().showMessage(f"Selected MSI file: {file_path}")
            
            # Auto-run metadata, streams, and tables when file is selected
            self.get_metadata()
            self.list_streams()
            self.list_tables()
            
    def get_output_directory(self):
        """Prompt for output directory if needed, using last directory as default"""
        start_dir = self.last_output_dir if self.last_output_dir else ""
        dir_path = QFileDialog.getExistingDirectory(
            self, "Select Output Directory", start_dir
        )
        if dir_path:
            self.last_output_dir = dir_path  # Remember this directory
            return dir_path
        return None
            
    def stream_selected(self, item):
        # Update button states when selection changes
        self.update_button_states()
        
    def identify_streams(self):
        """Identify the file types of all streams using Magika"""
        if not self.msi_file_path or not self.streams_data:
            return
            
        # Set progress bar to show progress
        self.progress_bar.setRange(0, len(self.streams_data))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage("Identifying stream file types...")
        
        # Disable identify button while running
        self.identify_streams_button.setEnabled(False)
        
        # Create and start the identification thread
        self.identify_thread = IdentifyStreamsThread(
            self.msiparse_path,
            self.msi_file_path,
            self.streams_data
        )
        self.active_threads.append(self.identify_thread)
        
        # Connect signals
        self.identify_thread.progress_updated.connect(self.update_identify_progress)
        self.identify_thread.stream_identified.connect(self.update_stream_file_type)
        self.identify_thread.finished_successfully.connect(self.identify_streams_finished)
        self.identify_thread.error_occurred.connect(lambda msg: self.handle_error("Identification Error", msg))
        self.identify_thread.finished.connect(lambda: self.cleanup_thread(self.identify_thread))
        
        # Start the thread
        self.identify_thread.start()
        
    def update_identify_progress(self, current, total):
        """Update the progress bar during stream identification"""
        self.progress_bar.setValue(current)
        self.statusBar().showMessage(f"Identifying stream types: {current}/{total}")
        
    def update_stream_file_type(self, stream_name, group, mime_type, file_size, sha1_hash):
        """Update the group, MIME type, size, and SHA1 hash for a stream in the tree"""
        # Temporarily disable sorting while updating
        was_sorting_enabled = self.streams_tree.isSortingEnabled()
        self.streams_tree.setSortingEnabled(False)
        
        # Find the item for this stream
        for i in range(self.streams_tree.topLevelItemCount()):
            item = self.streams_tree.topLevelItem(i)
            if item.text(0) == stream_name:
                item.setText(1, group)
                item.setText(2, mime_type)
                item.setText(3, file_size)
                item.setText(4, sha1_hash)
                
                # Set icon based on group
                self.set_icon_for_group(item, group)
                
                # Set data for proper sorting of file sizes
                if file_size != "Unknown":
                    try:
                        # Extract the numeric value for sorting
                        if file_size.endswith(" B"):
                            size_value = float(file_size.split(" ")[0])
                        elif file_size.endswith(" KB"):
                            size_value = float(file_size.split(" ")[0]) * 1024
                        elif file_size.endswith(" MB"):
                            size_value = float(file_size.split(" ")[0]) * 1024 * 1024
                        elif file_size.endswith(" GB"):
                            size_value = float(file_size.split(" ")[0]) * 1024 * 1024 * 1024
                        else:
                            size_value = 0
                        item.setData(3, Qt.UserRole, size_value)
                    except (ValueError, IndexError):
                        pass
                break
        
        # Restore sorting state
        self.streams_tree.setSortingEnabled(was_sorting_enabled)
        
    def set_icon_for_group(self, item, group):
        """Set an appropriate icon based on the file group"""
        if not group or group == "undefined":
            item.setIcon(0, self.group_icons['unknown'])
            return
            
        # Set the icon based on the group
        if group in self.group_icons:
            item.setIcon(0, self.group_icons[group])
        else:
            item.setIcon(0, self.group_icons['unknown'])
            
    def set_icon_for_mime_type(self, item, mime_type):
        """Legacy method - now just calls set_icon_for_group with 'unknown'"""
        # This is kept for backward compatibility
        self.set_icon_for_group(item, 'unknown')
        
    def identify_streams_finished(self):
        """Called when stream identification is complete"""
        self.progress_bar.setVisible(False)
        self.identify_streams_button.setEnabled(True)
        self.statusBar().showMessage("Stream identification completed")
        
        # Resize columns to fit content
        self.resize_streams_columns()
        
        # Don't automatically sort after identification
        # self.streams_tree.sortByColumn(2, Qt.DescendingOrder)
        
    def resize_streams_columns(self):
        """Resize the columns in the streams tree to fit content but not more than 33% of width"""
        if self.streams_tree.topLevelItemCount() == 0:
            return
            
        # Resize columns to fit content
        self.streams_tree.resizeColumnToContents(0)  # Stream Name column
        self.streams_tree.resizeColumnToContents(1)  # Group column
        self.streams_tree.resizeColumnToContents(2)  # MIME Type column
        self.streams_tree.resizeColumnToContents(3)  # File Size column
        self.streams_tree.resizeColumnToContents(4)  # SHA1 Hash column
        
        # Get total width
        total_width = self.streams_tree.width()
        
        # Limit each column to 33% of total width
        max_width = total_width // 5
        
        if self.streams_tree.columnWidth(0) > max_width:
            self.streams_tree.setColumnWidth(0, max_width)
            
        if self.streams_tree.columnWidth(1) > max_width:
            self.streams_tree.setColumnWidth(1, max_width)
            
        if self.streams_tree.columnWidth(2) > max_width:
            self.streams_tree.setColumnWidth(2, max_width)
            
        if self.streams_tree.columnWidth(3) > max_width:
            self.streams_tree.setColumnWidth(3, max_width)
            
        if self.streams_tree.columnWidth(4) > max_width:
            self.streams_tree.setColumnWidth(4, max_width)
        
    def table_selected(self, current, previous):
        """Handle table selection from the list"""
        if not current or not self.tables_data:
            return
            
        table_name = current.text()
        
        # Find the selected table in the data
        selected_table = None
        for table in self.tables_data:
            if table["name"] == table_name:
                selected_table = table
                break
                
        if not selected_table:
            return
            
        # Set up the table content view
        columns = selected_table["columns"]
        rows = selected_table["rows"]
        
        self.table_content.clear()
        self.table_content.setRowCount(len(rows))
        self.table_content.setColumnCount(len(columns))
        self.table_content.setHorizontalHeaderLabels(columns)
        
        # Fill the table with data
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_data in enumerate(row_data):
                if col_idx < len(columns):  # Safety check
                    item = QTableWidgetItem(cell_data)
                    self.table_content.setItem(row_idx, col_idx, item)
        
        self.statusBar().showMessage(f"Showing table: {table_name} ({len(rows)} rows)")
        
        # Update button states
        self.update_button_states()
        
    def run_command(self, command, callback):
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage("Running command...")
        
        self.thread = CommandThread(command)
        self.active_threads.append(self.thread)  # Track this thread
        self.thread.output_ready.connect(callback)
        self.thread.error_occurred.connect(lambda msg: self.handle_error("Command Error", msg))
        self.thread.finished_successfully.connect(lambda: self.command_finished(self.thread))
        self.thread.finished.connect(lambda: self.cleanup_thread(self.thread))
        self.thread.start()
        
    def command_finished(self, thread):
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Command completed successfully")
        
    def cleanup_thread(self, thread):
        """Remove thread from active threads list when it's done"""
        if thread in self.active_threads:
            self.active_threads.remove(thread)
        
    def handle_error(self, title, error, show_dialog=True):
        """Centralized error handling"""
        error_msg = str(error)
        self.statusBar().showMessage(f"Error: {error_msg[:100]}")
        if show_dialog:
            QMessageBox.critical(self, title, error_msg)

    def get_metadata(self):
        """Get MSI metadata"""
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_metadata", self.msi_file_path]
        output = self.run_command_safe(command)
        if output:
            self.display_metadata(output)
            
    def display_metadata(self, output):
        try:
            metadata = json.loads(output)
            formatted_text = "MSI Metadata:\n\n"
            
            for key, value in metadata.items():
                formatted_key = key.replace('_', ' ').title()
                if isinstance(value, list):
                    value_str = ", ".join(value) if value else "None"
                else:
                    value_str = str(value) if value else "None"
                    
                formatted_text += f"{formatted_key}: {value_str}\n"
                
            self.metadata_text.setText(formatted_text)
        except json.JSONDecodeError:
            self.handle_error("Parse Error", f"Error parsing metadata output:\n{output}", show_dialog=True)
            
    def list_streams(self):
        """List MSI streams"""
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_streams", self.msi_file_path]
        output = self.run_command_safe(command)
        if output:
            self.display_streams(output)
            
    def display_streams(self, output):
        try:
            streams = json.loads(output)
            self.streams_data = streams  # Store for later use
            
            # Disable sorting while populating
            self.streams_tree.setSortingEnabled(False)
            self.streams_tree.clear()
            
            for stream in streams:
                # Create item with four columns: Stream Name, Group, MIME Type, and File Size (empty for now)
                item = QTreeWidgetItem([stream, "", "", "", ""])
                # Set default icon
                item.setIcon(0, self.group_icons['unknown'])
                self.streams_tree.addTopLevelItem(item)
                
            # Re-enable sorting
            self.streams_tree.setSortingEnabled(True)
            
            # Mark as original order
            self.original_order = True
                
            # Resize columns to fit content
            self.resize_streams_columns()
                
            self.statusBar().showMessage(f"Found {len(streams)} streams")
        except json.JSONDecodeError:
            self.handle_error("Parse Error", "Error parsing streams output", show_dialog=True)
            
    def extract_all_streams(self):
        if not self.msi_file_path:
            return
            
        # Prompt for output directory
        output_dir = self.get_output_directory()
        if not output_dir:
            return  # User cancelled
            
        command = [self.msiparse_path, "extract_all", self.msi_file_path, output_dir]
        self.run_command(command, lambda output: self.handle_extraction_result(output, output_dir))
        
    def extract_stream(self):
        """Extract selected streams"""
        if not self.msi_file_path:
            return
            
        selected_items = self.streams_tree.selectedItems()
        if not selected_items:
            self.show_warning("Warning", "Please select streams to extract")
            return
            
        stream_names = [item.text(0) for item in selected_items]
        output_dir = self.get_output_directory()
        if not output_dir:
            return  # User cancelled
        
        with self.status_progress(f"Extracting {len(stream_names)} streams...", show_progress=True):
            self.progress_bar.setRange(0, len(stream_names))
            self.extract_multiple_streams(stream_names, output_dir)
            
    def extract_multiple_streams(self, stream_names, output_dir):
        """Extract multiple streams sequentially"""
        commands = [
            [self.msiparse_path, "extract", self.msi_file_path, output_dir, name]
            for name in stream_names
        ]
        
        self.current_extraction_index = 0
        self.extraction_commands = commands
        self.extraction_output_dir = output_dir
        self.extract_next_stream()
        
    def extract_next_stream(self):
        """Extract the next stream in the queue"""
        if self.current_extraction_index >= len(self.extraction_commands):
            self.show_status(self.STATUS_MESSAGES['extract_complete'])
            QMessageBox.information(
                self, 
                "Extraction Complete", 
                f"Files have been extracted to:\n{self.extraction_output_dir}"
            )
            return
            
        self.progress_bar.setValue(self.current_extraction_index + 1)
        command = self.extraction_commands[self.current_extraction_index]
        
        # Create and start the thread
        self.extract_thread = CommandThread(command)
        self.active_threads.append(self.extract_thread)
        
        self.extract_thread.finished_successfully.connect(self.on_stream_extracted)
        self.extract_thread.error_occurred.connect(lambda msg: self.show_error("Extraction Error", msg))
        self.extract_thread.finished.connect(lambda: self.cleanup_thread(self.extract_thread))
        
        self.extract_thread.start()
        
    def on_stream_extracted(self):
        """Called when a stream has been extracted successfully"""
        self.current_extraction_index += 1
        self.extract_next_stream()
        
    def handle_extraction_result(self, output, output_dir):
        self.statusBar().showMessage("Extraction completed")
        QMessageBox.information(
            self, 
            "Extraction Complete", 
            f"Files have been extracted to:\n{output_dir}"
        )
        
    def list_tables(self):
        """List MSI tables"""
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_tables", self.msi_file_path]
        output = self.run_command_safe(command)
        if output:
            self.display_tables(output)
            
    def display_tables(self, output):
        try:
            self.tables_data = json.loads(output)
            
            # Clear the table list and content
            self.table_list.clear()
            self.table_content.clear()
            self.table_content.setRowCount(0)
            self.table_content.setColumnCount(0)
            
            # Populate the table list
            for table in self.tables_data:
                table_name = table["name"]
                item = QListWidgetItem(table_name)
                
                # Add tooltip if table is in our description map
                if table_name in self.msi_tables:
                    item.setToolTip(self.msi_tables[table_name])
                    
                    # Create a custom widget with text and info button
                    widget = QWidget()
                    layout = QHBoxLayout(widget)
                    layout.setContentsMargins(4, 0, 4, 0)
                    layout.setSpacing(2)
                    
                    # Add table name label
                    label = QLabel(table_name)
                    layout.addWidget(label)
                    
                    # Add spacer to push the info button to the right
                    layout.addStretch()
                    
                    # Add info button
                    info_button = QToolButton()
                    info_button.setIcon(QApplication.style().standardIcon(QApplication.style().SP_MessageBoxInformation))
                    info_button.setToolTip(self.msi_tables[table_name])
                    info_button.setFixedSize(16, 16)
                    info_button.clicked.connect(lambda checked, name=table_name: self.show_table_info(name))
                    layout.addWidget(info_button)
                    
                    # Set the custom widget for this item
                    self.table_list.addItem(item)
                    self.table_list.setItemWidget(item, widget)
                else:
                    # Just add the regular item without an info button
                    self.table_list.addItem(item)
                
            self.statusBar().showMessage(f"Found {len(self.tables_data)} tables")
            
            # Update button states
            self.update_button_states()
        except json.JSONDecodeError:
            self.handle_error("Parse Error", "Error parsing tables output", show_dialog=True)
            
    def show_table_info(self, table_name):
        """Show a message box with information about the selected table"""
        if table_name in self.msi_tables:
            QMessageBox.information(
                self,
                f"Table Information: {table_name}",
                self.msi_tables[table_name]
            )
            
    def export_selected_table(self):
        """Export the currently selected table as JSON"""
        if not self.tables_data or not self.table_list.currentItem():
            return
            
        table_name = self.table_list.currentItem().text()
        selected_table = next(
            (table for table in self.tables_data if table["name"] == table_name),
            None
        )
        
        if not selected_table:
            self.show_warning("Warning", f"Table '{table_name}' not found in data")
            return
            
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            f"Save {table_name} as JSON", 
            f"{table_name}.json", 
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            with open(file_path, 'w') as f:
                json.dump(selected_table, f, indent=2)
                
            self.show_status(f"Table '{table_name}' exported to {file_path}")
            QMessageBox.information(
                self, 
                "Export Successful", 
                f"Table '{table_name}' has been exported to:\n{file_path}"
            )
        except Exception as e:
            self.show_error("Export Failed", e)
            
    def export_all_tables(self):
        """Export all tables as a single JSON file (exact output of list_tables command)"""
        if not self.tables_data:
            return
            
        # Prompt for save location
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            "Save All Tables as JSON", 
            "all_tables.json", 
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Write the complete tables data to the file
            with open(file_path, 'w') as f:
                json.dump(self.tables_data, f, indent=2)
                
            self.statusBar().showMessage(f"All tables exported to {file_path}")
            QMessageBox.information(
                self, 
                "Export Successful", 
                f"All tables have been exported to:\n{file_path}"
            )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Export Failed", 
                f"Failed to export tables: {str(e)}"
            )

    def on_sort_indicator_changed(self, logical_index, order):
        """Handle changes to the sort indicator"""
        # Update status message
        column_name = self.streams_tree.headerItem().text(logical_index)
        order_str = "ascending" if order == Qt.AscendingOrder else "descending"
        self.statusBar().showMessage(f"Sorted by {column_name} ({order_str})")
        
        # We're no longer in original order
        self.original_order = False
        
    def reset_to_original_order(self):
        """Reset the tree to the original order"""
        # Disable sorting temporarily
        self.streams_tree.setSortingEnabled(False)
        
        # Store current selection
        selected_streams = []
        for item in self.streams_tree.selectedItems():
            selected_streams.append(item.text(0))
        
        # Store current group, MIME type and sizes
        stream_data = {}
        for i in range(self.streams_tree.topLevelItemCount()):
            item = self.streams_tree.topLevelItem(i)
            stream_name = item.text(0)
            group = item.text(1)
            mime_type = item.text(2)
            file_size = item.text(3)
            size_value = item.data(3, Qt.UserRole)
            stream_data[stream_name] = (group, mime_type, file_size, size_value)
        
        # Clear and repopulate the tree
        self.streams_tree.clear()
        
        for stream in self.streams_data:
            # Get stored data if available
            group = ""
            mime_type = ""
            file_size = ""
            size_value = None
            
            if stream in stream_data:
                group, mime_type, file_size, size_value = stream_data[stream]
            
            # Create item with four columns
            item = QTreeWidgetItem([stream, group, mime_type, file_size])
            
            # Set the size value for proper sorting if available
            if size_value is not None:
                item.setData(3, Qt.UserRole, size_value)
                
            # Set icon based on group
            self.set_icon_for_group(item, group)
                
            self.streams_tree.addTopLevelItem(item)
            
            # Restore selection
            if stream in selected_streams:
                item.setSelected(True)
        
        # Re-enable sorting
        self.streams_tree.setSortingEnabled(True)
        
        # Mark as original order
        self.original_order = True
        
        # Resize columns to fit content
        self.resize_streams_columns()
        
        # Update status
        self.statusBar().showMessage("Restored original order")

    def show_streams_context_menu(self, position):
        """Show context menu for the streams tree"""
        # Get the item at the position
        item = self.streams_tree.itemAt(position)
        if not item:
            return
            
        # Get the stream name, group and MIME type
        stream_name = item.text(0)
        group = item.text(1)
        mime_type = item.text(2)
        sha1_hash = item.text(4)
        
        # Create context menu
        context_menu = QMenu(self)
        
        # Add Hex View action (always available)
        hex_view_action = QAction("Hex View", self)
        hex_view_action.triggered.connect(lambda: self.show_hex_view(stream_name))
        context_menu.addAction(hex_view_action)
        
        # Add group-specific actions
        if group == "image":
            # Image preview action
            preview_image_action = QAction("Preview Image", self)
            preview_image_action.triggered.connect(lambda: self.show_image_preview(stream_name))
            context_menu.addAction(preview_image_action)
            
        elif group == "text" or group == "code" or group == "document":
            # Text preview action
            preview_text_action = QAction("Preview Text", self)
            preview_text_action.triggered.connect(lambda: self.show_text_preview(stream_name))
            context_menu.addAction(preview_text_action)
            
        elif group == "archive" and self.archive_support:
            # Archive preview action
            preview_archive_action = QAction("Preview Archive", self)
            preview_archive_action.triggered.connect(lambda: self.show_archive_preview(stream_name))
            context_menu.addAction(preview_archive_action)
        
        # Add separator
        context_menu.addSeparator()
        
        # Add Extract File option
        extract_action = QAction("Extract File...", self)
        extract_action.triggered.connect(lambda: self.extract_single_stream(stream_name))
        context_menu.addAction(extract_action)
        
        # Add Hash Lookup option if hash is available
        if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
            # Add separator
            context_menu.addSeparator()
            
            # Add Hash Lookup submenu
            hash_menu = QMenu("Lookup Hash", self)
            
            # Add FileScan.io option (first)
            fs_action = QAction("FileScan.io", self)
            fs_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "filescan"))
            hash_menu.addAction(fs_action)
            
            # Add MetaDefender Cloud option (second)
            md_action = QAction("MetaDefender Cloud", self)
            md_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "metadefender"))
            hash_menu.addAction(md_action)
            
            # Add VirusTotal option (third)
            vt_action = QAction("VirusTotal", self)
            vt_action.triggered.connect(lambda: self.open_hash_lookup(sha1_hash, "virustotal"))
            hash_menu.addAction(vt_action)
            
            context_menu.addMenu(hash_menu)
        
        # Add separator before copy options
        context_menu.addSeparator()
        
        # Add Copy submenu at the bottom
        copy_menu = QMenu("Copy", self)
        
        # Add Copy options to the submenu
        copy_name_action = QAction("Stream Name", self)
        copy_name_action.triggered.connect(lambda: self.copy_to_clipboard(stream_name))
        copy_menu.addAction(copy_name_action)
        
        if mime_type:  # Only add if mime_type is not empty
            copy_type_action = QAction("MIME Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_to_clipboard(mime_type))
            copy_menu.addAction(copy_type_action)
            
        if sha1_hash and sha1_hash != "Error calculating hash" and sha1_hash != "":
            copy_hash_action = QAction("SHA1 Hash", self)
            copy_hash_action.triggered.connect(lambda: self.copy_to_clipboard(sha1_hash))
            copy_menu.addAction(copy_hash_action)
        
        context_menu.addMenu(copy_menu)
        
        # Show the context menu
        context_menu.exec_(self.streams_tree.mapToGlobal(position))
        
    def extract_single_stream(self, stream_name):
        """Extract a single stream to a user-specified location"""
        if not self.msi_file_path:
            return
            
        # Prompt for output directory
        output_dir = self.get_output_directory()
        if not output_dir:
            return  # User cancelled
            
        # Show progress
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage(f"Extracting stream: {stream_name}")
        
        # Create and run the command
        command = [
            self.msiparse_path,
            "extract",
            self.msi_file_path,
            output_dir,
            stream_name
        ]
        
        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=True
            )
            
            # Hide progress bar
            self.progress_bar.setVisible(False)
            
            # Show success message
            self.statusBar().showMessage(f"Stream '{stream_name}' extracted to {output_dir}")
            QMessageBox.information(
                self,
                "Extraction Complete",
                f"Stream '{stream_name}' has been extracted to:\n{output_dir}"
            )
            
        except Exception as e:
            # Hide progress bar
            self.progress_bar.setVisible(False)
            self.handle_error("Extraction Error", e)
        
    def copy_to_clipboard(self, text):
        """Copy the given text to the clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.statusBar().showMessage(f"Copied to clipboard: {text[:30]}{'...' if len(text) > 30 else ''}", 3000)
        
    def extract_file_to_temp(self, stream_name, temp_dir):
        """Extract a stream to a temporary directory and return the file path"""
        try:
            # Extract the stream to the temp directory
            command = [
                self.msiparse_path,
                "extract",
                self.msi_file_path,
                temp_dir,
                stream_name
            ]
            
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=True
            )
            
            # Path to the extracted file
            file_path = os.path.join(temp_dir, stream_name)
            
            # Check if file exists
            if os.path.exists(file_path):
                return file_path
            else:
                self.statusBar().showMessage(f"Failed to extract stream: {stream_name}")
                QMessageBox.warning(self, "Error", f"Failed to extract stream: {stream_name}")
                return None
                
        except Exception as e:
            self.statusBar().showMessage("Error during extraction")
            QMessageBox.critical(self, "Error", f"Error extracting file: {str(e)}")
            return None
            
    @contextlib.contextmanager
    def status_progress(self, message, show_progress=True, indeterminate=True):
        """Context manager for showing status message with optional progress bar"""
        if show_progress:
            if indeterminate:
                self.progress_bar.setRange(0, 0)
            self.progress_bar.setVisible(True)
        self.statusBar().showMessage(message)
        QApplication.processEvents()  # Keep UI responsive
        try:
            yield
        finally:
            if show_progress:
                self.progress_bar.setVisible(False)
            self.statusBar().showMessage(self.STATUS_MESSAGES['ready'])

    def show_status(self, message, timeout=0):
        """Show a status message with optional timeout"""
        self.statusBar().showMessage(message, timeout)
        
    def show_error(self, title, error, show_dialog=True, status_only=False):
        """Centralized error handling with optional dialog"""
        error_msg = str(error)
        self.show_status(f"Error: {error_msg[:100]}")
        if show_dialog and not status_only:
            QMessageBox.critical(self, title, error_msg)
            
    def show_warning(self, title, message, show_dialog=True, status_only=False):
        """Centralized warning handling with optional dialog"""
        self.show_status(message)
        if show_dialog and not status_only:
            QMessageBox.warning(self, title, message)
            
    def extract_file_safe(self, stream_name, output_dir=None, temp=False):
        """Safe file extraction with proper error handling and progress indication"""
        if temp and not output_dir:
            output_dir = tempfile.mkdtemp()
            
        try:
            with self.status_progress(f"Extracting stream: {stream_name}"):
                command = [
                    self.msiparse_path,
                    "extract",
                    self.msi_file_path,
                    output_dir,
                    stream_name
                ]
                
                result = subprocess.run(command, capture_output=True, text=True, check=True)
                file_path = os.path.join(output_dir, stream_name)
                
                if os.path.exists(file_path):
                    self.show_status(f"Stream '{stream_name}' extracted successfully")
                    return file_path
                    
                raise FileNotFoundError(f"Failed to extract stream: {stream_name}")
                
        except Exception as e:
            self.show_error("Extraction Error", e)
            return None

    def run_command_safe(self, command, success_message=None):
        """Run a command with proper error handling and progress indication"""
        try:
            with self.status_progress(self.STATUS_MESSAGES['running_command']):
                result = subprocess.run(command, capture_output=True, text=True, check=True)
                if success_message:
                    self.show_status(success_message)
                return result.stdout
        except Exception as e:
            self.show_error("Command Error", e)
            return None

    def show_preview(self, stream_name, preview_func):
        """Base method for showing previews with common functionality"""
        if not self.msi_file_path:
            return
                
        # Use context manager for temporary directory
        with temp_directory() as temp_dir:
            try:
                # Extract the file
                file_path = self.extract_file_to_temp(stream_name, temp_dir)
                
                if file_path:
                    # Call the specific preview function
                    preview_func(self, stream_name, file_path)
                    
            except Exception as e:
                self.handle_error("Preview Error", e)

    def show_hex_view(self, stream_name):
        """Show hex view of the stream content"""
        self.show_preview(stream_name, show_hex_view_dialog)
        
    def show_text_preview(self, stream_name):
        """Show text preview for text files"""
        self.show_preview(stream_name, show_text_preview_dialog)
        
    def show_image_preview(self, stream_name):
        """Show image preview for supported image formats"""
        self.show_preview(stream_name, show_image_preview_dialog)
        
    def show_archive_preview(self, stream_name):
        """Show archive preview for archive files"""
        # Check if archive support is enabled
        if not self.archive_support:
            QMessageBox.warning(
                self,
                "Archive Support Disabled",
                "Archive preview functionality is disabled because libarchive-c is not available.\n\n"
                "To enable archive support, install libarchive-c with:\n"
                "pip install libarchive-c"
            )
            return
            
        with self.status_progress(f"Extracting archive: {stream_name}..."):
            def show_archive(parent, name, path):
                self.show_status(f"Opening archive preview: {name}")
                archive_dialog = ArchivePreviewDialog(parent, name, path, self.group_icons)
                archive_dialog.exec_()
                
            self.show_preview(
                stream_name,
                show_archive
            )

    def has_theme_icons(self):
        """Check if theme icons are available"""
        # Try to get a common theme icon
        test_icon = QIcon.fromTheme('document-new')
        return not test_icon.isNull()

    def open_hash_lookup(self, hash_value, service):
        """Open the hash lookup in the specified service"""
        if service == "virustotal":
            url = f"https://www.virustotal.com/gui/file/{hash_value}"
            self.statusBar().showMessage(f"Opening hash in VirusTotal: {hash_value}")
        elif service == "metadefender":
            url = f"https://metadefender.com/results/hash/{hash_value}"
            self.statusBar().showMessage(f"Opening hash in MetaDefender Cloud: {hash_value}")
        elif service == "filescan":
            url = f"https://www.filescan.io/search-result?query={hash_value}"
            self.statusBar().showMessage(f"Opening hash in FileScan.io: {hash_value}")
        else:
            return
            
        try:
            webbrowser.open(url)
        except Exception as e:
            self.show_error("Browser Error", f"Failed to open browser: {str(e)}")
