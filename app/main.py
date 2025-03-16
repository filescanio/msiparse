# ruff: noqa: E722
import sys
import os
import json
import subprocess
import tempfile
import shutil
import contextlib
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QTabWidget, QTextEdit, 
                            QTreeWidget, QTreeWidgetItem, QMessageBox, QProgressBar,
                            QSplitter, QTableWidget, QTableWidgetItem, QHeaderView, QListWidget,
                            QToolButton, QListWidgetItem, QMenu, QAction, QDialog, QScrollArea)
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from PyQt5.QtGui import QIcon, QPixmap, QFont, QImage

# Try to import magika for file type identification
try:
    import magika
    MAGIKA_AVAILABLE = True
except ImportError:
    MAGIKA_AVAILABLE = False

# Try to import libarchive for archive handling
try:
    import libarchive as libarchive
    LIBARCHIVE_AVAILABLE = True
except ImportError:
    LIBARCHIVE_AVAILABLE = False

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

class CommandThread(QThread):
    """Thread for running msiparse commands without freezing the GUI"""
    output_ready = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    finished_successfully = pyqtSignal()
    
    def __init__(self, command):
        super().__init__()
        self.command = command
        self.running = True
        
    def run(self):
        try:
            if not self.running:
                return
                
            result = subprocess.run(
                self.command,
                capture_output=True,
                text=True,
                check=True
            )
            if self.running:
                self.output_ready.emit(result.stdout)
                self.finished_successfully.emit()
        except subprocess.CalledProcessError as e:
            if self.running:
                self.error_occurred.emit(f"Command failed with error: {e.stderr}")
        except Exception as e:
            if self.running:
                self.error_occurred.emit(f"Error: {str(e)}")
                
    def stop(self):
        """Stop the thread safely"""
        self.running = False

class IdentifyStreamsThread(QThread):
    """Thread for identifying stream file types"""
    progress_updated = pyqtSignal(int, int)  # current, total
    stream_identified = pyqtSignal(str, str, str, str)  # stream_name, group, mime_type, file_size
    finished_successfully = pyqtSignal()
    error_occurred = pyqtSignal(str)
    
    def __init__(self, msiparse_path, msi_file_path, streams):
        super().__init__()
        self.msiparse_path = msiparse_path
        self.msi_file_path = msi_file_path
        self.streams = streams
        self.running = True
        
        # Initialize magika if available
        if MAGIKA_AVAILABLE:
            self.magika_client = magika.Magika()
        
    def run(self):
        if not MAGIKA_AVAILABLE:
            self.error_occurred.emit("Magika library is not installed. Please install it with: pip install magika")
            return
            
        if not self.running:
            return
            
        # Create a temporary directory for stream extraction
        temp_dir = tempfile.mkdtemp()
        
        try:
            total_streams = len(self.streams)
            
            for i, stream_name in enumerate(self.streams):
                if not self.running:
                    break
                    
                # Update progress
                self.progress_updated.emit(i + 1, total_streams)
                
                # Extract the stream to the temp directory
                try:
                    command = [
                        self.msiparse_path,
                        "extract",
                        self.msi_file_path,
                        temp_dir,
                        stream_name
                    ]
                    
                    subprocess.run(
                        command,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    
                    # Path to the extracted file
                    file_path = Path(temp_dir) / stream_name
                    
                    # Check if file exists
                    if file_path.exists():
                        try:
                            # Get file size
                            file_size = file_path.stat().st_size
                            file_size_str = self.format_file_size(file_size)
                            
                            # Identify file type using magika with Path object
                            result = self.magika_client.identify_path(file_path)
                            mime_type = result.output.mime_type
                            group = result.output.group
                            
                            # Emit the result with group
                            self.stream_identified.emit(stream_name, group, mime_type, file_size_str)
                        except Exception as e:
                            self.stream_identified.emit(stream_name, "unknown", f"Error: {str(e)[:50]}", "Unknown")
                        
                        # Delete the temporary file
                        try:
                            file_path.unlink()
                        except:
                            pass
                    else:
                        self.stream_identified.emit(stream_name, "unknown", "Error: File not extracted", "Unknown")
                except Exception as e:
                    # Continue with next stream if one fails
                    self.stream_identified.emit(stream_name, "unknown", f"Error: {str(e)[:50]}", "Unknown")
                    
            if self.running:
                self.finished_successfully.emit()
                
        except Exception as e:
            if self.running:
                self.error_occurred.emit(f"Error during identification: {str(e)}")
        finally:
            # Clean up the temporary directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    def format_file_size(self, size_bytes):
        """Format file size in a human-readable format"""
        return format_file_size(size_bytes)
                
    def stop(self):
        """Stop the thread safely"""
        self.running = False

def read_text_file_with_fallback(file_path):
    """Read a text file with UTF-8 encoding, falling back to Latin-1 if needed.
    Returns the content or None if reading fails."""
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
    return content

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

def show_text_preview_dialog(parent, file_name, file_path):
    """Show a text preview dialog for the given file path"""
    # Try to read the file content
    content = read_text_file_with_fallback(file_path)
    
    if content is None:
        QMessageBox.warning(parent, "Error", "Failed to read text file")
        return False
        
    # Show text preview dialog
    text_dialog = TextPreviewDialog(parent, file_name, content)
    text_dialog.exec_()
    return True

def show_hex_view_dialog(parent, file_name, file_path):
    """Show a hex view dialog for the given file path"""
    try:
        # Read the file content
        with open(file_path, 'rb') as f:
            content = f.read()
        
        # Show hex view dialog
        hex_dialog = HexViewDialog(parent, file_name, content)
        hex_dialog.exec_()
        return True
    except Exception as e:
        QMessageBox.critical(parent, "Error", f"Error showing hex view: {str(e)}")
        return False

def show_image_preview_dialog(parent, file_name, file_path):
    """Show an image preview dialog for the given file path"""
    try:
        # Show image preview dialog
        image_dialog = ImagePreviewDialog(parent, file_name, file_path)
        image_dialog.exec_()
        return True
    except Exception as e:
        QMessageBox.critical(parent, "Error", f"Error showing image preview: {str(e)}")
        return False

class MSIParseGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.msi_file_path = None
        self.output_dir = None  # Will be set when needed
        self.last_output_dir = None  # Remember the last chosen directory
        self.msiparse_path = self.find_msiparse_executable()
        self.active_threads = []  # Keep track of active threads
        self.tables_data = None  # Store the tables data
        self.streams_data = []  # Store stream names
        
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
        self.streams_tree.setHeaderLabels(["Stream Name", "Group", "MIME Type", "File Size"])
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
        
    def check_library_available(self, library_available, library_name, install_command):
        """Check if a library is available and show a warning if not"""
        if not library_available:
            QMessageBox.warning(
                self, 
                f"{library_name} Not Found", 
                f"The {library_name} library is not installed. Please install it with:\n{install_command}"
            )
            return False
        return True
    
    def identify_streams(self):
        """Identify the file types of all streams using Magika"""
        if not self.msi_file_path or not self.streams_data:
            return
            
        # Check if Magika is available
        if not self.check_library_available(MAGIKA_AVAILABLE, "Magika", "pip install magika"):
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
        self.identify_thread.error_occurred.connect(self.handle_error)
        self.identify_thread.finished.connect(lambda: self.cleanup_thread(self.identify_thread))
        
        # Start the thread
        self.identify_thread.start()
        
    def update_identify_progress(self, current, total):
        """Update the progress bar during stream identification"""
        self.progress_bar.setValue(current)
        self.statusBar().showMessage(f"Identifying stream types: {current}/{total}")
        
    def update_stream_file_type(self, stream_name, group, mime_type, file_size):
        """Update the group, MIME type and size for a stream in the tree"""
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
        
        # Get total width
        total_width = self.streams_tree.width()
        
        # Limit each column to 33% of total width
        max_width = total_width // 3
        
        if self.streams_tree.columnWidth(0) > max_width:
            self.streams_tree.setColumnWidth(0, max_width)
            
        if self.streams_tree.columnWidth(1) > max_width:
            self.streams_tree.setColumnWidth(1, max_width)
            
        if self.streams_tree.columnWidth(2) > max_width:
            self.streams_tree.setColumnWidth(2, max_width)
            
        if self.streams_tree.columnWidth(3) > max_width:
            self.streams_tree.setColumnWidth(3, max_width)
        
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
        self.thread.error_occurred.connect(self.handle_error)
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
        
    def handle_error(self, error_message):
        self.progress_bar.setVisible(False)
        self.statusBar().showMessage("Error occurred")
        QMessageBox.critical(self, "Error", error_message)
        
    def get_metadata(self):
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_metadata", self.msi_file_path]
        self.run_command(command, self.display_metadata)
        
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
            self.metadata_text.setText(f"Error parsing JSON output:\n{output}")
            
    def list_streams(self):
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_streams", self.msi_file_path]
        self.run_command(command, self.display_streams)
        
    def display_streams(self, output):
        try:
            streams = json.loads(output)
            self.streams_data = streams  # Store for later use
            
            # Disable sorting while populating
            self.streams_tree.setSortingEnabled(False)
            self.streams_tree.clear()
            
            for stream in streams:
                # Create item with four columns: Stream Name, Group, MIME Type, and File Size (empty for now)
                item = QTreeWidgetItem([stream, "", "", ""])
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
            self.statusBar().showMessage("Error parsing streams output")
            QMessageBox.warning(self, "Error", f"Error parsing JSON output:\n{output}")
            
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
        if not self.msi_file_path:
            return
            
        # Get selected streams
        selected_items = self.streams_tree.selectedItems()
        stream_names = []
        
        if not selected_items:
            QMessageBox.warning(self, "Warning", "Please select streams to extract")
            return
            
        # Use selected items from the tree
        for item in selected_items:
            stream_names.append(item.text(0))
            
        # Prompt for output directory
        output_dir = self.get_output_directory()
        if not output_dir:
            return  # User cancelled
        
        # Show progress
        self.progress_bar.setRange(0, len(stream_names))
        self.progress_bar.setValue(0)
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage(f"Extracting {len(stream_names)} streams...")
        
        # Extract each stream
        self.extract_multiple_streams(stream_names, output_dir)
        
    def extract_multiple_streams(self, stream_names, output_dir):
        """Extract multiple streams sequentially"""
        # Create a list to store commands for each stream
        commands = []
        for stream_name in stream_names:
            command = [
                self.msiparse_path, 
                "extract", 
                self.msi_file_path, 
                output_dir,
                stream_name
            ]
            commands.append(command)
        
        # Start the first extraction
        self.current_extraction_index = 0
        self.extraction_commands = commands
        self.extraction_output_dir = output_dir
        self.extract_next_stream()
        
    def extract_next_stream(self):
        """Extract the next stream in the queue"""
        if self.current_extraction_index >= len(self.extraction_commands):
            # All extractions complete
            self.progress_bar.setVisible(False)
            self.statusBar().showMessage("Extraction completed")
            QMessageBox.information(
                self, 
                "Extraction Complete", 
                f"Files have been extracted to:\n{self.extraction_output_dir}"
            )
            return
            
        # Update progress
        self.progress_bar.setValue(self.current_extraction_index + 1)
        self.statusBar().showMessage(f"Extracting stream {self.current_extraction_index + 1}/{len(self.extraction_commands)}")
        
        # Run the command for the current stream
        command = self.extraction_commands[self.current_extraction_index]
        
        # Create and start the thread
        self.extract_thread = CommandThread(command)
        self.active_threads.append(self.extract_thread)
        
        # Connect signals
        self.extract_thread.finished_successfully.connect(self.on_stream_extracted)
        self.extract_thread.error_occurred.connect(self.handle_error)
        self.extract_thread.finished.connect(lambda: self.cleanup_thread(self.extract_thread))
        
        # Start the thread
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
        if not self.msi_file_path:
            return
            
        command = [self.msiparse_path, "list_tables", self.msi_file_path]
        self.run_command(command, self.display_tables)
        
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
            self.statusBar().showMessage("Error parsing tables output")
            QMessageBox.warning(self, "Error", f"Error parsing JSON output:\n{output}")
            
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
        
        # Find the selected table in the data
        selected_table = None
        for table in self.tables_data:
            if table["name"] == table_name:
                selected_table = table
                break
                
        if not selected_table:
            QMessageBox.warning(self, "Warning", f"Table '{table_name}' not found in data")
            return
            
        # Prompt for save location
        file_path, _ = QFileDialog.getSaveFileName(
            self, 
            f"Save {table_name} as JSON", 
            f"{table_name}.json", 
            "JSON Files (*.json);;All Files (*)"
        )
        
        if not file_path:
            return  # User cancelled
            
        try:
            # Write the table data to the file
            with open(file_path, 'w') as f:
                json.dump(selected_table, f, indent=2)
                
            self.statusBar().showMessage(f"Table '{table_name}' exported to {file_path}")
            QMessageBox.information(
                self, 
                "Export Successful", 
                f"Table '{table_name}' has been exported to:\n{file_path}"
            )
        except Exception as e:
            QMessageBox.critical(
                self, 
                "Export Failed", 
                f"Failed to export table: {str(e)}"
            )
            
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
            
        elif group == "archive":
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
        
        # Add separator before copy options
        context_menu.addSeparator()
        
        # Add Copy options at the bottom
        copy_name_action = QAction("Copy Stream Name", self)
        copy_name_action.triggered.connect(lambda: self.copy_to_clipboard(stream_name))
        context_menu.addAction(copy_name_action)
        
        if mime_type:  # Only add if mime_type is not empty
            copy_type_action = QAction("Copy MIME Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_to_clipboard(mime_type))
            context_menu.addAction(copy_type_action)
        
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
            
            # Show error message
            self.statusBar().showMessage("Error during extraction")
            QMessageBox.critical(
                self,
                "Extraction Failed",
                f"Failed to extract stream: {str(e)}"
            )
        
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
            
    def show_hex_view(self, stream_name):
        """Show hex view of the stream content"""
        if not self.msi_file_path:
            return
            
        # Use context manager for temporary directory
        with temp_directory() as temp_dir:
            try:
                # Extract the file
                file_path = self.extract_file_to_temp(stream_name, temp_dir)
                
                if file_path:
                    show_hex_view_dialog(self, stream_name, file_path)
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error showing hex view: {str(e)}")
        
    def show_image_preview(self, stream_name):
        """Show image preview for supported image formats"""
        if not self.msi_file_path:
            return
            
        # Use context manager for temporary directory
        with temp_directory() as temp_dir:
            try:
                # Extract the file
                file_path = self.extract_file_to_temp(stream_name, temp_dir)
                
                if file_path:
                    show_image_preview_dialog(self, stream_name, file_path)
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error showing image preview: {str(e)}")
        
    def show_text_preview(self, stream_name):
        """Show text preview for text files"""
        if not self.msi_file_path:
            return
            
        # Use context manager for temporary directory
        with temp_directory() as temp_dir:
            try:
                # Extract the file
                file_path = self.extract_file_to_temp(stream_name, temp_dir)
                
                if file_path:
                    show_text_preview_dialog(self, stream_name, file_path)
                    
            except Exception as e:
                QMessageBox.critical(self, "Error", f"Error showing text preview: {str(e)}")
                
    def show_archive_preview(self, stream_name):
        """Show archive preview for archive files"""
        if not self.msi_file_path:
            return
            
        # Check if libarchive is available
        if not self.check_library_available(LIBARCHIVE_AVAILABLE, "libarchive-c", "pip install libarchive-c"):
            return
            
        # Show progress bar during extraction
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setVisible(True)
        self.statusBar().showMessage(f"Extracting archive: {stream_name}...")
        QApplication.processEvents()  # Ensure UI updates
            
        # Use context manager for temporary directory
        with temp_directory() as temp_dir:
            try:
                # Extract the file
                file_path = self.extract_file_to_temp(stream_name, temp_dir)
                
                if file_path:
                    # Hide progress bar before showing dialog
                    self.progress_bar.setVisible(False)
                    self.statusBar().showMessage(f"Opening archive preview: {stream_name}")
                    
                    # Show archive preview dialog
                    archive_dialog = ArchivePreviewDialog(self, stream_name, file_path, self.group_icons)
                    archive_dialog.exec_()
                    
            except Exception as e:
                self.progress_bar.setVisible(False)
                QMessageBox.critical(self, "Error", f"Error showing archive preview: {str(e)}")
            finally:
                # Reset status bar
                self.statusBar().showMessage("Ready")

    def has_theme_icons(self):
        """Check if theme icons are available"""
        # Try to get a common theme icon
        test_icon = QIcon.fromTheme('document-new')
        return not test_icon.isNull()

class HexViewDialog(QDialog):
    """Dialog for displaying hex view of stream content"""
    def __init__(self, parent, stream_name, content):
        super().__init__(parent)
        self.stream_name = stream_name
        self.content = content
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle(f"Hex View: {self.stream_name}")
        self.setGeometry(100, 100, 800, 600)
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Create a text edit for the hex view (much more efficient than individual labels)
        self.hex_view = QTextEdit()
        self.hex_view.setReadOnly(True)
        self.hex_view.setLineWrapMode(QTextEdit.NoWrap)
        
        # Use a monospaced font for the hex view
        font = QFont("Courier New", 10)
        font.setFixedPitch(True)
        self.hex_view.setFont(font)
        
        # Add a status label to show position information
        self.status_label = QLabel()
        layout.addWidget(self.hex_view)
        layout.addWidget(self.status_label)
        
        # Set the content (after status_label is created)
        self.format_hex_view()
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)
        
    def format_hex_view(self):
        """Format the content as a hex view"""
        # Calculate number of bytes and set status
        num_bytes = len(self.content)
        
        # Calculate number of rows needed
        num_rows = (num_bytes + 15) // 16
        
        # Limit the number of rows to display for very large files
        display_limit = 10000  # Maximum number of rows to render at once
        if num_rows > display_limit:
            self.status_label.setText(f"File size: {num_bytes} bytes (showing first {display_limit * 16} bytes)")
            num_rows = display_limit
        else:
            self.status_label.setText(f"File size: {num_bytes} bytes")
        
        # Create the header
        header = "Offset   | 00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F | ASCII\n"
        header += "-" * 79 + "\n"
        
        # Set the header first to show something immediately
        self.hex_view.setText(header)
        QApplication.processEvents()  # Allow UI to update
        
        # Process the file in chunks to avoid memory issues with very large files
        chunk_size = 1000  # Number of rows to process at once
        
        for chunk_start in range(0, num_rows, chunk_size):
            chunk_end = min(chunk_start + chunk_size, num_rows)
            chunk_content = ""
            
            for row in range(chunk_start, chunk_end):
                # Offset
                offset = row * 16
                line = f"{offset:08X} | "
                
                # Hex values
                ascii_text = ""
                for col in range(16):
                    byte_pos = row * 16 + col
                    if byte_pos < num_bytes:
                        byte_val = self.content[byte_pos]
                        hex_val = f"{byte_val:02X}"
                        line += hex_val + " "
                        
                        # Add extra space after 8 bytes
                        if col == 7:
                            line += " "
                            
                        # Add to ASCII representation
                        if 32 <= byte_val <= 126:  # Printable ASCII
                            ascii_text += chr(byte_val)
                        else:
                            ascii_text += "."
                    else:
                        # Padding for incomplete rows
                        line += "   "
                        if col == 7:
                            line += " "
                        ascii_text += " "
                
                # Add ASCII representation
                line += "| " + ascii_text + "\n"
                chunk_content += line
            
            # Append this chunk to the text edit
            current_text = self.hex_view.toPlainText()
            self.hex_view.setText(current_text + chunk_content)
            QApplication.processEvents()  # Allow UI to update between chunks
            
        # Move cursor to the beginning
        cursor = self.hex_view.textCursor()
        cursor.movePosition(cursor.Start)
        self.hex_view.setTextCursor(cursor)

class ImagePreviewDialog(QDialog):
    """Dialog for displaying image preview"""
    def __init__(self, parent, stream_name, image_path):
        super().__init__(parent)
        self.stream_name = stream_name
        self.image_path = image_path
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle(f"Image Preview: {self.stream_name}")
        self.setGeometry(100, 100, 800, 600)
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Create a scroll area for the image
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        
        # Create a label to display the image
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        
        # Add the label to the scroll area
        scroll_area.setWidget(self.image_label)
        
        # Add the scroll area to the layout
        layout.addWidget(scroll_area)
        
        # Add status label to show image information - create this BEFORE loading the image
        self.status_label = QLabel("Loading image...")
        layout.addWidget(self.status_label)
        
        # Load the image (after status_label is created)
        self.load_image()
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)
        
    def load_image(self):
        """Load and display the image"""
        try:
            # Load the image
            image = QImage(self.image_path)
            
            if image.isNull():
                self.image_label.setText("Failed to load image")
                self.status_label.setText("Error: Invalid or unsupported image format")
                return
                
            # Get image dimensions
            width = image.width()
            height = image.height()
            
            # Create a pixmap from the image
            pixmap = QPixmap.fromImage(image)
            
            # Set the pixmap to the label
            self.image_label.setPixmap(pixmap)
            
            # Update status label with image information
            format_name = self.get_format_name(image.format())
            depth = image.depth()
            self.status_label.setText(f"Format: {format_name} | Size: {width}x{height} | Depth: {depth} bits")
            
        except Exception as e:
            self.image_label.setText(f"Error loading image: {str(e)}")
            self.status_label.setText("Error occurred while loading the image")
            
    def get_format_name(self, format_id):
        """Convert QImage format ID to a readable name"""
        format_names = {
            QImage.Format_Invalid: "Invalid",
            QImage.Format_Mono: "Mono",
            QImage.Format_MonoLSB: "MonoLSB",
            QImage.Format_Indexed8: "Indexed8",
            QImage.Format_RGB32: "RGB32",
            QImage.Format_ARGB32: "ARGB32",
            QImage.Format_ARGB32_Premultiplied: "ARGB32_Premultiplied",
            QImage.Format_RGB16: "RGB16",
            QImage.Format_ARGB8565_Premultiplied: "ARGB8565_Premultiplied",
            QImage.Format_RGB666: "RGB666",
            QImage.Format_ARGB6666_Premultiplied: "ARGB6666_Premultiplied",
            QImage.Format_RGB555: "RGB555",
            QImage.Format_ARGB8555_Premultiplied: "ARGB8555_Premultiplied",
            QImage.Format_RGB888: "RGB888",
            QImage.Format_RGB444: "RGB444",
            QImage.Format_ARGB4444_Premultiplied: "ARGB4444_Premultiplied",
            QImage.Format_RGBX8888: "RGBX8888",
            QImage.Format_RGBA8888: "RGBA8888",
            QImage.Format_RGBA8888_Premultiplied: "RGBA8888_Premultiplied"
        }
        
        return format_names.get(format_id, f"Unknown ({format_id})")
        
class TextPreviewDialog(QDialog):
    """Dialog for displaying text preview"""
    def __init__(self, parent, stream_name, content):
        super().__init__(parent)
        self.stream_name = stream_name
        self.content = content
        self.init_ui()
        
    def init_ui(self):
        self.setWindowTitle(f"Text Preview: {self.stream_name}")
        self.setGeometry(100, 100, 800, 600)
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Create a text edit for the content
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.WidgetWidth)  # Enable line wrapping
        
        # Set the content
        self.text_edit.setText(self.content)
        
        # Add the text edit to the layout
        layout.addWidget(self.text_edit)
        
        # Add status label to show text information
        self.status_label = QLabel(f"Length: {len(self.content)} characters")
        layout.addWidget(self.status_label)
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)

class ArchivePreviewDialog(QDialog):
    """Dialog for displaying archive contents"""
    def __init__(self, parent, archive_name, archive_path, group_icons):
        super().__init__(parent)
        self.parent = parent
        self.archive_name = archive_name
        self.archive_path = archive_path
        self.group_icons = group_icons
        self.archive_entries = []
        self.temp_dir = None
        self.magika_client = None
        self.extracted_files = set()  # Track extracted files for cleanup
        
        # Initialize magika if available
        if MAGIKA_AVAILABLE:
            self.magika_client = magika.Magika()
            
        self.init_ui()
        self.load_archive_contents()
        
    def init_ui(self):
        self.setWindowTitle(f"Archive Preview: {self.archive_name}")
        self.setGeometry(100, 100, 900, 600)
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Status label at the top
        self.status_label = QLabel("Loading archive contents...")
        layout.addWidget(self.status_label)
        
        # Progress bar
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        layout.addWidget(self.progress_bar)
        
        # Button layout
        button_layout = QHBoxLayout()
        
        # Add Identify Files button
        self.identify_button = QPushButton("Identify Files")
        self.identify_button.clicked.connect(self.auto_identify_files)
        button_layout.addWidget(self.identify_button)
        
        # Add spacer to push buttons to the left
        button_layout.addStretch()
        
        layout.addLayout(button_layout)
        
        # Create a tree widget for the archive contents
        self.contents_tree = QTreeWidget()
        self.contents_tree.setHeaderLabels(["Name", "Group", "MIME Type", "Size"])
        self.contents_tree.setSelectionMode(QTreeWidget.SingleSelection)
        
        # Enable context menu for contents tree
        self.contents_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.contents_tree.customContextMenuRequested.connect(self.show_context_menu)
        
        layout.addWidget(self.contents_tree)
        
        # Add a close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.close_and_cleanup)
        layout.addWidget(close_button)
        
    def close_and_cleanup(self):
        """Close the dialog and clean up temporary files"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                # Clean up all extracted files first
                for file_path in self.extracted_files:
                    if os.path.exists(file_path):
                        try:
                            os.remove(file_path)
                        except:
                            pass
                
                # Then remove the temp directory
                shutil.rmtree(self.temp_dir)
            except:
                pass
        self.accept()
        
    def closeEvent(self, event):
        """Handle dialog close event"""
        self.close_and_cleanup()
        event.accept()
        
    def load_archive_contents(self):
        """Load the contents of the archive file"""
        try:
            # Create a temporary directory for extracted files
            self.temp_dir = tempfile.mkdtemp()
            
            # Show progress bar (indeterminate) during loading
            self.progress_bar.setRange(0, 0)  # Indeterminate progress
            self.progress_bar.setVisible(True)
            self.status_label.setText(f"Loading archive: {self.archive_name}...")
            QApplication.processEvents()  # Ensure UI updates
            
            # List archive entries
            with libarchive.file_reader(self.archive_path) as archive:
                # Count entries for progress
                self.archive_entries = list(archive)
                
            # Update status
            self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files)")
            
            # Populate the tree
            self.populate_tree()
            
            # Hide progress bar when done
            self.progress_bar.setVisible(False)
            
            # Auto-resize columns
            self.auto_resize_columns()
            
        except Exception as e:
            self.status_label.setText(f"Error: {str(e)}")
            self.progress_bar.setVisible(False)
            QMessageBox.critical(self, "Error", f"Failed to read archive: {str(e)}")
            
    def check_library_available(self, library_available, library_name, install_command):
        """Check if a library is available and show a warning if not"""
        if not library_available:
            QMessageBox.warning(
                self, 
                f"{library_name} Not Found", 
                f"The {library_name} library is not installed. Please install it with:\n{install_command}"
            )
            return False
        return True
        
    def auto_identify_files(self):
        """Identify file types for all files in the archive"""
        if not self.check_library_available(MAGIKA_AVAILABLE, "Magika", "pip install magika"):
            return
            
        if not self.magika_client:
            self.magika_client = magika.Magika()
            
        # Disable identify button while running
        self.identify_button.setEnabled(False)
            
        # Show progress bar for identification
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, self.count_all_items())
        self.status_label.setText("Identifying file types...")
        QApplication.processEvents()  # Ensure UI updates
        
        # Process all top-level items
        self.identify_items_recursive(self.contents_tree.invisibleRootItem(), 0)
        
        # Hide progress bar when done
        self.progress_bar.setVisible(False)
        self.status_label.setText(f"Archive: {self.archive_name} ({len(self.archive_entries)} files)")
        
        # Auto-resize columns after identification
        self.auto_resize_columns()
        
        # Re-enable identify button
        self.identify_button.setEnabled(True)
        
    def count_all_items(self):
        """Count all file items in the tree recursively"""
        return self.count_items_recursive(self.contents_tree.invisibleRootItem())
        
    def count_items_recursive(self, parent_item):
        """Recursively count all file items under the parent item"""
        count = 0
        
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If it's a file (has path data)
            if child.data(0, Qt.UserRole):
                count += 1
                
            # Recursively count child items (for directories)
            count += self.count_items_recursive(child)
            
        return count
        
    def identify_items_recursive(self, parent_item, progress_count):
        """Recursively identify file types for all items"""
        # Process all child items
        for i in range(parent_item.childCount()):
            child = parent_item.child(i)
            
            # If it's a file (has path data)
            full_path = child.data(0, Qt.UserRole)
            if full_path:
                # Update progress
                self.progress_bar.setValue(progress_count)
                progress_count += 1
                QApplication.processEvents()  # Keep UI responsive
                
                # Identify the file
                self.identify_file(child, show_message=False)
            
            # Recursively process child items (for directories)
            progress_count = self.identify_items_recursive(child, progress_count)
            
        return progress_count
        
    def auto_resize_columns(self):
        """Automatically resize columns to fit content"""
        for i in range(4):
            self.contents_tree.resizeColumnToContents(i)
            
        # Ensure columns don't get too wide
        total_width = self.contents_tree.width()
        max_width = total_width // 4
        
        for i in range(4):
            if self.contents_tree.columnWidth(i) > max_width:
                self.contents_tree.setColumnWidth(i, max_width)
            
    def populate_tree(self):
        """Populate the tree with archive entries"""
        # Clear the tree
        self.contents_tree.clear()
        
        # Create a dictionary to store the tree structure
        tree_structure = {}
        
        # Process each entry
        for entry in self.archive_entries:
            path = entry.pathname
            is_dir = entry.isdir
            size = entry.size if not is_dir else 0
            
            # Skip directories for now
            if is_dir:
                continue
                
            # Split the path into components
            path_parts = path.split('/')
            
            # Create tree items for each part of the path
            current_dict = tree_structure
            for i, part in enumerate(path_parts):
                if i == len(path_parts) - 1:  # Last part (file)
                    if '' not in current_dict:
                        current_dict[''] = []
                    current_dict[''].append((part, size, path))
                else:  # Directory
                    if part not in current_dict:
                        current_dict[part] = {}
                    current_dict = current_dict[part]
        
        # Build the tree from the structure
        self.build_tree(tree_structure, self.contents_tree)
        
        # Auto-resize columns
        self.auto_resize_columns()
            
    def build_tree(self, structure, parent_item):
        """Recursively build the tree from the structure dictionary"""
        # Add directories first
        for key, value in structure.items():
            if key == '':  # Files at this level
                continue
                
            # Create a directory item
            dir_item = QTreeWidgetItem(parent_item if isinstance(parent_item, QTreeWidget) else parent_item)
            dir_item.setText(0, key)
            dir_item.setIcon(0, self.group_icons['inode'])
            
            # Recursively add children
            self.build_tree(value, dir_item)
            
        # Add files
        if '' in structure:
            for file_info in structure['']:
                file_name, size, full_path = file_info
                
                # Create a file item
                file_item = QTreeWidgetItem(parent_item if isinstance(parent_item, QTreeWidget) else parent_item)
                file_item.setText(0, file_name)
                file_item.setText(3, self.format_size(size))
                
                # Store the full path for later use
                file_item.setData(0, Qt.UserRole, full_path)
                
                # Set a default icon
                file_item.setIcon(0, self.group_icons['unknown'])
                
    def format_size(self, size_bytes):
        """Format file size in a human-readable format"""
        return format_file_size(size_bytes)
            
    def show_context_menu(self, position):
        """Show context menu for the contents tree"""
        # Get the item at the position
        item = self.contents_tree.itemAt(position)
        if not item:
            return
            
        # Check if it's a file (has path data)
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            return  # It's a directory
            
        # Get file name and group/mime type if available
        file_name = item.text(0)
        group = item.text(1)
        mime_type = item.text(2)
        
        # Create context menu
        context_menu = QMenu(self)
        
        # Add Copy options
        copy_name_action = QAction("Copy File Name", self)
        copy_name_action.triggered.connect(lambda: self.copy_to_clipboard(file_name))
        context_menu.addAction(copy_name_action)
        
        if mime_type:  # Only add if mime_type is not empty
            copy_type_action = QAction("Copy MIME Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_to_clipboard(mime_type))
            context_menu.addAction(copy_type_action)
        
        # Add separator
        context_menu.addSeparator()
        
        # Add Extract File option
        extract_action = QAction("Extract File...", self)
        extract_action.triggered.connect(lambda: self.extract_file_to_location(item))
        context_menu.addAction(extract_action)
        
        # Add separator
        context_menu.addSeparator()
        
        # Extract and identify the file if needed
        if not group or not mime_type:
            identify_action = QAction("Identify File Type", self)
            identify_action.triggered.connect(lambda: self.identify_file(item))
            context_menu.addAction(identify_action)
        
        # Add Hex View action (always available)
        hex_view_action = QAction("Hex View", self)
        hex_view_action.triggered.connect(lambda: self.show_hex_view(item))
        context_menu.addAction(hex_view_action)
        
        # Add group-specific actions
        if group == "image":
            # Image preview action
            preview_image_action = QAction("Preview Image", self)
            preview_image_action.triggered.connect(lambda: self.show_image_preview(item))
            context_menu.addAction(preview_image_action)
            
        elif group == "text" or group == "code" or group == "document":
            # Text preview action
            preview_text_action = QAction("Preview Text", self)
            preview_text_action.triggered.connect(lambda: self.show_text_preview(item))
            context_menu.addAction(preview_text_action)
            
        elif group == "archive":
            # Nested archive preview action
            preview_archive_action = QAction("Preview Archive", self)
            preview_archive_action.triggered.connect(lambda: self.show_nested_archive_preview(item))
            context_menu.addAction(preview_archive_action)
        
        # Show the context menu
        context_menu.exec_(self.contents_tree.mapToGlobal(position))
        
    def extract_file_to_location(self, item):
        """Extract a file to a user-specified location"""
        file_name = item.text(0)
        full_path = item.data(0, Qt.UserRole)
        
        if not full_path:
            return
            
        # Ask user for save location
        save_path, _ = QFileDialog.getSaveFileName(
            self,
            "Save File",
            file_name,
            "All Files (*)"
        )
        
        if not save_path:
            return  # User cancelled
            
        try:
            # Extract the file first
            temp_path = self.extract_file(item)
            
            if not temp_path or not os.path.exists(temp_path):
                QMessageBox.warning(self, "Error", "Failed to extract file")
                return
                
            # Copy to the destination
            shutil.copy2(temp_path, save_path)
            
            self.status_label.setText(f"Extracted: {file_name} to {save_path}")
            QMessageBox.information(
                self,
                "File Extracted",
                f"File has been extracted to:\n{save_path}"
            )
            
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
        
    def copy_to_clipboard(self, text):
        """Copy the given text to the clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.status_label.setText(f"Copied to clipboard: {text[:30]}{'...' if len(text) > 30 else ''}")
        
    def extract_file(self, item):
        """Extract a file from the archive to a temporary location"""
        full_path = item.data(0, Qt.UserRole)
        if not full_path:
            return None
            
        file_name = item.text(0)
        output_path = os.path.join(self.temp_dir, file_name)
        
        # Check if already extracted
        if os.path.exists(output_path):
            return output_path
            
        try:
            # Extract the file
            with libarchive.file_reader(self.archive_path) as archive:
                for entry in archive:
                    if entry.pathname == full_path:
                        # Create parent directories if needed
                        os.makedirs(os.path.dirname(output_path), exist_ok=True)
                        
                        # Extract the file
                        with open(output_path, 'wb') as f:
                            for block in entry.get_blocks():
                                f.write(block)
                        
                        # Add to the set of extracted files for cleanup
                        self.extracted_files.add(output_path)
                        break
                        
            return output_path
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to extract file: {str(e)}")
            return None
            
    def identify_file(self, item, show_message=True):
        """Identify the type of a file in the archive"""
        if not self.check_library_available(MAGIKA_AVAILABLE, "Magika", "pip install magika") and show_message:
            return
            
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            return
            
        try:
            # Identify file type using magika with Path object
            result = self.magika_client.identify_path(Path(file_path))
            group = result.output.group
            mime_type = result.output.mime_type
            
            # Update the item
            item.setText(1, group)
            item.setText(2, mime_type)
            
            # Set icon based on group
            if group in self.group_icons:
                item.setIcon(0, self.group_icons[group])
            else:
                item.setIcon(0, self.group_icons['unknown'])
                
            if show_message:
                self.status_label.setText(f"Identified: {item.text(0)} as {mime_type}")
            
        except Exception as e:
            if show_message:
                QMessageBox.critical(self, "Error", f"Failed to identify file: {str(e)}")
            
    def show_hex_view(self, item):
        """Show hex view of a file in the archive"""
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            return
            
        show_hex_view_dialog(self, item.text(0), file_path)
        
    def show_image_preview(self, item):
        """Show image preview for an image file in the archive"""
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            return
            
        show_image_preview_dialog(self, item.text(0), file_path)
        
    def show_text_preview(self, item):
        """Show text preview for a text file in the archive"""
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            return
            
        try:
            show_text_preview_dialog(self, item.text(0), file_path)
        except Exception as e:
            QMessageBox.critical(self, "Error", f"Error showing text preview: {str(e)}")
            
    def show_nested_archive_preview(self, item):
        """Show preview for a nested archive file"""
        if not self.check_library_available(LIBARCHIVE_AVAILABLE, "libarchive-c", "pip install libarchive-c"):
            return
            
        # Show progress bar during extraction
        self.progress_bar.setRange(0, 0)  # Indeterminate progress
        self.progress_bar.setVisible(True)
        self.status_label.setText(f"Extracting nested archive: {item.text(0)}...")
        QApplication.processEvents()  # Ensure UI updates
            
        # Extract the file
        file_path = self.extract_file(item)
        if not file_path:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Failed to extract: {item.text(0)}")
            return
            
        try:
            # Hide progress bar before showing dialog
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Opening nested archive: {item.text(0)}")
            
            # Show archive preview dialog
            archive_dialog = ArchivePreviewDialog(self.parent, item.text(0), file_path, self.group_icons)
            archive_dialog.exec_()
            
        except Exception as e:
            self.progress_bar.setVisible(False)
            self.status_label.setText(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", f"Error showing archive preview: {str(e)}")

def main():
    app = QApplication(sys.argv)
    window = MSIParseGUI()
    window.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()