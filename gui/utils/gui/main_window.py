"""
Main window class for the MSI Parser GUI
"""

import os
import webbrowser
import json
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QTabWidget, QTextEdit, 
                            QTreeWidget, QMessageBox, QProgressBar,
                            QSplitter, QTableWidget, QHeaderView, QListWidget,
                            QApplication, QLineEdit, QShortcut, QTextBrowser, QCheckBox, QMenu, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence

# Import custom modules
from threads.command import CommandThread

# Import helper functions
from utils.gui.helpers import (
    find_msiparse_executable, 
    get_group_icons, 
    get_msi_tables_descriptions,
    get_status_messages,
    status_progress,
    run_command_safe
)

# Import tab functionality
from utils.gui.metadata_tab import get_metadata, display_metadata
from utils.gui.streams_tab import (
    list_streams, 
    display_streams, 
    identify_streams,
    update_identify_progress,
    update_stream_file_type,
    identify_streams_finished,
    set_icon_for_group,
    resize_streams_columns,
    filter_streams,
    on_sort_indicator_changed,
    reset_to_original_order,
    show_streams_context_menu
)
from utils.gui.tables_tab import (
    list_tables,
    display_tables,
    table_selected,
    show_table_info,
    export_selected_table,
    export_all_tables,
    filter_tables
)
from utils.gui.workflow_tab import (
    display_workflow_analysis,
    analyze_install_sequence,
    evaluate_custom_action_impact,
    evaluate_standard_action_impact
)
from utils.gui.certificate_tab import (
    extract_certificates,
    handle_certificate_extraction_complete,
    analyze_certificate,
    _analyze_certificate_files,
    analyze_certificate_chain_simple,
    analyze_signer_info_simple,
    get_name_as_text
)
from utils.gui.extraction import (
    extract_stream_unified,
    extract_file_to_temp,
    extract_file_safe,
    extract_single_stream,
    extract_all_streams,
    handle_extraction_all_complete,
    extract_stream,
    extract_multiple_streams,
    extract_next_stream,
    get_output_directory
)
from utils.gui.preview import (
    show_preview,
    show_hex_view,
    show_text_preview,
    show_image_preview,
    show_archive_preview
)

class MSIParseGUI(QMainWindow):
    def __init__(self, archive_support=True):
        super().__init__()
        self.msi_file_path = None
        self.output_dir = None  # Will be set when needed
        self.last_output_dir = None  # Remember the last chosen directory
        self.msiparse_path = find_msiparse_executable()
        self.active_threads = []  # Keep track of active threads
        self.tables_data = None  # Store the tables data
        self.streams_data = []  # Store stream names
        self.archive_support = archive_support  # Store archive support status
        
        # Initialize icons for different file groups
        self.group_icons = get_group_icons()
        
        # Table descriptions for common MSI tables
        self.msi_tables = get_msi_tables_descriptions()
        
        # Status message constants
        self.STATUS_MESSAGES = get_status_messages()
        
        self.init_ui()
        
    def closeEvent(self, event):
        """Clean up threads when the application is closing"""
        # Stop all active threads
        for thread in self.active_threads:
            thread.stop()
            thread.wait(100)  # Wait a bit for thread to finish
            
        # Accept the close event
        event.accept()
        
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
        self.identify_streams_button = QPushButton("Identify Stream Types")
        self.identify_streams_button.clicked.connect(self.identify_streams)
        streams_button_layout.addWidget(self.identify_streams_button)
        
        # Add Reset Order button
        self.reset_order_button = QPushButton("Reset to Original Order")
        self.reset_order_button.clicked.connect(self.reset_to_original_order)
        streams_button_layout.addWidget(self.reset_order_button)
        
        # Add spacer to separate button groups
        streams_button_layout.addStretch()
        
        # Add Extract All button
        self.extract_all_button = QPushButton("Extract All Streams")
        self.extract_all_button.clicked.connect(self.extract_all_streams)
        streams_button_layout.addWidget(self.extract_all_button)
        
        # Move Extract Stream button to top row
        self.extract_stream_button = QPushButton("Extract Selected Streams")
        self.extract_stream_button.clicked.connect(self.extract_stream)
        streams_button_layout.addWidget(self.extract_stream_button)
        
        streams_layout.addLayout(streams_button_layout)
        
        # Add filter input for streams
        filter_layout = QHBoxLayout()
        filter_label = QLabel("Filter:")
        self.streams_filter = QLineEdit()
        self.streams_filter.setPlaceholderText("Type to filter streams... (Ctrl+F)")
        self.streams_filter.textChanged.connect(self.filter_streams)
        self.streams_filter.setClearButtonEnabled(True)  # Add clear button inside the field
        
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.streams_filter)
        streams_layout.addLayout(filter_layout)
        
        # Update streams tree to have four columns
        self.streams_tree = QTreeWidget()
        self.streams_tree.setHeaderLabels(["Stream Name", "Group", "MIME Type", "File Size", "SHA1 Hash"])
        self.streams_tree.itemClicked.connect(self.stream_selected)
        self.streams_tree.setSelectionMode(QTreeWidget.ExtendedSelection)  # Allow multiple selection
        
        # Connect to selectionChanged signal to handle multiple selection
        self.streams_tree.itemSelectionChanged.connect(self.on_stream_selection_changed)
        
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
        # Set up context menu policy for the export all button
        self.export_all_tables_button.setContextMenuPolicy(Qt.CustomContextMenu)
        self.export_all_tables_button.customContextMenuRequested.connect(self.show_export_all_context_menu)
        tables_button_layout.addWidget(self.export_all_tables_button)
        
        tables_layout.addLayout(tables_button_layout)
        
        # Create a splitter for the tables view
        tables_splitter = QSplitter(Qt.Horizontal)
        tables_layout.addWidget(tables_splitter)
        
        # Left side - Table list with info buttons
        table_list_widget = QWidget()
        table_list_layout = QVBoxLayout()
        table_list_widget.setLayout(table_list_layout)
        
        # Add filter input for tables
        table_filter_layout = QHBoxLayout()
        table_filter_label = QLabel("Filter:")
        self.table_filter = QLineEdit()
        self.table_filter.setPlaceholderText("Type to filter tables... (Ctrl+F)")
        self.table_filter.textChanged.connect(self.filter_tables)
        self.table_filter.setClearButtonEnabled(True)  # Add clear button inside the field
        
        table_filter_layout.addWidget(table_filter_label)
        table_filter_layout.addWidget(self.table_filter)
        table_list_layout.addLayout(table_filter_layout)
        
        self.table_list = QListWidget()
        self.table_list.setMaximumWidth(200)  # Limit width of the table list
        
        self.table_list.currentItemChanged.connect(self.table_selected)
        table_list_layout.addWidget(self.table_list)
        
        # Add checkbox to hide empty tables BELOW the table list
        self.hide_empty_tables_checkbox = QCheckBox("Hide Empty Tables")
        self.hide_empty_tables_checkbox.setChecked(False)  # Default unchecked
        self.hide_empty_tables_checkbox.stateChanged.connect(self.filter_tables)
        table_list_layout.addWidget(self.hide_empty_tables_checkbox)
        
        tables_splitter.addWidget(table_list_widget)
        
        # Right side - Table content
        self.table_content = QTableWidget()
        self.table_content.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tables_splitter.addWidget(self.table_content)
        
        # Set initial splitter sizes
        tables_splitter.setSizes([200, 600])
        
        self.tabs.addTab(self.tables_tab, "Tables")
        
        # Certificate tab
        self.certificate_tab = QWidget()
        certificate_layout = QVBoxLayout()
        self.certificate_tab.setLayout(certificate_layout)
        
        # Add description label
        cert_description = QLabel("Extract digital signatures from the MSI file. Digital signatures are used to verify the authenticity and integrity of the MSI package.")
        cert_description.setWordWrap(True)
        certificate_layout.addWidget(cert_description)
        
        # Add button to extract certificates
        cert_button_layout = QHBoxLayout()
        self.extract_cert_button = QPushButton("Extract Digital Signatures")
        self.extract_cert_button.clicked.connect(self.extract_certificates)
        cert_button_layout.addWidget(self.extract_cert_button)
        
        # Add analyze button
        self.analyze_cert_button = QPushButton("Analyze Signature")
        self.analyze_cert_button.clicked.connect(self.analyze_certificate)
        self.analyze_cert_button.setEnabled(False)  # Disabled until certificate is extracted
        cert_button_layout.addWidget(self.analyze_cert_button)
        
        cert_button_layout.addStretch()
        certificate_layout.addLayout(cert_button_layout)
        
        # Create a splitter for certificate information
        cert_splitter = QSplitter(Qt.Vertical)
        certificate_layout.addWidget(cert_splitter, 1)
        
        # Top part - Status and basic info
        self.cert_status = QTextEdit()
        self.cert_status.setReadOnly(True)
        cert_splitter.addWidget(self.cert_status)
        
        # Bottom part - Detailed certificate information
        self.cert_details = QTextEdit()
        self.cert_details.setReadOnly(True)
        cert_splitter.addWidget(self.cert_details)
        
        # Set initial splitter sizes
        cert_splitter.setSizes([200, 400])
        
        # Add the certificate tab to the tab widget
        self.tabs.addTab(self.certificate_tab, "Certificates")
        
        # Workflow Analysis tab
        self.workflow_tab = QWidget()
        workflow_layout = QVBoxLayout()
        self.workflow_tab.setLayout(workflow_layout)
        
        # Add description label
        workflow_description = QLabel("Analyze the MSI installation workflow to understand the actions and processes that will be performed during installation, highlighting potentially high-impact operations. For detailed information about MSI installation workflows, see the Help tab.")
        workflow_description.setWordWrap(True)
        workflow_layout.addWidget(workflow_description)
        
        # Button layout
        workflow_button_layout = QHBoxLayout()
        
        # Button to analyze workflow
        self.analyze_workflow_button = QPushButton("Analyze Installation Workflow")
        self.analyze_workflow_button.clicked.connect(self.analyze_install_sequence)
        self.analyze_workflow_button.setEnabled(False)  # Disabled until MSI is loaded
        workflow_button_layout.addWidget(self.analyze_workflow_button)
        
        workflow_button_layout.addStretch()
        workflow_layout.addLayout(workflow_button_layout)
        
        # Tree view for actions in sequence
        self.sequence_tree = QTreeWidget()
        self.sequence_tree.setColumnCount(5)
        self.sequence_tree.setHeaderLabels(["Sequence", "Action", "Condition", "Type", "Impact"])
        self.sequence_tree.setAlternatingRowColors(True)
        self.sequence_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        workflow_layout.addWidget(self.sequence_tree, 1)  # Give it stretch factor of 1
        
        # Add the workflow tab to the tab widget
        self.tabs.addTab(self.workflow_tab, "Workflow Analysis")
        
        # Help Tab
        self.help_tab = QWidget()
        help_layout = QVBoxLayout()
        self.help_tab.setLayout(help_layout)
        
        # Add title and introduction
        help_title = QLabel("MSI Parser Help & Documentation")
        help_title.setStyleSheet("font-size: 16px; font-weight: bold;")
        help_layout.addWidget(help_title)
        
        help_intro = QLabel("This section provides comprehensive documentation about MSI files and their analysis. The information below will help you understand how MSI files work and how to interpret the results displayed in other tabs.")
        help_intro.setWordWrap(True)
        help_layout.addWidget(help_intro)
        
        # Add a small vertical spacer
        help_layout.addSpacing(10)
        
        # Create text browser for help content
        self.help_html = QTextBrowser()
        self.help_html.setOpenExternalLinks(True)
        help_layout.addWidget(self.help_html, 1)  # Give it stretch factor of 1
        
        # Load the static workflow analysis documentation
        display_workflow_analysis(self, target_widget='help')
        
        # Add the help tab to the tab widget
        self.tabs.addTab(self.help_tab, "Help")
        
        # Status bar
        self.statusBar().showMessage("Ready")
        
        # Disable buttons until file is selected
        self.update_button_states()
        
        # Set up keyboard shortcuts
        self.setup_shortcuts()
        
    def setup_shortcuts(self):
        """Set up keyboard shortcuts for the application"""
        # Global shortcut for Ctrl+F that works on any tab
        self.global_filter_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.global_filter_shortcut.activated.connect(self.focus_current_filter)
        
        # Tab-specific shortcuts as fallbacks
        self.streams_filter_shortcut = QShortcut(QKeySequence("Ctrl+F"), self.streams_tab)
        self.streams_filter_shortcut.activated.connect(lambda: self.streams_filter.setFocus())
        
        self.tables_filter_shortcut = QShortcut(QKeySequence("Ctrl+F"), self.tables_tab)
        self.tables_filter_shortcut.activated.connect(lambda: self.table_filter.setFocus())
        
        # Escape key to clear filters when they have focus
        self.streams_filter_escape = QShortcut(QKeySequence("Escape"), self.streams_filter)
        self.streams_filter_escape.activated.connect(self.streams_filter.clear)
        
        self.table_filter_escape = QShortcut(QKeySequence("Escape"), self.table_filter)
        self.table_filter_escape.activated.connect(self.table_filter.clear)
        
    def focus_current_filter(self):
        """Focus on the filter field of the currently active tab"""
        current_tab = self.tabs.currentWidget()
        
        if current_tab == self.streams_tab:
            self.streams_filter.setFocus()
        elif current_tab == self.tables_tab:
            self.table_filter.setFocus()
        
    def update_button_states(self):
        """Update the enabled state of buttons based on current state"""
        try:
            has_file = self.msi_file_path is not None
            has_selected_streams = len(self.streams_tree.selectedItems()) > 0
            has_tables = self.tables_data is not None and len(self.tables_data) > 0
            has_selected_table = self.table_list.currentItem() is not None
            
            # Update button states
            self.identify_streams_button.setEnabled(has_file)
            self.extract_all_button.setEnabled(has_file)
            self.extract_stream_button.setEnabled(has_file and has_selected_streams)
            self.export_selected_table_button.setEnabled(has_file and has_selected_table)
            self.export_all_tables_button.setEnabled(has_file and has_tables)
            self.extract_cert_button.setEnabled(has_file)
            self.analyze_cert_button.setEnabled(has_file)
            self.analyze_workflow_button.setEnabled(has_file)
            
            # Update reset order button if it exists
            if hasattr(self, 'reset_order_button'):
                self.reset_order_button.setEnabled(has_file and not self.original_order)
                
        except Exception as e:
            # Log the error but don't show a dialog to avoid disrupting the UI
            print(f"Error updating button states: {str(e)}")
            # Try to enable essential buttons as a fallback
            try:
                self.identify_streams_button.setEnabled(True)
                self.extract_all_button.setEnabled(True)
                self.extract_stream_button.setEnabled(True)
            except:
                pass
        
    def browse_msi_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select MSI File", "", "MSI Files (*.msi);;All Files (*)"
        )
        if file_path:
            self.msi_file_path = file_path
            self.file_label.setText(os.path.basename(file_path))
            self.update_button_states()
            self.statusBar().showMessage(f"Selected MSI file: {file_path}")
            
            # Clear certificate status and details
            self.cert_status.clear()
            self.cert_details.clear()
            
            # Reset extracted certificate files
            if hasattr(self, 'extracted_cert_files'):
                delattr(self, 'extracted_cert_files')
            
            # Auto-run metadata, streams, and tables when file is selected
            self.get_metadata()
            self.list_streams()
            self.list_tables()
            
    def stream_selected(self, item):
        """Handle single item click in the streams tree"""
        try:
            # Update button states when selection changes
            self.update_button_states()
        except Exception as e:
            self.handle_error("Selection Error", f"Error handling stream selection: {str(e)}", show_dialog=False)
            
    def on_stream_selection_changed(self):
        """Handle selection changes in the streams tree"""
        try:
            # Get selected items
            selected_items = self.streams_tree.selectedItems()
            
            # Update status bar with selection info
            if selected_items:
                if len(selected_items) == 1:
                    self.statusBar().showMessage(f"Selected stream: {selected_items[0].text(0)}")
                else:
                    self.statusBar().showMessage(f"Selected {len(selected_items)} streams")
            
            # Update button states
            self.update_button_states()
        except Exception as e:
            self.handle_error("Selection Error", f"Error handling stream selection change: {str(e)}", show_dialog=False)
        
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

    def copy_to_clipboard(self, text):
        """Copy the given text to the clipboard"""
        clipboard = QApplication.clipboard()
        clipboard.setText(text)
        self.statusBar().showMessage(f"Copied to clipboard: {text[:30]}{'...' if len(text) > 30 else ''}", 3000)

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

    # Forward method calls to the appropriate modules
    def get_metadata(self):
        return get_metadata(self)
        
    def display_metadata(self, output):
        return display_metadata(self, output)
        
    def list_streams(self):
        return list_streams(self)
        
    def display_streams(self, output):
        return display_streams(self, output)
        
    def identify_streams(self):
        return identify_streams(self)
        
    def update_identify_progress(self, current, total):
        return update_identify_progress(self, current, total)
        
    def update_stream_file_type(self, stream_name, group, mime_type, file_size, sha1_hash):
        return update_stream_file_type(self, stream_name, group, mime_type, file_size, sha1_hash)
        
    def identify_streams_finished(self):
        return identify_streams_finished(self)
        
    def set_icon_for_group(self, item, group):
        return set_icon_for_group(self, item, group)
        
    def resize_streams_columns(self):
        return resize_streams_columns(self)
        
    def filter_streams(self, filter_text):
        return filter_streams(self, filter_text)
        
    def on_sort_indicator_changed(self, logical_index, order):
        return on_sort_indicator_changed(self, logical_index, order)
        
    def reset_to_original_order(self):
        return reset_to_original_order(self)
        
    def show_streams_context_menu(self, position):
        return show_streams_context_menu(self, position)
        
    def list_tables(self):
        return list_tables(self)
        
    def display_tables(self, output):
        return display_tables(self, output)
        
    def table_selected(self, current, previous):
        return table_selected(self, current, previous)
        
    def show_table_info(self, table_name):
        return show_table_info(self, table_name)
        
    def export_selected_table(self):
        return export_selected_table(self)
        
    def export_all_tables(self):
        return export_all_tables(self)
        
    def filter_tables(self, filter_text):
        return filter_tables(self, filter_text)
        
    def extract_certificates(self):
        return extract_certificates(self)
        
    def handle_certificate_extraction_complete(self, output):
        return handle_certificate_extraction_complete(self, output)
        
    def analyze_certificate(self):
        return analyze_certificate(self)
        
    def _analyze_certificate_files(self, certificate_files):
        return _analyze_certificate_files(self, certificate_files)
        
    def analyze_certificate_chain_simple(self, certificates):
        """Analyze certificate chain in a simple way"""
        analyze_certificate_chain_simple(self, certificates)
        
    def analyze_signer_info_simple(self, signed_data):
        """Analyze signer info in a simple way"""
        analyze_signer_info_simple(self, signed_data)
        
    def get_name_as_text(self, name):
        """Get name as text"""
        return get_name_as_text(name)
        
    def analyze_install_sequence(self):
        """Analyze the MSI installation workflow sequence"""
        analyze_install_sequence(self)
        
    def extract_stream_unified(self, stream_name, output_dir=None, temp=False, show_messages=True):
        """Extract a stream and return the path to the extracted file"""
        return extract_stream_unified(self, stream_name, output_dir, temp, show_messages)
        
    def extract_file_to_temp(self, stream_name, temp_dir):
        return extract_file_to_temp(self, stream_name, temp_dir)
        
    def extract_file_safe(self, stream_name, output_dir=None, temp=False):
        return extract_file_safe(self, stream_name, output_dir, temp)
        
    def extract_single_stream(self, stream_name):
        return extract_single_stream(self, stream_name)
        
    def extract_all_streams(self):
        return extract_all_streams(self)
        
    def handle_extraction_all_complete(self, output_dir):
        return handle_extraction_all_complete(self, output_dir)
        
    def extract_stream(self):
        return extract_stream(self)
        
    def extract_multiple_streams(self, stream_names, output_dir):
        return extract_multiple_streams(self, stream_names, output_dir)
        
    def extract_next_stream(self):
        return extract_next_stream(self)
        
    def get_output_directory(self):
        return get_output_directory(self)
        
    def show_preview(self, stream_name, preview_func):
        return show_preview(self, stream_name, preview_func)
        
    def show_hex_view(self, stream_name):
        return show_hex_view(self, stream_name)
        
    def show_text_preview(self, stream_name):
        return show_text_preview(self, stream_name)
        
    def show_image_preview(self, stream_name):
        return show_image_preview(self, stream_name)
        
    def show_archive_preview(self, stream_name):
        return show_archive_preview(self, stream_name)
        
    def status_progress(self, message, show_progress=True, indeterminate=True):
        return status_progress(self, message, show_progress, indeterminate)
        
    def run_command_safe(self, command, success_message=None):
        return run_command_safe(self, command, success_message)

    def show_export_all_context_menu(self, position):
        """Show context menu for the Export All Tables button with hidden features"""
        # Only show the menu if tables data is available
        if not self.tables_data:
            return
            
        # Create context menu
        context_menu = QMenu(self)
        
        # Add export individual tables action
        export_individual_action = QAction("Export Tables Individually", self)
        export_individual_action.triggered.connect(self.export_tables_individually)
        context_menu.addAction(export_individual_action)
        
        # Show the context menu
        context_menu.exec_(self.export_all_tables_button.mapToGlobal(position))
    
    def export_tables_individually(self):
        """Hidden feature: Export all tables as individual JSON files"""
        if not self.tables_data:
            return
            
        # Prompt for directory to save all table files
        export_dir = QFileDialog.getExistingDirectory(
            self,
            "Select Directory for Table Exports",
            os.path.expanduser("~")
        )
        
        if not export_dir:
            return  # User cancelled
            
        try:
            # Count for progress tracking
            total_tables = len(self.tables_data)
            exported_count = 0
            
            # Show progress
            self.progress_bar.setRange(0, total_tables)
            self.progress_bar.setValue(0)
            self.progress_bar.setVisible(True)
            
            # Export each table as a separate file
            for table in self.tables_data:
                table_name = table["name"]
                file_path = os.path.join(export_dir, f"{table_name}.json")
                
                # Update status
                self.statusBar().showMessage(f"Exporting table {exported_count+1}/{total_tables}: {table_name}")
                
                # Export the table
                with open(file_path, 'w') as f:
                    json.dump(table, f, indent=2)
                
                # Update progress
                exported_count += 1
                self.progress_bar.setValue(exported_count)
                QApplication.processEvents()  # Keep UI responsive
            
            # Hide progress bar when done
            self.progress_bar.setVisible(False)
            
            # Update status
            self.statusBar().showMessage(f"Exported {exported_count} tables to {export_dir}")
            
            # Show success message
            QMessageBox.information(
                self,
                "Export Successful",
                f"All tables have been exported as individual files to:\n{export_dir}"
            )
            
        except Exception as e:
            # Hide progress bar on error
            self.progress_bar.setVisible(False)
            
            # Show error
            QMessageBox.critical(
                self,
                "Export Failed",
                f"Failed to export tables: {str(e)}"
            ) 