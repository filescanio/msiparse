"""
Main window class for the MSI Parser GUI
"""

import os
import webbrowser
import json
from PyQt5.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QPushButton, QLabel, QFileDialog, QTabWidget, QTextEdit, 
                            QTreeWidget, QMessageBox,
                            QSplitter, QTableWidget, QHeaderView, QListWidget,
                            QApplication, QLineEdit, QShortcut, QCheckBox, QMenu, QAction)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QKeySequence, QFont

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
from utils.gui.footprint_tab import (
    analyze_installation_impact,
    display_installation_impact,
    create_footprint_tab
)
from utils.gui.execution_tab import (
    analyze_install_sequence,
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
    extract_file_to_temp,
    extract_file_safe,
    extract_single_stream,
    extract_all_streams,
    handle_extraction_all_complete,
    extract_stream,
    extract_multiple_streams
)
from utils.gui.preview import (
    show_preview,
    show_hex_view,
    show_text_preview,
    show_image_preview,
    show_archive_preview
)
from utils.gui.help_tab import create_help_tab

# Helper function to center dialogs
def center_dialog_on_parent_screen(dialog, parent_window):
    if not parent_window:
        # Try to get the active window if no explicit parent is given
        parent_window = QApplication.activeWindow()
    if not parent_window:
        # If still no parent, can't do much, let Qt decide
        return

    # Determine the screen of the parent window
    # Using mapToGlobal and rect().center() is robust for finding the parent's center
    try:
        parent_center_global = parent_window.mapToGlobal(parent_window.rect().center())
        parent_screen = QApplication.screenAt(parent_center_global)
        if not parent_screen: # Fallback if screenAt returns None
            parent_screen = parent_window.screen() # Relies on QWindow.screen(), might be less precise for QWidget
        if not parent_screen:
             # Last resort, use primary screen if others fail
            parent_screen = QApplication.primaryScreen()
        
        if not parent_screen: # If still no screen, abort
            return
            
    except AttributeError: # parent_window might not be a QWidget with mapToGlobal or rect
        # Fallback to primary screen if parent properties are not as expected
        parent_screen = QApplication.primaryScreen()
        if not parent_screen:
            return

    screen_geometry = parent_screen.geometry()
    # Use frameGeometry to account for window decorations for more accurate centering
    dialog_frame_geometry = dialog.frameGeometry()

    new_x = screen_geometry.x() + (screen_geometry.width() - dialog_frame_geometry.width()) / 2
    new_y = screen_geometry.y() + (screen_geometry.height() - dialog_frame_geometry.height()) / 2
    dialog.move(int(new_x), int(new_y))

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
        
        # Font scaling attributes
        self.current_font_scale = 1.0
        self.base_font_size = QApplication.font().pointSize() # Store default app font size
        self._original_widget_fonts = {} # To store original font of widgets for proper scaling
        
        # Initialize icons for different file groups
        self.group_icons = get_group_icons()
        
        # Table descriptions for common MSI tables
        self.msi_tables = get_msi_tables_descriptions()
        
        # Status message constants
        self.STATUS_MESSAGES = get_status_messages()
        
        self.init_ui()
        
    def closeEvent(self, event):
        """Handle application close event"""
        # Stop any running threads
        for thread in self.active_threads[:]:  # Create a copy of the list to avoid modification during iteration
            try:
                if hasattr(thread, 'cleanup'):
                    # Ensure cleanup happens in the main thread
                    thread.cleanup()
                    thread.wait()  # Wait for the thread to finish
                elif hasattr(thread, 'stop'):
                    thread.stop()
                    thread.wait()  # Wait for the thread to finish
            except Exception as e:
                print(f"Error cleaning up thread: {str(e)}")
        
        # Clear the active threads list
        self.active_threads.clear()
        
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
        
        # Enable drag and drop
        self.setAcceptDrops(True)
        
        # Create menus (including View menu for scaling)
        self.create_menus()
        
        # File selection area
        file_layout = QHBoxLayout()
        self.file_label = QLabel("No MSI file selected")
        self.browse_button = QPushButton("Browse MSI File")
        self.browse_button.clicked.connect(self.browse_msi_file)
        file_layout.addWidget(self.file_label)
        file_layout.addWidget(self.browse_button)
        main_layout.addLayout(file_layout)
        
        # Create and store tab references
        self.metadata_tab = self.create_metadata_tab()
        self.streams_tab = self.create_streams_tab()
        self.tables_tab = self.create_tables_tab()
        self.certificates_tab = self.create_certificates_tab()
        self.execution_tab = self.create_execution_tab()
        self.footprint_tab = self.create_footprint_tab()
        self.help_tab = self.create_help_tab()
        
        # Tab widget for different functions
        self.tabs = QTabWidget()
        self.tabs.addTab(self.metadata_tab, "Metadata")
        self.tabs.addTab(self.streams_tab, "Streams")
        self.tabs.addTab(self.tables_tab, "Tables")
        self.tabs.addTab(self.certificates_tab, "Certificates")
        self.tabs.addTab(self.execution_tab, "Execution")
        self.tabs.addTab(self.footprint_tab, "Footprint")
        self.tabs.addTab(self.help_tab, "Help")
        main_layout.addWidget(self.tabs)
        
        # Status bar
        self.statusBar().showMessage("Ready")
        
        # Disable buttons until file is selected
        self.update_button_states()
        
        # Set up keyboard shortcuts after tabs are created
        self.setup_shortcuts()
        
        # Apply initial font scaling (which will be 1.0 by default)
        # This also helps populate _original_widget_fonts
        self.apply_scaling()
        
    def create_menus(self):
        menu_bar = self.menuBar()

        # File menu (example, if you had one)
        # file_menu = menu_bar.addMenu("&File")
        # open_action = QAction("&Open MSI...", self)
        # open_action.triggered.connect(self.browse_msi_file)
        # file_menu.addAction(open_action)
        # ... add other file actions

        # Define zoom actions and their shortcuts
        zoom_in_action = QAction("Zoom In", self)
        zoom_in_action.setShortcut(QKeySequence.ZoomIn) # Standard shortcut (Ctrl++)
        zoom_in_action.triggered.connect(self.zoom_in)
        self.addAction(zoom_in_action) # Add action to window for shortcut to work

        zoom_out_action = QAction("Zoom Out", self)
        zoom_out_action.setShortcut(QKeySequence.ZoomOut) # Standard shortcut (Ctrl+-)
        zoom_out_action.triggered.connect(self.zoom_out)
        self.addAction(zoom_out_action) # Add action to window

        reset_zoom_action = QAction("Reset Zoom", self)
        reset_zoom_action.setShortcut(QKeySequence("Ctrl+0"))
        reset_zoom_action.triggered.connect(self.reset_zoom)
        self.addAction(reset_zoom_action) # Add action to window

    def zoom_in(self):
        self.current_font_scale += 0.1
        if self.current_font_scale > 3.0: # Max scale limit
            self.current_font_scale = 3.0
        self.apply_scaling()

    def zoom_out(self):
        self.current_font_scale -= 0.1
        if self.current_font_scale < 0.5: # Min scale limit
            self.current_font_scale = 0.5
        self.apply_scaling()

    def reset_zoom(self):
        self.current_font_scale = 1.0
        self.apply_scaling()

    def apply_scaling(self):
        """Apply current font scale to all relevant widgets."""
        # First, ensure all widgets have their original fonts stored if not already
        for widget in self.findChildren(QWidget):
            widget_id = id(widget)
            if widget_id not in self._original_widget_fonts:
                self._original_widget_fonts[widget_id] = widget.font()

        # Apply scaled font
        for widget_id, original_font in self._original_widget_fonts.items():
            try:
                # Attempt to find widget by its stored id - this is tricky if widgets are recreated
                # A better approach might be to re-traverse or apply to known persistent widgets
                # For now, we assume widgets persist or this function is called after UI setup.
                
                # This direct id lookup is not safe if widgets are destroyed and recreated.
                # We will rely on iterating findChildren again for safety here.
                pass # Placeholder for direct id lookup logic if it were safe
            except: # pragma: no cover
                # Widget might no longer exist, skip
                continue
        
        all_widgets = self.findChildren(QWidget) # Get current list of widgets
        
        for widget in all_widgets:
            widget_id = id(widget)
            original_font = self._original_widget_fonts.get(widget_id)

            if not original_font: # Should have been populated above
                original_font = widget.font() # Fallback
                self._original_widget_fonts[widget_id] = original_font

            scaled_font = QFont(original_font) # Create a new font object from the original
            
            # Get original point size if it was positive, otherwise use app base
            original_point_size = original_font.pointSize()
            if original_point_size <= 0: # If point size isn't set, use a default base
                # For widgets like QTreeWidget items, pointSize might be -1
                # Let's try with application's base font size for these.
                 original_point_size = QApplication.font().pointSize()
                 if original_point_size <=0: # If still invalid, use our stored base_font_size
                     original_point_size = self.base_font_size


            new_size = int(original_point_size * self.current_font_scale)
            
            if new_size <= 0:
                new_size = 1 # Minimum font size
            
            scaled_font.setPointSize(new_size)
            widget.setFont(scaled_font)

            # Special handling for some widget types
            if isinstance(widget, QTreeWidget):
                header_font = QFont(widget.header().font())
                original_header_point_size = self._original_widget_fonts.get(id(widget.header()), widget.header().font()).pointSize()
                if original_header_point_size <=0: original_header_point_size = self.base_font_size
                new_header_size = int(original_header_point_size * self.current_font_scale)
                if new_header_size <=0: new_header_size = 1
                header_font.setPointSize(new_header_size)
                widget.header().setFont(header_font)
                widget.header().resizeSections(QHeaderView.ResizeToContents) # or another mode
                 # For QTreeWidget items, font is often set per item.
                 # This needs to be handled where items are created or by iterating items.
                 # We'll address this by modifying item creation functions later.

            elif isinstance(widget, QTableWidget):
                header_font = QFont(widget.horizontalHeader().font())
                original_h_header_point_size = self._original_widget_fonts.get(id(widget.horizontalHeader()), widget.horizontalHeader().font()).pointSize()
                if original_h_header_point_size <=0: original_h_header_point_size = self.base_font_size
                new_h_header_size = int(original_h_header_point_size * self.current_font_scale)
                if new_h_header_size <=0: new_h_header_size = 1
                header_font.setPointSize(new_h_header_size)
                widget.horizontalHeader().setFont(header_font)
                widget.horizontalHeader().resizeSections(QHeaderView.ResizeToContents)

                v_header_font = QFont(widget.verticalHeader().font())
                original_v_header_point_size = self._original_widget_fonts.get(id(widget.verticalHeader()), widget.verticalHeader().font()).pointSize()
                if original_v_header_point_size <=0: original_v_header_point_size = self.base_font_size
                new_v_header_size = int(original_v_header_point_size * self.current_font_scale)
                if new_v_header_size <=0: new_v_header_size = 1
                v_header_font.setPointSize(new_v_header_size)
                widget.verticalHeader().setFont(v_header_font)
                widget.verticalHeader().resizeSections(QHeaderView.ResizeToContents)
                # Similar to QTreeWidget, QTableWidgetItems need individual font updates.
        
        # Re-populate/update views that manage their own item fonts
        if self.msi_file_path: # Only if a file is loaded
            # These functions internally clear and re-add items, so they should pick up new global font scale
            # if item creation logic is updated.
            if self.tabs.currentWidget() == self.execution_tab or self.execution_tab.isVisible():
                 self.analyze_install_sequence() # This will repopulate the sequence_tree
            
            if self.tabs.currentWidget() == self.footprint_tab or self.footprint_tab.isVisible():
                 self.analyze_installation_impact() # This will repopulate impact_tree
            
            # For tables, if a table is selected, re-display it and adjust row heights
            if self.tabs.currentWidget() == self.tables_tab or self.tables_tab.isVisible():
                current_table_item = self.table_list.currentItem()
                if current_table_item:
                    self.table_selected(current_table_item, None) # Force reload of table content
                # Adjust row heights after table content might have been repopulated and fonts scaled
                if hasattr(self, 'table_content') and self.table_content.rowCount() > 0:
                    self.table_content.resizeRowsToContents()
                    for r in range(self.table_content.rowCount()):
                       self.table_content.setRowHeight(r, int(self.table_content.rowHeight(r) * 1.1)) 

        self.update_button_states() # Some button text might change
        self.updateGeometry() # Request layout recalculation
        self.update() # Schedule a repaint

    def setup_shortcuts(self):
        """Set up keyboard shortcuts for the application"""
        # Global shortcut for Ctrl+F that works on any tab
        self.global_filter_shortcut = QShortcut(QKeySequence("Ctrl+F"), self)
        self.global_filter_shortcut.setContext(Qt.ApplicationShortcut)  # Make it work anywhere in the application
        self.global_filter_shortcut.activated.connect(self.focus_current_filter)
        
        # We don't need tab-specific shortcuts since the global one will handle all cases
        # via the focus_current_filter method
        
        # Escape key to clear filters when they have focus
        if hasattr(self, 'streams_filter'):
            self.streams_filter_escape = QShortcut(QKeySequence("Escape"), self.streams_filter)
            self.streams_filter_escape.activated.connect(self.streams_filter.clear)
        
        if hasattr(self, 'table_filter'):
            self.table_filter_escape = QShortcut(QKeySequence("Escape"), self.table_filter)
            self.table_filter_escape.activated.connect(self.table_filter.clear)
        
        # Add key sequences for zoom, ensuring they don't conflict if already added by menu
        # QKeySequence.ZoomIn and ZoomOut are usually handled by QAction.setShortcut
        # Ctrl+0 is also handled by QAction.setShortcut

    def focus_current_filter(self):
        """Focus on the filter field of the currently active tab or dialog"""
        try:
            # First check if there's an active dialog that has focus
            active_window = QApplication.activeWindow()
            
            # Check for ArchivePreviewDialog or any dialog with 'contents_filter'
            if active_window and active_window != self and hasattr(active_window, 'contents_filter'):
                # This is likely an ArchivePreviewDialog
                active_window.contents_filter.setFocus()
                return
                
            # If no active dialog with filter, focus on the current tab's filter
            current_tab = self.tabs.currentWidget()
            
            if current_tab == self.streams_tab and hasattr(self, 'streams_filter'):
                self.streams_filter.setFocus()
            elif current_tab == self.tables_tab and hasattr(self, 'table_filter'):
                self.table_filter.setFocus()
            # Can add more tab types with filters here in the future
            
        except Exception as e:
            # Log error but don't disrupt the UI
            print(f"Error focusing filter: {str(e)}")
            self.statusBar().showMessage("Error focusing on filter field", 3000)
        
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
            self.load_msi_file(file_path)
            
    def load_msi_file(self, file_path):
        """Load and process the selected MSI file"""
        self.msi_file_path = file_path
        self.file_label.setText(os.path.basename(file_path))
        self.update_button_states()
        self.statusBar().showMessage(f"Selected MSI file: {file_path}")
        
        # Clear previous data if any
        self.clear_previous_data()
        
        # Auto-run metadata, streams, and tables when file is selected
        self.get_metadata()
        self.list_streams()
        self.list_tables()
        
        # Auto-analyze certificates when file is selected (no dialogs)
        self.analyze_certificate(show_dialogs=False)

    def clear_previous_data(self):
        """Clear data from previous file analysis"""
        # Clear certificate status and details
        self.cert_details.clear()
        
        # Reset extracted certificate files
        if hasattr(self, 'extracted_cert_files'):
            delattr(self, 'extracted_cert_files')
            
        # Clear tables tab
        self.table_list.clear()
        self.table_content.clear()
        self.table_content.setRowCount(0)
        self.table_content.setColumnCount(0)
        self.tables_data = None
        
        # Clear streams tab
        self.streams_tree.clear()
        self.streams_data = []
        self.streams_filter.clear() # Clear filter
        
        # Clear metadata tab
        self.metadata_text.clear()
        
        # Clear execution tab
        self.sequence_tree.clear()
        
        # Clear footprint tab
        if hasattr(self, 'impact_tree'):
            self.impact_tree.clear()
            
        self.update_button_states() # Update button states after clearing

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
        self.statusBar().showMessage("Running command...")
        
        self.thread = CommandThread(command)
        self.active_threads.append(self.thread)  # Track this thread
        self.thread.output_ready.connect(callback)
        self.thread.error_occurred.connect(lambda msg: self.handle_error("Command Error", msg))
        self.thread.finished_successfully.connect(lambda: self.command_finished(self.thread))
        self.thread.finished.connect(lambda: self.cleanup_thread(self.thread))
        self.thread.start()
        
    def command_finished(self, thread):
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
            # QMessageBox.critical(self, title, error_msg)
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle(title)
            msg_box.setText(error_msg)
            center_dialog_on_parent_screen(msg_box, self)
            msg_box.exec_()

    def show_status(self, message, timeout=0):
        """Show a status message with optional timeout"""
        self.statusBar().showMessage(message, timeout)
        
    def show_error(self, title, error, show_dialog=True, status_only=False):
        """Centralized error handling with optional dialog"""
        error_msg = str(error)
        self.show_status(f"Error: {error_msg[:100]}")
        if show_dialog and not status_only:
            # QMessageBox.critical(self, title, error_msg)
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle(title)
            msg_box.setText(error_msg)
            center_dialog_on_parent_screen(msg_box, self)
            msg_box.exec_()
            
    def show_warning(self, title, message, show_dialog=True, status_only=False):
        """Centralized warning handling with optional dialog"""
        self.show_status(message)
        if show_dialog and not status_only:
            # QMessageBox.warning(self, title, message)
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Warning)
            msg_box.setWindowTitle(title)
            msg_box.setText(message)
            center_dialog_on_parent_screen(msg_box, self)
            msg_box.exec_()

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
            # QMessageBox.critical(self, "Browser Error", f"Failed to open browser: {str(e)}")
            msg_box = QMessageBox(self)
            msg_box.setIcon(QMessageBox.Critical)
            msg_box.setWindowTitle("Browser Error")
            msg_box.setText(f"Failed to open browser: {str(e)}")
            center_dialog_on_parent_screen(msg_box, self)
            msg_box.exec_()
            self.statusBar().showMessage("Failed to open browser")

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
        
    def analyze_certificate(self, show_dialogs=True):
        return analyze_certificate(self, show_dialogs)
        
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
        from utils.gui.extraction import extract_stream
        return extract_stream(self, stream_name, output_dir, temp)
        
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
        """Get the next stream to extract"""
        if not self.streams_data:
            return None
        return self.streams_data[0] if self.streams_data else None
        
    def get_output_directory(self):
        """Get the output directory for file operations"""
        if not self.output_dir:
            self.output_dir = QFileDialog.getExistingDirectory(
                self,
                "Select Output Directory",
                self.last_output_dir if self.last_output_dir else ""
            )
            if self.output_dir:
                self.last_output_dir = self.output_dir
        return self.output_dir
        
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
            self.last_output_dir if self.last_output_dir else ""
        )
        
        if not export_dir:
            return  # User cancelled
            
        self.last_output_dir = export_dir  # Remember the last chosen directory
        
        try:
            # Count for progress tracking
            total_tables = len(self.tables_data)
            exported_count = 0
            
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
                QApplication.processEvents()  # Keep UI responsive
            
            # Update status
            self.statusBar().showMessage(f"Exported {exported_count} tables to {export_dir}")
            
            # Show success message
            # QMessageBox.information(
            #     self,
            #     "Export Successful",
            #     f"All tables have been exported as individual files to:\n{export_dir}"
            # )
            msg_box_info = QMessageBox(self)
            msg_box_info.setIcon(QMessageBox.Information)
            msg_box_info.setWindowTitle("Export Successful")
            msg_box_info.setText(f"All tables have been exported as individual files to:\n{export_dir}")
            center_dialog_on_parent_screen(msg_box_info, self)
            msg_box_info.exec_()
            
        except Exception as e:
            # Show error
            # QMessageBox.critical(
            #     self,
            #     "Export Failed",
            #     f"Failed to export tables: {str(e)}"
            # )
            msg_box_err = QMessageBox(self)
            msg_box_err.setIcon(QMessageBox.Critical)
            msg_box_err.setWindowTitle("Export Failed")
            msg_box_err.setText(f"Failed to export tables: {str(e)}")
            center_dialog_on_parent_screen(msg_box_err, self)
            msg_box_err.exec_()

    def analyze_installation_impact(self):
        """Analyze the MSI package to identify all system changes that will occur during installation"""
        analyze_installation_impact(self)
        
    def display_installation_impact(self, output):
        """Display the results of the installation impact analysis"""
        display_installation_impact(self, output)

    def show_tables_context_menu(self, position):
        """Show context menu for tables list"""
        item = self.table_list.itemAt(position)
        if item:
            menu = QMenu()
            copy_action = QAction("Copy Table Name", self)
            copy_action.triggered.connect(lambda: self.copy_table_name(item))
            menu.addAction(copy_action)
            menu.exec_(self.table_list.mapToGlobal(position))
    
    def copy_table_name(self, item):
        """Copy the table name to clipboard"""
        if item:
            table_name = item.text()
            QApplication.clipboard().setText(table_name)
            self.statusBar().showMessage(f"Copied table name: {table_name}", 2000)

    def show_impact_context_menu(self, position):
        """Show context menu for the impact tree"""
        item = self.impact_tree.itemAt(position)
        if item:
            menu = QMenu()
            
            # Add copy actions in the specified order
            copy_entry_action = QAction("Copy Entry", self)
            copy_entry_action.triggered.connect(lambda: self.copy_impact_item(item, 1))  # Copy Entry column
            menu.addAction(copy_entry_action)
            
            copy_concern_action = QAction("Copy Concern", self)
            copy_concern_action.triggered.connect(lambda: self.copy_impact_item(item, 2))  # Copy Concern column
            menu.addAction(copy_concern_action)
            
            copy_details_action = QAction("Copy Details", self)
            copy_details_action.triggered.connect(lambda: self.copy_impact_item(item, 3))  # Copy Details column
            menu.addAction(copy_details_action)
            
            # Add separator
            menu.addSeparator()
            
            copy_full_line_action = QAction("Copy Full Line", self)
            copy_full_line_action.triggered.connect(lambda: self.copy_impact_full_line(item))
            menu.addAction(copy_full_line_action)
            
            menu.exec_(self.impact_tree.mapToGlobal(position))
    
    def copy_impact_item(self, item, column):
        """Copy a specific column from the impact item"""
        if item:
            text = item.text(column)
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage(f"Copied: {text[:50]}...", 2000)
    
    def copy_impact_full_line(self, item):
        """Copy all columns from the impact item as a tab-separated line"""
        if item:
            # Skip the Type column (0) and only include Entry, Concern, and Details
            text = f"{item.text(1)}\t{item.text(2)}\t{item.text(3)}"
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage("Copied full line", 2000)

    def show_sequence_context_menu(self, position):
        """Show context menu for the sequence tree"""
        item = self.sequence_tree.itemAt(position)
        if item:
            menu = QMenu()
            
            # Add copy actions for each column
            copy_sequence_action = QAction("Copy Sequence", self)
            copy_sequence_action.triggered.connect(lambda: self.copy_sequence_item(item, 0))
            menu.addAction(copy_sequence_action)
            
            copy_action_action = QAction("Copy Action", self)
            copy_action_action.triggered.connect(lambda: self.copy_sequence_item(item, 1))
            menu.addAction(copy_action_action)
            
            copy_condition_action = QAction("Copy Condition", self)
            copy_condition_action.triggered.connect(lambda: self.copy_sequence_item(item, 2))
            menu.addAction(copy_condition_action)
            
            copy_type_action = QAction("Copy Type", self)
            copy_type_action.triggered.connect(lambda: self.copy_sequence_item(item, 3))
            menu.addAction(copy_type_action)
            
            copy_impact_action = QAction("Copy Impact", self)
            copy_impact_action.triggered.connect(lambda: self.copy_sequence_item(item, 4))
            menu.addAction(copy_impact_action)
            
            # Add separator
            menu.addSeparator()
            
            copy_full_line_action = QAction("Copy Full Line", self)
            copy_full_line_action.triggered.connect(lambda: self.copy_sequence_full_line(item))
            menu.addAction(copy_full_line_action)
            
            menu.exec_(self.sequence_tree.mapToGlobal(position))
    
    def copy_sequence_item(self, item, column):
        """Copy a specific column from the sequence item"""
        if item:
            text = item.text(column)
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage(f"Copied: {text[:50]}...", 2000)
    
    def copy_sequence_full_line(self, item):
        """Copy all columns from the sequence item as a tab-separated line"""
        if item:
            # Copy all columns as tab-separated text
            text = "\t".join(item.text(i) for i in range(5))
            QApplication.clipboard().setText(text)
            self.statusBar().showMessage("Copied full line", 2000)

    def create_metadata_tab(self):
        """Create the metadata tab"""
        metadata_tab = QWidget()
        metadata_layout = QVBoxLayout()
        metadata_tab.setLayout(metadata_layout)
        
        self.metadata_text = QTextEdit()
        self.metadata_text.setReadOnly(True)
        metadata_layout.addWidget(self.metadata_text)
        
        return metadata_tab

    def create_streams_tab(self):
        """Create the streams tab"""
        streams_tab = QWidget()
        streams_layout = QVBoxLayout()
        streams_tab.setLayout(streams_layout)
        
        streams_button_layout = QHBoxLayout()
        
        # Add Identify Streams button
        self.identify_streams_button = QPushButton("Identify Stream Types")
        self.identify_streams_button.clicked.connect(self.identify_streams)
        streams_button_layout.addWidget(self.identify_streams_button)
        
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
        self.streams_filter.setClearButtonEnabled(True)
        
        filter_layout.addWidget(filter_label)
        filter_layout.addWidget(self.streams_filter)
        streams_layout.addLayout(filter_layout)
        
        # Update streams tree to have four columns
        self.streams_tree = QTreeWidget()
        self.streams_tree.setHeaderLabels(["Stream Name", "Group", "MIME Type", "File Size", "SHA1 Hash"])
        self.streams_tree.itemClicked.connect(self.stream_selected)
        self.streams_tree.setSelectionMode(QTreeWidget.ExtendedSelection)
        
        # Connect to selectionChanged signal to handle multiple selection
        self.streams_tree.itemSelectionChanged.connect(self.on_stream_selection_changed)
        
        # Enable context menu for streams tree
        self.streams_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.streams_tree.customContextMenuRequested.connect(self.show_streams_context_menu)
        
        # Enable sorting but keep original order initially
        self.streams_tree.setSortingEnabled(True)
        self.streams_tree.header().setSortIndicator(-1, Qt.AscendingOrder)
        
        streams_layout.addWidget(self.streams_tree)
        
        return streams_tab

    def create_tables_tab(self):
        """Create the tables tab"""
        tables_tab = QWidget()
        tables_layout = QVBoxLayout()
        tables_tab.setLayout(tables_layout)
        
        # Add buttons for table export
        tables_button_layout = QHBoxLayout()
        
        self.export_selected_table_button = QPushButton("Export Selected Table")
        self.export_selected_table_button.clicked.connect(self.export_selected_table)
        self.export_selected_table_button.setEnabled(False)
        tables_button_layout.addWidget(self.export_selected_table_button)
        
        self.export_all_tables_button = QPushButton("Export All Tables")
        self.export_all_tables_button.clicked.connect(self.export_all_tables)
        self.export_all_tables_button.setEnabled(False)
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
        self.table_filter.setClearButtonEnabled(True)
        
        table_filter_layout.addWidget(table_filter_label)
        table_filter_layout.addWidget(self.table_filter)
        table_list_layout.addLayout(table_filter_layout)
        
        self.table_list = QListWidget()
        self.table_list.setMaximumWidth(200)
        
        # Enable context menu for table list
        self.table_list.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table_list.customContextMenuRequested.connect(self.show_tables_context_menu)
        
        self.table_list.currentItemChanged.connect(self.table_selected)
        table_list_layout.addWidget(self.table_list)
        
        # Add checkbox to hide empty tables
        self.hide_empty_tables_checkbox = QCheckBox("Hide Empty Tables")
        self.hide_empty_tables_checkbox.setChecked(False)
        self.hide_empty_tables_checkbox.stateChanged.connect(self.filter_tables)
        table_list_layout.addWidget(self.hide_empty_tables_checkbox)
        
        tables_splitter.addWidget(table_list_widget)
        
        # Right side - Table content
        self.table_content = QTableWidget()
        self.table_content.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        tables_splitter.addWidget(self.table_content)
        
        # Set initial splitter sizes
        tables_splitter.setSizes([200, 600])
        
        return tables_tab

    def create_certificates_tab(self):
        """Create the certificates tab"""
        certificate_tab = QWidget()
        certificate_layout = QVBoxLayout()
        certificate_tab.setLayout(certificate_layout)
        
        # Add description label
        cert_description = QLabel("Digital signatures verify the authenticity and integrity of the MSI package. Certificates are analyzed automatically when loading an MSI file. Use the Save button to extract certificate files to disk.")
        cert_description.setWordWrap(True)
        certificate_layout.addWidget(cert_description)
        
        # Add button to extract certificates
        cert_button_layout = QHBoxLayout()
        self.extract_cert_button = QPushButton("Save Digital Signatures (DER Encoded PKCS#7)")
        self.extract_cert_button.clicked.connect(self.extract_certificates)
        cert_button_layout.addWidget(self.extract_cert_button)
        
        cert_button_layout.addStretch()
        certificate_layout.addLayout(cert_button_layout)
        
        # Create a splitter for certificate information
        cert_splitter = QSplitter(Qt.Vertical)
        certificate_layout.addWidget(cert_splitter, 1)
        
        # Certificate details section
        self.cert_details = QTextEdit()
        self.cert_details.setReadOnly(True)
        cert_splitter.addWidget(self.cert_details)
        
        return certificate_tab

    def create_execution_tab(self):
        """Create the workflow analysis tab"""
        execution_tab = QWidget()
        workflow_layout = QVBoxLayout()
        execution_tab.setLayout(workflow_layout)
        
        # Add description label
        workflow_description = QLabel("Analyze the MSI installation sequence to understand the actions and processes that will be performed, highlighting potentially high-impact operations. For detailed information about MSI execution sequence, see the Help tab.")
        workflow_description.setWordWrap(True)
        workflow_layout.addWidget(workflow_description)
        
        # Tree view for actions in sequence
        self.sequence_tree = QTreeWidget()
        self.sequence_tree.setColumnCount(5)
        self.sequence_tree.setHeaderLabels(["Sequence", "Action", "Condition", "Type", "Impact"])
        self.sequence_tree.setAlternatingRowColors(True)
        self.sequence_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
        
        # Enable context menu for sequence tree
        self.sequence_tree.setContextMenuPolicy(Qt.CustomContextMenu)
        self.sequence_tree.customContextMenuRequested.connect(self.show_sequence_context_menu)
        
        workflow_layout.addWidget(self.sequence_tree, 1)
        
        return execution_tab

    def create_footprint_tab(self):
        """Create the installation impact tab"""
        return create_footprint_tab(self)

    def create_help_tab(self):
        """Create the help tab"""
        return create_help_tab()

    # --- Drag and Drop Event Handlers ---

    def dragEnterEvent(self, event):
        """Handle drag enter event"""
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            # Check if any URL is a local file
            for url in mime_data.urls():
                if url.isLocalFile():
                    event.acceptProposedAction()
                    return
        event.ignore() # Ignore non-file drags

    def dropEvent(self, event):
        """Handle drop event"""
        mime_data = event.mimeData()
        if mime_data.hasUrls():
            for url in mime_data.urls():
                if url.isLocalFile():
                    file_path = url.toLocalFile()
                    # Attempt to load any dropped file
                    self.load_msi_file(file_path)
                    event.acceptProposedAction()
                    return # Process only the first dropped file
        event.ignore()

    # --- End Drag and Drop --- 