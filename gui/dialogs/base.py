from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QLabel, QApplication, QMainWindow
from utils.gui.main_window import center_dialog_on_parent_screen
from utils.gui.helpers import apply_scaling_to_dialog

class BasePreviewDialog(QDialog):
    """Base class for preview dialogs with common functionality"""
    def __init__(self, parent, title, content_widget):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(800, 600)
        self._original_widget_fonts = {}
        
        # Main layout
        layout = QVBoxLayout()
        self.setLayout(layout)
        
        # Add the main content widget
        layout.addWidget(content_widget)
        
        # Status label
        self.status_label = QLabel()
        layout.addWidget(self.status_label)
        
        # Close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(self.accept)
        layout.addWidget(close_button)
        
        # Center the dialog on the parent's screen.
        # The center_dialog_on_parent_screen function handles cases where parent is None.
        center_dialog_on_parent_screen(self, parent)
        
        # Apply scaling by finding the main application window and its scale factor
        scale_factor = 1.0  # Default scale
        for widget in QApplication.topLevelWidgets():
            if isinstance(widget, QMainWindow) and hasattr(widget, 'current_font_scale'):
                scale_factor = widget.current_font_scale
                break # Assume this is the main window we're looking for
        
        apply_scaling_to_dialog(self, scale_factor, self._original_widget_fonts)
        
    def set_status(self, text):
        """Update status label text"""
        self.status_label.setText(text)
