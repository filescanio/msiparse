from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QLabel, QApplication
from utils.gui.main_window import center_dialog_on_parent_screen

class BasePreviewDialog(QDialog):
    """Base class for preview dialogs with common functionality"""
    def __init__(self, parent, title, content_widget):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.resize(800, 600)
        
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
        
    def set_status(self, text):
        """Update status label text"""
        self.status_label.setText(text)
