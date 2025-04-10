from PyQt5.QtWidgets import QDialog, QVBoxLayout, QPushButton, QLabel

class BasePreviewDialog(QDialog):
    """Base class for preview dialogs with common functionality"""
    def __init__(self, parent, title, content_widget):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setGeometry(100, 100, 800, 600)
        
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
        
    def set_status(self, text):
        """Update status label text"""
        self.status_label.setText(text)
