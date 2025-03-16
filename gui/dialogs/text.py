from PyQt5.QtWidgets import QTextEdit
from dialogs.base import BasePreviewDialog

class TextPreviewDialog(BasePreviewDialog):
    """Dialog for displaying text preview"""
    def __init__(self, parent, stream_name, content):
        # Create text edit widget
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.text_edit.setText(content)
        
        # Initialize base dialog
        super().__init__(parent, f"Text Preview: {stream_name}", self.text_edit)
        
        # Set status
        self.set_status(f"Length: {len(content)} characters")