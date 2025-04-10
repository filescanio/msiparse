from PyQt5.QtWidgets import QTextEdit, QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QWidget, QPushButton
from PyQt5.QtGui import QFont
from PyQt5.QtCore import QUrl
from PyQt5.QtGui import QDesktopServices
from dialogs.base import BasePreviewDialog
from utils.gui.syntax_highlighter import CodeSyntaxHighlighter, detect_language
import base64

class TextPreviewDialog(BasePreviewDialog):
    """Dialog for displaying text content in a read-only format with syntax highlighting"""
    def __init__(self, parent, stream_name, content, mime_type=None):
        # Create text edit widget
        self.text_edit = QTextEdit()
        self.text_edit.setReadOnly(True)
        self.text_edit.setLineWrapMode(QTextEdit.WidgetWidth)
        self.text_edit.setText(content)
        
        # Set monospaced font for better code readability
        font = QFont("Courier New", 10)
        font.setFixedPitch(True)
        self.text_edit.setFont(font)
        
        # Create a container widget with layout
        container_widget = QWidget()
        container = QVBoxLayout(container_widget)
        
        # Add language selector at the top
        language_layout = QHBoxLayout()
        language_layout.addWidget(QLabel("Language:"))
        
        self.language_selector = QComboBox()
        self.language_selector.addItems(["auto-detect", "generic", "python", "javascript", "vbscript", "powershell", "xml", "html", "batch"])
        self.language_selector.currentTextChanged.connect(self.update_syntax_highlighting)
        language_layout.addWidget(self.language_selector)
        
        # Add CyberChef button
        self.cyberchef_button = QPushButton("Open in CyberChef")
        self.cyberchef_button.clicked.connect(self.open_in_cyberchef)
        # Disable button if content is too large
        if len(content.encode('utf-8')) > 8 * 1024:  # 8kB limit
            self.cyberchef_button.setEnabled(False)
            self.cyberchef_button.setToolTip("Content too large (>8kB)")
        language_layout.addWidget(self.cyberchef_button)
        
        language_layout.addStretch()
        
        container.addLayout(language_layout)
        container.addWidget(self.text_edit)
        
        # Initialize base dialog with the container widget
        super().__init__(parent, f"Text Preview: {stream_name}", container_widget)
        
        # Set status
        self.set_status(f"Length: {len(content)} characters")
        
        # Store content and MIME type
        self.content = content
        self.mime_type = mime_type
        
        # Detect language and apply syntax highlighting
        detected_language = detect_language(content, mime_type)
        self.language_selector.setCurrentText(detected_language)
        self.apply_syntax_highlighting(detected_language)
    
    def update_syntax_highlighting(self, language):
        """Update syntax highlighting when language selection changes"""
        if language == "auto-detect":
            language = detect_language(self.content, self.mime_type)
        self.apply_syntax_highlighting(language)
    
    def apply_syntax_highlighting(self, language):
        """Apply syntax highlighting for the specified language"""
        # Create and apply syntax highlighter
        self.highlighter = CodeSyntaxHighlighter(self.text_edit.document(), language)
        self.set_status(f"Length: {len(self.content)} characters | Language: {language}")
        
    def open_in_cyberchef(self):
        """Open the current text content in CyberChef"""
        # Check size again to be safe
        if len(self.content.encode('utf-8')) > 8 * 1024:  # 8kB limit
            return
            
        # Base64 encode the text
        encoded_text = base64.b64encode(self.content.encode('utf-8')).decode('ascii').rstrip('=')
        
        # Construct the CyberChef URL
        cyberchef_url = f"https://cyberchef.org/#input={encoded_text}"
        
        # Open URL in default browser
        QDesktopServices.openUrl(QUrl(cyberchef_url))