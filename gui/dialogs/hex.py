from PyQt5.QtWidgets import QTextEdit, QApplication, QMessageBox
from PyQt5.QtGui import QFont
from dialogs.base import BasePreviewDialog

class HexViewDialog(BasePreviewDialog):
    """Dialog for displaying hex view of stream content"""
    def __init__(self, parent, file_name, file_path_or_content):
        # Create hex view widget
        self.hex_view = QTextEdit()
        self.hex_view.setReadOnly(True)
        self.hex_view.setLineWrapMode(QTextEdit.NoWrap)
        
        # Use monospaced font
        self.hex_view.setFont(QFont("Courier New", 10, QFont.Monospace))
        
        # Initialize base dialog
        super().__init__(parent, f"Hex View: {file_name}", self.hex_view)
        
        try:
            content = file_path_or_content if isinstance(file_path_or_content, (bytes, bytearray)) else open(file_path_or_content, 'rb').read()
            self.format_hex_view(content)
        except Exception as e:
            self.set_status(f"Error: {str(e)}")
            QMessageBox.critical(self, "Error", str(e))
        
    def format_hex_view(self, content):
        """Format the content as a hex view"""
        num_bytes = len(content)
        display_limit = 10000
        num_rows = min((num_bytes + 15) // 16, display_limit)
        
        self.set_status(f"File size: {num_bytes} bytes" + 
                       (f" (showing first {display_limit * 16} bytes)" if num_rows == display_limit else ""))
        
        # Create header
        header = "Offset   | 00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F | ASCII\n"
        header += "-" * 79 + "\n"
        self.hex_view.setText(header)
        
        # Process in chunks
        for row in range(num_rows):
            offset = row * 16
            hex_values = " ".join(f"{content[offset + i]:02X}" if offset + i < num_bytes else "  " 
                                for i in range(16))
            ascii_values = "".join(chr(content[offset + i]) if 32 <= content[offset + i] <= 126 else "."
                                 for i in range(16) if offset + i < num_bytes)
            
            self.hex_view.append(f"{offset:08X} | {hex_values[:23]} {hex_values[24:]} | {ascii_values}")
            if row % 100 == 0:
                QApplication.processEvents()  # Keep UI responsive
