from PyQt5.QtWidgets import QTextEdit, QApplication
from PyQt5.QtGui import QFont
from dialogs.base import BasePreviewDialog

class HexViewDialog(BasePreviewDialog):
    """Dialog for displaying hex view of stream content"""
    def __init__(self, parent, stream_name, content):
        # Create hex view widget
        self.hex_view = QTextEdit()
        self.hex_view.setReadOnly(True)
        self.hex_view.setLineWrapMode(QTextEdit.NoWrap)
        
        # Use monospaced font
        font = QFont("Courier New", 10)
        font.setFixedPitch(True)
        self.hex_view.setFont(font)
        
        # Initialize base dialog
        super().__init__(parent, f"Hex View: {stream_name}", self.hex_view)
        
        # Format and display content
        self.format_hex_view(content)
        
    def format_hex_view(self, content):
        """Format the content as a hex view"""
        num_bytes = len(content)
        num_rows = (num_bytes + 15) // 16
        
        # Limit display for large files
        display_limit = 10000
        if num_rows > display_limit:
            self.set_status(f"File size: {num_bytes} bytes (showing first {display_limit * 16} bytes)")
            num_rows = display_limit
        else:
            self.set_status(f"File size: {num_bytes} bytes")
        
        # Create header
        header = "Offset   | 00 01 02 03 04 05 06 07  08 09 0A 0B 0C 0D 0E 0F | ASCII\n"
        header += "-" * 79 + "\n"
        self.hex_view.setText(header)
        QApplication.processEvents()
        
        # Process in chunks
        chunk_size = 1000
        for chunk_start in range(0, num_rows, chunk_size):
            chunk_end = min(chunk_start + chunk_size, num_rows)
            chunk_content = ""
            
            for row in range(chunk_start, chunk_end):
                offset = row * 16
                line = f"{offset:08X} | "
                ascii_text = ""
                
                for col in range(16):
                    byte_pos = row * 16 + col
                    if byte_pos < num_bytes:
                        byte_val = content[byte_pos]
                        line += f"{byte_val:02X} "
                        if col == 7:
                            line += " "
                        ascii_text += chr(byte_val) if 32 <= byte_val <= 126 else "."
                    else:
                        line += "   "
                        if col == 7:
                            line += " "
                        ascii_text += " "
                
                line += "| " + ascii_text + "\n"
                chunk_content += line
            
            current_text = self.hex_view.toPlainText()
            self.hex_view.setText(current_text + chunk_content)
            QApplication.processEvents()
        
        # Move cursor to start
        cursor = self.hex_view.textCursor()
        cursor.movePosition(cursor.Start)
        self.hex_view.setTextCursor(cursor)
