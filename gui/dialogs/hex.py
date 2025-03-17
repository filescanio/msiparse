from PyQt5.QtWidgets import QTextEdit, QApplication, QMessageBox
from PyQt5.QtGui import QFont
from dialogs.base import BasePreviewDialog
import os

class HexViewDialog(BasePreviewDialog):
    """Dialog for displaying hex view of stream content"""
    def __init__(self, parent, file_name, file_path_or_content):
        # Create hex view widget
        self.hex_view = QTextEdit()
        self.hex_view.setReadOnly(True)
        self.hex_view.setLineWrapMode(QTextEdit.NoWrap)
        
        # Use monospaced font
        font = QFont("Courier New", 10)
        font.setFixedPitch(True)
        self.hex_view.setFont(font)
        
        # Initialize base dialog
        super().__init__(parent, f"Hex View: {file_name}", self.hex_view)
        
        # Determine if we have a file path or content
        if isinstance(file_path_or_content, (bytes, bytearray)):
            # We have binary content directly
            self.format_hex_view(file_path_or_content)
        else:
            # We have a file path, read the content
            try:
                if os.path.exists(file_path_or_content):
                    with open(file_path_or_content, 'rb') as f:
                        content = f.read()
                    self.format_hex_view(content)
                else:
                    self.set_status(f"Error: File not found: {file_path_or_content}")
                    QMessageBox.critical(self, "Error", f"File not found: {file_path_or_content}")
            except Exception as e:
                self.set_status(f"Error reading file: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to read file: {str(e)}")
        
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
                
                # Format offset
                line = f"{offset:08X} | "
                
                # Format hex values
                hex_values = ""
                ascii_values = ""
                
                for col in range(16):
                    pos = offset + col
                    if pos < num_bytes:
                        byte = content[pos]
                        hex_values += f"{byte:02X} "
                        
                        # Add extra space after 8 bytes
                        if col == 7:
                            hex_values += " "
                            
                        # Format ASCII representation
                        if 32 <= byte <= 126:  # Printable ASCII
                            ascii_values += chr(byte)
                        else:
                            ascii_values += "."
                    else:
                        # Padding for incomplete rows
                        hex_values += "   "
                        if col == 7:
                            hex_values += " "
                        ascii_values += " "
                
                # Combine parts
                line += hex_values + "| " + ascii_values + "\n"
                chunk_content += line
            
            # Append chunk to display
            self.hex_view.append(chunk_content)
            QApplication.processEvents()  # Keep UI responsive
