from PyQt5.QtWidgets import QVBoxLayout, QLabel, QScrollArea, QWidget
from PyQt5.QtGui import QPixmap
from PyQt5.QtCore import Qt
from dialogs.base import BasePreviewDialog
import os
import tempfile

class PDFPreviewDialog(BasePreviewDialog):
    """Simple dialog for displaying PDF content using PyMuPDF (fitz)"""
    def __init__(self, parent, file_name, file_path):
        # Create scroll area for the PDF page
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        
        # Create image label
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        self.scroll_area.setWidget(self.image_label)
        
        # Create container
        container = QVBoxLayout()
        container.addWidget(self.scroll_area)
        
        # Create a widget to hold the layout
        container_widget = QWidget()
        container_widget.setLayout(container)
        
        # Initialize base dialog
        super().__init__(parent, f"PDF Preview: {file_name}", container_widget)
        
        # Set status initially
        self.set_status(f"Loading PDF: {file_name}")
        
        # Try to render the PDF with PyMuPDF
        try:
            # Import PyMuPDF
            import fitz
            
            # Create a temporary file for the output
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                temp_png = temp_file.name
            
            # Open the PDF file
            doc = fitz.open(file_path)
            
            # Load the first page
            page = doc.load_page(0)
            
            # Render page to pixmap
            pix = page.get_pixmap()
            
            # Save pixmap to temporary PNG file
            pix.save(temp_png)
            
            # Close the document
            doc.close()
            
            # Display the image
            pixmap = QPixmap(temp_png)
            if not pixmap.isNull():
                # Scale pixmap if needed
                screen_size = self.screen().size()
                max_width = min(pixmap.width(), screen_size.width() - 100)
                max_height = min(pixmap.height(), screen_size.height() - 100)
                
                if pixmap.width() > max_width or pixmap.height() > max_height:
                    pixmap = pixmap.scaled(
                        max_width, max_height,
                        Qt.KeepAspectRatio, 
                        Qt.SmoothTransformation
                    )
                
                # Set the image to the label
                self.image_label.setPixmap(pixmap)
                self.set_status(f"Showing first page of: {file_name}")
            else:
                self.set_status("Failed to create PDF preview image")
                self.image_label.setText("Failed to create PDF preview")
            
            # Clean up the temporary file
            try:
                if os.path.exists(temp_png):
                    os.unlink(temp_png)
            except:
                pass
                
        except ImportError:
            self.set_status("PyMuPDF not available")
            self.image_label.setText("PDF preview requires PyMuPDF\n\n"
                                   "To enable PDF preview, install:\n"
                                   "pip install PyMuPDF")
        except Exception as e:
            self.set_status(f"Error: {str(e)}")
            self.image_label.setText(f"Failed to preview PDF: {str(e)}\n\n"
                                   "Please make sure PyMuPDF is installed correctly.") 