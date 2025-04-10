from PyQt5.QtWidgets import QScrollArea, QLabel
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QPixmap, QImage
from dialogs.base import BasePreviewDialog

class ImagePreviewDialog(BasePreviewDialog):
    """Dialog for displaying image preview"""
    def __init__(self, parent, stream_name, image_path):
        # Create scroll area and image label
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        scroll_area.setWidget(self.image_label)
        
        # Initialize base dialog
        super().__init__(parent, f"Image Preview: {stream_name}", scroll_area)
        self.set_status("Loading image...")
        
        # Load the image
        self.load_image(image_path)
        
    def load_image(self, image_path):
        """Load and display the image"""
        try:
            image = QImage(image_path)
            if image.isNull():
                self.image_label.setText("Failed to load image")
                self.set_status("Error: Invalid image")
                return
                
            self.image_label.setPixmap(QPixmap.fromImage(image))
            self.set_status(f"{self.get_format_name(image.format())} | {image.width()}x{image.height()} | {image.depth()}b")
            
        except Exception as e:
            self.image_label.setText(f"Error: {str(e)}")
            self.set_status("Failed to load image")
            
    def get_format_name(self, format_id):
        formats = {
            QImage.Format_RGB32: "RGB32",
            QImage.Format_ARGB32: "ARGB32",
            QImage.Format_RGB888: "RGB888",
            QImage.Format_RGBA8888: "RGBA8888"
        }
        return formats.get(format_id, f"Format {format_id}")
