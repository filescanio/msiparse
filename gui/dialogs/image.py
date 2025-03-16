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
                self.set_status("Error: Invalid or unsupported image format")
                return
                
            width = image.width()
            height = image.height()
            pixmap = QPixmap.fromImage(image)
            self.image_label.setPixmap(pixmap)
            
            format_name = self.get_format_name(image.format())
            depth = image.depth()
            self.set_status(f"Format: {format_name} | Size: {width}x{height} | Depth: {depth} bits")
            
        except Exception as e:
            self.image_label.setText(f"Error loading image: {str(e)}")
            self.set_status("Error occurred while loading the image")
            
    def get_format_name(self, format_id):
        """Convert QImage format ID to a readable name"""
        format_names = {
            QImage.Format_Invalid: "Invalid",
            QImage.Format_Mono: "Mono",
            QImage.Format_MonoLSB: "MonoLSB",
            QImage.Format_Indexed8: "Indexed8",
            QImage.Format_RGB32: "RGB32",
            QImage.Format_ARGB32: "ARGB32",
            QImage.Format_ARGB32_Premultiplied: "ARGB32_Premultiplied",
            QImage.Format_RGB16: "RGB16",
            QImage.Format_ARGB8565_Premultiplied: "ARGB8565_Premultiplied",
            QImage.Format_RGB666: "RGB666",
            QImage.Format_ARGB6666_Premultiplied: "ARGB6666_Premultiplied",
            QImage.Format_RGB555: "RGB555",
            QImage.Format_ARGB8555_Premultiplied: "ARGB8555_Premultiplied",
            QImage.Format_RGB888: "RGB888",
            QImage.Format_RGB444: "RGB444",
            QImage.Format_ARGB4444_Premultiplied: "ARGB4444_Premultiplied",
            QImage.Format_RGBX8888: "RGBX8888",
            QImage.Format_RGBA8888: "RGBA8888",
            QImage.Format_RGBA8888_Premultiplied: "RGBA8888_Premultiplied"
        }
        return format_names.get(format_id, f"Unknown ({format_id})")
