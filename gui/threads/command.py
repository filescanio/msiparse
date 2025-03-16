import subprocess
from PyQt5.QtCore import QThread, pyqtSignal

class CommandThread(QThread):
    """Thread for running msiparse commands without freezing the GUI"""
    output_ready = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    finished_successfully = pyqtSignal()
    
    def __init__(self, command):
        super().__init__()
        self.command = command
        self.running = True
        
    def run(self):
        try:
            if not self.running:
                return
                
            result = subprocess.run(
                self.command,
                capture_output=True,
                text=True,
                check=True
            )
            if self.running:
                self.output_ready.emit(result.stdout)
                self.finished_successfully.emit()
        except subprocess.CalledProcessError as e:
            if self.running:
                self.error_occurred.emit(f"Command failed with error: {e.stderr}")
        except Exception as e:
            if self.running:
                self.error_occurred.emit(f"Error: {str(e)}")
                
    def stop(self):
        """Stop the thread safely"""
        self.running = False
