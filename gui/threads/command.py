import subprocess
from PyQt5.QtCore import QThread, pyqtSignal

class CommandThread(QThread):
    """Thread for running msiparse commands without freezing the GUI.
    
    Args:
        command: List of command arguments to execute
    """
    output_ready = pyqtSignal(str)
    error_occurred = pyqtSignal(str)
    finished_successfully = pyqtSignal()
    
    def __init__(self, command):
        super().__init__()
        self.command = command
        
    def run(self):
        try:
            result = subprocess.run(
                self.command,
                capture_output=True,
                text=True,
                check=True
            )
            self.output_ready.emit(result.stdout)
            self.finished_successfully.emit()
        except subprocess.CalledProcessError as e:
            self.error_occurred.emit(f"Command failed with error: {e.stderr}")
        except Exception as e:
            self.error_occurred.emit(f"Error: {str(e)}")
