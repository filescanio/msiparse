import subprocess
from PyQt5.QtCore import QThread, pyqtSignal
from utils.subprocess_utils import run_subprocess

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
            # Use the utility function
            result = run_subprocess(
                self.command,
                capture_output=True,
                text=True,
                check=True,
            )
            self.output_ready.emit(result.stdout)
            self.finished_successfully.emit()
        except subprocess.CalledProcessError as e:
            # Check if stderr exists before trying to access it
            error_message = f"Command failed with exit code {e.returncode}"
            if e.stderr:
                error_message += f": {e.stderr}"
            self.error_occurred.emit(error_message)
        except Exception as e:
            self.error_occurred.emit(f"Error: {str(e)}")
