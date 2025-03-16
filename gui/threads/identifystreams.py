import subprocess
import tempfile
import shutil
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal
import magika
from utils.common import format_file_size, calculate_sha1

class IdentifyStreamsThread(QThread):
    """Thread for identifying stream file types"""
    progress_updated = pyqtSignal(int, int)  # current, total
    stream_identified = pyqtSignal(str, str, str, str, str)  # stream_name, group, mime_type, file_size, sha1_hash
    finished_successfully = pyqtSignal()
    error_occurred = pyqtSignal(str)
    
    def __init__(self, msiparse_path, msi_file_path, streams):
        super().__init__()
        self.msiparse_path = msiparse_path
        self.msi_file_path = msi_file_path
        self.streams = streams
        self.running = True
        
        # Initialize magika
        self.magika_client = magika.Magika()
        
    def run(self):
        if not self.running:
            return
            
        # Create a temporary directory for stream extraction
        temp_dir = tempfile.mkdtemp()
        
        try:
            total_streams = len(self.streams)
            
            for i, stream_name in enumerate(self.streams):
                if not self.running:
                    break
                    
                # Update progress
                self.progress_updated.emit(i + 1, total_streams)
                
                # Extract the stream to the temp directory
                try:
                    command = [
                        self.msiparse_path,
                        "extract",
                        self.msi_file_path,
                        temp_dir,
                        stream_name
                    ]
                    
                    subprocess.run(
                        command,
                        capture_output=True,
                        text=True,
                        check=True
                    )
                    
                    # Path to the extracted file
                    file_path = Path(temp_dir) / stream_name
                    
                    # Check if file exists
                    if file_path.exists():
                        try:
                            # Get file size
                            file_size = file_path.stat().st_size
                            file_size_str = format_file_size(file_size)
                            
                            # Calculate SHA1 hash
                            sha1_hash = calculate_sha1(file_path)
                            
                            # Identify file type using magika with Path object
                            result = self.magika_client.identify_path(file_path)
                            mime_type = result.output.mime_type
                            group = result.output.group
                            
                            # Emit the result with group and SHA1 hash
                            self.stream_identified.emit(stream_name, group, mime_type, file_size_str, sha1_hash)
                        except Exception as e:
                            self.stream_identified.emit(stream_name, "unknown", f"Error: {str(e)[:50]}", "Unknown", "")
                        
                        # Delete the temporary file
                        try:
                            file_path.unlink()
                        except:
                            pass
                    else:
                        self.stream_identified.emit(stream_name, "unknown", "Error: File not extracted", "Unknown", "")
                except Exception as e:
                    # Continue with next stream if one fails
                    self.stream_identified.emit(stream_name, "unknown", f"Error: {str(e)[:50]}", "Unknown", "")
                    
            if self.running:
                self.finished_successfully.emit()
                
        except Exception as e:
            if self.running:
                self.error_occurred.emit(f"Error during identification: {str(e)}")
        finally:
            # Clean up the temporary directory
            try:
                shutil.rmtree(temp_dir)
            except:
                pass
                
    def stop(self):
        """Stop the thread safely"""
        self.running = False
