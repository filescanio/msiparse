# ruff: noqa: E722
import subprocess
import tempfile
import shutil
from pathlib import Path
from PyQt5.QtCore import QThread, pyqtSignal, QObject
import magika
from utils.common import format_file_size, calculate_sha1
import os

class IdentifyStreamsWorker(QObject):
    """Worker object for identifying stream file types"""
    progress_updated = pyqtSignal(int, int)  # current, total
    stream_identified = pyqtSignal(str, str, str, str, str)  # stream_name, group, mime_type, file_size, sha1_hash
    error_occurred = pyqtSignal(str)
    finished = pyqtSignal()
    
    def __init__(self, parent, streams, temp_dir):
        super().__init__()
        self.parent = parent
        self.streams = streams
        self.temp_dir = temp_dir
        self.is_running = True
        # Initialize magika client
        self.magika_client = magika.Magika()
        
    def stop(self):
        """Stop the worker safely"""
        self.is_running = False

    def cleanup(self):
        """Clean up resources"""
        try:
            self.is_running = False
            if hasattr(self, 'magika_client') and self.magika_client:
                self.magika_client = None
        except Exception as e:
            print(f"Error during worker cleanup: {str(e)}")

    def __del__(self):
        """Ensure cleanup happens when the worker is deleted"""
        self.cleanup()

    def run(self):
        try:
            total_streams = len(self.streams)
            for i, stream_name in enumerate(self.streams):
                if not self.is_running:
                    break

                try:
                    # Extract stream to temp directory
                    output_dir = os.path.join(self.temp_dir, stream_name)
                    os.makedirs(output_dir, exist_ok=True)
                    
                    # Extract the stream using the unified method
                    from utils.gui.extraction import extract_stream_unified
                    extracted_file = extract_stream_unified(
                        self.parent,
                        stream_name,
                        output_dir=output_dir,
                        temp=True
                    )
                    
                    if not extracted_file or not os.path.exists(extracted_file):
                        self.error_occurred.emit(f"Failed to extract stream: {stream_name}")
                        continue

                    # Get file size
                    file_size = os.path.getsize(extracted_file)
                    file_size_str = f"{file_size:,} bytes"

                    # Calculate SHA1 hash
                    sha1_hash = calculate_sha1(extracted_file)

                    # Identify file type using Magika
                    try:
                        result = self.magika_client.identify_path(Path(extracted_file))
                        group = result.output.group
                        mime_type = result.output.mime_type
                    except Exception as e:
                        self.error_occurred.emit(f"Error identifying file type for {stream_name}: {str(e)}")
                        group = "unknown"
                        mime_type = "application/octet-stream"

                    # Emit results through signals
                    self.stream_identified.emit(stream_name, group, mime_type, file_size_str, sha1_hash)
                    self.progress_updated.emit(i + 1, total_streams)

                except Exception as e:
                    self.error_occurred.emit(f"Error processing stream {stream_name}: {str(e)}")
                    continue

            self.finished.emit()

        except Exception as e:
            self.error_occurred.emit(f"Error in identification thread: {str(e)}")
            self.finished.emit()
        finally:
            self.cleanup()

class IdentifyStreamsThread(QThread):
    """Thread for running the identify streams worker"""
    # Define signals in the thread class
    progress_updated = pyqtSignal(int, int)  # current, total
    stream_identified = pyqtSignal(str, str, str, str, str)  # stream_name, group, mime_type, file_size, sha1_hash
    error_occurred = pyqtSignal(str)
    finished = pyqtSignal()

    def __init__(self, parent, streams, temp_dir):
        super().__init__()
        self.streams = streams
        self.temp_dir = temp_dir
        self.parent = parent
        self._worker = None
        
        # Create worker in the thread's context
        self._worker = IdentifyStreamsWorker(parent, streams, temp_dir)
        
        # Connect signals
        self.started.connect(self._worker.run)
        self._worker.finished.connect(self.quit)
        self._worker.finished.connect(self._worker.deleteLater)
        self.finished.connect(self.deleteLater)
        
        # Forward signals from worker
        self._worker.progress_updated.connect(self.progress_updated)
        self._worker.stream_identified.connect(self.stream_identified)
        self._worker.error_occurred.connect(self.error_occurred)
        self._worker.finished.connect(self.finished)

    def stop(self):
        """Stop the thread and worker safely"""
        try:
            if self._worker:
                self._worker.stop()
            if not self.isFinished():
                self.quit()
                self.wait()
        except Exception as e:
            print(f"Error stopping thread: {str(e)}")

    def cleanup(self):
        """Clean up resources"""
        try:
            # Disconnect signals first
            if self._worker:
                try:
                    self.started.disconnect()
                    self._worker.finished.disconnect()
                    self._worker.progress_updated.disconnect()
                    self._worker.stream_identified.disconnect()
                    self._worker.error_occurred.disconnect()
                except Exception as e:
                    print(f"Error disconnecting signals: {str(e)}")
                
                # Stop the worker if it's still running
                self._worker.stop()
                self._worker.cleanup()
                self._worker = None
                
            # Quit and wait for the thread to finish
            if not self.isFinished():
                self.quit()
                self.wait()
                
        except Exception as e:
            print(f"Error during thread cleanup: {str(e)}")
            # Ensure thread is stopped even if cleanup fails
            try:
                if not self.isFinished():
                    self.quit()
                    self.wait()
            except:
                pass

    def __del__(self):
        """Ensure cleanup happens when the thread is deleted"""
        try:
            # Only attempt cleanup if the C++ object still exists
            if not self.isFinished():
                self.cleanup()
        except:
            pass  # Ignore any errors during final cleanup
