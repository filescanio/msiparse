"""
Stream extraction functionality for the MSI Parser GUI
"""

import os
import subprocess
import tempfile
from PyQt5.QtWidgets import QMessageBox, QFileDialog

def extract_file_to_temp(parent, stream_name, temp_dir):
    """Extract a stream to a temporary directory and return the file path"""
    return extract_stream(parent, stream_name, temp_dir, temp=False)

def extract_file_safe(parent, stream_name, output_dir=None, temp=False):
    """Safe file extraction with proper error handling"""
    return extract_stream(parent, stream_name, output_dir, temp)

def extract_single_stream(parent, stream_name):
    """Extract a single stream to a user-specified location"""
    if not parent.msi_file_path:
        return
        
    output_dir = QFileDialog.getExistingDirectory(
        parent, 
        "Select Output Directory", 
        parent.last_output_dir if parent.last_output_dir else ""
    )
    if not output_dir:
        return
        
    parent.last_output_dir = output_dir
    file_path = extract_stream(parent, stream_name, output_dir)
    
    if file_path:
        QMessageBox.information(
            parent,
            "Extraction Complete",
            f"Stream '{stream_name}' has been extracted to:\n{output_dir}"
        )

def extract_stream(parent, stream_name, output_dir=None, temp=False):
    """
    Extract a stream from the MSI file.
    
    Args:
        parent: The parent window
        stream_name: Name of the stream to extract
        output_dir: Directory to extract to (if None and temp=True, creates temp dir)
        temp: Whether to create a temporary directory if output_dir is None
        
    Returns:
        Path to the extracted file or None if extraction failed
    """
    if not parent.msi_file_path:
        return None
        
    if temp and not output_dir:
        output_dir = tempfile.mkdtemp()
        
    try:
        parent.progress_bar.setVisible(True)
        parent.statusBar().showMessage(f"Extracting stream: {stream_name}")
        
        command = [parent.msiparse_path, "extract", parent.msi_file_path, output_dir, stream_name]
        subprocess.run(command, capture_output=True, text=True, check=True)
        
        file_path = os.path.join(output_dir, stream_name)
        if os.path.exists(file_path):
            parent.statusBar().showMessage(f"Stream '{stream_name}' extracted successfully")
            return file_path
            
        parent.statusBar().showMessage(f"Failed to extract stream: {stream_name}")
        QMessageBox.warning(parent, "Error", f"Failed to extract stream: {stream_name}")
        return None
            
    except Exception as e:
        parent.statusBar().showMessage("Error during extraction")
        QMessageBox.critical(parent, "Error", f"Error extracting file: {str(e)}")
        return None
    finally:
        parent.progress_bar.setVisible(False)

def extract_streams(parent, stream_names=None):
    """
    Extract one or more streams from the MSI file.
    If stream_names is None, extracts all streams.
    """
    if not parent.msi_file_path:
        parent.show_error("Error", "No MSI file selected")
        return
        
    output_dir = QFileDialog.getExistingDirectory(
        parent, 
        "Select Output Directory", 
        parent.last_output_dir if parent.last_output_dir else ""
    )
    if not output_dir:
        return
        
    parent.last_output_dir = output_dir
    parent.progress_bar.setVisible(True)
    
    if stream_names is None:
        # Extract all streams
        command = [parent.msiparse_path, "extract_all", parent.msi_file_path, output_dir]
        parent.run_command(command, lambda output: handle_extraction_complete(parent, output_dir))
    else:
        # Extract selected streams
        parent.progress_bar.setRange(0, len(stream_names))
        parent.progress_bar.setValue(0)
        parent.statusBar().showMessage(f"Extracting {len(stream_names)} streams...")
        
        errors = []
        for i, name in enumerate(stream_names):
            parent.progress_bar.setValue(i + 1)
            parent.statusBar().showMessage(f"Extracting stream {i + 1}/{len(stream_names)}: {name}")
            
            if not extract_stream(parent, name, output_dir):
                errors.append(f"Failed to extract '{name}'")
        
        handle_extraction_complete(parent, output_dir, errors)

def handle_extraction_complete(parent, output_dir, errors=None):
    """Handle completion of extraction process"""
    parent.progress_bar.setVisible(False)
    
    if errors:
        error_msg = f"Completed with {len(errors)} errors:\n\n"
        for i, error in enumerate(errors[:5]):
            error_msg += f"{i+1}. {error}\n"
        if len(errors) > 5:
            error_msg += f"\n...and {len(errors) - 5} more errors."
        QMessageBox.warning(parent, "Extraction Completed with Errors", error_msg)
    else:
        parent.show_status(parent.STATUS_MESSAGES['extract_complete'])
        QMessageBox.information(
            parent, 
            "Extraction Complete", 
            f"All files have been extracted to:\n{output_dir}"
        )

# Aliases for backward compatibility
extract_stream_unified = extract_stream
extract_all_streams = extract_streams
handle_extraction_all_complete = handle_extraction_complete
extract_multiple_streams = extract_streams 