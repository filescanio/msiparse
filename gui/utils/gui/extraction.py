"""
Stream extraction functionality for the MSI Parser GUI
"""

import os
import subprocess
import tempfile
from PyQt5.QtWidgets import QMessageBox, QFileDialog

def extract_stream_unified(parent, stream_name, output_dir=None, temp=False, show_messages=True):
    """
    Unified method for extracting streams with consistent error handling.
    
    Args:
        parent: The parent window
        stream_name: Name of the stream to extract
        output_dir: Directory to extract to (if None and temp=True, creates temp dir)
        temp: Whether to create a temporary directory if output_dir is None
        show_messages: Whether to show status messages and dialogs
        
    Returns:
        Path to the extracted file or None if extraction failed
    """
    if not parent.msi_file_path:
        return None
        
    # Create temp directory if needed
    if temp and not output_dir:
        output_dir = tempfile.mkdtemp()
        
    try:
        # Show progress if requested
        if show_messages:
            parent.progress_bar.setVisible(True)
            parent.statusBar().showMessage(f"Extracting stream: {stream_name}")
        
        # Create and run the command
        command = [
            parent.msiparse_path,
            "extract",
            parent.msi_file_path,
            output_dir,
            stream_name
        ]
        
        result = subprocess.run(
            command,
            capture_output=True,
            text=True,
            check=True
        )
        
        # Path to the extracted file
        file_path = os.path.join(output_dir, stream_name)
        
        # Check if file exists
        if os.path.exists(file_path):
            if show_messages:
                parent.statusBar().showMessage(f"Stream '{stream_name}' extracted successfully")
            return file_path
        else:
            if show_messages:
                parent.statusBar().showMessage(f"Failed to extract stream: {stream_name}")
                QMessageBox.warning(parent, "Error", f"Failed to extract stream: {stream_name}")
            return None
            
    except Exception as e:
        if show_messages:
            parent.statusBar().showMessage("Error during extraction")
            QMessageBox.critical(parent, "Error", f"Error extracting file: {str(e)}")
        return None
    finally:
        # Hide progress bar if shown
        if show_messages:
            parent.progress_bar.setVisible(False)

def extract_file_to_temp(parent, stream_name, temp_dir):
    """Extract a stream to a temporary directory and return the file path"""
    return extract_stream_unified(parent, stream_name, temp_dir, temp=False, show_messages=True)

def extract_file_safe(parent, stream_name, output_dir=None, temp=False):
    """Safe file extraction with proper error handling and progress indication"""
    return extract_stream_unified(parent, stream_name, output_dir, temp)

def extract_single_stream(parent, stream_name):
    """Extract a single stream to a user-specified location"""
    if not parent.msi_file_path:
        return
        
    # Prompt for output directory
    output_dir = get_output_directory(parent)
    if not output_dir:
        return  # User cancelled
        
    # Use the unified extraction method
    file_path = extract_stream_unified(parent, stream_name, output_dir)
    
    # Show success message if extraction succeeded
    if file_path:
        QMessageBox.information(
            parent,
            "Extraction Complete",
            f"Stream '{stream_name}' has been extracted to:\n{output_dir}"
        )

def extract_all_streams(parent):
    """Extract all streams from the MSI file"""
    if not parent.msi_file_path:
        parent.show_error("Error", "No MSI file selected")
        return
        
    # Get output directory
    output_dir = get_output_directory(parent)
    if not output_dir:
        return
        
    # Show progress
    parent.progress_bar.setVisible(True)
    
    # Build command
    command = [
        parent.msiparse_path,
        "extract_all",
        parent.msi_file_path,
        output_dir
    ]
    
    # Run command
    parent.run_command(command, lambda output: handle_extraction_all_complete(parent, output_dir))

def handle_extraction_all_complete(parent, output_dir):
    """Handle completion of extract all command"""
    parent.progress_bar.setVisible(False)
    
    # Show success message
    QMessageBox.information(
        parent,
        "Extraction Complete",
        f"All streams have been extracted to:\n{output_dir}"
    )

def extract_stream(parent):
    """Extract selected streams"""
    if not parent.msi_file_path:
        return
        
    selected_items = parent.streams_tree.selectedItems()
    if not selected_items:
        parent.show_warning("Warning", "Please select streams to extract")
        return
        
    stream_names = [item.text(0) for item in selected_items]
    output_dir = get_output_directory(parent)
    if not output_dir:
        return  # User cancelled
    
    try:
        # Show progress
        parent.progress_bar.setRange(0, len(stream_names))
        parent.progress_bar.setValue(0)
        parent.progress_bar.setVisible(True)
        parent.statusBar().showMessage(f"Extracting {len(stream_names)} streams...")
        
        # Setup extraction
        parent.current_extraction_index = 0
        parent.extraction_commands = [
            [parent.msiparse_path, "extract", parent.msi_file_path, output_dir, name]
            for name in stream_names
        ]
        parent.extraction_output_dir = output_dir
        parent.extraction_errors = []
        
        # Start extraction process
        extract_next_stream(parent)
    except Exception as e:
        parent.progress_bar.setVisible(False)
        parent.handle_error("Extraction Setup Error", f"Failed to setup extraction: {str(e)}", show_dialog=True)

def extract_multiple_streams(parent, stream_names, output_dir):
    """Extract multiple streams sequentially"""
    try:
        # Create commands for each stream
        commands = [
            [parent.msiparse_path, "extract", parent.msi_file_path, output_dir, name]
            for name in stream_names
        ]
        
        parent.current_extraction_index = 0
        parent.extraction_commands = commands
        parent.extraction_output_dir = output_dir
        parent.extraction_errors = []
        
        # Start extraction process
        extract_next_stream(parent)
    except Exception as e:
        parent.progress_bar.setVisible(False)
        parent.handle_error("Extraction Setup Error", f"Failed to setup extraction: {str(e)}", show_dialog=True)

def extract_next_stream(parent):
    """Extract the next stream in the queue"""
    try:
        # Check if we've processed all streams
        if parent.current_extraction_index >= len(parent.extraction_commands):
            parent.progress_bar.setVisible(False)
            
            # Show completion message
            if not parent.extraction_errors:
                parent.show_status(parent.STATUS_MESSAGES['extract_complete'])
                QMessageBox.information(
                    parent, 
                    "Extraction Complete", 
                    f"All files have been extracted to:\n{parent.extraction_output_dir}"
                )
            else:
                # Show error message if there were any errors
                error_msg = f"Completed with {len(parent.extraction_errors)} errors:\n\n"
                for i, error in enumerate(parent.extraction_errors[:5]):  # Show first 5 errors
                    error_msg += f"{i+1}. {error}\n"
                
                if len(parent.extraction_errors) > 5:
                    error_msg += f"\n...and {len(parent.extraction_errors) - 5} more errors."
                    
                QMessageBox.warning(
                    parent,
                    "Extraction Completed with Errors",
                    error_msg
                )
            return
            
        # Update progress
        parent.progress_bar.setValue(parent.current_extraction_index + 1)
        command = parent.extraction_commands[parent.current_extraction_index]
        stream_name = command[4]  # The stream name is the 5th element in the command
        
        parent.statusBar().showMessage(f"Extracting stream {parent.current_extraction_index + 1}/{len(parent.extraction_commands)}: {stream_name}")
        
        # Extract the stream using the unified method, but don't show individual messages
        file_path = extract_stream_unified(
            parent,
            stream_name, 
            parent.extraction_output_dir,
            show_messages=False
        )
        
        if file_path:
            # Stream extracted successfully
            parent.current_extraction_index += 1
            extract_next_stream(parent)
        else:
            # Record the error
            parent.extraction_errors.append(f"Failed to extract '{stream_name}'")
            
            # Continue with next stream
            parent.current_extraction_index += 1
            extract_next_stream(parent)
            
    except Exception as e:
        # Handle any unexpected errors
        parent.extraction_errors.append(f"Unexpected error: {str(e)}")
        parent.current_extraction_index += 1
        extract_next_stream(parent)  # Continue with next stream

def get_output_directory(parent):
    """Prompt for output directory if needed, using last directory as default"""
    start_dir = parent.last_output_dir if parent.last_output_dir else ""
    dir_path = QFileDialog.getExistingDirectory(
        parent, "Select Output Directory", start_dir
    )
    if dir_path:
        parent.last_output_dir = dir_path  # Remember this directory
        return dir_path
    return None 