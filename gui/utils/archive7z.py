import os
import subprocess
import shutil
import tempfile
import re
from pathlib import Path

class Archive7z:
    """A simple 7z-based archive handler using the 7z command-line tool."""
    
    def __init__(self):
        """Initialize the Archive7z handler."""
        self.temp_dirs = set()  # Track temp directories for cleanup
        self._find_7z_binary()
    
    def _find_7z_binary(self):
        """Find the 7z binary on the system."""
        # Check common paths for 7z
        possible_paths = [
            "7z",              # If in PATH
            "/usr/bin/7z",     # Linux
            "/usr/local/bin/7z", # macOS with Homebrew
            "C:\\Program Files\\7-Zip\\7z.exe",  # Windows
        ]
        
        for path in possible_paths:
            try:
                # Test if 7z is callable
                subprocess.run([path, "--help"], stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)
                self.binary = path
                return
            except (FileNotFoundError, PermissionError):
                continue
        
        # If we get here, 7z wasn't found
        raise FileNotFoundError("7z command-line tool not found. Please install 7-Zip.")
    
    def list_contents(self, archive_path):
        """List the contents of an archive."""
        try:
            # Run 7z list command
            process = subprocess.run(
                [self.binary, "l", "-slt", archive_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True
            )
            
            # Debug: print the raw output to help diagnose issues
            # print(f"Raw 7z output: {process.stdout}")
            
            # Parse the output to get file information
            entries = []
            current_entry = {}
            in_entry = False  # Track if we're parsing an entry
            
            for line in process.stdout.splitlines():
                line = line.strip()
                
                # Skip empty lines outside entry data
                if not line and not in_entry:
                    continue
                
                # Look for the start of entry data (after the header)
                if line.startswith('----------'):
                    in_entry = True
                    continue
                
                # Handle entry separation
                if not line and in_entry and current_entry:
                    # Process complete entry
                    if 'Path' in current_entry:
                        # Ensure all required fields have defaults
                        if 'Size' not in current_entry:
                            current_entry['Size'] = 0
                        if 'IsDir' not in current_entry:
                            # Try to infer if it's a directory from the path
                            if current_entry['Path'].endswith('/') or current_entry['Path'].endswith('\\'):
                                current_entry['IsDir'] = True
                            else:
                                current_entry['IsDir'] = False
                        
                        # Only add non-directory entries to the result
                        if not current_entry.get('IsDir', False):
                            entries.append(current_entry)
                    current_entry = {}
                    continue
                
                # Parse key-value pairs
                if '=' in line and in_entry:
                    key, value = line.split('=', 1)
                    key = key.strip()
                    value = value.strip()
                    
                    # Convert some values
                    if key == 'Size':
                        try:
                            current_entry['Size'] = int(value)
                        except ValueError:
                            current_entry['Size'] = 0
                    elif key == 'IsDir' or key == 'Folder':  # Handle both possible keys
                        current_entry['IsDir'] = value.lower() in ('1', '+', 'yes', 'true')
                    elif key == 'Path' or key == 'Name':  # Handle both possible keys
                        current_entry['Path'] = value
            
            # Add the last entry if it exists
            if current_entry and 'Path' in current_entry:
                # Ensure all required fields have defaults
                if 'Size' not in current_entry:
                    current_entry['Size'] = 0
                if 'IsDir' not in current_entry:
                    # Try to infer if it's a directory from the path
                    if current_entry['Path'].endswith('/') or current_entry['Path'].endswith('\\'):
                        current_entry['IsDir'] = True
                    else:
                        current_entry['IsDir'] = False
                
                # Only add non-directory entries to the result  
                if not current_entry.get('IsDir', False):
                    entries.append(current_entry)
            
            # If no entries were found, try alternative parsing approach
            if not entries:
                # Try parsing with a simpler approach (for older 7z versions)
                entries = self._parse_simple_list_output(process.stdout)
            
            return entries
            
        except subprocess.CalledProcessError as e:
            raise Exception(f"Failed to list archive contents: {e.stderr}")
    
    def _parse_simple_list_output(self, output_text):
        """Parse 7z output with a simpler approach for older 7z versions."""
        entries = []
        lines = output_text.splitlines()
        
        # Find the start of the file list (after the line of dashes)
        start_idx = -1
        for i, line in enumerate(lines):
            if '-------------------' in line:
                start_idx = i + 1
                break
        
        if start_idx < 0:
            return []
        
        # Parse the file list
        for i in range(start_idx, len(lines)):
            line = lines[i].strip()
            
            # Skip empty lines
            if not line:
                continue
                
            # Skip the summary line at the end
            if line.startswith('---') or line.lower().startswith('total'):
                break
            
            # Try to extract information from the line
            # Format is typically: Date Time Attr Size Compressed Name
            parts = re.split(r'\s+', line, 5)
            if len(parts) >= 6:
                path = parts[5].strip()
                
                # Skip directory entries
                if path.endswith('/') or path.endswith('\\'):
                    continue
                
                try:
                    size = int(parts[3])
                except (ValueError, IndexError):
                    size = 0
                
                entries.append({
                    'Path': path,
                    'Size': size,
                    'IsDir': False
                })
        
        return entries
    
    def extract_file(self, archive_path, file_path, output_path=None):
        """Extract a specific file from the archive."""
        if not output_path:
            # Create a temporary directory if no output path is specified
            temp_dir = tempfile.mkdtemp()
            self.temp_dirs.add(temp_dir)
            output_path = os.path.join(temp_dir, os.path.basename(file_path))
        
        output_dir = os.path.dirname(output_path)
        if output_dir:
            os.makedirs(output_dir, exist_ok=True)
        
        # Get the basename of the file to be extracted
        base_name = os.path.basename(file_path)
        
        # Create a list of paths to try in order
        paths_to_try = [
            file_path,  # Original path
            base_name,  # Just the filename
            file_path.replace('\\', '/'),  # Forward slashes
            file_path.replace('\\', '/').lstrip('/'),  # Remove leading slash
            f"*{base_name}",  # Wildcard + basename
            f"*{base_name}*",  # Wildcard + basename + wildcard
        ]
        
        extraction_errors = []
        
        for path in paths_to_try:
            try:
                # Run 7z extract command for a specific file
                process = subprocess.run(
                    [self.binary, "e", archive_path, path, f"-o{output_dir}", "-y"],
                    stdout=subprocess.PIPE,
                    stderr=subprocess.PIPE,
                    text=True,
                    check=False  # Don't raise exception, handle errors below
                )
                
                # Check if extraction was successful
                if process.returncode == 0:
                    # Try to find the extracted file
                    all_files = [f for f in os.listdir(output_dir) 
                              if os.path.isfile(os.path.join(output_dir, f))]
                    
                    # If we have multiple files, try to find the right one
                    extracted_file = None
                    
                    # First try exact match
                    if base_name in all_files:
                        extracted_file = base_name
                    # Then try case-insensitive match
                    else:
                        for f in all_files:
                            if f.lower() == base_name.lower():
                                extracted_file = f
                                break
                    # Lastly, use the first file if there's only one
                    if not extracted_file and len(all_files) == 1:
                        extracted_file = all_files[0]
                    
                    if extracted_file:
                        extracted_path = os.path.join(output_dir, extracted_file)
                        
                        # If extracted file is not at the desired output path, move it
                        if extracted_path != output_path:
                            shutil.move(extracted_path, output_path)
                        
                        return output_path
                    
                    extraction_errors.append(f"Extraction successful but file not found: {path}")
                else:
                    extraction_errors.append(f"Failed to extract path '{path}': {process.stderr}")
            
            except Exception as e:
                extraction_errors.append(f"Error extracting path '{path}': {str(e)}")
        
        # If all methods failed, try a more aggressive approach
        try:
            # Extract all files to a temporary directory
            temp_extract_dir = tempfile.mkdtemp()
            self.temp_dirs.add(temp_extract_dir)
            
            extract_all_process = subprocess.run(
                [self.binary, "x", archive_path, f"-o{temp_extract_dir}", "-y"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=False
            )
            
            if extract_all_process.returncode == 0:
                # Look for the file we want in the extracted files
                for root, dirs, files in os.walk(temp_extract_dir):
                    for filename in files:
                        if filename.lower() == base_name.lower():
                            found_path = os.path.join(root, filename)
                            shutil.copy2(found_path, output_path)
                            return output_path
                
                # If we can't find an exact match, try a fuzzy match
                for root, dirs, files in os.walk(temp_extract_dir):
                    for filename in files:
                        if base_name.lower() in filename.lower():
                            found_path = os.path.join(root, filename)
                            shutil.copy2(found_path, output_path)
                            return output_path
            
                extraction_errors.append("Full extract succeeded but couldn't find the file")
            else:
                extraction_errors.append(f"Full extract failed: {extract_all_process.stderr}")
                
        except Exception as e:
            extraction_errors.append(f"Error during full extract: {str(e)}")
        
        # If we get here, all attempts failed
        raise Exception(f"Failed to extract file: {extraction_errors[-1] if extraction_errors else 'Unknown error'}")
    
    def extract_all(self, archive_path, output_dir=None):
        """Extract all files from the archive."""
        if not output_dir:
            # Create a temporary directory if no output directory is specified
            output_dir = tempfile.mkdtemp()
            self.temp_dirs.add(output_dir)
        
        os.makedirs(output_dir, exist_ok=True)
        
        try:
            # Run 7z extract command for all files
            process = subprocess.run(
                [self.binary, "x", archive_path, f"-o{output_dir}", "-y"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
                check=True
            )
            
            return output_dir
            
        except subprocess.CalledProcessError as e:
            raise Exception(f"Failed to extract archive: {e.stderr}")
    
    def cleanup(self):
        """Clean up temporary directories."""
        for temp_dir in self.temp_dirs:
            if os.path.exists(temp_dir):
                try:
                    shutil.rmtree(temp_dir)
                except Exception:
                    pass
        
        self.temp_dirs.clear()
    
    def __del__(self):
        """Ensure cleanup when the object is garbage collected."""
        self.cleanup()

# Compatibility functions to work as drop-in replacement for libarchive

def is_available():
    """Check if 7z is available on the system."""
    try:
        Archive7z()
        return True
    except FileNotFoundError:
        return False

class ArchiveEntry:
    """A simple class to mimic libarchive's entry object."""
    
    def __init__(self, entry_dict):
        self.pathname = entry_dict['Path']
        self.size = entry_dict['Size']
        self.isdir = entry_dict.get('IsDir', False)
        self._data = None
    
    def get_blocks(self):
        """Mimic libarchive's get_blocks method."""
        if self._data is None:
            return []
        return [self._data]

class ArchiveReader:
    """A simple class to mimic libarchive's file_reader."""
    
    def __init__(self, archive_path):
        self.archive_path = archive_path
        self.archive = Archive7z()
        self.entries = []
        self.temp_dir = tempfile.mkdtemp()
        self.archive.temp_dirs.add(self.temp_dir)
    
    def __enter__(self):
        # List the archive contents and create entry objects
        entry_dicts = self.archive.list_contents(self.archive_path)
        self.entries = [ArchiveEntry(entry_dict) for entry_dict in entry_dicts]
        return self.entries
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        # Clean up the temporary directory
        self.archive.cleanup()

def file_reader(archive_path):
    """Create an ArchiveReader for the specified archive."""
    return ArchiveReader(archive_path)

class Archive:
    """A simple class to mimic libarchive's Archive class."""
    
    def __init__(self, file_obj):
        # Save the file to a temporary location if it's a file object
        self.temp_file = None
        if hasattr(file_obj, 'read'):
            self.temp_file = tempfile.NamedTemporaryFile(delete=False)
            self.temp_file.write(file_obj.read())
            self.temp_file.close()
            self.archive_path = self.temp_file.name
        else:
            # Assume it's a path
            self.archive_path = str(file_obj)
        
        self.archive = Archive7z()
        self.entries = []
    
    def __enter__(self):
        # List the archive contents and create entry objects
        entry_dicts = self.archive.list_contents(self.archive_path)
        self.entries = [ArchiveEntry(entry_dict) for entry_dict in entry_dicts]
        return self.entries
    
    def __exit__(self, exc_type, exc_val, exc_tb):
        # Clean up the temporary file if created
        if self.temp_file:
            try:
                os.unlink(self.temp_file.name)
            except Exception:
                pass
        
        # Clean up the archive
        self.archive.cleanup()

# Define available functions for compatibility with the original code
__all__ = ['is_available', 'file_reader', 'Archive', 'ArchiveEntry'] 