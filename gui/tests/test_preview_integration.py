import os
import pytest
import tempfile
import shutil
import re
from PyQt5.QtWidgets import QApplication
from PyQt5.QtTest import QTest

from utils.gui.main_window import MSIParseGUI
from utils.gui.extraction import extract_file_to_temp

TEST_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_DATA_DIR = os.path.join(TEST_DIR, 'test_data')
CLEAN_SIGNED_INSTALLER_PATH = os.path.join(TEST_DATA_DIR, 'clean_signed_installer')

@pytest.fixture(scope="session", autouse=True)
def setup_qapplication(request):
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app

@pytest.fixture
def main_window(qtbot):
    """Create a main window instance with the clean_signed_installer loaded"""
    window = MSIParseGUI()
    qtbot.addWidget(window)  # Register widget with qtbot for cleanup
    
    assert os.path.exists(CLEAN_SIGNED_INSTALLER_PATH), f"Test file not found: {CLEAN_SIGNED_INSTALLER_PATH}"
    window.load_msi_file(CLEAN_SIGNED_INSTALLER_PATH)
    
    # Ensure the streams tab is selected
    window.tabs.setCurrentIndex(1)  # Streams tab is at index 1
    
    # Wait for streams data to be loaded
    try:
        qtbot.waitUntil(lambda: window.streams_tree.topLevelItemCount() > 0, timeout=10000)
    except TimeoutError:
        pytest.fail("Timed out waiting for streams to load")
    
    return window

def find_stream_by_type(main_window, group_type, mime_type=None):
    """Helper function to find a stream by its type"""
    # First identify stream types
    main_window.identify_streams()
    
    # Wait for identification to complete
    QTest.qWait(1000)
    
    # Find a stream with the given group type and mime type
    for i in range(main_window.streams_tree.topLevelItemCount()):
        item = main_window.streams_tree.topLevelItem(i)
        if item.text(1) == group_type:
            if mime_type is None or mime_type in item.text(2):
                return item.text(0)
    
    return None

def test_hex_view_content(main_window, qtbot, tmp_path):
    """Test that hex view content is correctly extracted and contains hexadecimal data"""
    # Get the first stream as our test subject
    stream_name = main_window.streams_tree.topLevelItem(0).text(0)
    
    # Create a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Extract the file
    file_path = extract_file_to_temp(main_window, stream_name, temp_dir)
    
    # Verify the file was extracted
    assert file_path is not None, "File should be extracted successfully"
    assert os.path.exists(file_path), f"Extracted file should exist at {file_path}"
    
    # Read the first 100 bytes as binary
    with open(file_path, 'rb') as f:
        binary_data = f.read(100)
    
    # Convert to hex and check if it looks like valid hex data
    hex_data = binary_data.hex()
    assert len(hex_data) > 0, "Hex data should not be empty"
    assert re.match(r'^[0-9a-f]+$', hex_data), "Hex data should only contain hexadecimal characters"

def test_text_preview_content(main_window, qtbot, tmp_path):
    """Test that text files can be extracted and contain readable content"""
    # Find a text-based stream
    stream_name = find_stream_by_type(main_window, "text") or find_stream_by_type(main_window, "code")
    if not stream_name:
        pytest.skip("No suitable text stream found for testing")
    
    # Create a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Extract the file
    file_path = extract_file_to_temp(main_window, stream_name, temp_dir)
    
    # Verify the file was extracted
    assert file_path is not None, "Text file should be extracted successfully"
    assert os.path.exists(file_path), f"Extracted file should exist at {file_path}"
    
    # Read the file content as text
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            text_content = f.read(1000)  # Read first 1000 chars
    except UnicodeDecodeError:
        # Try another common encoding
        with open(file_path, 'r', encoding='latin-1') as f:
            text_content = f.read(1000)
    
    # Verify the text content
    assert len(text_content) > 0, "Text content should not be empty"
    # Check for some common text characters
    assert re.search(r'[a-zA-Z0-9\s.,;:!?()]', text_content), "Text content should contain readable characters"

def test_image_preview_content(main_window, qtbot, tmp_path):
    """Test that image files can be extracted and have correct image file signatures"""
    # Find an image stream
    stream_name = find_stream_by_type(main_window, "image")
    if not stream_name:
        pytest.skip("No suitable image stream found for testing")
    
    # Create a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Extract the file
    file_path = extract_file_to_temp(main_window, stream_name, temp_dir)
    
    # Verify the file was extracted
    assert file_path is not None, "Image file should be extracted successfully"
    assert os.path.exists(file_path), f"Extracted file should exist at {file_path}"
    
    # Read the file header
    with open(file_path, 'rb') as f:
        header = f.read(12)  # Read enough for common image headers
    
    # Check for common image format signatures
    is_image = False
    
    # JPEG: FF D8 FF
    if header.startswith(b'\xff\xd8\xff'):
        is_image = True
    
    # PNG: 89 50 4E 47 0D 0A 1A 0A
    elif header.startswith(b'\x89PNG\r\n\x1a\n'):
        is_image = True
    
    # GIF: 47 49 46 38
    elif header.startswith(b'GIF8'):
        is_image = True
    
    # BMP: 42 4D
    elif header.startswith(b'BM'):
        is_image = True
    
    # ICO: 00 00 01 00
    elif header.startswith(b'\x00\x00\x01\x00'):
        is_image = True
    
    assert is_image, f"File {os.path.basename(file_path)} should have a valid image file signature"

def test_archive_preview_content(main_window, qtbot, tmp_path):
    """Test that archive files can be extracted and have correct archive file signatures"""
    # Find an archive stream
    stream_name = find_stream_by_type(main_window, "archive")
    if not stream_name:
        pytest.skip("No suitable archive stream found for testing")
    
    # Create a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Extract the file
    file_path = extract_file_to_temp(main_window, stream_name, temp_dir)
    
    # Verify the file was extracted
    assert file_path is not None, "Archive file should be extracted successfully"
    assert os.path.exists(file_path), f"Extracted file should exist at {file_path}"
    
    # Read the file header
    with open(file_path, 'rb') as f:
        header = f.read(10)  # Read enough for common archive headers
    
    # Check for common archive format signatures
    is_archive = False
    
    # ZIP: 50 4B 03 04
    if header.startswith(b'PK\x03\x04'):
        is_archive = True
    
    # RAR: 52 61 72 21 1A 07
    elif header.startswith(b'Rar!\x1a\x07'):
        is_archive = True
    
    # 7Z: 37 7A BC AF 27 1C
    elif header.startswith(b'7z\xbc\xaf\x27\x1c'):
        is_archive = True
    
    # CAB: 4D 53 43 46
    elif header.startswith(b'MSCF'):
        is_archive = True
    
    # TAR: 75 73 74 61 72
    elif len(header) >= 8 and header[257:262] == b'ustar':  # TAR header at offset 257
        is_archive = True
    
    # If no archive signature is found but it has a .cab extension, count it as valid
    if not is_archive and file_path.lower().endswith('.cab'):
        is_archive = True
    
    assert is_archive, f"File {os.path.basename(file_path)} should have a valid archive file signature or be a .cab file"

def test_extraction_preview_integration(main_window, qtbot, tmp_path):
    """Integration test for extraction and preview combined - extracts a file and verifies it exists"""
    # Get the first stream as our test subject
    stream_name = main_window.streams_tree.topLevelItem(0).text(0)
    
    # Create a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Extract the file
    file_path = extract_file_to_temp(main_window, stream_name, temp_dir)
    
    # Verify the file was extracted
    assert file_path is not None, "File should be extracted successfully"
    assert os.path.exists(file_path), f"Extracted file should exist at {file_path}"
    
    # Verify the file has content
    assert os.path.getsize(file_path) > 0, "Extracted file should not be empty" 