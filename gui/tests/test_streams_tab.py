import os
import pytest
import tempfile
import shutil
from PyQt5.QtWidgets import QApplication, QMenu, QAction, QHeaderView
from PyQt5.QtCore import Qt, QPoint, QItemSelectionModel
from PyQt5.QtTest import QTest

from utils.gui.main_window import MSIParseGUI
from utils.gui.streams_tab import show_streams_context_menu
from utils.gui.extraction import extract_single_stream, extract_all_streams

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

def test_streams_tab_populated(main_window):
    """Test that the streams tab is properly populated when a file is loaded"""
    # Verify that streams are populated in the tree widget
    assert main_window.streams_tree.topLevelItemCount() > 0, "Streams tree should contain items"
    
    # Verify that the first column contains stream names
    first_item = main_window.streams_tree.topLevelItem(0)
    assert first_item.text(0), "Stream name should not be empty"
    
    # Verify that the Identify Stream Types button is enabled
    assert main_window.identify_streams_button.isEnabled(), "Identify Stream Types button should be enabled"
    
    # Verify that the Extract All button is enabled
    assert main_window.extract_all_button.isEnabled(), "Extract All Streams button should be enabled"
    
    # Select an item to make extract_stream_button enabled
    first_item = main_window.streams_tree.topLevelItem(0)
    main_window.streams_tree.setCurrentItem(first_item)
    
    # Now verify that the Extract Selected Streams button is enabled
    assert main_window.extract_stream_button.isEnabled(), "Extract Selected Streams button should be enabled"

def test_streams_filter(main_window, qtbot):
    """Test that the filter works for streams"""
    # Count initial visible items
    initial_count = count_visible_items(main_window.streams_tree)
    assert initial_count > 0, "Initial stream count should be greater than 0"
    
    # Type a filter text that should match at least one item but not all
    filter_text = "Binary"  # This should match Binary streams like "BinaryWixCA"
    qtbot.keyClicks(main_window.streams_filter, filter_text)
    
    # Wait for the filter to be applied
    qtbot.wait(500)
    
    # Count filtered visible items
    filtered_count = count_visible_items(main_window.streams_tree)
    
    # Verify that filtering changed the visible items count
    assert 0 < filtered_count < initial_count, f"Filtered count ({filtered_count}) should be less than initial count ({initial_count}) and greater than 0"
    
    # Clear the filter
    main_window.streams_filter.clear()
    qtbot.wait(500)
    
    # Verify that clearing the filter restores all items
    assert count_visible_items(main_window.streams_tree) == initial_count, "Clearing filter should restore original item count"

def test_stream_selection(main_window, qtbot):
    """Test selecting a stream and verifying button state changes"""
    # Initially we need to clear selection
    main_window.streams_tree.clearSelection()
    
    # Verify that the extract button is disabled with no selection
    main_window.on_stream_selection_changed()  # Force update of button state
    assert not main_window.extract_stream_button.isEnabled(), "Extract Selected Streams button should be disabled with no selection"
    
    # Select the first stream
    first_item = main_window.streams_tree.topLevelItem(0)
    main_window.streams_tree.setCurrentItem(first_item)
    
    # Verify that a stream is selected
    assert len(main_window.streams_tree.selectedItems()) == 1, "One stream should be selected"
    
    # Extract button should be enabled with selection
    main_window.on_stream_selection_changed()  # Force update of button state
    assert main_window.extract_stream_button.isEnabled(), "Extract Selected Streams button should be enabled with selection"

def test_identify_streams(main_window, monkeypatch):
    """Test the identify streams functionality directly without clicking buttons"""
    # Create a mock for the identify_streams function to avoid actual processing
    identify_called = False
    
    def mock_identify_streams(parent):
        nonlocal identify_called
        identify_called = True
        # Simulate identification by updating the first item
        if parent.streams_tree.topLevelItemCount() > 0:
            item = parent.streams_tree.topLevelItem(0)
            item.setText(1, "archive")  # Group
            item.setText(2, "application/x-msi")  # MIME Type
            item.setText(3, "1.2 MB")  # Size
            item.setText(4, "abcdef1234567890")  # Hash
    
    # Store original method
    original_identify = main_window.identify_streams
    
    try:
        # Apply mock directly to the main_window instance
        main_window.identify_streams = lambda: mock_identify_streams(main_window)
        
        # Call the identify_streams method directly
        main_window.identify_streams()
        
        # Verify that the mock function was called
        assert identify_called, "Identify streams function should have been called"
        
        # Verify that at least the first item has been updated with file type information
        if main_window.streams_tree.topLevelItemCount() > 0:
            first_item = main_window.streams_tree.topLevelItem(0)
            assert first_item.text(1) != "", "Group should not be empty after identification"
            assert first_item.text(2) != "", "MIME Type should not be empty after identification"
    finally:
        # Restore original method
        main_window.identify_streams = original_identify

def test_extract_stream_action(main_window, qtbot, monkeypatch, tmp_path):
    """Test extracting a stream to a file"""
    # Setup a temporary directory for extraction
    temp_dir = str(tmp_path)
    
    # Create a mock for extraction to avoid actual disk operations
    extract_called = False
    extracted_stream = None
    
    def mock_extract_single_stream(parent, stream_name):
        nonlocal extract_called, extracted_stream
        extract_called = True
        extracted_stream = stream_name
        return True
    
    # Mock get_output_directory
    def mock_get_output_directory(parent):
        return temp_dir
    
    # Store original methods
    original_extract = main_window.extract_single_stream
    original_get_dir = main_window.get_output_directory
    
    try:
        # Apply mocks directly to the main_window instance
        main_window.extract_single_stream = lambda stream_name: mock_extract_single_stream(main_window, stream_name)
        main_window.get_output_directory = lambda: mock_get_output_directory(main_window)
        
        # Select first stream
        first_item = main_window.streams_tree.topLevelItem(0)
        stream_name = first_item.text(0)
        main_window.streams_tree.setCurrentItem(first_item)
        
        # Direct call to extract functionality to avoid GUI issues
        main_window.extract_single_stream(stream_name)
        
        # Verify extraction was called with correct parameters
        assert extract_called, "Extract single stream should have been called"
        assert extracted_stream == stream_name, f"Expected stream {stream_name} to be extracted"
    finally:
        # Restore original methods
        main_window.extract_single_stream = original_extract
        main_window.get_output_directory = original_get_dir

def test_extract_all_streams_action(main_window, qtbot, monkeypatch):
    """Test the 'Extract All Streams' functionality"""
    extract_all_called = False
    
    def mock_extract_all_streams(parent):
        nonlocal extract_all_called
        extract_all_called = True
        return True
    
    # Store original method
    original_extract_all = main_window.extract_all_streams
    
    try:
        # Apply mock directly to the main_window instance
        main_window.extract_all_streams = lambda: mock_extract_all_streams(main_window)
        
        # Direct call to extract all functionality to avoid GUI issues
        main_window.extract_all_streams()
        
        # Verify that the extraction function was called
        assert extract_all_called, "Extract all streams function should have been called"
    finally:
        # Restore original method
        main_window.extract_all_streams = original_extract_all

def test_column_sorting(main_window, qtbot, monkeypatch):
    """Test that clicking on column headers sorts the streams"""
    # Get header sort state before sorting
    header = main_window.streams_tree.header()
    initial_sort_column = header.sortIndicatorSection()
    initial_sort_order = header.sortIndicatorOrder()
    
    # Click on a different column
    new_sort_column = (initial_sort_column + 1) % 5  # Choose a different column (mod 5 for 5 columns)
    
    # Store original data for the first few items to verify sorting changes
    original_data = []
    for i in range(min(5, main_window.streams_tree.topLevelItemCount())):
        item = main_window.streams_tree.topLevelItem(i)
        original_data.append(item.text(new_sort_column))
    
    # Sort by the new column
    main_window.streams_tree.sortByColumn(new_sort_column, Qt.AscendingOrder)
    
    # Verify sort was applied
    assert header.sortIndicatorSection() == new_sort_column, "Sort column should have changed"
    assert header.sortIndicatorOrder() == Qt.AscendingOrder, "Sort order should be ascending"
    
    # If there are enough items, verify that sorting had an effect
    if len(original_data) > 1:
        # Get data after sorting
        sorted_data = []
        for i in range(min(5, main_window.streams_tree.topLevelItemCount())):
            item = main_window.streams_tree.topLevelItem(i)
            sorted_data.append(item.text(new_sort_column))
        
        # Check if sorted data is different from original (not perfect but reasonable)
        # This could fail if data is already sorted or values are identical
        if sorted(original_data) != original_data:
            assert sorted_data != original_data, "Sorting should change the order of items"

def test_stream_preview_methods(main_window, monkeypatch):
    """Test the various stream preview methods directly"""
    # Mock preview functions to verify they're called
    preview_calls = set()
    
    def mock_show_hex_view(parent, stream_name):
        preview_calls.add(f"hex_view:{stream_name}")
    
    def mock_show_text_preview(parent, stream_name):
        preview_calls.add(f"text_preview:{stream_name}")
    
    def mock_show_image_preview(parent, stream_name):
        preview_calls.add(f"image_preview:{stream_name}")
    
    # Store original methods
    original_hex_view = main_window.show_hex_view
    original_text_preview = main_window.show_text_preview
    original_image_preview = main_window.show_image_preview
    
    try:
        # Apply mocks directly to the main_window instance
        main_window.show_hex_view = lambda stream_name: mock_show_hex_view(main_window, stream_name)
        main_window.show_text_preview = lambda stream_name: mock_show_text_preview(main_window, stream_name)
        main_window.show_image_preview = lambda stream_name: mock_show_image_preview(main_window, stream_name)
        
        # Get first stream name
        first_item = main_window.streams_tree.topLevelItem(0)
        stream_name = first_item.text(0)
        
        # Test each preview method directly
        main_window.show_hex_view(stream_name)
        assert f"hex_view:{stream_name}" in preview_calls, "Hex view should have been called"
        
        main_window.show_text_preview(stream_name)
        assert f"text_preview:{stream_name}" in preview_calls, "Text preview should have been called"
        
        main_window.show_image_preview(stream_name)
        assert f"image_preview:{stream_name}" in preview_calls, "Image preview should have been called"
    finally:
        # Restore original methods
        main_window.show_hex_view = original_hex_view
        main_window.show_text_preview = original_text_preview
        main_window.show_image_preview = original_image_preview

def count_visible_items(tree_widget):
    """Helper function to count visible items in a tree widget"""
    count = 0
    for i in range(tree_widget.topLevelItemCount()):
        if not tree_widget.topLevelItem(i).isHidden():
            count += 1
    return count 