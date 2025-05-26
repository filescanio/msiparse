import os
import pytest
from PyQt5.QtWidgets import QApplication

from utils.gui.main_window import MSIParseGUI

TEST_DIR = os.path.dirname(os.path.abspath(__file__))
TEST_DATA_DIR = os.path.join(TEST_DIR, 'test_data')
DUMMY_MSI_PATH = os.path.join(TEST_DATA_DIR, 'registry_persistence')

@pytest.fixture(scope="session", autouse=True)
def setup_qapplication(request):
    app = QApplication.instance()
    if app is None:
        app = QApplication([])
    return app

def test_load_msi_populates_metadata(qtbot):
    """
    Test that loading an MSI file populates the metadata tab.
    """
    window = MSIParseGUI()
    qtbot.addWidget(window) # Register widget with qtbot for cleanup

    assert os.path.exists(DUMMY_MSI_PATH), f"Test file not found: {DUMMY_MSI_PATH}"
    window.load_msi_file(DUMMY_MSI_PATH)

    try:
        qtbot.waitUntil(lambda: window.metadata_text.toPlainText() != "", timeout=10000) # 10 seconds timeout
    except TimeoutError:
        pytest.fail("Timed out waiting for metadata to load. Check msiparse command execution and metadata_tab population.")

  
    expected_text = "MSI Metadata:\n\nTitle: MyCompany\nSubject: MyProduct\nAuthor: None\nUuid: a2e9cdc2-0714-4580-a887-2ee4d0efd60c\nArch: Intel\nLanguages: en-US\nCreated At: 2025-03-18 9:10:42.0 +00:00:00\nCreated With: Windows Installer XML Toolset (3.14.1.8722)\nIs Signed: None\nCodepage: Windows Latin 1\nCodepage Id: 1252\nWord Count: 10\nComments: This installer database contains the logic and data required to install MyProduct.\n"
    actual_text = window.metadata_text.toPlainText()

    assert actual_text == expected_text, "Metadata text area should not be empty after loading the file."
