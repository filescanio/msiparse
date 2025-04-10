"""
MSI Parser GUI Package

This package contains the GUI implementation for the MSI Parser tool.
The implementation has been split into multiple modules for better maintainability.
"""

from utils.gui.main_window import MSIParseGUI
from utils.gui.footprint_tab import analyze_installation_impact, create_footprint_tab

__all__ = [
    'MSIParseGUI',
    'analyze_installation_impact',
    'create_footprint_tab'
] 