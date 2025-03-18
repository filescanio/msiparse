"""
Installation Impact tab functionality for the MSI Parser GUI
"""

import json
import re
from PyQt5.QtWidgets import (QTreeWidgetItem, QTableWidgetItem, QHeaderView,
                            QMessageBox, QTreeWidget, QMenu, QAction, QApplication)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor

# Registry root mapping for better display
REGISTRY_ROOTS = {
    "-1": "UNKNOWN",
    "0": "CLASSES_ROOT (HKCR)",
    "1": "CURRENT_USER (HKCU)",
    "2": "LOCAL_MACHINE (HKLM)",
    "3": "USERS (HKU)",
    "4": "PERFORMANCE_DATA",
    "5": "CURRENT_CONFIG (HKCC)"
}

# Define service types
SERVICE_TYPES = {
    "0x00000001": "Kernel Driver",
    "0x00000002": "File System Driver",
    "0x00000010": "Win32 Own Process",
    "0x00000020": "Win32 Share Process",
    "0x00000100": "Interactive",
    "0x00000110": "Interactive, Win32 Own Process",
    "0x00000120": "Interactive, Win32 Share Process"
}

# Define service start types
SERVICE_START_TYPES = {
    "0x00000000": "Boot Start",
    "0x00000001": "System Start",
    "0x00000002": "Auto Start",
    "0x00000003": "Demand Start",
    "0x00000004": "Disabled"
}

# Define directory mapping (commonly used MSI directory properties)
MSI_DIRECTORIES = {
    "TARGETDIR": "Root Installation Directory",
    "ProgramFilesFolder": "Program Files",
    "ProgramFiles64Folder": "Program Files (x64)",
    "ProgramMenuFolder": "Start Menu",
    "SystemFolder": "System32",
    "System64Folder": "System64",
    "AppDataFolder": "AppData",
    "CommonAppDataFolder": "Common AppData",
    "CommonFiles64Folder": "Common Files (x64)",
    "CommonFilesFolder": "Common Files",
    "DesktopFolder": "Desktop",
    "FavoritesFolder": "Favorites",
    "FontsFolder": "Fonts",
    "PersonalFolder": "Personal",
    "SendToFolder": "SendTo",
    "StartMenuFolder": "Start Menu",
    "StartupFolder": "Startup",
    "TempFolder": "Temp",
    "TemplateFolder": "Templates",
    "WindowsFolder": "Windows",
    "WindowsVolume": "Windows Volume"
}

def analyze_installation_impact(parent):
    """Analyze the MSI package for installation impact"""
    if not parent.tables_data:
        parent.show_warning("Warning", "No MSI file loaded or tables not parsed.")
        return
    
    # Clear current items
    parent.impact_tree.clear()
    
    # Setup headers - reorder to put Security Concern before Details
    parent.impact_tree.setHeaderLabels(["Type", "Entry", "Concern", "Details"])
    
    # Collect information from various tables
    registry_entries = []
    file_entries = []
    folder_entries = []
    shortcut_entries = []
    extension_entries = []
    service_entries = []
    env_var_entries = []
    
    # Look up tables we'll need for relationships
    component_table = None
    
    # Find the Component table
    for table in parent.tables_data:
        if table["name"] == "Component":
            component_table = table
    
    # Create component to directory mapping
    component_directories = {}
    
    # If we found the component table, populate the component -> directory mapping
    if component_table:
        for row in component_table["rows"]:
            if len(row) >= 3:  # Component, ComponentId, Directory_
                component_id = row[0]
                directory_id = row[2]
                component_directories[component_id] = directory_id
    
    # Process tables
    for table in parent.tables_data:
        table_name = table["name"]
        
        # Registry entries - using the improved handling
        if table_name == "Registry":
            for row in table["rows"]:
                if len(row) >= 5:  # Registry, Root, Key, Name, Value
                    registry = row[0].strip()
                    root = row[1].strip()
                    key = row[2].strip()
                    name = row[3].strip()
                    value = row[4].strip()
                    
                    # Map registry root numbers to names
                    root_name = REGISTRY_ROOTS.get(root, f"Unknown ({root})")
                    
                    # Build registry path
                    if root and key:
                        reg_path = f"{root_name}\\{key}"
                        if name:
                            reg_path = f"{reg_path}\\{name}"
                        
                        # Check for persistence mechanisms
                        is_persistence = False
                        persistence_type = ""
                        
                        # Common autorun locations
                        autorun_patterns = [
                            r".*\\Run\\.*",
                            r".*\\RunOnce\\.*",
                            r".*\\Windows\\CurrentVersion\\Run\\.*",
                            r".*\\Windows\\CurrentVersion\\RunOnce\\.*",
                            r".*\\Explorer\\ShellExecuteHooks\\.*",
                            r".*\\Shell\\Open\\Command\\.*",
                            r".*\\ShellIconOverlayIdentifiers\\.*",
                            r".*\\ShellServiceObjects\\.*"
                        ]
                        
                        for pattern in autorun_patterns:
                            if re.match(pattern, reg_path, re.IGNORECASE):
                                is_persistence = True
                                persistence_type = "Autorun Registry Key"
                                break
                        
                        # Add to registry entries
                        entry_type = "Registry Entry"
                        if is_persistence:
                            entry_type = f"Registry Entry (Persistence Mechanism)"
                        
                        registry_entries.append({
                            "type": entry_type,
                            "entry": normalize_registry_path(reg_path),
                            "concern": "High - Persistence Mechanism" if is_persistence else "",
                            "details": f"Value: {value}"
                        })
        
        # File entries
        elif table_name == "File":
            for row in table["rows"]:
                if len(row) >= 3:  # File, Component_, FileName
                    file_key = row[0]
                    component = row[1]
                    filename = row[2]
                    file_entries.append({
                        "File": file_key,
                        "Component": component,
                        "FileName": filename
                    })
        
        # Service entries
        elif table_name == "ServiceInstall":
            for row in table["rows"]:
                if len(row) >= 4:  # ServiceInstall, Name, DisplayName, ServiceType
                    service_key = row[0]
                    service_name = row[1]
                    display_name = row[2]
                    service_type = row[3]
                    start_type = row[4] if len(row) > 4 else ""
                    
                    # Check if this is a critical service (autostart or system)
                    is_critical = False
                    if start_type == "2" or service_type == "16":  # Auto start or system service
                        is_critical = True
                    
                    service_entries.append({
                        "ServiceInstall": service_key,
                        "Name": service_name,
                        "DisplayName": display_name,
                        "Type": service_type,
                        "StartType": start_type,
                        "IsCritical": is_critical
                    })
        
        # Shortcut entries
        elif table_name == "Shortcut":
            for row in table["rows"]:
                if len(row) >= 4:  # Shortcut, Directory_, Name, Component_
                    shortcut_key = row[0]
                    dir_key = row[1]
                    shortcut_name = row[2]
                    component = row[3]
                    shortcut_entries.append({
                        "Shortcut": shortcut_key,
                        "Directory": dir_key,
                        "Name": shortcut_name,
                        "Component": component
                    })
        
        # Extension associations
        elif table_name == "Extension":
            for row in table["rows"]:
                if len(row) >= 2:  # Extension, Component_
                    extension = row[0]
                    component = row[1]
                    extension_entries.append({
                        "Extension": extension,
                        "Component": component
                    })
        
        # Environment variables
        elif table_name == "Environment":
            for row in table["rows"]:
                if len(row) >= 3:  # Environment, Name, Value
                    env_key = row[0]
                    env_name = row[1]
                    env_value = row[2]
                    env_var_entries.append({
                        "Environment": env_key,
                        "Name": env_name,
                        "Value": env_value
                    })
    
    # Track total changes for status message
    total_changes = 0
    
    # Files section
    if file_entries:
        files_count = len(file_entries)
        files_header = QTreeWidgetItem(["Files", f"{files_count} files will be installed", "", ""])
        files_header.setBackground(0, QColor(240, 248, 255))  # Light blue
        files_header.setBackground(1, QColor(240, 248, 255))
        files_header.setBackground(2, QColor(240, 248, 255))
        files_header.setBackground(3, QColor(240, 248, 255))
        font = files_header.font(0)
        font.setBold(True)
        files_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(files_header)
        
        # Add all files
        for file_entry in file_entries:
            # Get installation directory from component
            install_dir = "UNKNOWN"
            if file_entry["Component"] in component_directories:
                install_dir = component_directories[file_entry["Component"]]
            
            # Handle filenames with pipe character (|)
            filename = file_entry['FileName']
            display_name = filename
            install_path = filename
            
            if "|" in filename:
                parts = filename.split("|", 1)
                display_name = parts[0]  # Part before the | for display
                install_path = parts[0]
                if len(parts) > 1 and parts[1]:
                    install_path = parts[0] + "\\" + parts[1]  # Append part after | to path
            
            # Now put the full installation path in the Item column without descriptive text
            install_path_display = f"[{install_dir}]\\{install_path}"
            
            # And put the original filename in the Details column with note about archived name
            file_details = f"Archived name: {display_name}"
            
            file_item = QTreeWidgetItem([
                "", 
                install_path_display,  # Path now in Item column
                "",  # No security concern for most files
                file_details  # Original filename now in Details column
            ])
            files_header.addChild(file_item)
    
    # Registry section
    if registry_entries:
        registry_count = len(registry_entries)
        registry_header = QTreeWidgetItem(["Registry", f"{registry_count} registry entries will be created/modified", "", ""])
        registry_header.setBackground(0, QColor(255, 240, 245))  # Light pink
        registry_header.setBackground(1, QColor(255, 240, 245))
        registry_header.setBackground(2, QColor(255, 240, 245))
        registry_header.setBackground(3, QColor(255, 240, 245))
        font = registry_header.font(0)
        font.setBold(True)
        registry_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(registry_header)
        
        # Add all registry entries
        for reg_entry in registry_entries:
            reg_item = QTreeWidgetItem([
                reg_entry["type"],
                reg_entry["entry"],
                reg_entry["concern"],
                reg_entry["details"]
            ])
            
            # Highlight persistence items
            if "Persistence Mechanism" in reg_entry["type"]:
                reg_item.setForeground(2, QColor("red"))
                
                # Add light red background for persistence items
                light_red = QColor(255, 235, 235)
                for i in range(4):
                    reg_item.setBackground(i, light_red)
            
            registry_header.addChild(reg_item)
    
    # Services section
    if service_entries:
        services_header = QTreeWidgetItem(["Services", f"{len(service_entries)} services will be installed", "", ""])
        services_header.setBackground(0, QColor(230, 230, 250))  # Lavender
        services_header.setBackground(1, QColor(230, 230, 250))
        services_header.setBackground(2, QColor(230, 230, 250))
        services_header.setBackground(3, QColor(230, 230, 250))
        font = services_header.font(0)
        font.setBold(True)
        services_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(services_header)
        
        for service in service_entries:
            security_concern = ""
            if service["IsCritical"]:
                security_concern = "Auto-start system service"
            
            service_item = QTreeWidgetItem([
                "",
                f"{service['Name']} ({service['DisplayName']})",
                security_concern,
                f"Type: {service['Type']}, Start: {service['StartType']}"  # Keep service info
            ])
            
            # Highlight critical services
            if service["IsCritical"]:
                service_item.setForeground(2, QColor("blue"))
                
                # Light blue background for critical services
                light_blue = QColor(240, 245, 255)
                for i in range(4):
                    service_item.setBackground(i, light_blue)
            
            services_header.addChild(service_item)
    
    # Shortcuts section
    if shortcut_entries:
        shortcuts_header = QTreeWidgetItem(["Shortcuts", f"{len(shortcut_entries)} shortcuts will be created", "", ""])
        shortcuts_header.setBackground(0, QColor(240, 255, 240))  # Light green
        shortcuts_header.setBackground(1, QColor(240, 255, 240))
        shortcuts_header.setBackground(2, QColor(240, 255, 240))
        shortcuts_header.setBackground(3, QColor(240, 255, 240))
        font = shortcuts_header.font(0)
        font.setBold(True)
        shortcuts_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(shortcuts_header)
        
        for shortcut in shortcut_entries:
            # Handle shortcut names with pipe character (|)
            shortcut_name = shortcut['Name']
            display_name = shortcut_name
            install_path = shortcut_name
            
            if "|" in shortcut_name:
                parts = shortcut_name.split("|", 1)
                display_name = parts[0]  # Part before the | for display
                install_path = parts[0]
                if len(parts) > 1 and parts[1]:
                    install_path = parts[0] + "\\" + parts[1]  # Append part after | to path
            
            # Put the full path in the Item column without descriptive text
            path_display = f"[{shortcut['Directory']}]\\{install_path}"
            
            # And put the original shortcut name in the Details column
            shortcut_details = f"Archived name: {display_name}"
            
            shortcut_item = QTreeWidgetItem([
                "",
                path_display,  # Path now in Item column
                "",  # No security concern
                shortcut_details  # Original shortcut name now in Details column
            ])
            shortcuts_header.addChild(shortcut_item)
    
    # Environment variables section
    if env_var_entries:
        env_header = QTreeWidgetItem(["Environment", f"{len(env_var_entries)} environment variables will be set", "", ""])
        env_header.setBackground(0, QColor(255, 250, 205))  # Light yellow
        env_header.setBackground(1, QColor(255, 250, 205))
        env_header.setBackground(2, QColor(255, 250, 205))
        env_header.setBackground(3, QColor(255, 250, 205))
        font = env_header.font(0)
        font.setBold(True)
        env_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(env_header)
        
        for env in env_var_entries:
            env_item = QTreeWidgetItem([
                "",
                env["Name"],
                "",  # No security concern
                f"Value: {env['Value']}"  # Keep value as is, it's relevant
            ])
            env_header.addChild(env_item)
    
    # File extensions section
    if extension_entries:
        ext_header = QTreeWidgetItem(["File Associations", f"{len(extension_entries)} file extensions will be registered", "", ""])
        ext_header.setBackground(0, QColor(255, 240, 245))  # Light pink
        ext_header.setBackground(1, QColor(255, 240, 245))
        ext_header.setBackground(2, QColor(255, 240, 245))
        ext_header.setBackground(3, QColor(255, 240, 245))
        font = ext_header.font(0)
        font.setBold(True)
        ext_header.setFont(0, font)
        parent.impact_tree.addTopLevelItem(ext_header)
        
        for ext in extension_entries:
            # Get installation directory from component
            install_dir = "UNKNOWN"
            if ext["Component"] in component_directories:
                install_dir = component_directories[ext["Component"]]
                
            # Put the directory path in the Item column
            path_display = f"[{install_dir}]"
            
            # And put the extension info in the Details column
            ext_details = f"Extension: .{ext['Extension']}"
            
            ext_item = QTreeWidgetItem([
                "",
                path_display,  # Path now in Item column
                "",  # No security concern
                ext_details  # Extension info now in Details column
            ])
            ext_header.addChild(ext_item)
    
    # Keep items collapsed by default - except Execution section which is most important for security analysis
    for i in range(parent.impact_tree.topLevelItemCount()):
        item = parent.impact_tree.topLevelItem(i)
        if item.text(0) == "Execution":
            item.setExpanded(True)  # Only expand the Execution section
        else:
            item.setExpanded(False)  # Collapse all other sections
    
    # Resize columns to content
    for i in range(4):
        parent.impact_tree.resizeColumnToContents(i)
    
    # Show status
    parent.show_status(f"Found {total_changes} installation impact items")

def normalize_registry_path(path):
    """Normalize registry path to use single backslashes"""
    return path.replace('\\\\', '\\') if path else ""

def display_installation_impact(parent, output):
    """Display the results of the installation impact analysis"""
    try:
        # Clear existing items
        parent.impact_tree.clear()
        
        # Parse the output
        current_section = None
        current_subsection = None
        
        for line in output.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            # Check for section headers
            if line.startswith('=== '):
                current_section = line[4:].strip()
                section_item = QTreeWidgetItem([current_section])
                section_item.setFlags(section_item.flags() & ~Qt.ItemIsSelectable)
                parent.impact_tree.addTopLevelItem(section_item)
                continue
                
            # Check for subsection headers
            if line.startswith('--- '):
                current_subsection = line[4:].strip()
                subsection_item = QTreeWidgetItem([current_subsection])
                subsection_item.setFlags(subsection_item.flags() & ~Qt.ItemIsSelectable)
                if current_section:
                    section_item = parent.impact_tree.topLevelItem(parent.impact_tree.topLevelItemCount() - 1)
                    section_item.addChild(subsection_item)
                else:
                    parent.impact_tree.addTopLevelItem(subsection_item)
                continue
            
            # Skip empty lines and section headers
            if not line or line.startswith('===') or line.startswith('---'):
                continue
                
            # Parse the line into columns
            parts = line.split('|')
            if len(parts) >= 3:
                item_type = parts[0].strip()
                entry = normalize_registry_path(parts[1].strip())  # Normalize registry paths
                concern = parts[2].strip() if len(parts) > 2 else ""
                details = normalize_registry_path(parts[3].strip()) if len(parts) > 3 else ""  # Normalize registry paths in details too
                
                # Create item with all columns
                item = QTreeWidgetItem([item_type, entry, concern, details])
                
                # Add to appropriate parent
                if current_subsection:
                    subsection_item.addChild(item)
                elif current_section:
                    section_item.addChild(item)
                else:
                    parent.impact_tree.addTopLevelItem(item)
                    
                # Highlight persistence mechanisms
                if "Persistence Mechanism" in item_type:
                    item.setForeground(0, Qt.red)  # Type column
                    item.setForeground(1, Qt.red)  # Entry column
                    item.setForeground(2, Qt.red)  # Concern column
                    item.setForeground(3, Qt.red)  # Details column
                    
        # Expand all items by default
        parent.impact_tree.expandAll()
        
    except Exception as e:
        parent.handle_error("Display Error", f"Error displaying installation impact: {str(e)}") 