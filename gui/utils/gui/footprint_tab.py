"""
Installation Impact tab functionality for the MSI Parser GUI
"""

import json
import re
from PyQt5.QtWidgets import (QTreeWidgetItem, QTableWidgetItem, QHeaderView,
                            QMessageBox, QTreeWidget, QMenu, QAction, QApplication)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QFont, QColor

# Constants for better readability and maintenance
REGISTRY_ROOTS = {
    "-1": "UNKNOWN",
    "0": "HKCR",
    "1": "HKCU",
    "2": "HKLM",
    "3": "HKU",
    "4": "PERF",
    "5": "HKCC"
}

SERVICE_TYPES = {
    "0x00000001": "Kernel Driver",
    "0x00000002": "File System Driver",
    "0x00000010": "Win32 Own Process",
    "0x00000020": "Win32 Share Process",
    "0x00000100": "Interactive",
    "0x00000110": "Interactive, Win32 Own Process",
    "0x00000120": "Interactive, Win32 Share Process"
}

SERVICE_START_TYPES = {
    "0x00000000": "Boot Start",
    "0x00000001": "System Start",
    "0x00000002": "Auto Start",
    "0x00000003": "Demand Start",
    "0x00000004": "Disabled"
}

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

AUTORUN_PATTERNS = [
    r".*\\Run\\.*",
    r".*\\RunOnce\\.*",
    r".*\\Windows\\CurrentVersion\\Run\\.*",
    r".*\\Windows\\CurrentVersion\\RunOnce\\.*",
    r".*\\Explorer\\ShellExecuteHooks\\.*",
    r".*\\Shell\\Open\\Command\\.*",
    r".*\\ShellIconOverlayIdentifiers\\.*",
    r".*\\ShellServiceObjects\\.*"
]

def create_section_header(text, count, color):
    """Create a section header with consistent styling"""
    header = QTreeWidgetItem([text, f"{count} items", "", ""])
    for i in range(4):
        header.setBackground(i, color)
    font = header.font(0)
    font.setBold(True)
    header.setFont(0, font)
    return header

def analyze_installation_impact(parent):
    """Analyze the MSI package for installation impact"""
    if not parent.tables_data:
        parent.show_warning("Warning", "No MSI file loaded or tables not parsed.")
        return
    
    parent.impact_tree.clear()
    parent.impact_tree.setHeaderLabels(["Type", "Entry", "Concern", "Details"])
    
    # Initialize collections
    registry_entries = []
    file_entries = []
    service_entries = []
    shortcut_entries = []
    extension_entries = []
    env_var_entries = []
    
    # Get component table for directory mapping
    component_table = next((table for table in parent.tables_data if table["name"] == "Component"), None)
    component_directories = {}
    
    if component_table:
        for row in component_table["rows"]:
            if len(row) >= 3:
                component_directories[row[0]] = row[2]
    
    # Process tables
    for table in parent.tables_data:
        table_name = table["name"]
        
        if table_name == "Registry":
            for row in table["rows"]:
                if len(row) >= 5:
                    registry, root, key, name, value = [x.strip() for x in row[:5]]
                    root_name = REGISTRY_ROOTS.get(root, f"Unknown ({root})")
                    
                    if root and key:
                        reg_path = f"{root_name}\\{key}"
                        if name:
                            reg_path = f"{reg_path}\\{name}"
                        
                        is_persistence = any(re.match(pattern, reg_path, re.IGNORECASE) for pattern in AUTORUN_PATTERNS)
                        
                        registry_entries.append({
                            "type": "Persistence Mechanism" if is_persistence else "",
                            "entry": normalize_registry_path(reg_path),
                            "concern": "High - Persistence Mechanism" if is_persistence else "",
                            "details": f"Value: {value}"
                        })
        
        elif table_name == "File":
            for row in table["rows"]:
                if len(row) >= 3:
                    file_entries.append({
                        "File": row[0],
                        "Component": row[1],
                        "FileName": row[2]
                    })
        
        elif table_name == "ServiceInstall":
            for row in table["rows"]:
                if len(row) >= 4:
                    is_critical = row[4] == "2" or row[3] == "16" if len(row) > 4 else False
                    service_entries.append({
                        "ServiceInstall": row[0],
                        "Name": row[1],
                        "DisplayName": row[2],
                        "Type": row[3],
                        "StartType": row[4] if len(row) > 4 else "",
                        "IsCritical": is_critical
                    })
        
        elif table_name == "Shortcut":
            for row in table["rows"]:
                if len(row) >= 4:
                    shortcut_entries.append({
                        "Shortcut": row[0],
                        "Directory": row[1],
                        "Name": row[2],
                        "Component": row[3]
                    })
        
        elif table_name == "Extension":
            for row in table["rows"]:
                if len(row) >= 2:
                    extension_entries.append({
                        "Extension": row[0],
                        "Component": row[1]
                    })
        
        elif table_name == "Environment":
            for row in table["rows"]:
                if len(row) >= 3:
                    env_var_entries.append({
                        "Environment": row[0],
                        "Name": row[1],
                        "Value": row[2]
                    })
    
    # Add sections to tree
    if file_entries:
        files_header = create_section_header("Files", len(file_entries), QColor(240, 248, 255))
        parent.impact_tree.addTopLevelItem(files_header)
        
        for file_entry in file_entries:
            install_dir = component_directories.get(file_entry["Component"], "UNKNOWN")
            filename = file_entry['FileName']
            display_name = filename.split("|")[0] if "|" in filename else filename
            install_path = filename.split("|")[1] if "|" in filename and filename.split("|")[1] else filename.split("|")[0]
            
            file_item = QTreeWidgetItem([
                "", 
                f"[{install_dir}]\\{install_path}",
                "",
                f"Archived name: {display_name}"
            ])
            files_header.addChild(file_item)
    
    if registry_entries:
        registry_header = create_section_header("Registry", len(registry_entries), QColor(255, 240, 245))
        parent.impact_tree.addTopLevelItem(registry_header)
        
        for reg_entry in registry_entries:
            reg_item = QTreeWidgetItem([
                reg_entry["type"],
                reg_entry["entry"],
                reg_entry["concern"],
                reg_entry["details"]
            ])
            
            if reg_entry["type"] == "Persistence Mechanism":
                reg_item.setForeground(2, QColor("red"))
                for i in range(4):
                    reg_item.setBackground(i, QColor(255, 235, 235))
            
            registry_header.addChild(reg_item)
    
    if service_entries:
        services_header = create_section_header("Services", len(service_entries), QColor(230, 230, 250))
        parent.impact_tree.addTopLevelItem(services_header)
        
        for service in service_entries:
            service_item = QTreeWidgetItem([
                "",
                f"{service['Name']} ({service['DisplayName']})",
                "Auto-start system service" if service["IsCritical"] else "",
                f"Type: {service['Type']}, Start: {service['StartType']}"
            ])
            
            if service["IsCritical"]:
                service_item.setForeground(2, QColor("blue"))
                for i in range(4):
                    service_item.setBackground(i, QColor(240, 245, 255))
            
            services_header.addChild(service_item)
    
    if shortcut_entries:
        shortcuts_header = create_section_header("Shortcuts", len(shortcut_entries), QColor(240, 255, 240))
        parent.impact_tree.addTopLevelItem(shortcuts_header)
        
        for shortcut in shortcut_entries:
            shortcut_name = shortcut['Name']
            display_name = shortcut_name.split("|")[0] if "|" in shortcut_name else shortcut_name
            install_path = shortcut_name.split("|")[0] + ("\\" + shortcut_name.split("|")[1] if "|" in shortcut_name and shortcut_name.split("|")[1] else "")
            
            shortcut_item = QTreeWidgetItem([
                "",
                f"[{shortcut['Directory']}]\\{install_path}",
                "",
                f"Archived name: {display_name}"
            ])
            shortcuts_header.addChild(shortcut_item)
    
    if env_var_entries:
        env_header = create_section_header("Environment", len(env_var_entries), QColor(255, 250, 205))
        parent.impact_tree.addTopLevelItem(env_header)
        
        for env in env_var_entries:
            env_item = QTreeWidgetItem([
                "",
                env["Name"],
                "",
                f"Value: {env['Value']}"
            ])
            env_header.addChild(env_item)
    
    if extension_entries:
        ext_header = create_section_header("File Associations", len(extension_entries), QColor(255, 240, 245))
        parent.impact_tree.addTopLevelItem(ext_header)
        
        for ext in extension_entries:
            install_dir = component_directories.get(ext["Component"], "UNKNOWN")
            ext_item = QTreeWidgetItem([
                "",
                f"[{install_dir}]",
                "",
                f"Extension: .{ext['Extension']}"
            ])
            ext_header.addChild(ext_item)
    
    # Keep items collapsed by default
    for i in range(parent.impact_tree.topLevelItemCount()):
        parent.impact_tree.topLevelItem(i).setExpanded(False)
    
    # Resize columns to content
    for i in range(4):
        parent.impact_tree.resizeColumnToContents(i)
    
    parent.show_status(f"Found {len(file_entries) + len(registry_entries) + len(service_entries) + len(shortcut_entries) + len(env_var_entries) + len(extension_entries)} installation impact items")

def normalize_registry_path(path):
    """Normalize registry path to use single backslashes"""
    return path.replace('\\\\', '\\') if path else ""

def display_installation_impact(parent, output):
    """Display the results of the installation impact analysis"""
    try:
        parent.impact_tree.clear()
        current_section = None
        current_subsection = None
        
        for line in output.split('\n'):
            line = line.strip()
            if not line:
                continue
                
            if line.startswith('=== '):
                current_section = line[4:].strip()
                section_item = QTreeWidgetItem([current_section])
                section_item.setFlags(section_item.flags() & ~Qt.ItemIsSelectable)
                parent.impact_tree.addTopLevelItem(section_item)
                continue
                
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
            
            if not line or line.startswith('===') or line.startswith('---'):
                continue
                
            parts = line.split('|')
            if len(parts) >= 3:
                item_type = parts[0].strip()
                entry = normalize_registry_path(parts[1].strip())
                concern = parts[2].strip() if len(parts) > 2 else ""
                details = normalize_registry_path(parts[3].strip()) if len(parts) > 3 else ""
                
                # Remove redundant "Registry Entry" type
                if item_type == "Registry Entry":
                    item_type = ""
                
                item = QTreeWidgetItem([item_type, entry, concern, details])
                
                if current_subsection:
                    subsection_item.addChild(item)
                elif current_section:
                    section_item.addChild(item)
                else:
                    parent.impact_tree.addTopLevelItem(item)
                    
                if "Persistence Mechanism" in item_type:
                    for i in range(4):
                        item.setForeground(i, Qt.red)
        
        parent.impact_tree.expandAll()
        
    except Exception as e:
        parent.handle_error("Display Error", f"Error displaying installation impact: {str(e)}") 