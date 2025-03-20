"""
Installation Impact tab functionality for the MSI Parser GUI
"""

import re
import os
from PyQt5.QtWidgets import (QTreeWidgetItem, QHeaderView, QTreeWidget, QApplication,
                            QCheckBox, QHBoxLayout, QWidget, QVBoxLayout, QLabel)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# Constants for better readability and maintenance
REGISTRY_ROOTS = {
    "-1": "HKCU",
    "0": "HKCR",
    "1": "HKCU",
    "2": "HKLM",
    "3": "HKU"
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

SUSPICIOUS_FILE_PATTERNS = [
    r".*\\windows\\system32\\.*\.(exe|dll|sys|drv)$",
    r".*\\windows\\syswow64\\.*\.(exe|dll|sys|drv)$",
    r".*\\appdata\\.*\\microsoft\\windows\\start menu\\programs\\startup\\.*",
    r".*\\programdata\\.*\\microsoft\\windows\\start menu\\programs\\startup\\.*",
    r".*\\program files\\.*\\common files\\.*\.(exe|dll|sys|drv)$",
    r".*\\program files \(x86\)\\common files\\.*\.(exe|dll|sys|drv)$"
]

SUSPICIOUS_REGISTRY_PATTERNS = [
    # Expand existing AUTORUN_PATTERNS with more specific entries
    r".*\\Windows\\CurrentVersion\\Explorer\\Browser Helper Objects\\.*",
    r".*\\Browser Extensions\\.*",
    r".*\\Winlogon\\.*",
    r".*\\SecurityProviders\\.*",
    r".*\\Services\\.*",
    r".*\\NetworkProvider\\.*",
    r".*\\Schedule\\TaskCache\\.*",
    r".*\\PolicyManager\\.*",
    r".*\\Group Policy\\Scripts\\.*",
    r".*\\Windows\\CurrentVersion\\ShellServiceObjectDelayLoad\\.*",
    r".*\\Windows\\CurrentVersion\\Shell Extensions\\Approved\\.*",
    r".*\\Windows\\CurrentVersion\\Shell Extensions\\Blocked\\.*",
    r".*\\Windows\\CurrentVersion\\ShellServiceObjectDelayLoad\\.*",
    r".*\\Windows\\CurrentVersion\\Shell Extensions\\Approved\\.*",
    r".*\\Windows\\CurrentVersion\\Shell Extensions\\Blocked\\.*"
]

CRITICAL_DIRECTORIES = {
    "SystemFolder": "System directory modification",
    "System64Folder": "System directory modification",
    "WindowsFolder": "Windows directory modification",
    "StartupFolder": "Startup folder modification",
    "CommonFilesFolder": "Common Files directory modification",
    "CommonFiles64Folder": "Common Files (x64) directory modification"
}

# Add new constants for file extensions
HIGH_RISK_EXTENSIONS = {
    '.vbs': 'VBScript file',
    '.vbe': 'VBScript file',
    '.wsf': 'Windows Script File',
    '.wsh': 'Windows Script Host file',
    '.hta': 'HTML Application',
    '.scr': 'Screen Saver',
    '.cpl': 'Control Panel Extension',
    '.msc': 'Microsoft Management Console file',
    '.lnk': 'Windows Shortcut file',
    '.pif': 'Program Information File',
}

MEDIUM_RISK_EXTENSIONS = {
    '.ps1': 'PowerShell script',
    '.psm1': 'PowerShell module',
    '.psd1': 'PowerShell data file',
    '.bat': 'Batch file',
    '.cmd': 'Command file',
    '.reg': 'Registry file',
    '.sys': 'System driver file',
    '.xll': 'Excel Add-in file'
}

# Example paths for MSI directories (for demonstration)
MSI_DIRECTORY_EXAMPLES = {
    "TARGETDIR": "C:\\Program Files\\[ProductName]",
    "ProgramFilesFolder": "C:\\Program Files (x86)",
    "ProgramFiles64Folder": "C:\\Program Files",
    "SystemFolder": "C:\\Windows\\System32",
    "AppDataFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming",
    "CommonAppDataFolder": "C:\\ProgramData",
    "CommonFilesFolder": "C:\\Program Files\\Common Files",
    "DesktopFolder": "C:\\Users\\%USERNAME%\\Desktop",
    "FavoritesFolder": "C:\\Users\\%USERNAME%\\Favorites",
    "FontsFolder": "C:\\Windows\\Fonts",
    "PersonalFolder": "C:\\Users\\%USERNAME%\\Documents",
    "SendToFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Microsoft\\Windows\\SendTo",
    "StartMenuFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu",
    "StartupFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",
    "TempFolder": "C:\\Windows\\Temp",
    "TemplateFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Microsoft\\Windows\\Templates",
    "WindowsFolder": "C:\\Windows",
    "WindowsVolume": "C:\\",
    "SourceDir": "D:\\SetupFiles\\[ProductName]",
    "INSTALLDIR": "C:\\Program Files\\[ProductName]",
    "LocalAppDataFolder": "C:\\Users\\%USERNAME%\\AppData\\Local",
    "CommonDesktopFolder": "C:\\Users\\Public\\Desktop",
    "CommonStartMenuFolder": "C:\\ProgramData\\Microsoft\\Windows\\Start Menu",
    "CommonStartupFolder": "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Startup",
    "AdminToolsFolder": "C:\\Users\\%USERNAME%\\AppData\\Roaming\\Microsoft\\Windows\\Start Menu\\Programs\\Administrative Tools",
    "CommonAdminToolsFolder": "C:\\ProgramData\\Microsoft\\Windows\\Start Menu\\Programs\\Administrative Tools",
    "NetworkFolder": "C:\\Windows\\Network Shortcuts",
    "MyPicturesFolder": "C:\\Users\\%USERNAME%\\Pictures",
    "MyMusicFolder": "C:\\Users\\%USERNAME%\\Music",
    "MyVideoFolder": "C:\\Users\\%USERNAME%\\Videos",
    "RecycleBinFolder": "C:\\$Recycle.Bin"
}

# Add a flag to control whether to use example paths
USE_EXAMPLE_PATHS = False

def create_section_header(text, count, color):
    """Create a section header with consistent styling"""
    header = QTreeWidgetItem([text, f"{count} items", "", ""])
    for i in range(4):
        header.setBackground(i, color)
    font = header.font(0)
    font.setBold(True)
    header.setFont(0, font)
    return header

def resolve_directory_path(dir_id, directory_table=None):
    """Resolve a directory ID to its final path by following the directory chain.
    
    Args:
        dir_id: The directory ID to resolve (e.g., "INSTALLDIR")
        directory_table: Optional directory table data from MSI. If not provided,
                        will return the original directory ID.
    
    Returns:
        The resolved directory path or the original dir_id if no resolution found
    """
    if not dir_id:
        return "NULL"
        
    # If directory table is provided, use it for resolution
    if directory_table:
        # Create a mapping of directory IDs to their parent directories
        dir_map = {}
        for row in directory_table.get("rows", []):
            if len(row) >= 2:
                dir_map[row[0]] = row[1]
        
        # Follow the chain until we reach NULL or a directory not in the map
        current_dir = dir_id
        last_valid_dir = dir_id  # Keep track of the last valid directory
        
        while current_dir and current_dir != "NULL" and current_dir in dir_map:
            last_valid_dir = current_dir  # Update last valid directory
            current_dir = dir_map[current_dir]
        
        # If we hit NULL, return the last valid directory
        if current_dir == "NULL":
            return last_valid_dir
        # If we hit a directory not in the map, return that directory
        return current_dir
    
    # If no directory table provided, just return the original ID
    return dir_id

def get_directory_path(dir_id, directory_table=None, parent=None):
    """Get the directory path, either using the example path or the resolved directory"""
    # First resolve the directory path
    resolved_dir = resolve_directory_path(dir_id, directory_table)
    
    # Then check if we should use example paths
    if USE_EXAMPLE_PATHS and resolved_dir in MSI_DIRECTORY_EXAMPLES:
        path = MSI_DIRECTORY_EXAMPLES[resolved_dir]
        
        # Replace [ProductName] placeholder with actual product name if available
        if "[ProductName]" in path and parent and hasattr(parent, 'tables_data'):
            # Find the Property table
            property_table = next((table for table in parent.tables_data if table["name"] == "Property"), None)
            
            if property_table:
                # Find the ProductName property
                product_name = "ExampleApp"  # Default fallback
                for row in property_table["rows"]:
                    if len(row) >= 2 and row[0] == "ProductName":
                        product_name = row[1]
                        break
                
                # Replace the placeholder with the actual product name
                path = path.replace("[ProductName]", product_name)
        
        return path
    
    # If not using example paths or no example exists, return the directory ID with brackets
    return f"[{resolved_dir}]"  # Return the resolved directory keyword with brackets

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
    
    # Get component and directory tables for mapping
    component_table = next((table for table in parent.tables_data if table["name"] == "Component"), None)
    directory_table = next((table for table in parent.tables_data if table["name"] == "Directory"), None)
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
                        if name and name != "NULL":
                            reg_path = f"{reg_path}\\{name}"
                        
                        # Only resolve placeholders if using example paths
                        if USE_EXAMPLE_PATHS:
                            # Resolve any property placeholders in the registry path and value
                            resolved_reg_path = resolve_property_values(reg_path, parent, directory_table)
                            resolved_value = resolve_property_values(value, parent, directory_table)
                            
                            # Determine registry value type
                            value_type, processed_value = determine_registry_value_type(resolved_value)
                            
                            # Don't display "NULL" value
                            if processed_value == "NULL":
                                display_value = ""
                            else:
                                display_value = processed_value
                                
                            # Format details with type in brackets
                            details = f"[{value_type}] {display_value}" if display_value else f"[{value_type}]"
                        else:
                            # Otherwise, keep original values
                            resolved_reg_path = reg_path
                            resolved_value = value
                            
                            # Don't display "NULL" value
                            if resolved_value == "NULL":
                                details = ""
                            else:
                                details = f"Value: {resolved_value}"
                        
                        is_persistence = any(re.match(pattern, resolved_reg_path, re.IGNORECASE) for pattern in AUTORUN_PATTERNS)
                        
                        registry_entries.append({
                            "type": "Persistence Mechanism" if is_persistence else "",
                            "entry": normalize_registry_path(resolved_reg_path),
                            "concern": "High - Persistence Mechanism" if is_persistence else "",
                            "details": details
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
            
            risk = assess_file_risk(install_path, install_dir, directory_table, parent)
            
            file_item = QTreeWidgetItem([
                "", 
                f"{get_directory_path(install_dir, directory_table, parent)}\\{install_path}",
                risk["concern"],
                f"Archived name: {display_name}"
            ])
            
            if risk["color"]:
                for i in range(4):
                    file_item.setForeground(i, risk["color"])
                    file_item.setBackground(i, risk["color"].lighter(180))
            
            files_header.addChild(file_item)
    
    if registry_entries:
        registry_header = create_section_header("Registry", len(registry_entries), QColor(255, 240, 245))
        parent.impact_tree.addTopLevelItem(registry_header)
        
        for reg_entry in registry_entries:
            # Extract the actual value from the details string (format: "Value: xxx")
            value = reg_entry["details"].replace("Value: ", "") if reg_entry["details"].startswith("Value: ") else reg_entry["details"]
            
            risk = assess_registry_risk(reg_entry["entry"], value, parent, directory_table)
            
            reg_item = QTreeWidgetItem([
                reg_entry["type"],
                reg_entry["entry"],
                risk["concern"],
                reg_entry["details"]
            ])
            
            if risk["color"]:
                for i in range(4):
                    reg_item.setForeground(i, risk["color"])
                    reg_item.setBackground(i, risk["color"].lighter(180))
            
            registry_header.addChild(reg_item)
    
    if service_entries:
        services_header = create_section_header("Services", len(service_entries), QColor(230, 230, 250))
        parent.impact_tree.addTopLevelItem(services_header)
        
        for service in service_entries:
            risk = assess_service_risk(service['Type'], service.get('StartType', ''), service["IsCritical"])
            
            service_item = QTreeWidgetItem([
                "",
                f"{service['Name']} ({service['DisplayName']})",
                risk["concern"],
                f"Type: {SERVICE_TYPES.get(service['Type'], service['Type'])}, Start: {SERVICE_START_TYPES.get(service['StartType'], service['StartType'])}"
            ])
            
            if risk["color"]:
                for i in range(4):
                    service_item.setForeground(i, risk["color"])
                    service_item.setBackground(i, risk["color"].lighter(180))
            
            services_header.addChild(service_item)
    
    if shortcut_entries:
        shortcuts_header = create_section_header("Shortcuts", len(shortcut_entries), QColor(240, 255, 240))
        parent.impact_tree.addTopLevelItem(shortcuts_header)
        
        for shortcut in shortcut_entries:
            shortcut_name = shortcut['Name']
            display_name = shortcut_name.split("|")[0] if "|" in shortcut_name else shortcut_name
            install_path = shortcut_name.split("|")[0] + ("\\" + shortcut_name.split("|")[1] if "|" in shortcut_name and shortcut_name.split("|")[1] else "")
            
            dir_path = get_directory_path(shortcut['Directory'], directory_table, parent)
            shortcut_item = QTreeWidgetItem([
                "",
                f"{get_directory_path(shortcut['Directory'], directory_table, parent)}\\{install_path}",
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
                get_directory_path(install_dir, directory_table, parent),
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
    """Normalize registry path to use single backslashes and remove trailing dashes"""
    if not path:
        return ""
    
    # Replace double backslashes with single backslashes
    path = path.replace('\\\\', '\\') 
    
    # Remove trailing dash if present
    if path.endswith('\\-'):
        path = path[:-2]  # Remove the last two characters ('\-')
    elif path.endswith('-'):
        path = path[:-1]  # Remove just the dash
        
    return path

def resolve_property_values(text, parent=None, directory_table=None):
    """Replace property placeholders in text with their actual values.
    
    Args:
        text: The text containing property placeholders like [PropertyName]
        parent: The parent object with tables_data attribute
        directory_table: Optional directory table for resolving directory IDs
        
    Returns:
        The text with placeholders replaced by their actual values
    """
    if not text:
        return text
        
    # Find all property placeholders in the format [PropertyName]
    placeholders = re.findall(r'\[(.*?)\]', text)
    
    if not placeholders:
        return text
    
    result = text
    
    # Try to resolve directory IDs first if using example paths
    if USE_EXAMPLE_PATHS and directory_table:
        for placeholder in placeholders:
            # Check if this placeholder might be a directory ID
            directory_path = get_directory_path(placeholder, directory_table, parent)
            if directory_path != f"[{placeholder}]":  # If successfully resolved to a path
                # Remove the brackets from directory_path if it's just returning the original in brackets
                if directory_path.startswith('[') and directory_path.endswith(']'):
                    directory_path = directory_path[1:-1]
                result = result.replace(f"[{placeholder}]", directory_path + "\\")
    
    # Then try to resolve property values if using example paths
    if USE_EXAMPLE_PATHS and parent and hasattr(parent, 'tables_data'):
        # Get the Property table
        property_table = next((table for table in parent.tables_data if table["name"] == "Property"), None)
        
        if property_table:
            # Create a map of property names to values
            property_map = {}
            for row in property_table["rows"]:
                if len(row) >= 2:
                    property_map[row[0]] = row[1]
            
            # Replace each placeholder with its value if available
            remaining_placeholders = re.findall(r'\[(.*?)\]', result)
            for placeholder in remaining_placeholders:
                if placeholder in property_map:
                    result = result.replace(f"[{placeholder}]", property_map[placeholder])
    
    return result

def assess_file_risk(filepath, install_dir, directory_table=None, parent=None):
    """Assess the risk level of a file based on its path and characteristics"""
    risk = {
        "level": "Low",
        "concern": "",
        "color": None
    }
    
    # Get the actual directory path
    dir_path = get_directory_path(install_dir, directory_table, parent)
    
    # Check critical directories
    if install_dir in CRITICAL_DIRECTORIES:
        risk["level"] = "High" if install_dir in ["SystemFolder", "System64Folder", "WindowsFolder", "StartupFolder"] else "Medium"
        risk["concern"] = f"{CRITICAL_DIRECTORIES[install_dir]} ({dir_path})"
        risk["color"] = QColor("red") if risk["level"] == "High" else QColor(255, 140, 0)  # Orange for Medium
        return risk
    
    # Check file extension
    ext = os.path.splitext(filepath)[1].lower()
    if ext in HIGH_RISK_EXTENSIONS:
        risk["level"] = "High"
        risk["concern"] = f"Rare script/executable type: {HIGH_RISK_EXTENSIONS[ext]}"
        risk["color"] = QColor("red")
        return risk
    elif ext in MEDIUM_RISK_EXTENSIONS:
        risk["level"] = "Medium"
        risk["concern"] = f"Script file: {MEDIUM_RISK_EXTENSIONS[ext]}"
        risk["color"] = QColor(255, 140, 0)  # Orange
        return risk
    
    # Check suspicious patterns
    if any(re.match(pattern, filepath, re.IGNORECASE) for pattern in SUSPICIOUS_FILE_PATTERNS):
        if "system32" in filepath.lower() or "syswow64" in filepath.lower():
            risk["level"] = "High"
            risk["concern"] = "System file modification"
            risk["color"] = QColor("red")
        elif "startup" in filepath.lower():
            risk["level"] = "High"
            risk["concern"] = "Startup persistence"
            risk["color"] = QColor("red")
        elif "common files" in filepath.lower():
            risk["level"] = "Medium"
            risk["concern"] = "Common Files directory modification"
            risk["color"] = QColor(255, 140, 0)  # Orange
        else:
            risk["level"] = "Medium"
            risk["concern"] = "Suspicious file location"
            risk["color"] = QColor(255, 140, 0)  # Orange
    
    return risk

def assess_registry_risk(reg_path, value, parent=None, directory_table=None):
    """Assess the risk level of a registry entry"""
    # If value contains property placeholders, resolve them only when using example paths
    if USE_EXAMPLE_PATHS:
        resolved_value = resolve_property_values(value, parent, directory_table)
    else:
        resolved_value = value
    
    risk = {
        "level": "Low",
        "concern": "",
        "color": None
    }
    
    # Check for persistence mechanisms (existing check)
    if any(re.match(pattern, reg_path, re.IGNORECASE) for pattern in AUTORUN_PATTERNS):
        risk["level"] = "High"
        risk["concern"] = "Persistence Mechanism"
        risk["color"] = QColor("red")
        return risk
    
    # Check for suspicious patterns
    if any(re.match(pattern, reg_path, re.IGNORECASE) for pattern in SUSPICIOUS_REGISTRY_PATTERNS):
        if "browser helper" in reg_path.lower() or "shell extension" in reg_path.lower():
            risk["level"] = "High"
            risk["concern"] = "Browser/Shell Extension Modification"
            risk["color"] = QColor("red")
        elif "security" in reg_path.lower() or "policy" in reg_path.lower():
            risk["level"] = "High"
            risk["concern"] = "Security Policy Modification"
            risk["color"] = QColor("red")
        else:
            risk["level"] = "Medium"
            risk["concern"] = "Security-sensitive registry modification"
            risk["color"] = QColor(255, 140, 0)  # Orange
    
    # Check for command execution in values (only for system executables)
    if resolved_value and any(ext in resolved_value.lower() for ext in [".exe", ".dll", ".sys", ".drv"]):
        if "system32" in resolved_value.lower() or "syswow64" in resolved_value.lower():
            risk["level"] = "High"
            risk["concern"] = "System executable in registry value"
            risk["color"] = QColor("red")
    
    return risk

def assess_service_risk(service_type, start_type, is_critical):
    """Assess the risk level of a service"""
    risk = {
        "level": "Low",
        "concern": "",
        "color": None
    }
    
    if is_critical:
        risk["level"] = "High"
        risk["concern"] = "Critical system service"
        risk["color"] = QColor("red")
    elif start_type == "0x00000002":  # Auto Start
        if service_type in ["0x00000001", "0x00000002"]:  # Kernel or File System Driver
            risk["level"] = "High"
            risk["concern"] = "Auto-start system driver"
            risk["color"] = QColor("red")
        else:
            risk["level"] = "Medium"
            risk["concern"] = "Automatic startup service"
            risk["color"] = QColor(255, 140, 0)  # Orange
    
    return risk

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

def toggle_example_paths(checked):
    """Toggle between resolved values and raw placeholders"""
    global USE_EXAMPLE_PATHS
    USE_EXAMPLE_PATHS = checked
    # Get the parent window and refresh the analysis
    parent = QApplication.activeWindow()
    if parent and hasattr(parent, 'tables_data') and parent.tables_data:
        analyze_installation_impact(parent)

def create_footprint_tab(parent):
    """Create the installation impact tab"""
    footprint_tab = QWidget()
    impact_layout = QVBoxLayout()
    footprint_tab.setLayout(impact_layout)
    
    # Add description label
    impact_description = QLabel("Analyze the MSI package to identify all system changes and artifacts that will be created during installation, including files, registry entries, services, and more.")
    impact_description.setWordWrap(True)
    impact_layout.addWidget(impact_description)
    
    # Add checkbox for resolving paths and properties
    checkbox_layout = QHBoxLayout()
    example_paths_checkbox = QCheckBox("Resolve paths and properties")
    example_paths_checkbox.setChecked(USE_EXAMPLE_PATHS)
    example_paths_checkbox.stateChanged.connect(toggle_example_paths)
    checkbox_layout.addWidget(example_paths_checkbox)
    checkbox_layout.addStretch()
    impact_layout.addLayout(checkbox_layout)
    
    # Tree view for installation impact details
    parent.impact_tree = QTreeWidget()
    parent.impact_tree.setColumnCount(4)
    parent.impact_tree.setHeaderLabels(["Type", "Entry", "Concern", "Details"])
    parent.impact_tree.setAlternatingRowColors(True)
    parent.impact_tree.header().setSectionResizeMode(QHeaderView.ResizeToContents)
    
    # Enable context menu for impact tree
    parent.impact_tree.setContextMenuPolicy(Qt.CustomContextMenu)
    parent.impact_tree.customContextMenuRequested.connect(parent.show_impact_context_menu)
    
    impact_layout.addWidget(parent.impact_tree, 1)
    
    return footprint_tab 

def determine_registry_value_type(value):
    """Determine the registry value type based on its prefix
    
    Prefixes:
    #x - REG_BINARY (hexadecimal)
    #% - REG_EXPAND_SZ (expandable string)
    # - REG_DWORD (integer)
    [~] - REG_MULTI_SZ (null-delimited string list)
    (none) - REG_SZ (string)
    
    Returns:
        Tuple of (value_type, processed_value)
    """
    if not value or value == "NULL":
        return "REG_SZ", value
    
    # Check for hexadecimal (REG_BINARY)
    if value.startswith('#x'):
        return "REG_BINARY", value[2:]
    
    # Check for expandable string (REG_EXPAND_SZ)
    if value.startswith('#%'):
        return "REG_EXPAND_SZ", value[2:]
    
    # Check for integer (REG_DWORD)
    if value.startswith('#'):
        # If it has multiple # signs, only the first one is ignored
        if value.startswith('##'):
            return "REG_SZ", value[1:]
        return "REG_DWORD", value[1:]
    
    # Check for multi-string (REG_MULTI_SZ)
    if '[~]' in value:
        return "REG_MULTI_SZ", value
    
    # Default to string (REG_SZ)
    return "REG_SZ", value 