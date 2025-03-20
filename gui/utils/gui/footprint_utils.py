"""
Utility functions for the Installation Impact tab functionality
"""

import re
import os
from PyQt5.QtWidgets import QTreeWidgetItem
from PyQt5.QtGui import QColor

from utils.gui.footprint_constants import (
    AUTORUN_PATTERNS, SUSPICIOUS_FILE_PATTERNS, SUSPICIOUS_REGISTRY_PATTERNS, CRITICAL_DIRECTORIES,
    HIGH_RISK_EXTENSIONS, MEDIUM_RISK_EXTENSIONS, MSI_DIRECTORY_EXAMPLES
)

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
    # Import this at function level to avoid circular imports
    from utils.gui.footprint_tab import USE_EXAMPLE_PATHS
    
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

def create_section_header(text, count, color):
    """Create a section header with consistent styling"""
    header = QTreeWidgetItem([text, f"{count} items", "", ""])
    for i in range(4):
        header.setBackground(i, color)
    font = header.font(0)
    font.setBold(True)
    header.setFont(0, font)
    return header

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
    # Import this at function level to avoid circular imports
    from utils.gui.footprint_tab import USE_EXAMPLE_PATHS
    
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
    # Import this at function level to avoid circular imports
    from utils.gui.footprint_tab import USE_EXAMPLE_PATHS
    
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