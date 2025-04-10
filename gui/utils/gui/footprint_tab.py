"""
Installation Impact tab functionality for the MSI Parser GUI
"""

import re
from PyQt5.QtWidgets import (QTreeWidgetItem, QHeaderView, QTreeWidget, QApplication,
                            QCheckBox, QHBoxLayout, QWidget, QVBoxLayout, QLabel)
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor

# Import the constants with specific names
from utils.gui.footprint_constants import (
    REGISTRY_ROOTS, SERVICE_TYPES, SERVICE_START_TYPES, AUTORUN_PATTERNS,
)

# Import utility functions with specific names
from utils.gui.footprint_utils import (
    create_section_header, get_directory_path, normalize_registry_path, 
    resolve_property_values, assess_file_risk, assess_registry_risk,
    assess_service_risk, determine_registry_value_type
)

# Add a flag to control whether to use example paths
USE_EXAMPLE_PATHS = False

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
                f"{dir_path}\\{install_path}",
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