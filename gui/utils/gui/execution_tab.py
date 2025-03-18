"""
MSI Workflow Analysis tab functionality for the MSI Parser GUI
"""

from PyQt5.QtWidgets import (QTreeWidgetItem, QApplication)
from PyQt5.QtGui import QColor
from PyQt5.QtCore import Qt

# Define severity levels for highlighting
SEVERITY_LEVELS = {
    "NONE": {"color": "black", "icon": "SP_DialogApplyButton"},
    "LOW": {"color": "blue", "icon": "SP_MessageBoxInformation"},
    "LOW-MEDIUM": {"color": "darkblue", "icon": "SP_MessageBoxInformation"},
    "MEDIUM": {"color": "darkorange", "icon": "SP_MessageBoxWarning"},
    "MEDIUM-HIGH": {"color": "orangered", "icon": "SP_MessageBoxWarning"},
    "HIGH": {"color": "red", "icon": "SP_MessageBoxCritical"},
    "CRITICAL": {"color": "darkred", "icon": "SP_MessageBoxCritical"}
}

# Phase boundaries for installation sequence
PHASE_BOUNDARIES = {
    "Initialization Phase": 1000,
    "Validation Phase": 2000,
    "Preparation Phase": 3000,
    "Execution Phase": 4000,
    "Commit Phase": 5000,
    "Rollback Phase": 6000,
    "Finalization Phase": float('inf')  # For any sequence numbers beyond other phases
}

# Colors for different phases
PHASE_COLORS = {
    "Initialization Phase": "#E3F2FD",  # Light Blue
    "Validation Phase": "#F3E5F5",      # Light Purple
    "Preparation Phase": "#E8F5E9",     # Light Green
    "Execution Phase": "#FFF3E0",       # Light Orange
    "Commit Phase": "#EFEBE9",          # Light Brown
    "Rollback Phase": "#FFEBEE",        # Light Red
    "Finalization Phase": "#F5F5F5"     # Light Gray
}

# Custom actions mapping
custom_actions = {}  # Will be populated during analysis

def clean_action_name(action_name):
    """Clean action names by removing non-printable characters"""
    # Replace any non-printable characters with empty string
    return ''.join(char for char in action_name if char.isprintable())

def create_phase_header(phase_name, parent):
    if phase_name != "Initialization Phase":
        spacer = QTreeWidgetItem(["", "", "", "", ""])
        parent.sequence_tree.addTopLevelItem(spacer)
    
    header = QTreeWidgetItem([f"📋 {phase_name.upper()}", "", "", "", ""])
    header_color = QColor(PHASE_COLORS[phase_name])
    header_color.setAlpha(80)
    
    for i in range(5):
        header.setBackground(i, header_color)
    
    font = header.font(0)
    font.setBold(True)
    font.setPointSize(font.pointSize() + 1)
    header.setFont(0, font)
    header.setTextAlignment(0, Qt.AlignCenter)
    
    parent.sequence_tree.addTopLevelItem(header)

def create_sequence_item(sequence, action, condition, impact, severity, phase_name, parent):
    item = QTreeWidgetItem([sequence, action, condition, "", impact])
    
    light_color = QColor(PHASE_COLORS[phase_name])
    light_color.setAlpha(40)
    for i in range(5):
        item.setBackground(i, light_color)
    
    if severity in SEVERITY_LEVELS:
        item.setForeground(4, QColor(SEVERITY_LEVELS[severity]["color"]))
        if SEVERITY_LEVELS[severity]["icon"]:
            icon = QApplication.style().standardIcon(getattr(QApplication.style(), SEVERITY_LEVELS[severity]["icon"]))
            item.setIcon(4, icon)
    
    if severity in ["HIGH", "CRITICAL"]:
        for col in range(5):
            item.setBackground(col, QColor(255, 240, 240))
    
    parent.sequence_tree.addTopLevelItem(item)

def display_workflow_analysis(parent, target_widget='workflow'):
    """Display the MSI Installation Workflow Analysis
    
    Args:
        parent: The parent window object
        target_widget: Which widget to display the content in ('workflow' or 'help')
    """
    try:
        # Clear existing items
        if target_widget == 'workflow':
            parent.sequence_tree.clear()
        
        # Get the InstallExecuteSequence table
        if not parent.tables_data:
            return
            
        sequence_table = next((table for table in parent.tables_data if table["name"] == "InstallExecuteSequence"), None)
        if not sequence_table:
            return
            
        current_phase = None
        
        # Process each action in the sequence
        for row in sequence_table.get("rows", []):
            action = clean_action_name(row.get("Action", "").strip())
            sequence = row.get("Sequence", "")
            condition = row.get("Condition", "").strip()
            
            # Skip empty actions
            if not action:
                continue
            
            # Determine which phase this action belongs to
            seq_num_int = int(sequence) if sequence.isdigit() else 0
            
            # Check if we need to add a new phase header
            phase_name = next((phase for phase, boundary in PHASE_BOUNDARIES.items() if seq_num_int < boundary), "Finalization Phase")
            
            if phase_name != current_phase:
                # Add a phase header
                current_phase = phase_name
                create_phase_header(phase_name, parent)
            
            # Evaluate the impact of the action
            impact, severity = evaluate_action_impact(action)
            create_sequence_item(sequence, action, condition, impact, severity, current_phase, parent)
    
    except Exception as e:
        parent.show_error("Error", f"Failed to load workflow analysis: {str(e)}")

def analyze_install_sequence(parent):
    """Analyze the InstallExecuteSequence from the loaded MSI file"""
    if not parent.tables_data:
        parent.show_warning("Warning", "No MSI file loaded or tables not parsed.")
        return
    
    # Find the key tables
    sequence_table = next((table for table in parent.tables_data if table["name"] == "InstallExecuteSequence"), None)
    if not sequence_table:
        parent.show_warning("Warning", "InstallExecuteSequence table not found in this MSI.")
        return
    
    # Clear the tree widget
    parent.sequence_tree.clear()
    
    # Set up the headers
    parent.sequence_tree.setHeaderLabels(["Sequence", "Action", "Condition", "Type", "Impact"])
    
    # Initialize current phase
    current_phase = None
    
    # Create a dictionary of custom actions for quick lookup
    custom_actions = {}
    if any(table["name"] == "CustomAction" for table in parent.tables_data):
        for row in next((table["rows"] for table in parent.tables_data if table["name"] == "CustomAction"), []):
            if len(row) >= 2:  # Ensure we have at least Action and Type
                action_name = row[0]
                action_type = row[1] if len(row) > 1 else ""
                source = row[2] if len(row) > 2 else ""
                target = row[3] if len(row) > 3 else ""
                custom_actions[action_name] = {
                    "Type": action_type,
                    "Source": source,
                    "Target": target
                }
    
    # Collect registry operations for analysis
    registry_operations = []
    if any(table["name"] == "Registry" for table in parent.tables_data):
        for row in next((table["rows"] for table in parent.tables_data if table["name"] == "Registry"), []):
            if len(row) >= 5:  # Registry, Root, Key, Name, Value
                registry_key = row[0]
                root = row[1]
                key_path = row[2]
                name = row[3]
                value = row[4]
                
                # Check for autorun keys, which are commonly used for persistence
                is_persistence = False
                if (root == "2" and  # HKLM
                    ("Run" in key_path or 
                     "RunOnce" in key_path or 
                     "\\Microsoft\\Windows\\CurrentVersion\\Run" in key_path or
                     "StartupApproved" in key_path or
                     "Shell Extensions" in key_path)):
                    is_persistence = True
                
                registry_operations.append({
                    "Registry": registry_key,
                    "Root": root,
                    "Key": key_path,
                    "Name": name,
                    "Value": value,
                    "IsPersistence": is_persistence
                })
    
    # Check for service installations
    service_installs = []
    if any(table["name"] == "ServiceInstall" for table in parent.tables_data):
        for row in next((table["rows"] for table in parent.tables_data if table["name"] == "ServiceInstall"), []):
            if len(row) >= 4:  # Basic service information
                service_name = row[1] if len(row) > 1 else ""
                display_name = row[2] if len(row) > 2 else ""
                service_type = row[3] if len(row) > 3 else ""
                start_type = row[4] if len(row) > 4 else ""
                
                # Auto-start services (2) and system services (type 16) are interesting
                is_critical = (start_type == "2" or service_type == "16")
                
                service_installs.append({
                    "Name": service_name,
                    "DisplayName": display_name,
                    "Type": service_type,
                    "StartType": start_type,
                    "IsCritical": is_critical
                })
    
    # Sort the sequence by the Sequence number
    sorted_sequence = sorted(sequence_table["rows"], key=lambda x: int(x[2]) if x[2].isdigit() else 0)
    
    # Process each action in the sequence
    for row in sorted_sequence:
        if len(row) >= 3:  # Ensure we have Action, Condition, and Sequence
            action_name = row[0]
            condition = row[1]
            sequence_num = row[2]
            
            # Clean the action name
            action_name = clean_action_name(action_name)
            
            # Skip empty actions
            if not action_name:
                continue
            
            # Determine which phase this action belongs to
            seq_num_int = int(sequence_num) if sequence_num.isdigit() else 0
            
            # Check if we need to add a new phase header
            phase_name = next((phase for phase, boundary in PHASE_BOUNDARIES.items() if seq_num_int < boundary), "Finalization Phase")
            
            if phase_name != current_phase:
                # Add a phase header
                current_phase = phase_name
                create_phase_header(phase_name, parent)
            
            # Evaluate the impact of the action
            impact, severity = evaluate_action_impact(action_name)
            create_sequence_item(sequence_num, action_name, condition, impact, severity, current_phase, parent)
    
    # Add system services installations directly in the tree
    system_services = sum(1 for service in service_installs if service["IsCritical"])
    
    if system_services > 0:
        # Add a spacer
        spacer = QTreeWidgetItem(["", "", "", "", ""])
        parent.sequence_tree.addTopLevelItem(spacer)
        
        # Add a header
        services_header = QTreeWidgetItem(["⚙️ SYSTEM SERVICES", "", "", "", ""])
        services_header.setBackground(0, QColor(240, 240, 255))  # Light purple
        services_header.setBackground(1, QColor(240, 240, 255))
        services_header.setBackground(2, QColor(240, 240, 255))
        services_header.setBackground(3, QColor(240, 240, 255))
        services_header.setBackground(4, QColor(240, 240, 255))
        font = services_header.font(0)
        font.setBold(True)
        services_header.setFont(0, font)
        parent.sequence_tree.addTopLevelItem(services_header)
        
        # Add items for each service
        for service in service_installs:
            if service["IsCritical"]:
                service_item = QTreeWidgetItem(["", f"Service: {service['Name']}", service["DisplayName"], service["Type"], service["StartType"]])
                service_item.setForeground(1, QColor("blue"))
                
                # Light blue background for services
                light_blue = QColor(240, 245, 255)
                for i in range(5):
                    service_item.setBackground(i, light_blue)
                
                parent.sequence_tree.addTopLevelItem(service_item)
    
    # Resize columns to content
    for i in range(5):
        parent.sequence_tree.resizeColumnToContents(i)
    
    total_actions = len(sorted_sequence)
    parent.show_status(f"Analyzed InstallExecuteSequence: {total_actions} actions")

def evaluate_custom_action_impact(ca_type):
    """Evaluate the impact of a custom action based on its type (backward compatibility)"""
    try:
        ca_type_int = int(ca_type)
        
        # Base type is lower 6 bits (0x003F)
        base_type = ca_type_int & 0x003F
        
        # Source location is bits 7-8 (0x00C0)
        source_type = (ca_type_int & 0x00C0) >> 6
        
        # Target location is bits 9-10 (0x0300)
        target_type = (ca_type_int & 0x0300) >> 8
        
        # Execution mode is bits 11-12 (0x0C00)
        execution_mode = (ca_type_int & 0x0C00) >> 10
        
        # Check for no impersonation flag (bit 13 or 0x1000)
        no_impersonation = (ca_type_int & 0x1000) != 0
        
        # Check for 64-bit flag (bit 14 or 0x2000)
        is_64bit = (ca_type_int & 0x2000) != 0
        
        # Check for hidden flag (bit 15 or 0x4000)
        is_hidden = (ca_type_int & 0x4000) != 0
        
        # Extended type information (bits 16-31, rarely used)
        extended_type = (ca_type_int & 0xFFFF0000) >> 16
        
        # Determine base action type
        base_type_desc = {
            1: "DLL function call",
            2: "EXE execution",
            5: "JScript execution",
            6: "VBScript execution",
            7: "Install operations",
            19: "Error message",
            34: "Directory creation",
            35: "Registry operation",
            37: "Formatting data",
            38: "Command-line environment",
            50: "System-wide command execution",
            51: "System-wide property setting",
            65: "UI validation",
            70: "Custom setup execution",
            98: "PowerShell script execution",
            210: "External program execution",
            226: "Custom installer extension",
            257: "Software detection"
        }.get(base_type, "Unknown action")
        
        # Special handling for certain combined types
        if ca_type_int == 51:
            return "Sets system property", "MEDIUM"
        elif ca_type_int == 307:
            return "Directory operation with custom data", "MEDIUM"
        elif ca_type_int == 70:
            return "Custom setup execution", "HIGH"
        elif ca_type_int == 257:
            return "Software detection routine", "MEDIUM"
        elif ca_type_int == 210:
            return "External program execution", "HIGH"
        
        # Determine primary impact and severity
        primary_impact = "Unknown impact"
        severity = "MEDIUM"
        
        # Executable execution types (most dangerous)
        if base_type in [2, 18, 34, 50, 210]:
            if no_impersonation:
                primary_impact = "Executes external program with elevated privileges"
                severity = "CRITICAL"
            else:
                primary_impact = "Executes external program"
                severity = "HIGH"
        
        # Script execution types (also dangerous)
        elif base_type in [5, 6, 98]:  # JScript, VBScript, PowerShell
            if no_impersonation:
                primary_impact = f"Executes {base_type_desc} with elevated privileges"
                severity = "CRITICAL"
            else:
                primary_impact = f"Executes {base_type_desc}"
                severity = "HIGH"
        
        # DLL execution types (potentially dangerous)
        elif base_type in [1, 17, 33, 49, 257]:
            if no_impersonation:
                primary_impact = "Calls DLL function with elevated privileges"
                severity = "HIGH"
            else:
                primary_impact = "Calls DLL function"
                severity = "MEDIUM"
        
        # Registry operations (medium risk)
        elif base_type in [35]:
            primary_impact = "Modifies registry"
            severity = "MEDIUM"
        
        # Command-line environment (variable risk)
        elif base_type in [38]:
            if no_impersonation:
                primary_impact = "Executes command-line with elevated privileges"
                severity = "HIGH"
            else:
                primary_impact = "Executes command-line operations"
                severity = "MEDIUM"
        
        # UI operations (typically low risk)
        elif base_type in [65]:
            primary_impact = "UI validation operation"
            severity = "LOW"
            
        # Setup operations (higher risk)
        elif base_type in [70, 226]:
            primary_impact = "Custom setup operation"
            severity = "HIGH"
        
        # Other types
        else:
            if no_impersonation:
                primary_impact = f"{base_type_desc} with elevated privileges"
                severity = "MEDIUM"
            else:
                primary_impact = base_type_desc
                severity = "LOW"
        
        # Adjust severity for deferred execution with no impersonation (system context)
        if execution_mode == 2 and no_impersonation:  # Deferred execution + no impersonation
            if "with elevated privileges" not in primary_impact:
                primary_impact += " with elevated privileges (deferred)"
            severity = "CRITICAL"
        
        # Adjust severity for hidden actions
        if is_hidden:
            primary_impact += " (hidden action)"
            
            # Upgrade severity if hidden
            if severity == "MEDIUM":
                severity = "HIGH"
            elif severity == "LOW":
                severity = "MEDIUM"
        
        return primary_impact, severity
        
    except (ValueError, TypeError):
        return "Unknown custom action", "LOW"

def evaluate_standard_action_impact(action_name):
    """Evaluate the impact of a standard action based on its name (backward compatibility)"""
    return evaluate_action_impact(action_name)

def evaluate_action_impact(action):
    if action.startswith("_") and len(action) > 30 and any(c.isdigit() for c in action):
        return "Custom action with generated name", "MEDIUM"
    
    if action.startswith("DIRCA_"):
        return "Directory customization action", "LOW"
    if action.startswith("ERRCA_"):
        return "Error handling action", "LOW"
    if action.startswith("AI_"):
        if "DETECT" in action or "CHECK" in action:
            return "Advanced Installer detection action", "LOW"
        if "SET" in action or "WRITE" in action:
            return "Advanced Installer property setting", "MEDIUM"
        if "INSTALL" in action or "EXECUTE" in action:
            return "Advanced Installer execution action", "MEDIUM"
        return "Advanced Installer custom action", "MEDIUM"
    
    high_impact = {
        "RemoveExistingProducts": "Removes existing products",
        "RegisterServices": "Installs services",
        "StartServices": "Starts services",
        "DeleteServices": "Removes services from system"
    }
    
    medium_high = {
        "InstallFiles": "Copies files to system",
        "WriteRegistryValues": "Modifies registry",
        "RemoveRegistryValues": "Removes registry entries",
        "RemoveFiles": "Removes files",
        "MoveFiles": "Moves files within the system",
        "PatchFiles": "Applies patches to files",
        "SelfRegModules": "Self-registers modules"
    }
    
    medium = {
        "CreateShortcuts": "Creates shortcuts",
        "RegisterClassInfo": "Registers COM classes",
        "RegisterExtensionInfo": "Registers file extensions",
        "RegisterProgIdInfo": "Registers ProgIDs",
        "RegisterMIMEInfo": "Registers MIME types",
        "RegisterTypeLibraries": "Registers type libraries",
        "PublishComponents": "Publishes components",
        "DuplicateFiles": "Creates file copies",
        "RegisterComPlus": "Registers COM+ applications",
        "WriteEnvironmentStrings": "Modifies environment variables",
        "InstallODBC": "Installs ODBC drivers",
        "AllocateRegistrySpace": "Reserves registry space"
    }
    
    low_medium = {
        "PublishFeatures": "Publishes features",
        "PublishProduct": "Publishes product info",
        "WriteIniValues": "Writes INI values",
        "SetODBCFolders": "Configures ODBC directories",
        "RemoveODBC": "Removes ODBC drivers",
        "UnregisterComPlus": "Unregisters COM+ applications",
        "SelfUnregModules": "Self-unregisters modules",
        "UnregisterTypeLibraries": "Unregisters type libraries",
        "RemoveEnvironmentStrings": "Removes environment variables",
        "RemoveDuplicateFiles": "Removes duplicate files",
        "UnregisterClassInfo": "Unregisters COM classes",
        "UnregisterExtensionInfo": "Unregisters file extensions",
        "UnregisterProgIdInfo": "Unregisters ProgIDs",
        "UnregisterMIMEInfo": "Unregisters MIME types",
        "BindImage": "Binds executable to imported DLLs",
        "ProcessComponents": "Processes component registrations"
    }
    
    low = {
        "RegisterUser": "Registers user info",
        "RegisterProduct": "Registers product",
        "RegisterFonts": "Registers fonts",
        "CreateFolders": "Creates folders",
        "RemoveFolders": "Removes folders",
        "UnregisterFonts": "Unregisters fonts",
        "RemoveIniValues": "Removes INI values",
        "RemoveShortcuts": "Removes shortcuts",
        "IsolateComponents": "Manages shared components",
        "UnpublishComponents": "Unpublishes components",
        "UnpublishFeatures": "Unpublishes features",
        "MsiPublishAssemblies": "Publishes .NET assemblies",
        "MsiUnpublishAssemblies": "Unpublishes .NET assemblies"
    }
    
    no_impact = {
        "CostInitialize": "Initializes costing",
        "FileCost": "Calculates disk space",
        "CostFinalize": "Finalizes cost calculation",
        "InstallValidate": "Validates installation",
        "AppSearch": "Searches for applications",
        "LaunchConditions": "Checks prerequisites",
        "FindRelatedProducts": "Finds related products",
        "RMCCPSearch": "Finds RMS CCPs",
        "CCPSearch": "Searches for compatible products",
        "ValidateProductID": "Validates product ID"
    }
    
    if action in high_impact:
        return high_impact[action], "HIGH"
    if action in medium_high:
        return medium_high[action], "MEDIUM-HIGH"
    if action in medium:
        return medium[action], "MEDIUM"
    if action in low_medium:
        return low_medium[action], "LOW-MEDIUM"
    if action in low:
        return low[action], "LOW"
    if action in no_impact:
        return no_impact[action], "NONE"
    if action == "ExecuteAction":
        return "Triggers Execute Sequence", "CRITICAL"
    
    return "Standard MSI action", "LOW" 