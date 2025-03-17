"""
MSI Workflow Analysis tab functionality for the MSI Parser GUI
"""

import os
import markdown
from PyQt5.QtWidgets import (QTreeWidgetItem, QApplication)
from PyQt5.QtGui import QColor

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

def display_workflow_analysis(parent, target_widget='workflow'):
    """Display the MSI Installation Workflow Analysis
    
    Args:
        parent: The parent window object
        target_widget: Which widget to display the content in ('workflow' or 'help')
    """
    # Load the workflow markdown file
    markdown_path = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(__file__))), 
                                "MSI_Installation_Workflow_Analysis.md")
    
    # Determine which widget to use
    if target_widget == 'workflow':
        display_widget = parent.workflow_html if hasattr(parent, 'workflow_html') else None
    else:
        display_widget = parent.help_html if hasattr(parent, 'help_html') else None
    
    # Check if the widget exists
    if display_widget is None:
        print(f"Warning: Target widget '{target_widget}' not found")
        return
    
    if not os.path.exists(markdown_path):
        display_widget.setHtml("<h1>MSI Workflow Analysis File Not Found</h1>"
                                   "<p>The workflow analysis file could not be found. "
                                   "Please make sure the file exists at:</p>"
                                   f"<code>{markdown_path}</code>")
        return
    
    try:
        with open(markdown_path, 'r') as f:
            md_content = f.read()
            
        # Convert markdown to HTML
        html_content = markdown.markdown(md_content, extensions=['tables', 'fenced_code'])
        
        # Add some CSS styling
        styled_html = f"""
        <!DOCTYPE html>
        <html>
        <head>
        <style>
            body {{
                font-family: Arial, sans-serif;
                margin: 20px;
                line-height: 1.5;
            }}
            h1, h2, h3 {{
                color: #2c3e50;
            }}
            h2 {{
                border-bottom: 1px solid #eee;
                padding-bottom: 5px;
            }}
            blockquote {{
                background-color: #f6f8fa;
                border-left: 4px solid #dfe2e5;
                padding: 10px 15px;
                margin: 15px 0;
            }}
            blockquote p {{
                margin: 0;
            }}
            code {{
                font-family: Consolas, monospace;
                background-color: #f6f8fa;
                padding: 2px 4px;
                border-radius: 3px;
            }}
            strong {{
                font-weight: bold;
            }}
            .security-note {{
                background-color: #fff8dc;
                border: 1px solid #ffeb99;
                padding: 10px;
                margin: 10px 0;
                border-radius: 5px;
            }}
            .high-impact {{
                color: red;
                font-weight: bold;
            }}
            .medium-impact {{
                color: darkorange;
                font-weight: bold;
            }}
            .low-impact {{
                color: blue;
            }}
        </style>
        </head>
        <body>
        {html_content}
        </body>
        </html>
        """
        
        # Set the HTML content
        display_widget.setHtml(styled_html)
        
        # Make links clickable (for any external links in the markdown)
        display_widget.setOpenExternalLinks(True)
        
    except Exception as e:
        parent.show_error("Error", f"Failed to load workflow analysis: {str(e)}")

def analyze_install_sequence(parent):
    """Analyze the InstallExecuteSequence from the loaded MSI file"""
    if not parent.tables_data:
        parent.show_warning("Warning", "No MSI file loaded or tables not parsed.")
        return
    
    # Find the key tables
    install_sequence = None
    custom_action_table = None
    registry_table = None
    file_table = None
    service_install_table = None
    
    for table in parent.tables_data:
        if table["name"] == "InstallExecuteSequence":
            install_sequence = table
        elif table["name"] == "CustomAction":
            custom_action_table = table
        elif table["name"] == "Registry":
            registry_table = table
        elif table["name"] == "File":
            file_table = table
        elif table["name"] == "ServiceInstall":
            service_install_table = table
    
    if not install_sequence:
        parent.show_warning("Warning", "InstallExecuteSequence table not found in this MSI.")
        return
    
    # Clear the tree widget
    parent.sequence_tree.clear()
    
    # Set up the headers
    parent.sequence_tree.setHeaderLabels(["Sequence", "Action", "Condition", "Type", "Impact"])
    
    # Create a dictionary of custom actions for quick lookup
    custom_actions = {}
    if custom_action_table:
        for row in custom_action_table["rows"]:
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
    if registry_table:
        for row in registry_table["rows"]:
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
    if service_install_table:
        for row in service_install_table["rows"]:
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
    sorted_sequence = sorted(install_sequence["rows"], key=lambda x: int(x[2]) if x[2].isdigit() else 0)
    
    # Define phase boundaries for different installation stages
    phase_boundaries = {
        "Initialization Phase": 900,
        "Validation Phase": 1500,
        "Execution Phase": 6000, 
        "Finalization Phase": float('inf')
    }
    
    # Reference to track which phase we're in as we process items
    current_phase = None
    
    # Add headers for each phase
    phase_colors = {
        "Initialization Phase": QColor(240, 248, 255),  # Light blue
        "Validation Phase": QColor(240, 255, 240),      # Light green
        "Execution Phase": QColor(255, 248, 240),       # Light orange
        "Finalization Phase": QColor(245, 245, 245)     # Light grey
    }
    
    # Process each action in the sequence
    for row in sorted_sequence:
        if len(row) >= 3:  # Ensure we have Action, Condition, and Sequence
            action_name = row[0]
            condition = row[1]
            sequence_num = row[2]
            
            # Determine which phase this action belongs to
            seq_num_int = int(sequence_num) if sequence_num.isdigit() else 0
            
            # Check if we need to add a new phase header
            phase_name = None
            for phase, boundary in phase_boundaries.items():
                if seq_num_int < boundary:
                    phase_name = phase
                    break
            
            if phase_name != current_phase:
                # Add a phase header
                current_phase = phase_name
                phase_color = phase_colors[phase_name]
                
                # Add blank spacer row before each new phase (except the first)
                if phase_name != "Initialization Phase":
                    spacer = QTreeWidgetItem(["", "", "", "", ""])
                    parent.sequence_tree.addTopLevelItem(spacer)
                
                phase_header = QTreeWidgetItem(["", f"📋 {phase_name.upper()}", "", "", ""])
                for i in range(5):
                    phase_header.setBackground(i, phase_color)
                
                font = phase_header.font(0)
                font.setBold(True)
                phase_header.setFont(1, font)
                parent.sequence_tree.addTopLevelItem(phase_header)
            
            # Create a tree item
            item = QTreeWidgetItem([sequence_num, action_name, condition, "", ""])
            
            # Add light background color based on the phase
            light_color = QColor(phase_colors[current_phase])
            light_color.setAlpha(40)  # Make it very light
            for i in range(5):
                item.setBackground(i, light_color)
            
            # Check if this is a custom action
            if action_name in custom_actions:
                # This is a custom action
                ca_info = custom_actions[action_name]
                ca_type = ca_info["Type"]
                
                # Add the custom action type
                item.setText(3, ca_type)
                
                # Evaluate the impact based on the custom action type
                impact, severity = evaluate_custom_action_impact(ca_type)
                item.setText(4, impact)
                
                # Set text color based on severity
                if severity in SEVERITY_LEVELS:
                    item.setForeground(4, QColor(SEVERITY_LEVELS[severity]["color"]))
                    # Add an icon if available
                    if SEVERITY_LEVELS[severity]["icon"]:
                        icon_name = SEVERITY_LEVELS[severity]["icon"]
                        icon = QApplication.style().standardIcon(getattr(QApplication.style(), icon_name))
                        item.setIcon(4, icon)
                
                # For custom actions, add details prefix to the action name
                item.setText(1, f"{action_name} 🔧")
                
                # Highlight the entire row for high-impact custom actions
                if severity in ["HIGH", "CRITICAL"]:
                    for col in range(5):
                        item.setBackground(col, QColor(255, 240, 240))  # Light red background
                
                # For suspicious targets, append a warning indicator
                if ca_info["Target"] and isinstance(ca_info["Target"], str):
                    target_lower = ca_info["Target"].lower()
                    
                    # Look for suspicious commands or parameters
                    suspicious_patterns = [
                        "powershell", "cmd.exe", "http://", "https://", "ftp://", 
                        "regsvr32", "rundll32", "wscript", "cscript", "certutil",
                        "bitsadmin", "reg add", "reg delete", "regedit", "sc create"
                    ]
                    
                    for pattern in suspicious_patterns:
                        if pattern in target_lower:
                            item.setText(1, f"{action_name} 🔧 ⚠️")
                            break
            else:
                # This is a standard action
                impact, severity = evaluate_standard_action_impact(action_name)
                item.setText(4, impact)
                
                # Set text color based on severity
                if severity in SEVERITY_LEVELS:
                    item.setForeground(4, QColor(SEVERITY_LEVELS[severity]["color"]))
                    # Add an icon if available
                    if SEVERITY_LEVELS[severity]["icon"]:
                        icon_name = SEVERITY_LEVELS[severity]["icon"]
                        icon = QApplication.style().standardIcon(getattr(QApplication.style(), icon_name))
                        item.setIcon(4, icon)
                
                # Highlight the entire row for high-impact standard actions
                if severity in ["HIGH", "CRITICAL"]:
                    for col in range(5):
                        item.setBackground(col, QColor(255, 240, 240))  # Light red background
            
            # Add the item to the tree as a top-level item
            parent.sequence_tree.addTopLevelItem(item)
    
    # Add registry operations for persistence directly in the tree
    persistence_mechanisms = 0
    for reg_op in registry_operations:
        if reg_op["IsPersistence"]:
            persistence_mechanisms += 1
            
    if persistence_mechanisms > 0:
        # Add a spacer
        spacer = QTreeWidgetItem(["", "", "", "", ""])
        parent.sequence_tree.addTopLevelItem(spacer)
        
        # Add a header
        registry_header = QTreeWidgetItem(["", "🔐 PERSISTENCE MECHANISMS", "", "", ""])
        registry_header.setBackground(0, QColor(255, 245, 240))  # Light amber
        registry_header.setBackground(1, QColor(255, 245, 240))
        registry_header.setBackground(2, QColor(255, 245, 240))
        registry_header.setBackground(3, QColor(255, 245, 240))
        registry_header.setBackground(4, QColor(255, 245, 240))
        font = registry_header.font(0)
        font.setBold(True)
        registry_header.setFont(1, font)
        parent.sequence_tree.addTopLevelItem(registry_header)
        
        # Add items for each persistence registry key
        for reg_op in registry_operations:
            if reg_op["IsPersistence"]:
                root_name = {
                    "0": "HKCR", "1": "HKCU", "2": "HKLM", "3": "HKU", "4": "HKCC"
                }.get(reg_op["Root"], f"Unknown ({reg_op['Root']})")
                
                reg_item = QTreeWidgetItem(["", f"Registry: {root_name}\\{reg_op['Key']}", reg_op["Name"], "", reg_op["Value"]])
                reg_item.setForeground(4, QColor("red"))
                
                # Red background for persistence items
                light_red = QColor(255, 240, 240)
                for i in range(5):
                    reg_item.setBackground(i, light_red)
                
                parent.sequence_tree.addTopLevelItem(reg_item)
    
    # Add service installations to the tree if present
    system_services = 0
    for service in service_installs:
        if service["IsCritical"]:
            system_services += 1
            
    if system_services > 0:
        # Add a spacer
        spacer = QTreeWidgetItem(["", "", "", "", ""])
        parent.sequence_tree.addTopLevelItem(spacer)
        
        # Add a header
        services_header = QTreeWidgetItem(["", "⚙️ SYSTEM SERVICES", "", "", ""])
        services_header.setBackground(0, QColor(240, 240, 255))  # Light purple
        services_header.setBackground(1, QColor(240, 240, 255))
        services_header.setBackground(2, QColor(240, 240, 255))
        services_header.setBackground(3, QColor(240, 240, 255))
        services_header.setBackground(4, QColor(240, 240, 255))
        font = services_header.font(0)
        font.setBold(True)
        services_header.setFont(1, font)
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
    """Evaluate the impact of a custom action based on its type"""
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
    """Evaluate the impact of a standard action based on its name"""
    # High impact actions - genuinely risky operations that modify system state significantly
    high_impact_actions = {
        "RemoveExistingProducts": "Removes existing products",
        "RegisterServices": "Installs services",
        "StartServices": "Starts services",
        "DeleteServices": "Removes services from system"
    }
    
    # Medium-high impact actions - significant system changes but standard during installations
    medium_high_actions = {
        "InstallFiles": "Copies files to system",
        "WriteRegistryValues": "Modifies registry",
        "RemoveRegistryValues": "Removes registry entries",
        "RemoveFiles": "Removes files",
        "MoveFiles": "Moves files within the system",
        "PatchFiles": "Applies patches to files",
        "SelfRegModules": "Self-registers modules"
    }
    
    # Medium impact actions - system modifications within expected ranges
    medium_impact_actions = {
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
    
    # Low-medium impact actions - minor system changes
    low_medium_actions = {
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
    
    # Low impact actions - minimal system changes
    low_impact_actions = {
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
    
    # Transaction control actions - fundamental installation operations
    transaction_actions = {
        "InstallInitialize": "Begins installation transaction",
        "InstallFinalize": "Commits changes to system",
        "InstallExecute": "Executes installation script",
        "StopServices": "Stops services during update"
    }
    
    # No impact actions (information only)
    no_impact_actions = {
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
    
    # Critical impact (special attention)
    critical_impact_actions = {
        "ExecuteAction": "Triggers Execute Sequence"
    }
    
    # Special handling for known patterns of custom actions
    # Use more precise patterns to avoid false positives
    if action_name.startswith("_") and len(action_name) > 30 and any(c.isdigit() for c in action_name):
        # More specifically identify GUID-like patterns
        if action_name.count("_") > 3 or action_name.count("-") > 0:
            return "Custom action with GUID identifier", "MEDIUM"
        return "Custom action with generated name", "MEDIUM"
    elif action_name.startswith("DIRCA_"):
        # Directory customization actions are generally for directory settings
        return "Directory customization action", "LOW"
    elif action_name.startswith("ERRCA_"):
        # Error handling actions are informational
        return "Error handling action", "LOW"
    elif action_name.startswith("AI_"):
        # Analyze the AI_ action more specifically
        if "DETECT" in action_name or "CHECK" in action_name:
            return "Advanced Installer detection action", "LOW"
        elif "SET" in action_name or "WRITE" in action_name:
            return "Advanced Installer property setting", "MEDIUM"
        elif "INSTALL" in action_name or "EXECUTE" in action_name:
            return "Advanced Installer execution action", "MEDIUM"
        return "Advanced Installer custom action", "MEDIUM"
    
    # Return appropriate impact level based on action category
    if action_name in high_impact_actions:
        return high_impact_actions[action_name], "HIGH"
    elif action_name in medium_high_actions:
        return medium_high_actions[action_name], "MEDIUM-HIGH"
    elif action_name in medium_impact_actions:
        return medium_impact_actions[action_name], "MEDIUM"
    elif action_name in low_medium_actions:
        return low_medium_actions[action_name], "LOW-MEDIUM"
    elif action_name in low_impact_actions:
        return low_impact_actions[action_name], "LOW"
    elif action_name in transaction_actions:
        return transaction_actions[action_name], "MEDIUM"
    elif action_name in no_impact_actions:
        return no_impact_actions[action_name], "NONE"
    elif action_name in critical_impact_actions:
        return critical_impact_actions[action_name], "CRITICAL"
    else:
        # For unknown actions, assume a safer default
        return "Standard MSI action", "LOW" 