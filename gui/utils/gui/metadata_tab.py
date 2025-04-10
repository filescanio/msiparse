"""
Functions for retrieving and displaying MSI package metadata including properties like
ProductName, ProductVersion, Manufacturer, etc.
"""

import json

def get_metadata(parent):
    """Retrieve MSI package metadata using msiparse CLI tool"""
    if not parent.msi_file_path:
        return
        
    command = [parent.msiparse_path, "list_metadata", parent.msi_file_path]
    output = parent.run_command_safe(command)
    if output:
        display_metadata(parent, output)
        
def display_metadata(parent, output):
    """Display formatted MSI metadata in the metadata tab"""
    try:
        metadata = json.loads(output)
        formatted_text = "MSI Metadata:\n\n"
        
        for key, value in metadata.items():
            formatted_key = key.replace('_', ' ').title()
            if isinstance(value, list):
                value_str = ", ".join(value) if value else "None"
            else:
                value_str = str(value) if value else "None"
                
            formatted_text += f"{formatted_key}: {value_str}\n"
            
        parent.metadata_text.setText(formatted_text)
    except json.JSONDecodeError as e:
        parent.handle_error("Metadata Parse Error", 
                          f"Failed to parse metadata output: {str(e)}\nRaw output:\n{output}", 
                          show_dialog=True) 