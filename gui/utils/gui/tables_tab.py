"""
Tables tab functionality for the MSI Parser GUI
"""

import json
from PyQt5.QtWidgets import (QTableWidgetItem, QListWidgetItem, QWidget, QHBoxLayout, 
                            QLabel, QToolButton, QMessageBox, QFileDialog, QApplication)
from PyQt5.QtGui import QFont

def list_tables(parent):
    """List MSI tables"""
    if not parent.msi_file_path:
        return
        
    command = [parent.msiparse_path, "list_tables", parent.msi_file_path]
    output = parent.run_command_safe(command)
    if output:
        display_tables(parent, output)
        
def display_tables(parent, output):
    """Display tables in the tables tab"""
    try:
        parent.tables_data = json.loads(output)
        
        # Clear any existing filter
        parent.table_filter.clear()
        
        # Clear the table list and content
        parent.table_list.clear()
        parent.table_content.clear()
        parent.table_content.setRowCount(0)
        parent.table_content.setColumnCount(0)
        
        # Count empty tables for checkbox label
        empty_tables_count = sum(1 for table in parent.tables_data if len(table.get("rows", [])) == 0)
        parent.hide_empty_tables_checkbox.setText(f"Hide Empty Tables ({empty_tables_count})")
        
        # Populate the table list
        for table in parent.tables_data:
            table_name = table["name"]
            item = QListWidgetItem(table_name)
            
            # Add tooltip if table is in our description map
            if table_name in parent.msi_tables:
                item.setToolTip(parent.msi_tables[table_name])
                
                # Create a custom widget with text and info button
                widget = QWidget()
                layout = QHBoxLayout(widget)
                layout.setContentsMargins(4, 0, 4, 0)
                layout.setSpacing(2)
                
                # Add table name label
                label = QLabel(table_name)
                
                # Remove monospaced font for the label
                # mono_font = QFont("Courier New", 10)
                # mono_font.setFixedPitch(True)
                # label.setFont(mono_font)
                
                layout.addWidget(label)
                
                # Add spacer to push the info button to the right
                layout.addStretch()
                
                # Add info button
                info_button = QToolButton()
                info_button.setIcon(QApplication.style().standardIcon(QApplication.style().SP_MessageBoxInformation))
                info_button.setToolTip(parent.msi_tables[table_name])
                info_button.setFixedSize(16, 16)
                info_button.clicked.connect(lambda checked, name=table_name: show_table_info(parent, name))
                layout.addWidget(info_button)
                
                # Set the custom widget for this item
                parent.table_list.addItem(item)
                parent.table_list.setItemWidget(item, widget)
            else:
                # Just add the regular item without an info button
                parent.table_list.addItem(item)
            
        parent.statusBar().showMessage(f"Found {len(parent.tables_data)} tables")
        
        # Update button states
        parent.update_button_states()
        
        # Apply hide empty tables filter if checked
        filter_tables(parent)
        
    except json.JSONDecodeError:
        parent.handle_error("Parse Error", "Error parsing tables output", show_dialog=True)
        
def table_selected(parent, current, previous):
    """Handle table selection from the list"""
    if not current or not parent.tables_data:
        return
        
    table_name = current.text()
    
    # Find the selected table in the data
    selected_table = None
    for table in parent.tables_data:
        if table["name"] == table_name:
            selected_table = table
            break
            
    if not selected_table:
        return
        
    # Set up the table content view
    columns = selected_table["columns"]
    rows = selected_table["rows"]
    
    parent.table_content.clear()
    parent.table_content.setRowCount(len(rows))
    parent.table_content.setColumnCount(len(columns))
    parent.table_content.setHorizontalHeaderLabels(columns)
    
    # Create monospaced font for hash columns
    mono_font = QFont("Courier New", 10)
    mono_font.setFixedPitch(True)
    
    # Fill the table with data
    for row_idx, row_data in enumerate(rows):
        for col_idx, cell_data in enumerate(row_data):
            if col_idx < len(columns):  # Safety check
                item = QTableWidgetItem(cell_data)
                
                # Apply monospaced font only to hash columns
                # Check if column name contains "hash" or if the data looks like a hash
                column_name = columns[col_idx].lower()
                if ("hash" in column_name or 
                    (len(cell_data) >= 32 and all(c in "0123456789abcdefABCDEF" for c in cell_data))):
                    item.setFont(mono_font)
                    
                parent.table_content.setItem(row_idx, col_idx, item)
    
    parent.statusBar().showMessage(f"Showing table: {table_name} ({len(rows)} rows)")
    
    # Update button states
    parent.update_button_states()
    
def show_table_info(parent, table_name):
    """Show a message box with information about the selected table"""
    if table_name in parent.msi_tables:
        QMessageBox.information(
            parent,
            f"Table Information: {table_name}",
            parent.msi_tables[table_name]
        )
        
def export_selected_table(parent):
    """Export the currently selected table as JSON"""
    if not parent.tables_data or not parent.table_list.currentItem():
        return
        
    table_name = parent.table_list.currentItem().text()
    selected_table = next(
        (table for table in parent.tables_data if table["name"] == table_name),
        None
    )
    
    if not selected_table:
        parent.show_warning("Warning", f"Table '{table_name}' not found in data")
        return
        
    file_path, _ = QFileDialog.getSaveFileName(
        parent, 
        f"Save {table_name} as JSON", 
        f"{table_name}.json", 
        "JSON Files (*.json);;All Files (*)"
    )
    
    if not file_path:
        return  # User cancelled
        
    try:
        with open(file_path, 'w') as f:
            json.dump(selected_table, f, indent=2)
            
        parent.show_status(f"Table '{table_name}' exported to {file_path}")
        QMessageBox.information(
            parent, 
            "Export Successful", 
            f"Table '{table_name}' has been exported to:\n{file_path}"
        )
    except Exception as e:
        parent.show_error("Export Failed", e)
        
def export_all_tables(parent):
    """Export all tables as a single JSON file (exact output of list_tables command)"""
    if not parent.tables_data:
        return
        
    # Prompt for save location
    file_path, _ = QFileDialog.getSaveFileName(
        parent, 
        "Save All Tables as JSON", 
        "all_tables.json", 
        "JSON Files (*.json);;All Files (*)"
    )
    
    if not file_path:
        return  # User cancelled
        
    try:
        # Write the complete tables data to the file
        with open(file_path, 'w') as f:
            json.dump(parent.tables_data, f, indent=2)
            
        parent.statusBar().showMessage(f"All tables exported to {file_path}")
        QMessageBox.information(
            parent, 
            "Export Successful", 
            f"All tables have been exported to:\n{file_path}"
        )
    except Exception as e:
        QMessageBox.critical(
            parent, 
            "Export Failed", 
            f"Failed to export tables: {str(e)}"
        )

def filter_tables(parent, filter_text=None):
    """Filter the tables list based on the input text and checkbox state"""
    # If called from checkbox, the filter_text will be an integer (checkbox state)
    # So we need to get the actual filter text from the line edit
    if isinstance(filter_text, int) or filter_text is None:
        filter_text = parent.table_filter.text()
    
    # Get checkbox state
    hide_empty = parent.hide_empty_tables_checkbox.isChecked()
    
    # Convert filter text to lowercase for case-insensitive matching
    filter_text_lower = filter_text.lower() if filter_text else ""
    
    # Count visible items and empty tables for status message
    visible_count = 0
    empty_tables_count = 0
    
    # First, count all empty tables
    for table in parent.tables_data:
        if len(table.get("rows", [])) == 0:
            empty_tables_count += 1
    
    # Update checkbox label with empty tables count
    parent.hide_empty_tables_checkbox.setText(f"Hide Empty Tables ({empty_tables_count})")
    
    # Check each item against the filters
    for i in range(parent.table_list.count()):
        item = parent.table_list.item(i)
        table_name = item.text()
        
        # Text filter match
        if filter_text_lower:
            # Check if the table name contains the filter text
            match_found = filter_text_lower in table_name.lower()
            
            # Also check in the table description if available
            if not match_found and table_name in parent.msi_tables:
                description = parent.msi_tables[table_name].lower()
                match_found = filter_text_lower in description
        else:
            match_found = True  # No text filter, so it matches
        
        # Check for empty tables if the checkbox is checked
        if hide_empty and match_found:
            # Find the table in the data
            table_data = next((table for table in parent.tables_data if table["name"] == table_name), None)
            if table_data and len(table_data.get("rows", [])) == 0:
                match_found = False  # Hide empty tables
        
        # Show or hide the item based on the combined match
        item.setHidden(not match_found)
        
        # Count visible items
        if match_found:
            visible_count += 1
    
    # Update status message with both filters
    if hide_empty:
        parent.statusBar().showMessage(f"Showing {visible_count} of {parent.table_list.count()} tables (hiding {empty_tables_count} empty tables)")
    else:
        parent.statusBar().showMessage(f"Showing {visible_count} of {parent.table_list.count()} tables") 