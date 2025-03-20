"""
Help tab functionality for the MSI Parser GUI
"""

import os
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QLabel, QTextBrowser
import markdown

def create_help_tab():
    """Create the help tab"""
    help_tab = QWidget()
    help_layout = QVBoxLayout()
    help_tab.setLayout(help_layout)
    
    # Add title and introduction
    help_title = QLabel("MSI Parser Help & Documentation")
    help_title.setStyleSheet("font-size: 16px; font-weight: bold;")
    help_layout.addWidget(help_title)
    
    help_intro = QLabel("This section provides comprehensive documentation about MSI files and their analysis. The information below will help you understand how MSI files work and how to interpret the results displayed in other tabs.")
    help_intro.setWordWrap(True)
    help_layout.addWidget(help_intro)
    
    # Add a small vertical spacer
    help_layout.addSpacing(10)
    
    # Create text browser for help content
    help_html = QTextBrowser()
    help_html.setOpenExternalLinks(True)
    help_layout.addWidget(help_html, 1)
    
    # Read and convert the markdown file
    try:
        # Get the absolute path to the markdown file
        current_dir = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        md_file_path = os.path.join(current_dir, "MSI_Installation_Workflow_Analysis.md")
        
        # Read the markdown file
        with open(md_file_path, 'r', encoding='utf-8') as f:
            md_content = f.read()
        
        # Convert markdown to HTML
        html_content = markdown.markdown(md_content)
        
        # Add some basic CSS styling
        styled_html = f"""
        <style>
            body {{ font-family: Arial, sans-serif; line-height: 1.6; padding: 20px; }}
            h1 {{ color: #2c3e50; border-bottom: 2px solid #eee; padding-bottom: 10px; }}
            h2 {{ color: #34495e; margin-top: 20px; }}
            p {{ margin: 10px 0; }}
            ul {{ margin: 10px 0; padding-left: 20px; }}
            li {{ margin: 5px 0; }}
            code {{ background-color: #f8f9fa; padding: 2px 5px; border-radius: 3px; }}
            pre {{ background-color: #f8f9fa; padding: 15px; border-radius: 5px; overflow-x: auto; }}
        </style>
        {html_content}
        """
        
        help_html.setHtml(styled_html)
        
    except Exception as e:
        # Fallback content if file reading fails
        error_content = f"""
        <div style="color: red; padding: 20px;">
            <h2>Error Loading Documentation</h2>
            <p>Failed to load the MSI Installation Workflow Analysis documentation:</p>
            <pre>{str(e)}</pre>
            <p>Please ensure the documentation file exists and is accessible.</p>
        </div>
        """
        help_html.setHtml(error_content)
    
    return help_tab 