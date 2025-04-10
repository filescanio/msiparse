"""
Syntax highlighting for code preview in the MSI Parser GUI
"""

from PyQt5.QtCore import QRegExp, Qt
from PyQt5.QtGui import QTextCharFormat, QFont, QColor, QSyntaxHighlighter

class CodeSyntaxHighlighter(QSyntaxHighlighter):
    """Syntax highlighter for code preview"""
    def __init__(self, document, language="generic"):
        super().__init__(document)
        self.language = language.lower()
        
        # Create formatting for different syntax elements
        self.formats = {
            'keyword': self._create_format(QColor("#569CD6"), bold=True),  # Blue
            'string': self._create_format(QColor("#CE9178")),              # Brown
            'comment': self._create_format(QColor("#6A9955"), italic=True), # Green
            'function': self._create_format(QColor("#DCDCAA")),            # Yellow
            'number': self._create_format(QColor("#B5CEA8")),              # Light green
            'operator': self._create_format(QColor("#D4D4D4")),            # Light grey
            'variable': self._create_format(QColor("#9CDCFE")),            # Light blue
        }
        
        # Set up highlighting rules based on language
        self.highlighting_rules = []
        self._setup_rules()
    
    def _create_format(self, color, bold=False, italic=False):
        """Create a text format with specified properties"""
        fmt = QTextCharFormat()
        fmt.setForeground(color)
        if bold:
            fmt.setFontWeight(QFont.Bold)
        if italic:
            fmt.setFontItalic(True)
        return fmt
    
    def _setup_rules(self):
        """Set up syntax highlighting rules based on language"""
        self.highlighting_rules = []
        
        # Add generic rules (useful for most languages)
        # Numbers
        self.highlighting_rules.append((QRegExp(r'\b[0-9]+\b'), self.formats['number']))
        
        # String literals
        self.highlighting_rules.append((QRegExp(r'"[^"\\]*(\\.[^"\\]*)*"'), self.formats['string']))
        self.highlighting_rules.append((QRegExp(r"'[^'\\]*(\\.[^'\\]*)*'"), self.formats['string']))
        
        # Single-line comments
        self.highlighting_rules.append((QRegExp(r'//.*'), self.formats['comment']))
        self.highlighting_rules.append((QRegExp(r'#.*'), self.formats['comment']))
        
        # Add language-specific rules
        if self.language == "python":
            # Python keywords
            keywords = [
                'and', 'as', 'assert', 'break', 'class', 'continue', 'def',
                'del', 'elif', 'else', 'except', 'False', 'finally', 'for',
                'from', 'global', 'if', 'import', 'in', 'is', 'lambda', 'None',
                'nonlocal', 'not', 'or', 'pass', 'raise', 'return', 'True',
                'try', 'while', 'with', 'yield'
            ]
            self.highlighting_rules.append((
                QRegExp('\\b(' + '|'.join(keywords) + ')\\b'), 
                self.formats['keyword']
            ))
            
            # Python function definitions
            self.highlighting_rules.append((
                QRegExp(r'\bdef\s+([A-Za-z_][A-Za-z0-9_]*)\b'), 
                self.formats['function']
            ))
            
        elif self.language == "javascript" or self.language == "js":
            # JavaScript keywords
            keywords = [
                'break', 'case', 'catch', 'class', 'const', 'continue',
                'debugger', 'default', 'delete', 'do', 'else', 'export',
                'extends', 'false', 'finally', 'for', 'function', 'if',
                'import', 'in', 'instanceof', 'new', 'null', 'return',
                'super', 'switch', 'this', 'throw', 'true', 'try', 'typeof',
                'var', 'void', 'while', 'with', 'yield', 'let', 'static',
                'async', 'await'
            ]
            self.highlighting_rules.append((
                QRegExp('\\b(' + '|'.join(keywords) + ')\\b'), 
                self.formats['keyword']
            ))
            
            # JavaScript function definitions
            self.highlighting_rules.append((
                QRegExp(r'\bfunction\s+([A-Za-z_][A-Za-z0-9_]*)\b'), 
                self.formats['function']
            ))
            
        elif self.language == "vbscript" or self.language == "vbs":
            # VBScript keywords (case-insensitive)
            keywords = [
                'And', 'As', 'Boolean', 'ByRef', 'Byte', 'ByVal', 'Call',
                'Case', 'Class', 'Const', 'Currency', 'Debug', 'Dim', 'Do',
                'Double', 'Each', 'Else', 'ElseIf', 'Empty', 'End', 'Enum',
                'Eqv', 'Event', 'Exit', 'False', 'For', 'Function', 'Get',
                'GoTo', 'If', 'Imp', 'Implements', 'In', 'Integer', 'Is',
                'Let', 'Like', 'Long', 'Loop', 'LSet', 'Me', 'Mod', 'New',
                'Next', 'Not', 'Nothing', 'Null', 'On', 'Option', 'Or',
                'ParamArray', 'Preserve', 'Private', 'Public', 'ReDim',
                'Rem', 'Resume', 'RSet', 'Select', 'Set', 'Shared', 'Single',
                'Static', 'Stop', 'Sub', 'Then', 'To', 'True', 'Type',
                'TypeOf', 'Until', 'Variant', 'Wend', 'While', 'With', 'Xor'
            ]
            # Case-insensitive regex for VBScript
            self.highlighting_rules.append((
                QRegExp('\\b(' + '|'.join(keywords) + ')\\b', Qt.CaseInsensitive), 
                self.formats['keyword']
            ))
            
            # VBScript comment (single quote)
            self.highlighting_rules.append((
                QRegExp(r"'.*"), 
                self.formats['comment']
            ))
            
        # Add PowerShell highlighting if needed
        elif self.language == "powershell" or self.language == "ps1":
            # PowerShell keywords
            keywords = [
                'Begin', 'Break', 'Catch', 'Class', 'Continue', 'Data',
                'Define', 'Do', 'DynamicParam', 'Else', 'ElseIf', 'End',
                'Exit', 'Filter', 'Finally', 'For', 'ForEach', 'From',
                'Function', 'If', 'In', 'Param', 'Process', 'Return',
                'Switch', 'Throw', 'Trap', 'Try', 'Until', 'Using',
                'Var', 'While', 'Workflow'
            ]
            self.highlighting_rules.append((
                QRegExp('\\b(' + '|'.join(keywords) + ')\\b', Qt.CaseInsensitive), 
                self.formats['keyword']
            ))
            
            # PowerShell comment
            self.highlighting_rules.append((
                QRegExp(r'#.*'), 
                self.formats['comment']
            ))
        
        elif self.language == "xml" or self.language == "html":
            # XML tags
            self.highlighting_rules.append((
                QRegExp(r'<\/?[a-zA-Z0-9_:-]+\s*[^>]*>', Qt.CaseInsensitive),
                self.formats['keyword']
            ))
            
            # XML attributes
            self.highlighting_rules.append((
                QRegExp(r'\s+[a-zA-Z0-9_:-]+=', Qt.CaseInsensitive),
                self.formats['variable']
            ))
            
            # XML comments
            self.highlighting_rules.append((
                QRegExp(r'<!--.*-->'),
                self.formats['comment']
            ))
            
        elif self.language == "batch" or self.language == "bat":
            # Batch file keywords
            keywords = [
                'echo', 'set', 'if', 'else', 'for', 'in', 'do', 'call', 'goto',
                'exit', 'rem', 'pause', 'choice', 'cls', 'color', 'copy', 'del',
                'dir', 'mkdir', 'rmdir', 'type', 'cd', 'findstr', 'errorlevel',
                'not', 'exist', 'defined', 'equ', 'neq', 'lss', 'leq', 'gtr', 'geq'
            ]
            self.highlighting_rules.append((
                QRegExp('\\b(' + '|'.join(keywords) + ')\\b', Qt.CaseInsensitive), 
                self.formats['keyword']
            ))
            
            # Batch file variables
            self.highlighting_rules.append((
                QRegExp(r'%[^%]+%'), 
                self.formats['variable']
            ))
            
            # Batch file labels
            self.highlighting_rules.append((
                QRegExp(r':[A-Za-z0-9_]+'), 
                self.formats['function']
            ))
            
            # Batch file comments (REM and ::)
            self.highlighting_rules.append((
                QRegExp(r'(^|\s)rem\s+.*$', Qt.CaseInsensitive), 
                self.formats['comment']
            ))
            self.highlighting_rules.append((
                QRegExp(r'::.*$'), 
                self.formats['comment']
            ))

    def highlightBlock(self, text):
        """Apply highlighting to a block of text"""
        for pattern, format in self.highlighting_rules:
            expression = QRegExp(pattern)
            index = expression.indexIn(text)
            while index >= 0:
                length = expression.matchedLength()
                self.setFormat(index, length, format)
                index = expression.indexIn(text, index + length)
        
        # Handle multi-line comments if needed
        self.setCurrentBlockState(0)


def detect_language(content, mime_type=None):
    """Attempt to detect the language based on content and optional Magika MIME type
    
    Args:
        content: The text content to analyze
        mime_type: Optional MIME type from Magika detection
        
    Returns:
        String representing the detected language
    """
    # First check if we have a Magika MIME type that maps to a supported language
    if mime_type:
        # Handle text/plain with specific extensions
        if "." in mime_type:
            extension = mime_type.split(".")[-1].lower()
            if extension == "py":
                return "python"
            elif extension in ["js", "json"]:
                return "javascript"
            elif extension in ["vbs", "vbe"]:
                return "vbscript"
            elif extension in ["ps1", "psm1", "psd1"]:
                return "powershell"
            elif extension in ["xml", "xsl", "xsd", "svg"]:
                return "xml"
            elif extension in ["htm", "html", "xhtml"]:
                return "html"
            elif extension in ["bat", "cmd"]:
                return "batch"
        
        # Check direct MIME type mappings
        if mime_type.startswith("text/"):
            if "javascript" in mime_type:
                return "javascript"
            elif "python" in mime_type:
                return "python"
            elif "xml" in mime_type:
                return "xml"
            elif "html" in mime_type:
                return "html"
            elif "vbscript" in mime_type:
                return "vbscript"
            elif "powershell" in mime_type:
                return "powershell"
    
    # Fallback to content-based detection
    if content.strip().startswith('<?xml') or '</' in content:
        return "xml"
    elif 'function' in content and ('var ' in content or 'let ' in content or 'const ' in content):
        return "javascript"
    elif 'def ' in content and ':' in content:
        return "python"
    elif "Sub " in content or "Function " in content or "Dim " in content:
        return "vbscript"
    elif "$" in content and ("-eq" in content or "Write-Host" in content):
        return "powershell"
    elif "<!--" in content or "<html" in content:
        return "html"
    elif "@echo" in content.lower() or "goto" in content.lower() or "set " in content.lower():
        return "batch"
    elif ".exe" in content.lower() or ".dll" in content.lower():
        # This might be a binary file mistakenly opened as text
        return "generic"
    else:
        return "generic" 