"""
Constants used by the Installation Impact tab functionality
"""

# Registry root mappings
REGISTRY_ROOTS = {
    "-1": "HKCU",
    "0": "HKCR",
    "1": "HKCU",
    "2": "HKLM",
    "3": "HKU"
}

# Service type mappings
SERVICE_TYPES = {
    "0x00000001": "Kernel Driver",
    "0x00000002": "File System Driver",
    "0x00000010": "Win32 Own Process",
    "0x00000020": "Win32 Share Process",
    "0x00000100": "Interactive",
    "0x00000110": "Interactive, Win32 Own Process",
    "0x00000120": "Interactive, Win32 Share Process"
}

# Service start type mappings
SERVICE_START_TYPES = {
    "0x00000000": "Boot Start",
    "0x00000001": "System Start",
    "0x00000002": "Auto Start",
    "0x00000003": "Demand Start",
    "0x00000004": "Disabled"
}

# Registry patterns for autorun locations
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

# File patterns that are suspicious when modified
SUSPICIOUS_FILE_PATTERNS = [
    r".*\\windows\\system32\\.*\.(exe|dll|sys|drv)$",
    r".*\\windows\\syswow64\\.*\.(exe|dll|sys|drv)$",
    r".*\\appdata\\.*\\microsoft\\windows\\start menu\\programs\\startup\\.*",
    r".*\\programdata\\.*\\microsoft\\windows\\start menu\\programs\\startup\\.*",
    r".*\\program files\\.*\\common files\\.*\.(exe|dll|sys|drv)$",
    r".*\\program files \(x86\)\\common files\\.*\.(exe|dll|sys|drv)$"
]

# Registry patterns that are suspicious when modified
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

# Critical system directories
CRITICAL_DIRECTORIES = {
    "SystemFolder": "System directory modification",
    "System64Folder": "System directory modification",
    "WindowsFolder": "Windows directory modification",
    "StartupFolder": "Startup folder modification",
    "CommonFilesFolder": "Common Files directory modification",
    "CommonFiles64Folder": "Common Files (x64) directory modification"
}

# High-risk file extensions
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

# Medium-risk file extensions
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

# Example MSI directory paths
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