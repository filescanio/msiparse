# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_data_files

# Collect Magika models explicitly
magika_datas = collect_data_files('magika')

a = Analysis(
    ['__main__.py'],
    pathex=['.'],
    binaries=[],
    datas=[
        ('resources', 'resources'),
        ('help_dialog_template.md', '.'),
        # (magika_models_path, 'magika/models') # Removed explicit Magika models, relying on hooks
        ] + magika_datas, # Add Magika data files here
    hiddenimports=[], # Removed, relying on analysis
    hooksconfig={},
    runtime_hooks=['pyi_rth_onnxruntime.py'], # Keep the essential runtime hook
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

# Switched back to one-file EXE block
exe = EXE(
    pyz,
    a.scripts,
    a.binaries, # Include auto-detected binaries
    a.datas,    # Include auto-detected datas + resources
    [],
    name='msiparse-gui',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True, # Re-enabled UPX for smaller executable size
    runtime_tmpdir=None,
    console=False, # Windowed application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='resources/icon.ico' # Added application icon
)
