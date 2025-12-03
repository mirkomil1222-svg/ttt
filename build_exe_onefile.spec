# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for building Rash Manager v2 as SINGLE-FILE Windows executable
This creates one exe file that bundles everything (slower startup, but easier distribution)
"""

block_cipher = None

a = Analysis(
    ['rash_manager_v2.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include data files - they will be extracted at runtime
        ('titul_bubble_koordinatalar_2480x3508.xlsx', '.'),
        ('Titul.pdf', '.'),
    ],
    hiddenimports=[
        'pandas',
        'numpy',
        'PIL',
        'cv2',
        'qrcode',
        'pyzbar',
        'scipy',
        'scipy.optimize',
        'openpyxl',
        'openpyxl.chart',
        'openpyxl.styles',
        'pdf2image',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.simpledialog',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'IPython',
        'jupyter',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='RashManager',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Compress executable (may slow down startup slightly)
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # GUI application - no console window
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon file path here if you have one
    onefile=True,  # Create single-file executable
)

