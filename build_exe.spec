# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for building Rash Manager v2 Windows executable
"""

block_cipher = None

a = Analysis(
    ['rash_manager_v2.py'],
    pathex=[],
    binaries=[],
    datas=[
        # Include data files if they exist
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
    excludes=[],
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
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to False for GUI application (no console window)
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Add icon file path here if you have one
)

