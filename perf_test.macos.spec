# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        # In CI we generate config.ini from config.example.ini; keep it in the bundle.
        ('config.ini', '.'),
        ('lighting.ico', '.'),
        ('excel_template', 'excel_template'),
    ],
    hiddenimports=[
        'openpyxl.cell._writer',
        'aiohttp',
        'paramiko',
        'psycopg2',
        'pandas',
        'tkinter',
        'tkinter.ttk',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'configparser',
        'ctypes',
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
    [],
    exclude_binaries=True,
    name='perf_test',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # macOS runners typically require .icns for app icon; omit to keep the build portable.
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='perf_test',
)

app = BUNDLE(
    coll,
    name='perf_test.app',
    icon=None,
    bundle_identifier=None,
)
