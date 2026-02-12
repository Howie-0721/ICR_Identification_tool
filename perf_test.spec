# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
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
    upx=True,
    console=False,  # 設為 False 以隱藏命令提示字元視窗
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='lighting.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='perf_test',
)
