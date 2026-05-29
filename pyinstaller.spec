# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for SheinExtractAU — produces dist/SheinExtractAU.exe.

Build with:
    pyinstaller pyinstaller.spec --clean --noconfirm

Outputs a single console-mode .exe (employees see a console with progress).
"""

block_cipher = None

a = Analysis(
    ['app_main.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        # openpyxl pulls these dynamically
        'openpyxl.styles.alignment',
        'openpyxl.styles.borders',
        'openpyxl.styles.fills',
        'openpyxl.styles.fonts',
        'openpyxl.drawing.image',
        # PIL via openpyxl image support
        'PIL',
        'PIL.Image',
        # Tkinter (wizard + update dialog)
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        # websocket-client (used by Chrome CDP code in shein_scraper)
        'websocket._abnf',
        'websocket._app',
        'websocket._core',
        'websocket._exceptions',
        'websocket._handshake',
        'websocket._http',
        'websocket._logging',
        'websocket._socket',
        'websocket._ssl_compat',
        'websocket._url',
        'websocket._utils',
        # our own modules
        'config',
        'shein_scraper',
        'run_excel',
        'notify',
        'setup_wizard',
        'update_check',
        'auth',
        'version',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
        'sphinx',
        'pytest',
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
    name='SheinExtractAU',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    # icon='resources\\app.ico',  # add later if a .ico is provided
)
