# -*- mode: python ; coding: utf-8 -*-


block_cipher = None


a = Analysis(
    ['sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=[('src', 'src'), ('recursos', 'recursos'), ('src/config', 'src/config'), ('src/config/parametros_sistema.json', 'src/config')],
    hiddenimports=['src.config', 'src.config.utils', 'src.config.logger_config', 'src.config.window_config', 'src.config.config', 'src.config.dialogs', 'src.configuracoes_sistema', 'tkinter', 'tkinter.ttk', 'tkcalendar', 'pandas', 'openpyxl', 'xlwings', 'babel', 'babel.numbers', 'dateutil.relativedelta', 'reportlab', 'reportlab.lib', 'reportlab.lib.pagesizes', 'reportlab.pdfgen', 'reportlab.pdfgen.canvas', 'reportlab.lib.styles', 'reportlab.platypus', 'reportlab.lib.enums', 'reportlab.lib.colors', 'validate_docbr'],
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
    name='SistemaEntradaDados_Debug',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
