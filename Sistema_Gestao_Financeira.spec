# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.png', '.'), ('src', 'src')],
    hiddenimports=['tkinter', 'tkinter.ttk', 'PIL', 'pandas', 'openpyxl', 'reportlab', 'matplotlib', 'xlwings', 'src.Sistema_Entrada_Dados', 'src.relatorios_interface', 'src.relatorio_despesas_aprimorado', 'src.despesas_rateadas', 'src.gestao_medicoes', 'src.controle_pagamentos_taxas', 'src.correcao_monetaria', 'src.configuracoes_sistema', 'src.version_control'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Sistema_Gestao_Financeira',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo.ico'],
)
