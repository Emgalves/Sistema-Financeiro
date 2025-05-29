# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.png', '.'), ('src', 'src')],
    hiddenimports=['tkinter', 'tkinter.ttk', 'PIL', 'pandas', 'openpyxl', 'reportlab', 'matplotlib', 'xlwings', 'Sistema_Entrada_Dados', 'src.Sistema_Entrada_Dados', 'relatorios_interface', 'src.relatorios_interface', 'relatorio_despesas_aprimorado', 'src.relatorio_despesas_aprimorado', 'despesas_rateadas', 'src.despesas_rateadas', 'gestao_medicoes', 'src.gestao_medicoes', 'controle_pagamentos_taxas', 'src.controle_pagamentos_taxas', 'configuracoes_sistema', 'src.configuracoes_sistema', 'version_control', 'src.version_control'],
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
