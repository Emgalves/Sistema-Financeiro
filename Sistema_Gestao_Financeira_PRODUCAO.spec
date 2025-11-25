# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.png', '.'), ('src', 'src'), ('config', 'config')],
    hiddenimports=['tkinter', 'tkinter.ttk', 'tkinter.scrolledtext', 'tkinter.filedialog', 'tkinter.messagebox', 'PIL', 'PIL.Image', 'pandas', 'openpyxl', 'reportlab', 'matplotlib', 'xlwings', 'tkcalendar', 'babel', 'dotenv', 'ambiente_config', 'version_control', 'Sistema_Entrada_Dados', 'src.ambiente_config', 'src.version_control', 'src.Sistema_Entrada_Dados', 'src.relatorios_interface', 'src.relatorio_despesas_aprimorado', 'src.relatorio_despesas_service', 'src.relatorio_tipo_despesa', 'src.relatorio_categoria', 'src.relatorio_fornecedores', 'src.relatorio_contratos_medicoes', 'src.despesas_rateadas', 'src.gestao_medicoes', 'src.gestao_taxas', 'src.configuracoes_sistema', 'src.controle_pagamentos_taxas', 'src.controle_pagamentos', 'src.pagamentos_eventos', 'src.verificador_sistema', 'src.corrigir_imports_sistema', 'src.finalizacao_quinzena', 'src.correcao_monetaria', 'src.teste_certificado_automatico', 'src.config.utils', 'src.config.dialogs', 'src.config.logger_config', 'src.config.window_config', 'src.config.config'],
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
    name='Sistema_Gestao_Financeira_PRODUCAO',
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
    icon=['logo1.ico'],
)
