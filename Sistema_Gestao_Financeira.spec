# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['src\\sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=[('logo.png', '.'), ('src', 'src')],
    hiddenimports=['tkinter', 'tkinter.ttk', 'PIL', 'pandas', 'openpyxl', 'reportlab', 'matplotlib', 'xlwings', 'Sistema_Entrada_Dados', 'src.relatorios_interface', 'src.relatorio_despesas_aprimorado', 'src.relatorio_despesas_service', 'src.despesas_rateadas', 'src.gestao_medicoes', 'src.controle_pagamentos_taxas', 'src.controle_pagamentos', 'src.relatorio_tipo_despesa', 'src.verificador_sistema', 'src.gestao_taxas', 'src.pagamentos_eventos', 'src.relatorio_categoria', 'src.relatorio_fornecedores', 'src.relatorio_contratos_medicoes', 'src.corrigir_imports_sistema', 'src.finalizacao_quinzena', 'src.correcao_monetaria', 'src.configuracoes_sistema', 'src.teste_certificado_automatico', 'src.version_control'],
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
    icon=['logo1.ico'],
)
