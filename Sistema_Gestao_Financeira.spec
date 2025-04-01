# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Adicionar o diretório src ao PYTHONPATH
src_path = os.path.abspath('src')
sys.path.insert(0, src_path)

block_cipher = None

a = Analysis(
    ['src/sistema_principal.py'],
    pathex=[
        'src',
        os.path.abspath('src'),
        os.path.abspath('src/config'),
        os.path.dirname(os.path.abspath('src/sistema_principal.py'))
    ],
    binaries=[],
    datas=[
        ('logo.png', '.'),
        ('logo1.png', '.'),
        # Arquivos de configuração
        ('src/config/*.py', 'src/config'),
        ('src/config/parametros_sistema.json', 'src/config'),
        # Módulos principais
        ('src/gestao_taxas.py', 'src'),
        ('src/Sistema_Entrada_Dados.py', 'src'),
        ('src/relatorio_despesas_aprimorado.py', 'src'),
        ('src/controle_pagamentos.py', 'src'),
        ('src/finalizacao_quinzena.py', 'src'),
        ('src/configuracoes_sistema.py', 'src'),
        ('testes/Financeiro/Planilhas_Base/*.*', 'testes/Financeiro/Planilhas_Base'),
    ],
    
    hiddenimports=[
        'babel.numbers',
        'validate_docbr',
        'tkcalendar',
        'dateutil.relativedelta',
        'numpy',
        'numpy.core._dtype_ctypes',
        'pandas',
        'src.config',
        'src.config.utils',
        'config',
        'config.utils',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'openpyxl',
        'babel',
        'src.Sistema_Entrada_Dados',
        'src.configuracoes_sistema',
        'src.controle_pagamentos',
        'src.finalizacao_quinzena',
        'src.gestao_taxas',
        'src.relatorio_despesas_aprimorado',
        'src.version_control',
        'Sistema_Entrada_Dados',
        'configuracoes_sistema',
        'controle_pagamentos',
        'finalizacao_quinzena',
        'gestao_taxas',
        'relatorio_despesas_aprimorado',
        'version_control',
        'xlwings',  # Adicionado
        'xlwings.main',  # Adicionado
    ] + collect_submodules('numpy') + collect_submodules('pandas'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Sistema_Gestao_Financeira',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Alterado de True para False
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)