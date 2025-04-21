# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path

# Adicionar o diretório src ao PYTHONPATH
src_path = os.path.abspath('src')
sys.path.insert(0, src_path)

block_cipher = None

a = Analysis(
    ['src/sistema_principal.py'],
    pathex=[src_path],  # Adicionado src_path aqui
    binaries=[],
    datas=[
        # Arquivos de configuração
        ('src/config/utils.py', 'src/config'),
        ('src/config/config.py', 'src/config'),
        ('src/config/logger_config.py', 'src/config'),
        ('src/config/window_config.py', 'src/config'),
        ('src/config/parametros_sistema.json', 'src/config'),
        # Logo
        ('src/logo.png', 'src'),
        ('src/logo1.png', 'src'),
        # Módulos principais
        ('src/gestao_taxas.py', 'src'),
        ('src/Sistema_Entrada_Dados.py', 'src'),
        ('src/relatorio_despesas_aprimorado.py', 'src'),
        ('src/controle_pagamentos.py', 'src'),
        ('src/finalizacao_quinzena.py', 'src'),
        ('src/configuracoes_sistema.py', 'src'),
    ],
    hiddenimports=[
        'babel.numbers',
        'validate_docbr',
        'tkcalendar',
        'dateutil.relativedelta',
        'config',  # Adicionado o módulo config
        'config.utils',
        'config.config',
        'config.logger_config',
        'config.window_config',
    ],
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
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None
)