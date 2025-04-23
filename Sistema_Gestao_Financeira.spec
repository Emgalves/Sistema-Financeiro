# -*- mode: python ; coding: utf-8 -*-
import os
import sys
from pathlib import Path
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

# Usar caminhos absolutos e relativos sem depender de __file__
current_dir = os.path.abspath('.')
src_path = os.path.join(current_dir, 'src')
sys.path.insert(0, current_dir)  # Adiciona o diretório raiz
sys.path.insert(0, src_path)     # Adiciona o diretório src

# Definir uma função para obter a versão sem depender de importações
def get_version():
    try:
        # Tentar carregar o arquivo diretamente
        version_file = os.path.join(src_path, 'version_control.py')
        
        # Extrair a versão do arquivo usando execução segura
        version_info = {}
        with open(version_file, 'r', encoding='utf-8') as f:
            content = f.read()
            # Procurar pela definição de VERSION_INFO
            import re
            match = re.search(r'VERSION_INFO\s*=\s*{.*?"major":\s*(\d+),.*?"minor":\s*(\d+),.*?"patch":\s*(\d+)', 
                             content, re.DOTALL)
            if match:
                major = match.group(1)
                minor = match.group(2)
                patch = match.group(3)
                return f"{major}.{minor}.{patch}"
        
        # Se não conseguir extrair, retornar um valor padrão
        return "1.2.0"
    except Exception as e:
        # Em caso de erro, usar valor padrão
        print(f"Erro ao obter versão: {str(e)}")
        return "1.2.0"

# Obter a versão
version_string = get_version()
print(f"Versão detectada: {version_string}")

block_cipher = None

a = Analysis(
    ['src/sistema_principal.py'],
    pathex=[
        current_dir,
        src_path,
        os.path.join(src_path, 'config'),
        os.path.dirname(os.path.join(src_path, 'sistema_principal.py'))
    ],
    binaries=[],
    datas=[
        ('logo.png', '.'),
        ('logo1.png', '.'),
        # Arquivos de configuração
        ('src/config/*.py', 'src/config'),
        ('src/config/parametros_sistema.json', 'src/config'),
        # Módulos principais
        ('src/configuracoes_sistema.py', 'src'),
        ('src/controle_pagamentos.py', 'src'),
        ('src/despesas_rateadas.py', 'src'),
        ('src/finalizacao_quinzena.py', 'src'),
        ('src/gestao_medicoes.py', 'src'),
        ('src/pagamentos_eventos.py', 'src'),
        ('src/relatorio_contratos_medicoes.py', 'src'),
        ('src/relatorio_despesas_aprimorado.py', 'src'),
        ('src/relatorio_fornecedores.py', 'src'),
        ('src/relatorio_tipo_despesa.py', 'src'),
        ('src/relatorios_interface.py', 'src'),
        ('src/Sistema_Entrada_Dados.py', 'src'),
        # Adicionar explicitamente o arquivo de controle de versão
        ('src/version_control.py', 'src'),
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
        'src.controle_pagamentos_taxas',
        'src.despesas_rateadas',
        'src.finalizacao_quinzena',
        'src.gestao_medicoes',
        'src.pagamentos_eventos',
        'src.relatorio_contratos_medicoes',
        'src.relatorio_despesas_aprimorado',
        'src.relatorio_fornecedores',
        'src.relatorio_tipo_despesa',
        'src.relatorios_interface',
        'src.version_control',
        'Sistema_Entrada_Dados',
        'configuracoes_sistema',
        'controle_pagamentos',
        'controle_pagamentos_taxas',
        'despesas_rateadas',
        'finalizacao_quinzena',
        'gestao_medicoes',
        'pagamentos_eventos',
        'relatorio_contratos_medicoes',
        'relatorio_despesas_aprimorado',
        'relatorio_fornecedores',
        'relatorio_tipo_despesa',
        'relatorios_interface',
        'version_control',
        'xlwings',
        'xlwings.main',
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.backends.backend_tkagg',
        'matplotlib.backends.backend_agg',
        'matplotlib.figure',
        'matplotlib.font_manager',
        'matplotlib.text',
        'matplotlib.backends',
        'pylab',
    ] + collect_submodules('numpy') + collect_submodules('pandas') + collect_submodules('matplotlib'),
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False
)

# Adicionar arquivos de dados do matplotlib
matplotlib_data = collect_data_files('matplotlib')
a.datas += matplotlib_data

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
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    version=version_string,
)