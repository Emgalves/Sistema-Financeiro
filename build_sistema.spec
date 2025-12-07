# -*- mode: python ; coding: utf-8 -*-
"""
Arquivo SPEC otimizado para Sistema de Gestão Financeira
Resolve problemas com python-docx e garante todos os módulos necessários
"""

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_submodules, collect_data_files

block_cipher = None

# ====================================================================
# CONFIGURAÇÕES BÁSICAS
# ====================================================================

NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_PRODUCAO"  # Altere para _TESTE se necessário
ARQUIVO_PRINCIPAL = "src/sistema_principal.py"
ICONE = "logo1.ico"  # ou logo1.png

# ====================================================================
# COLETA COMPLETA DE MÓDULOS PROBLEMÁTICOS
# ====================================================================

# Python-docx (SOLUÇÃO COMPLETA)
docx_datas = []
docx_binaries = []
docx_hiddenimports = []

try:
    # Coletar TUDO do python-docx
    tmp_datas, tmp_binaries, tmp_hiddenimports = collect_all('docx')
    docx_datas += tmp_datas
    docx_binaries += tmp_binaries
    docx_hiddenimports += tmp_hiddenimports
    
    # Garantir submódulos específicos
    docx_hiddenimports += collect_submodules('docx')
    docx_hiddenimports += [
        'docx',
        'docx.shared',
        'docx.enum',
        'docx.enum.text',
        'docx.enum.style',
        'docx.oxml',
        'docx.oxml.ns',
        'docx.oxml.shared',
        'docx.oxml.text',
        'docx.oxml.table',
        'docx.oxml.section',
        'docx.opc',
        'docx.opc.constants',
        'docx.opc.packuri',
        'docx.opc.phys_pkg',
        'docx.parts',
        'docx.parts.document',
        'docx.text',
        'docx.text.paragraph',
        'docx.text.run',
        'docx.document',
        'docx.table',
        'docx.section',
        'docx.styles',
    ]
    
    print(f"[OK] python-docx: {len(docx_datas)} arquivos de dados coletados")
    print(f"[OK] python-docx: {len(docx_hiddenimports)} imports encontrados")
    
except Exception as e:
    print(f"[AVISO] Erro ao coletar python-docx: {e}")

# LXML (dependência crítica do python-docx)
lxml_datas = []
lxml_binaries = []
lxml_hiddenimports = []

try:
    tmp_datas, tmp_binaries, tmp_hiddenimports = collect_all('lxml')
    lxml_datas += tmp_datas
    lxml_binaries += tmp_binaries
    lxml_hiddenimports += tmp_hiddenimports
    lxml_hiddenimports += collect_submodules('lxml')
    
    print(f"[OK] lxml: {len(lxml_binaries)} binários coletados")
    
except Exception as e:
    print(f"[AVISO] Erro ao coletar lxml: {e}")

# Outras bibliotecas com coleta automática
outras_datas = []
outras_binaries = []
outras_hiddenimports = []

bibliotecas_auto = [
    'openpyxl',
    'pandas',
    'PIL',
    'reportlab',
    'matplotlib',
    'xlwings',
    'tkcalendar',
    'babel',
    'num2words',
    'dotenv',
]

for lib in bibliotecas_auto:
    try:
        tmp_datas, tmp_binaries, tmp_hiddenimports = collect_all(lib)
        outras_datas += tmp_datas
        outras_binaries += tmp_binaries
        outras_hiddenimports += tmp_hiddenimports
        print(f"[OK] {lib} coletado")
    except:
        print(f"[AVISO] {lib} não encontrado (pode não estar instalado)")

# ====================================================================
# HIDDEN IMPORTS COMPLETOS
# ====================================================================

hidden_imports = [
    # === TKINTER ===
    'tkinter',
    'tkinter.ttk',
    'tkinter.scrolledtext',
    'tkinter.filedialog',
    'tkinter.messagebox',
    
    # === MÓDULOS DO SISTEMA ===
    'ambiente_config',
    'version_control',
    'Sistema_Entrada_Dados',
    'src.ambiente_config',
    'src.version_control',
    'src.Sistema_Entrada_Dados',
    
    # === RELATÓRIOS ===
    'src.relatorios_interface',
    'src.relatorio_despesas_aprimorado',
    'src.relatorio_despesas_service',
    'src.relatorio_tipo_despesa',
    'src.relatorio_categoria',
    'src.relatorio_fornecedores',
    'src.relatorio_contratos_medicoes',
    'src.relatorio_gerencial_engenheiro',
    'src.relatorio_gerencial_pdf',
    
    # === GESTÃO ===
    'src.despesas_rateadas',
    'src.gestao_medicoes',
    'src.gestao_taxas',
    'src.configuracoes_sistema',
    
    # === MODULES (Gerador de Contratos) ===
    'src.modules',
    'src.modules.gerador_contrato',
    
    # === CONTROLE ===
    'src.controle_pagamentos_taxas',
    'src.controle_pagamentos',
    'src.pagamentos_eventos',
    
    # === UTILITÁRIOS ===
    'src.verificador_sistema',
    'src.corrigir_imports_sistema',
    'src.finalizacao_quinzena',
    'src.correcao_monetaria',
    'src.teste_certificado_automatico',
    
    # === CONFIG ===
    'src.config.utils',
    'src.config.dialogs',
    'src.config.logger_config',
    'src.config.window_config',
    'src.config.config',
]

# Combinar todos os hidden imports
hidden_imports += docx_hiddenimports
hidden_imports += lxml_hiddenimports
hidden_imports += outras_hiddenimports

# Remover duplicatas
hidden_imports = list(set(hidden_imports))

print(f"\n[INFO] Total de hidden imports: {len(hidden_imports)}")

# ====================================================================
# DADOS ADICIONAIS
# ====================================================================

datas = []

# Logo
if os.path.exists('logo.png'):
    datas.append(('logo.png', '.'))
    print("[OK] logo.png adicionado")

# Diretório src completo
if os.path.exists('src'):
    datas.append(('src', 'src'))
    print("[OK] Diretório src/ adicionado")

# Diretório config
if os.path.exists('config'):
    datas.append(('config', 'config'))
    print("[OK] Diretório config/ adicionado")

# Adicionar dados coletados
datas += docx_datas
datas += lxml_datas
datas += outras_datas

# Combinar binários
binaries = []
binaries += docx_binaries
binaries += lxml_binaries
binaries += outras_binaries

print(f"[INFO] Total de arquivos de dados: {len(datas)}")
print(f"[INFO] Total de binários: {len(binaries)}")

# ====================================================================
# ANÁLISE
# ====================================================================

a = Analysis(
    [ARQUIVO_PRINCIPAL],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib.tests',
        'numpy.tests',
        'pandas.tests',
        'PIL.tests',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# ====================================================================
# PYZ (arquivo compactado Python)
# ====================================================================

pyz = PYZ(
    a.pure,
    a.zipped_data,
    cipher=block_cipher
)

# ====================================================================
# EXE (executável)
# ====================================================================

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name=NOME_EXECUTAVEL,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Sem console para interface gráfica
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=ICONE if os.path.exists(ICONE) else None,
)

print("\n" + "="*70)
print("ARQUIVO SPEC CONFIGURADO COM SUCESSO".center(70))
print("="*70)
print(f"\nExecutável: {NOME_EXECUTAVEL}.exe")
print(f"Ícone: {ICONE if os.path.exists(ICONE) else 'Não definido'}")
print(f"Console: Desabilitado (interface gráfica)")
print("\n" + "="*70)
