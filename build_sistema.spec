# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_submodules

block_cipher = None

NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_TESTE"
ARQUIVO_PRINCIPAL = "src/sistema_principal.py"
ICONE = "logo3.ico"

# ====================================================================
# PYTHON-DOCX — incluir pasta inteira como dado + submodulos como imports
# ====================================================================

docx_datas = []
docx_hiddenimports = []

try:
    import docx as _docx
    docx_dir = os.path.dirname(_docx.__file__)
    # Incluir TODA a pasta docx como dado — garante que os .py ficam acessíveis
    docx_datas.append((docx_dir, 'docx'))
    print(f"[OK] pasta docx incluída: {docx_dir}")
except Exception as e:
    print(f"[AVISO] Erro ao localizar docx: {e}")

try:
    docx_hiddenimports += collect_submodules('docx')
    docx_hiddenimports += [
        'docx', 'docx.api', 'docx.shared', 'docx.enum', 'docx.enum.text',
        'docx.enum.style', 'docx.enum.dml', 'docx.enum.section', 'docx.enum.table',
        'docx.oxml', 'docx.oxml.ns', 'docx.oxml.shared', 'docx.oxml.text',
        'docx.oxml.table', 'docx.oxml.section', 'docx.oxml.styles',
        'docx.opc', 'docx.opc.constants', 'docx.opc.packuri', 'docx.opc.part',
        'docx.opc.parts', 'docx.opc.parts.coreprops',
        'docx.parts', 'docx.parts.document', 'docx.parts.image',
        'docx.text', 'docx.text.paragraph', 'docx.text.run',
        'docx.document', 'docx.table', 'docx.section',
        'docx.styles', 'docx.styles.styles',
        'docx.image', 'docx.image.png', 'docx.image.jpeg',
        'docx.blkcntnr', 'docx.drawing', 'docx.shape',
        'docx.comments', 'docx.settings', 'docx.package',
    ]
    print("[OK] python-docx hiddenimports configurados")
except Exception as e:
    print(f"[AVISO] Erro nos hiddenimports do docx: {e}")

# ====================================================================
# LXML
# ====================================================================

from PyInstaller.utils.hooks import collect_all

lxml_datas, lxml_binaries, lxml_hiddenimports = [], [], []
try:
    lxml_datas, lxml_binaries, lxml_hiddenimports = collect_all('lxml')
    lxml_hiddenimports += collect_submodules('lxml')
    print("[OK] lxml coletado")
except Exception as e:
    print(f"[AVISO] lxml: {e}")

# ====================================================================
# OUTROS MÓDULOS
# ====================================================================

outras_datas, outras_binaries, outras_hiddenimports = [], [], []
for modulo in ['openpyxl', 'pandas', 'PIL', 'reportlab', 'num2words', 'holidays']:
    try:
        d, b, h = collect_all(modulo)
        outras_datas += d
        outras_binaries += b
        outras_hiddenimports += h
        print(f"[OK] {modulo} coletado")
    except Exception as e:
        # NUNCA engolir esse erro silenciosamente — foi exatamente isso
        # que escondeu a ausência do num2words num build anterior.
        print(f"[AVISO] Falha ao coletar {modulo}: {e}")

# ====================================================================
# HIDDEN IMPORTS
# ====================================================================

hidden_imports = list(set(
    docx_hiddenimports + lxml_hiddenimports + outras_hiddenimports + [
        'tkinter', 'tkinter.ttk', 'tkinter.scrolledtext',
        'tkinter.filedialog', 'tkinter.messagebox',
        'matplotlib', 'xlwings', 'tkcalendar',
        'babel', 'babel.numbers', 'dotenv',
        'ambiente_config', 'version_control', 'Sistema_Entrada_Dados',
        'src.ambiente_config', 'src.version_control', 'src.Sistema_Entrada_Dados',
        'src.relatorios_interface', 'src.relatorio_despesas_aprimorado',
        'src.relatorio_despesas_service', 'src.relatorio_tipo_despesa',
        'src.relatorio_categoria', 'src.relatorio_fornecedores',
        'src.relatorio_contratos_medicoes', 'src.relatorio_gerencial_engenheiro',
        'src.relatorio_gerencial_pdf',
        'src.despesas_rateadas', 'src.gestao_medicoes', 'src.gestao_taxas',
        'src.configuracoes_sistema',
        'src.modules', 'src.modules.gerador_contrato',
        'src.controle_pagamentos_taxas', 'src.controle_pagamentos',
        'src.pagamentos_eventos', 'src.verificador_sistema',
        'src.corrigir_imports_sistema', 'src.finalizacao_quinzena',
        'src.correcao_monetaria', 'src.teste_certificado_automatico',
        'src.config.utils', 'src.config.dialogs', 'src.config.logger_config',
        'src.config.window_config', 'src.config.config',
    ]
))

print(f"[INFO] Total hidden imports: {len(hidden_imports)}")

# ====================================================================
# DADOS
# ====================================================================

datas = []

if os.path.exists('logo.png'):
    datas.append(('logo.png', '.'))
    print("[OK] logo.png adicionado")

if os.path.exists('logo3.png'):
    datas.append(('logo3.png', '.'))
    print("[OK] logo3.png adicionado")
else:
    print("[AVISO] logo3.png não encontrado na raiz do projeto — sistema_principal.py e window_config.py não vão achá-lo no .exe")

if os.path.exists('logo3.ico'):
    datas.append(('logo3.ico', '.'))
    print("[OK] logo3.ico adicionado")
else:
    print("[AVISO] logo3.ico não encontrado na raiz do projeto — window_config.py não vai achar o ícone no .exe")

if os.path.exists('src'):
    datas.append(('src', 'src'))
    print("[OK] src/ adicionado")

if os.path.exists('config'):
    datas.append(('config', 'config'))
    print("[OK] config/ adicionado")

datas += docx_datas + lxml_datas + outras_datas

binaries = [] + lxml_binaries + outras_binaries

print(f"[INFO] Total datas: {len(datas)}")
print(f"[INFO] Total binaries: {len(binaries)}")

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
    runtime_hooks=['hook_docx_runtime.py'],
    excludes=['matplotlib.tests', 'numpy.tests', 'pandas.tests', 'PIL.tests'],
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
    name=NOME_EXECUTAVEL,
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
    icon=ICONE if os.path.exists(ICONE) else None,
)
