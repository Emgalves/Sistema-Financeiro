# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Imports ocultos (módulos que podem não ser detectados automaticamente)
hidden_imports = ['tkinter', 'tkinter.ttk', 'tkinter.messagebox', 'tkinter.filedialog', 'PIL', 'PIL.Image', 'PIL.ImageTk', 'pandas', 'numpy', 'openpyxl', 'xlwings', 'babel', 'babel.numbers', 'dateutil.relativedelta', 'tkcalendar', 'validate_docbr', 'reportlab', 'reportlab.pdfgen', 'reportlab.pdfgen.canvas', 'reportlab.lib', 'reportlab.lib.pagesizes', 'reportlab.lib.styles', 'reportlab.lib.enums', 'reportlab.lib.colors', 'reportlab.platypus', 'matplotlib', 'matplotlib.pyplot', 'matplotlib.backends.backend_tkagg', 'python-dotenv', 'dotenv', 'version_control', 'controle_pagamentos_taxas', 'Sistema_Entrada_Dados', 'relatorios_interface', 'relatório_interface', 'relatorio_interface', 'relatorios_sistema', 'relatório_despesas_aprimorado', 'relatorio_despesas_aprimorado', 'relatorio_despesas', 'sistema_relatorios', 'despesas_rateadas', 'gestao_medicoes', 'configuracoes_sistema', 'src.relatorios_interface', 'src.relatório_interface', 'src.relatorio_interface', 'src.relatório_despesas_aprimorado', 'src.relatorio_despesas_aprimorado', 'src.Sistema_Entrada_Dados', 'src.despesas_rateadas', 'src.gestao_medicoes', 'src.configuracoes_sistema', 'src.controle_pagamentos_taxas', 'src.version_control']

# Arquivos de dados
datas = [('logo.png', '.'), ('logo.ico', '.'), ('.env', '.'), ('src\\configuracoes_sistema.py', 'src'), ('src\\controle_pagamentos.py', 'src'), ('src\\controle_pagamentos_taxas.py', 'src'), ('src\\despesas_rateadas.py', 'src'), ('src\\finalizacao_quinzena.py', 'src'), ('src\\gestao_medicoes.py', 'src'), ('src\\gestao_taxas.py', 'src'), ('src\\pagamentos_eventos.py', 'src'), ('src\\relatorios_interface.py', 'src'), ('src\\relatorio_categoria.py', 'src'), ('src\\relatorio_contratos_medicoes.py', 'src'), ('src\\relatorio_despesas_aprimorado.py', 'src'), ('src\\relatorio_fornecedores.py', 'src'), ('src\\relatorio_tipo_despesa.py', 'src'), ('src\\requirements.txt', 'src'), ('src\\Sistema_Entrada_Dados.py', 'src'), ('src\\sistema_principal.py', 'src'), ('src\\verificador_sistema.py', 'src'), ('src\\version_control.py', 'src'), ('src\\__init__.py', 'src'), ('src\\config\\config.py', 'src\\config'), ('src\\config\\dialogs.py', 'src\\config'), ('src\\config\\logger_config.py', 'src\\config'), ('src\\config\\parametros_sistema.json', 'src\\config'), ('src\\config\\paths.py', 'src\\config'), ('src\\config\\utils.py', 'src\\config'), ('src\\config\\window_config.py', 'src\\config'), ('src\\config\\__init__.py', 'src\\config')]

a = Analysis(
    ['src/sistema_principal.py'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
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
    name='Sistema_Gestao_Financeira',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Mude para True se precisar ver erros no console
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',
)

# Versão DEBUG - descomente as linhas abaixo para debug
# exe_debug = EXE(
#     pyz,
#     a.scripts,
#     a.binaries,
#     a.zipfiles,
#     a.datas,
#     [],
#     name='Sistema_Gestao_Financeira_DEBUG',
#     debug=True,
#     bootloader_ignore_signals=False,
#     strip=False,
#     upx=False,
#     upx_exclude=[],
#     runtime_tmpdir=None,
#     console=True,  # Console habilitado para debug
#     disable_windowed_traceback=False,
#     argv_emulation=False,
#     target_arch=None,
#     codesign_identity=None,
#     entitlements_file=None,
#     icon='logo.ico',
# )
