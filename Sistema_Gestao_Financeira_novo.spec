# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['src/sistema_principal.py'],  # Script principal
    pathex=[],
    binaries=[],
    datas=[
        ('.env', '.'),  # Arquivo de ambiente
        ('src/', 'src/'),  # Pasta src completa
    ],
    hiddenimports=[
        'pandas',
        'openpyxl',
        'dotenv',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'tkinter.filedialog',
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        'src.gestao_taxas',
        'src.config',
        'src.config.config',
        'src.config.window_config',
        'src.config.logger_config',
        'src.config.utils',
        'src.version_control',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Criar um script de correção para importações
with open('src_fix.py', 'w') as f:
    f.write("""
# Fix para importações relativas vs. absolutas
import sys
import importlib

# Criar alias para config -> src.config
sys.modules['config'] = importlib.import_module('src.config')
for submodule in ['config', 'utils', 'window_config', 'logger_config']:
    full_name = f'src.config.{submodule}'
    alias = f'config.{submodule}'
    if alias not in sys.modules:
        try:
            sys.modules[alias] = importlib.import_module(full_name)
        except ImportError:
            pass
""")

# Adicionar o script de correção aos dados
a.datas += [('src_fix.py', 'src_fix.py', 'DATA')]

# Modificar o script principal para incluir a correção
a.scripts = [(None, 'import sys; exec(open("src_fix.py").read()); ' + open(a.scripts[0][1]).read(), a.scripts[0][1])]

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='SGF_Nova_Versao',  # Nome diferente para evitar conflitos
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,  # Mantenha True até resolver o problema
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='SGF_Nova_Versao',  # Nome diferente para evitar conflitos
)