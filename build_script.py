#!/usr/bin/env python3
"""
Script de build para o Sistema de Gestão Financeira
Execute este script para gerar o executável
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# Configurações do build
APP_NAME = "Sistema_Gestao_Financeira"
MAIN_SCRIPT = "src/sistema_principal.py"
ICON_FILE = "logo.ico"  # Será criado se logo.png existir

# Diretórios
BUILD_DIR = "build"
DIST_DIR = "dist"
SPEC_DIR = "."

def clean_build():
    """Remove diretórios de build anteriores"""
    print("🧹 Limpando builds anteriores...")
    
    dirs_to_clean = [BUILD_DIR, DIST_DIR, "__pycache__"]
    
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"   Removido: {dir_name}")
    
    # Remove arquivos .spec antigos
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"   Removido: {spec_file}")

def convert_png_to_ico():
    """Converte logo.png para logo.ico se necessário"""
    png_path = Path("logo.png")
    ico_path = Path(ICON_FILE)
    
    if png_path.exists() and not ico_path.exists():
        try:
            from PIL import Image
            print("🖼️  Convertendo logo.png para .ico...")
            
            img = Image.open(png_path)
            # Redimensionar para tamanhos padrão de ícone
            img = img.resize((256, 256), Image.Resampling.LANCZOS)
            img.save(ico_path, format='ICO', sizes=[(256, 256), (128, 128), (64, 64), (32, 32), (16, 16)])
            print(f"   ✅ Ícone criado: {ico_path}")
            return True
            
        except ImportError:
            print("   ⚠️  Pillow não instalado. Executável será criado sem ícone.")
            print("   💡 Instale com: pip install Pillow")
            return False
        except Exception as e:
            print(f"   ❌ Erro ao converter ícone: {e}")
            return False
    
    return ico_path.exists()

def get_hidden_imports():
    """Retorna lista de imports que podem não ser detectados automaticamente"""
    hidden_imports = [
        # GUI
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        
        # Imagens
        'PIL',
        'PIL.Image',
        'PIL.ImageTk',
        
        # Manipulação de dados
        'pandas',
        'numpy',
        'openpyxl',
        'xlwings',
        
        # Datas e localização
        'babel',
        'babel.numbers',
        'dateutil.relativedelta',
        'tkcalendar',
        
        # Validação
        'validate_docbr',
        
        # Relatórios PDF
        'reportlab',
        'reportlab.pdfgen',
        'reportlab.pdfgen.canvas',
        'reportlab.lib',
        'reportlab.lib.pagesizes',
        'reportlab.lib.styles',
        'reportlab.lib.enums',
        'reportlab.lib.colors',
        'reportlab.platypus',
        
        # Gráficos
        'matplotlib',
        'matplotlib.pyplot',
        'matplotlib.backends.backend_tkagg',
        
        # Configurações
        'python-dotenv',
        'dotenv',
        
        # Módulos específicos do sistema
        'version_control',
        'controle_pagamentos_taxas', 
        'Sistema_Entrada_Dados',
        'relatorios_interface',
        'relatorio_despesas_aprimorado',
        'despesas_rateadas',
        'gestao_medicoes',
        'configuracoes_sistema',
    ]
    
    return hidden_imports

def get_data_files():
    """Retorna lista de arquivos de dados para incluir"""
    data_files = []
    
    # Arquivos essenciais na raiz
    files_to_include = [
        "logo.png",
        "logo.ico",
        ".env",
        "requirements.txt"
    ]
    
    for file_name in files_to_include:
        if os.path.exists(file_name):
            data_files.append((file_name, "."))
    
    # Todo o diretório src (que contém tudo)
    if os.path.exists("src"):
        # Incluir todo o conteúdo do src, mantendo a estrutura
        for root, dirs, files in os.walk("src"):
            for file in files:
                if file.endswith(('.py', '.txt', '.json', '.csv', '.xlsx', '.png', '.ico', '.env')):
                    src_path = os.path.join(root, file)
                    # Manter a estrutura de diretórios
                    dest_path = root
                    data_files.append((src_path, dest_path))
    
    # Diretórios adicionais na raiz (se existirem)
    additional_dirs = [
        "templates",
        "data", 
        "assets",
        "resources"
    ]
    
    for dir_name in additional_dirs:
        if os.path.exists(dir_name):
            data_files.append((f"{dir_name}/*", dir_name))
    
    return data_files

def create_spec_file():
    """Cria arquivo .spec personalizado"""
    spec_content = f'''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

# Imports ocultos (módulos que podem não ser detectados automaticamente)
hidden_imports = {get_hidden_imports()}

# Arquivos de dados
datas = {get_data_files()}

a = Analysis(
    ['{MAIN_SCRIPT}'],
    pathex=[],
    binaries=[],
    datas=datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={{}},
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
    name='{APP_NAME}',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False para aplicação GUI
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='{ICON_FILE if os.path.exists(ICON_FILE) else None}',
)
'''
    
    spec_filename = f"{APP_NAME}.spec"
    with open(spec_filename, 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    return spec_filename

def build_executable():
    """Executa o PyInstaller"""
    print("🔨 Iniciando build do executável...")
    
    # Verificar se arquivo principal existe
    if not os.path.exists(MAIN_SCRIPT):
        print(f"❌ Arquivo principal não encontrado: {MAIN_SCRIPT}")
        print("   Ajuste a variável MAIN_SCRIPT no início deste arquivo")
        return False
    
    # Converter ícone se necessário
    has_icon = convert_png_to_ico()
    
    # Criar arquivo .spec
    spec_file = create_spec_file()
    print(f"📄 Arquivo spec criado: {spec_file}")
    
    # Comando PyInstaller
    cmd = [
        "pyinstaller",
        "--clean",  # Limpar cache
        spec_file
    ]
    
    print(f"🚀 Executando: {' '.join(cmd)}")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        print("✅ Build concluído com sucesso!")
        
        # Mostrar localização do executável
        exe_path = Path(DIST_DIR) / f"{APP_NAME}.exe"
        if exe_path.exists():
            print(f"📦 Executável criado em: {exe_path.absolute()}")
            print(f"💾 Tamanho: {exe_path.stat().st_size / (1024*1024):.1f} MB")
        
        return True
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro durante o build:")
        print(f"   Código de saída: {e.returncode}")
        if e.stdout:
            print(f"   Saída: {e.stdout}")
        if e.stderr:
            print(f"   Erro: {e.stderr}")
        return False
    
    except FileNotFoundError:
        print("❌ PyInstaller não encontrado!")
        print("   Instale com: pip install pyinstaller")
        return False

def post_build_cleanup():
    """Limpeza pós-build"""
    print("🧹 Fazendo limpeza pós-build...")
    
    # Manter apenas o executável em dist/
    dist_path = Path(DIST_DIR)
    if dist_path.exists():
        for item in dist_path.iterdir():
            if item.name != f"{APP_NAME}.exe" and item.is_dir():
                shutil.rmtree(item)
                print(f"   Removido: {item}")

def main():
    """Função principal"""
    print("=" * 60)
    print(f"🏗️  BUILD SCRIPT - {APP_NAME}")
    print("=" * 60)
    
    # Verificar se estamos no diretório correto
    if not os.path.exists("src") and not os.path.exists(MAIN_SCRIPT):
        print("❌ Execute este script no diretório raiz do projeto!")
        return
    
    try:
        # 1. Limpeza
        clean_build()
        
        # 2. Build
        success = build_executable()
        
        if success:
            # 3. Limpeza pós-build (opcional)
            # post_build_cleanup()
            
            print("\n" + "=" * 60)
            print("🎉 BUILD CONCLUÍDO COM SUCESSO!")
            print("=" * 60)
            print(f"📦 Executável disponível em: dist/{APP_NAME}.exe")
            print("💡 Teste o executável antes de distribuir")
        else:
            print("\n❌ Build falhou. Verifique os erros acima.")
    
    except KeyboardInterrupt:
        print("\n⚠️  Build cancelado pelo usuário")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")

if __name__ == "__main__":
    main()