# -*- coding: utf-8 -*-
"""Build automatizado do Sistema de Gestao Financeira"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
from datetime import datetime

# ====================================================================
# CONFIGURACOES
# ====================================================================

NOME_BASE = "Sistema_Gestao_Financeira"
ARQUIVO_PRINCIPAL = "src/sistema_principal.py"
ICONE = "logo1.png"
LOGO = "logo.png"

# Modulos essenciais do sistema
MODULOS_SISTEMA = [
    # === BIBLIOTECAS BASE ===
    "tkinter",
    "tkinter.ttk",
    "tkinter.scrolledtext",
    "tkinter.filedialog",
    "tkinter.messagebox",
    
    # === BIBLIOTECAS PYTHON ===
    "PIL",
    "PIL.Image",
    "pandas",
    "openpyxl",
    "reportlab",
    "matplotlib",
    "xlwings",
    "tkcalendar",
    "babel",
    
    # === DOTENV ===
    "dotenv",
    
    # === MODULOS DO SISTEMA (ambas formas) ===
    "ambiente_config",
    "version_control",
    "Sistema_Entrada_Dados",
    "src.ambiente_config",
    "src.version_control",
    "src.Sistema_Entrada_Dados",
    
    # === RELATORIOS ===
    "src.relatorios_interface",
    "src.relatorio_despesas_aprimorado",
    "src.relatorio_despesas_service",
    "src.relatorio_tipo_despesa",
    "src.relatorio_categoria",
    "src.relatorio_fornecedores",
    "src.relatorio_contratos_medicoes",
    "src.relatorio_gerencial_engenheiro",
    
    # === GESTAO ===
    "src.despesas_rateadas",
    "src.gestao_medicoes",
    "src.gestao_taxas",
    "src.configuracoes_sistema",
    
    # === CONTROLE ===
    "src.controle_pagamentos_taxas",
    "src.controle_pagamentos",
    "src.pagamentos_eventos",
    
    # === UTILITARIOS ===
    "src.verificador_sistema",
    "src.corrigir_imports_sistema",
    "src.finalizacao_quinzena",
    "src.correcao_monetaria",
    "src.teste_certificado_automatico",
    
    # === CONFIG ===
    "src.config.utils",
    "src.config.dialogs",
    "src.config.logger_config",
    "src.config.window_config",
    "src.config.config",
]

# ====================================================================
# FUNCOES AUXILIARES
# ====================================================================

def print_header(texto):
    """Imprime cabecalho formatado"""
    print("\n" + "=" * 70)
    print(texto.center(70))
    print("=" * 70)

def print_step(texto):
    """Imprime passo da execucao"""
    print(f"\n>> {texto}")

def print_success(texto):
    """Imprime mensagem de sucesso"""
    print(f"   [OK] {texto}")

def print_warning(texto):
    """Imprime mensagem de aviso"""
    print(f"   [AVISO] {texto}")

def print_error(texto):
    """Imprime mensagem de erro"""
    print(f"   [ERRO] {texto}")

def verificar_ambiente():
    """Verifica se o ambiente esta configurado corretamente"""
    print_step("Verificando ambiente...")
    
    # Verificar diretorio
    if not os.path.exists(ARQUIVO_PRINCIPAL):
        print_error(f"Arquivo principal nao encontrado: {ARQUIVO_PRINCIPAL}")
        print_error("Execute este script no diretorio raiz do projeto!")
        return False
    
    print_success("Diretorio correto")
    
    # Verificar Python
    python_version = sys.version_info
    print_success(f"Python {python_version.major}.{python_version.minor}.{python_version.micro}")
    
    # Verificar PyInstaller
    try:
        result = subprocess.run(["pyinstaller", "--version"], 
                              capture_output=True, text=True)
        print_success(f"PyInstaller instalado")
    except FileNotFoundError:
        print_error("PyInstaller nao encontrado!")
        print_error("Instale com: pip install pyinstaller")
        return False
    
    return True

def verificar_modulos_importantes():
    """Verifica se modulos criticos existem"""
    print_step("Verificando modulos do sistema...")
    
    modulos_criticos = [
        "src/relatorios_interface.py",
        "src/ambiente_config.py",
        "src/version_control.py",
        "src/Sistema_Entrada_Dados.py",
    ]
    
    todos_ok = True
    for modulo in modulos_criticos:
        if os.path.exists(modulo):
            print_success(modulo)
        else:
            print_warning(f"{modulo} - NAO ENCONTRADO")
            todos_ok = False
    
    return todos_ok

def limpar_builds_anteriores():
    """Remove builds anteriores e cache"""
    print_step("Limpando builds anteriores...")
    
    # Diretorios
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print_success(f"Removido: {dir_name}/")
            except Exception as e:
                print_warning(f"Nao foi possivel remover {dir_name}/: {e}")
    
    # Arquivos .spec
    spec_files = list(Path(".").glob("*.spec"))
    for spec_file in spec_files:
        try:
            spec_file.unlink()
            print_success(f"Removido: {spec_file}")
        except Exception as e:
            print_warning(f"Nao foi possivel remover {spec_file}: {e}")
    
    # Arquivos .pyc em src
    pyc_files = list(Path("src").rglob("*.pyc"))
    if pyc_files:
        for pyc in pyc_files:
            pyc.unlink()
        print_success(f"Removidos {len(pyc_files)} arquivos .pyc")

def converter_icone():
    """Converte PNG para ICO se necessario"""
    if not os.path.exists(ICONE):
        return None
    
    try:
        from PIL import Image
        ico_path = "logo1.ico"
        
        if not os.path.exists(ico_path):
            print_step("Convertendo icone PNG -> ICO...")
            img = Image.open(ICONE)
            img.save(ico_path)
            print_success(f"Icone criado: {ico_path}")
        
        return ico_path
    except ImportError:
        print_warning("PIL nao instalado - icone nao sera incluido")
        return None
    except Exception as e:
        print_warning(f"Erro ao converter icone: {e}")
        return None

def escolher_ambientes():
    """Pergunta quais ambientes construir"""
    print_step("Escolha o(s) ambiente(s) para build:")
    print("  1 - TESTE apenas")
    print("  2 - PRODUCAO apenas")
    print("  3 - AMBOS (recomendado)")
    
    while True:
        try:
            escolha = input("\nOpcao (1-3): ").strip()
            
            if escolha == "1":
                return ["TESTE"]
            elif escolha == "2":
                return ["PRODUCAO"]
            elif escolha == "3":
                return ["TESTE", "PRODUCAO"]
            else:
                print_warning("Opcao invalida! Digite 1, 2 ou 3")
        except KeyboardInterrupt:
            print("\n\nBuild cancelado pelo usuario")
            sys.exit(0)

def construir_comando_pyinstaller(nome_exe, icone_path):
    """Constroi o comando do PyInstaller"""
    cmd = [
        "pyinstaller",
        "--clean",
        "--onefile",
        "--windowed",
        f"--name={nome_exe}",
    ]
    
    # Adicionar dados
    if os.path.exists(LOGO):
        cmd.append(f"--add-data={LOGO};.")
    
    cmd.append("--add-data=src;src")
    
    if os.path.exists("config"):
        cmd.append("--add-data=config;config")
    
    # Adicionar icone
    if icone_path and os.path.exists(icone_path):
        cmd.append(f"--icon={icone_path}")
    
    # Adicionar hidden imports
    for modulo in MODULOS_SISTEMA:
        cmd.append(f"--hidden-import={modulo}")
    
    # Arquivo principal
    cmd.append(ARQUIVO_PRINCIPAL)
    
    return cmd

def executar_build(ambiente):
    """Executa o build para um ambiente especifico"""
    print_header(f"BUILD: {ambiente}")
    
    nome_exe = f"{NOME_BASE}_{ambiente}"
    
    # Preparar icone
    icone_path = converter_icone()
    
    # Construir comando
    cmd = construir_comando_pyinstaller(nome_exe, icone_path)
    
    print_step(f"Executando PyInstaller...")
    print(f"   Modulos incluidos: {len(MODULOS_SISTEMA)}")
    print(f"   Nome do executavel: {nome_exe}.exe")
    
    # Executar
    try:
        result = subprocess.run(
            cmd,
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )
        
        # Verificar se foi criado
        exe_path = Path(f"dist/{nome_exe}.exe")
        
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024*1024)
            print_success(f"BUILD CONCLUIDO!")
            print(f"   Arquivo: {exe_path.name}")
            print(f"   Tamanho: {size_mb:.1f} MB")
            return exe_path
        else:
            print_error("Executavel nao foi criado!")
            return None
            
    except subprocess.CalledProcessError as e:
        print_error("Falha no build!")
        
        if e.stdout:
            print("\n--- SAIDA DO PYINSTALLER ---")
            print(e.stdout[-1000:])
        
        if e.stderr:
            print("\n--- ERROS DO PYINSTALLER ---")
            print(e.stderr[-1000:])
        
        return None

def mostrar_resumo(executaveis):
    """Mostra resumo final dos builds"""
    if not executaveis:
        print_header("NENHUM EXECUTAVEL FOI CRIADO")
        return False
    
    print_header("BUILD CONCLUIDO COM SUCESSO!")
    
    print(f"\n{len(executaveis)} executavel(is) criado(s):\n")
    
    for exe in executaveis:
        size_mb = exe.stat().st_size / (1024*1024)
        ambiente = "TESTE" if "TESTE" in exe.name else "PRODUCAO"
        simbolo = "[TESTE]" if ambiente == "TESTE" else "[PROD] "
        print(f"  {simbolo} {exe.name}")
        print(f"          Tamanho: {size_mb:.1f} MB")
        print(f"          Caminho: {exe.absolute()}")
        print()
    
    print("=" * 70)
    print("\nRECURSOS INCLUIDOS:")
    print("  [OK] Ambiente detectado pelo nome do arquivo")
    print("  [OK] relatorios_interface incluido")
    print("  [OK] version_control incluido")
    print(f"  [OK] {len(MODULOS_SISTEMA)} modulos do sistema incluidos")
    print("=" * 70)
    
    return True

def testar_executavel(executaveis):
    """Oferece opcao de testar um executavel"""
    print("\nDeseja testar algum executavel? (s/n): ", end="")
    
    try:
        resposta = input().strip().lower()
    except KeyboardInterrupt:
        print("\nSaindo...")
        return
    
    if resposta != 's':
        return
    
    if len(executaveis) == 1:
        print(f"\nIniciando {executaveis[0].name}...")
        subprocess.Popen([str(executaveis[0])])
        return
    
    # Multiplos executaveis
    print("\nEscolha qual testar:")
    for i, exe in enumerate(executaveis, 1):
        ambiente = "TESTE" if "TESTE" in exe.name else "PRODUCAO"
        print(f"  {i} - {exe.name} ({ambiente})")
    
    try:
        escolha = input("\nNumero: ").strip()
        idx = int(escolha) - 1
        
        if 0 <= idx < len(executaveis):
            print(f"\nIniciando {executaveis[idx].name}...")
            subprocess.Popen([str(executaveis[idx])])
        else:
            print_warning("Opcao invalida!")
    except (ValueError, KeyboardInterrupt):
        print("\nSaindo...")

# ====================================================================
# FUNCAO PRINCIPAL
# ====================================================================

def main():
    """Funcao principal do script de build"""
    print_header(f"SISTEMA DE BUILD - {NOME_BASE}")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    
    # 1. Verificar ambiente
    if not verificar_ambiente():
        return 1
    
    # 2. Verificar modulos importantes
    verificar_modulos_importantes()
    
    # 3. Limpar builds anteriores
    limpar_builds_anteriores()
    
    # 4. Escolher ambientes
    ambientes = escolher_ambientes()
    
    # 5. Executar builds
    executaveis_criados = []
    
    for ambiente in ambientes:
        exe = executar_build(ambiente)
        if exe:
            executaveis_criados.append(exe)
    
    # 6. Mostrar resumo
    if not mostrar_resumo(executaveis_criados):
        return 1
    
    # 7. Oferecer teste
    testar_executavel(executaveis_criados)
    
    print("\n" + "=" * 70)
    print("Script finalizado!".center(70))
    print("=" * 70 + "\n")
    
    return 0

# ====================================================================
# PONTO DE ENTRADA
# ====================================================================

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nBuild cancelado pelo usuario (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print_error(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)