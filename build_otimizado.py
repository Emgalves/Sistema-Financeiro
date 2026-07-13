# -*- coding: utf-8 -*-
"""
Script de Build Otimizado - Sistema de Gestão Financeira
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path
from datetime import datetime

# ====================================================================
# CONFIGURAÇÕES
# ====================================================================

SPEC_FILE = "build_sistema.spec"

# Pasta do docx no venv (fonte para copiar para dist/)
DOCX_ORIGEM = Path("venv/Lib/site-packages/docx")

# Arquivo de configuração de caminhos (distribuído junto com o exe)
CONFIG_JSON_ORIGEM = Path("config_caminhos.json")

# ====================================================================
# FUNÇÕES AUXILIARES
# ====================================================================

def print_header(texto):
    print("\n" + "=" * 70)
    print(texto.center(70))
    print("=" * 70)

def print_step(texto):
    print(f"\n>> {texto}")

def print_success(texto):
    print(f"   [OK] {texto}")

def print_warning(texto):
    print(f"   [AVISO] {texto}")

def print_error(texto):
    print(f"   [ERRO] {texto}")

# ====================================================================
# VERIFICAÇÕES PRÉ-BUILD
# ====================================================================

def verificar_dependencias():
    print_step("Verificando dependências críticas...")

    dependencias = {
        'docx': 'python-docx',
        'lxml': 'lxml',
        'openpyxl': 'openpyxl',
        'pandas': 'pandas',
        'PIL': 'Pillow',
        'reportlab': 'reportlab',
        'num2words': 'num2words',
    }

    faltando = []
    for modulo, nome_pip in dependencias.items():
        try:
            __import__(modulo)
            print_success(f"{nome_pip} instalado")
        except ImportError:
            print_error(f"{nome_pip} NÃO INSTALADO!")
            faltando.append(nome_pip)

    if faltando:
        print("\n" + "="*70)
        print_error("DEPENDÊNCIAS FALTANDO!")
        print(f"\nInstale com: pip install {' '.join(faltando)}")
        print("="*70)
        return False

    return True

def verificar_pyinstaller():
    print_step("Verificando PyInstaller...")
    try:
        result = subprocess.run(["pyinstaller", "--version"], capture_output=True, text=True)
        print_success(f"PyInstaller {result.stdout.strip()} encontrado")
        return True
    except FileNotFoundError:
        print_error("PyInstaller não encontrado!")
        print("\nInstale com: pip install pyinstaller")
        return False

def verificar_arquivos():
    print_step("Verificando arquivos do projeto...")

    arquivos_necessarios = [
        "src/sistema_principal.py",
        "src/gestao_medicoes.py",
        "src/modules/gerador_contrato.py",
        "hook_docx_runtime.py",
    ]

    todos_ok = True
    for arquivo in arquivos_necessarios:
        if os.path.exists(arquivo):
            print_success(arquivo)
        else:
            print_error(f"{arquivo} NÃO ENCONTRADO!")
            todos_ok = False

    # Verificar pasta docx no venv
    if DOCX_ORIGEM.exists():
        print_success(f"Pasta docx encontrada: {DOCX_ORIGEM}")
    else:
        print_warning(f"Pasta docx não encontrada em {DOCX_ORIGEM}")
        print_warning("O executável precisará da pasta docx copiada manualmente")

    # Verificar config_caminhos.json
    if CONFIG_JSON_ORIGEM.exists():
        print_success(f"config_caminhos.json encontrado")
    else:
        print_warning("config_caminhos.json não encontrado — será criado um padrão em dist/")

    return todos_ok

def limpar_builds_anteriores():
    print_step("Encerrando processos do executável (se rodando)...")
    for nome_exe in ["Sistema_Gestao_Financeira_PRODUCAO.exe", "Sistema_Gestao_Financeira_TESTE.exe"]:
        try:
            subprocess.run(["taskkill", "/F", "/IM", nome_exe], capture_output=True)
        except:
            pass

    print_step("Limpando builds anteriores...")
    for dir_name in ["build", "__pycache__"]:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print_success(f"Removido: {dir_name}/")
            except Exception as e:
                print_warning(f"Não foi possível remover {dir_name}/: {e}")

    # Limpar dist\ mas preservar config_caminhos.json se existir
    if os.path.exists("dist"):
        config_backup = None
        config_em_dist = Path("dist/config_caminhos.json")
        if config_em_dist.exists():
            config_backup = config_em_dist.read_text(encoding='utf-8')

        try:
            shutil.rmtree("dist")
            print_success("Removido: dist/")
        except Exception as e:
            print_warning(f"Não foi possível remover dist/: {e}")

        if config_backup:
            os.makedirs("dist", exist_ok=True)
            config_em_dist.write_text(config_backup, encoding='utf-8')
            print_success("config_caminhos.json preservado em dist/")

    # Limpar .pyc em src
    pyc_files = list(Path("src").rglob("*.pyc"))
    if pyc_files:
        for pyc in pyc_files:
            try:
                pyc.unlink()
            except:
                pass
        print_success(f"Removidos {len(pyc_files)} arquivos .pyc")

def converter_icone():
    # IMPORTANTE: usa logo3_icone.png (versão quadrada, com o desenho
    # centralizado), NÃO logo3.png. logo3.png é a versão retangular (~2:1)
    # usada em banners/cabeçalhos — convertê-la direto para .ico esmagaria
    # a imagem forçando-a num quadrado.
    png_path = "logo3_icone.png"
    ico_path = "logo3.ico"

    if not os.path.exists(png_path):
        print_warning("logo3_icone.png não encontrado - build sem ícone")
        return False

    # Regenerar sempre que o .png for mais novo que o .ico (ou o .ico não existir),
    # em vez de pular silenciosamente quando já existe um .ico desatualizado.
    if os.path.exists(ico_path) and os.path.getmtime(ico_path) >= os.path.getmtime(png_path):
        print_success("Ícone .ico já existe e está atualizado")
        return True

    try:
        from PIL import Image
        print_step("Convertendo PNG → ICO...")
        img = Image.open(png_path).convert("RGBA")
        # logo3_icone.png já é quadrada (desenho centralizado com margem) —
        # gerar múltiplas resoluções no .ico para que Windows escolha o
        # tamanho certo em cada contexto (ícone de arquivo, barra de
        # tarefas, título da janela etc.)
        tamanhos = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
        img.save(ico_path, sizes=tamanhos)
        print_success(f"Ícone criado: {ico_path}")
        return True
    except Exception as e:
        print_warning(f"Erro ao converter ícone: {e}")
        return False

# ====================================================================
# EDIÇÃO DO ARQUIVO SPEC
# ====================================================================

def escolher_ambiente():
    print_step("Escolha o ambiente:")
    print("  1 - TESTE")
    print("  2 - PRODUÇÃO")

    while True:
        try:
            escolha = input("\nOpção (1-2): ").strip()
            if escolha == "1":
                return "TESTE"
            elif escolha == "2":
                return "PRODUCAO"
            else:
                print_warning("Opção inválida! Digite 1 ou 2")
        except KeyboardInterrupt:
            print("\n\nBuild cancelado")
            sys.exit(0)

def atualizar_spec_para_ambiente(ambiente):
    print_step(f"Configurando .spec para ambiente {ambiente}...")

    if not os.path.exists(SPEC_FILE):
        print_error(f"Arquivo {SPEC_FILE} não encontrado!")
        return False

    try:
        with open(SPEC_FILE, 'r', encoding='utf-8') as f:
            conteudo = f.read()

        if ambiente == "TESTE":
            conteudo = conteudo.replace(
                'NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_PRODUCAO"',
                'NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_TESTE"'
            )
        else:
            conteudo = conteudo.replace(
                'NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_TESTE"',
                'NOME_EXECUTAVEL = "Sistema_Gestao_Financeira_PRODUCAO"'
            )

        with open(SPEC_FILE, 'w', encoding='utf-8') as f:
            f.write(conteudo)

        print_success(f"Arquivo .spec configurado para {ambiente}")
        return True

    except Exception as e:
        print_error(f"Erro ao atualizar .spec: {e}")
        return False

# ====================================================================
# PÓS-BUILD — copiar arquivos necessários para dist/
# ====================================================================

def copiar_arquivos_pos_build(exe_path):
    """
    Copia para dist/ os arquivos que precisam ficar junto com o executável:
      - pasta docx/ (python-docx não é extraído pelo PyInstaller)
      - config_caminhos.json (configuração de caminho do servidor)
    """
    print_step("Copiando arquivos necessários para dist/...")
    dist_dir = exe_path.parent

    # 1. Copiar pasta docx
    docx_destino = dist_dir / "docx"
    if DOCX_ORIGEM.exists():
        try:
            if docx_destino.exists():
                shutil.rmtree(docx_destino)
            shutil.copytree(str(DOCX_ORIGEM), str(docx_destino))
            print_success(f"Pasta docx copiada → {docx_destino}")
        except Exception as e:
            print_error(f"Erro ao copiar pasta docx: {e}")
            print_warning("Copie manualmente: xcopy venv\\Lib\\site-packages\\docx dist\\docx /E /I /Y")
    else:
        print_warning(f"Pasta docx não encontrada em {DOCX_ORIGEM}")
        print_warning("Copie manualmente: xcopy venv\\Lib\\site-packages\\docx dist\\docx /E /I /Y")

    # 2. Copiar config_caminhos.json
    config_destino = dist_dir / "config_caminhos.json"
    if CONFIG_JSON_ORIGEM.exists():
        try:
            shutil.copy2(str(CONFIG_JSON_ORIGEM), str(config_destino))
            print_success(f"config_caminhos.json copiado → {config_destino}")
        except Exception as e:
            print_error(f"Erro ao copiar config_caminhos.json: {e}")
    else:
        # Criar um padrão se não existir
        config_padrao = '''{
    "_instrucoes": [
        "Edite 'caminho_dados' se o servidor ou letra de drive mudar.",
        "NAO e necessario rebuild apos editar este arquivo.",
        "O caminho deve apontar para a pasta que contem 'Planilhas_Base' e 'Clientes'."
    ],
    "caminho_dados": "Z:/Servidor/Relatórios/Financeiro",
    "caminho_dados_alternativo": "//servidor/Servidor/Relatórios/Financeiro"
}'''
        try:
            config_destino.write_text(config_padrao, encoding='utf-8')
            print_success(f"config_caminhos.json padrão criado em dist/")
            print_warning("Edite o arquivo com o caminho correto do servidor antes de distribuir")
        except Exception as e:
            print_error(f"Erro ao criar config_caminhos.json: {e}")

    print_success("Arquivos pós-build concluídos")

# ====================================================================
# BUILD
# ====================================================================

def executar_build():
    print_header("INICIANDO BUILD")

    cmd = ["pyinstaller", "--clean", "--noconfirm", SPEC_FILE]

    print_step("Executando PyInstaller...")
    print(f"   Comando: {' '.join(cmd)}")

    try:
        result = subprocess.run(
            cmd,
            check=True,
            capture_output=True,
            text=True,
            encoding='utf-8',
            errors='replace'
        )

        print_success("PyInstaller executado com sucesso!")

        exe_files = list(Path("dist").glob("*.exe"))
        if exe_files:
            exe_path = exe_files[0]
            size_mb = exe_path.stat().st_size / (1024*1024)
            print_header("BUILD CONCLUÍDO!")
            print(f"\n   Arquivo: {exe_path.name}")
            print(f"   Tamanho: {size_mb:.1f} MB")
            print(f"   Caminho: {exe_path.absolute()}")
            return exe_path
        else:
            print_error("Nenhum executável foi criado em dist/")
            return None

    except subprocess.CalledProcessError as e:
        print_error("Falha no build!")
        if e.stdout:
            print("\n--- ÚLTIMAS LINHAS DA SAÍDA ---")
            linhas = e.stdout.split('\n')
            print('\n'.join(linhas[-30:]))
        if e.stderr:
            print("\n--- ERROS ---")
            print(e.stderr[-1000:])
        return None

def testar_executavel(exe_path):
    print("\nDeseja testar o executável? (s/n): ", end="")
    try:
        resposta = input().strip().lower()
    except KeyboardInterrupt:
        print("\nSaindo...")
        return

    if resposta == 's':
        print(f"\nIniciando {exe_path.name}...")
        try:
            subprocess.Popen([str(exe_path)])
            print_success("Executável iniciado!")
        except Exception as e:
            print_error(f"Erro ao iniciar: {e}")

# ====================================================================
# FUNÇÃO PRINCIPAL
# ====================================================================

def main():
    print_header("BUILD SISTEMA DE GESTÃO FINANCEIRA")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")

    if not verificar_dependencias():
        return 1

    if not verificar_pyinstaller():
        return 1

    if not verificar_arquivos():
        print_warning("Alguns arquivos não foram encontrados, mas continuando...")

    limpar_builds_anteriores()
    converter_icone()
    ambiente = escolher_ambiente()

    if not atualizar_spec_para_ambiente(ambiente):
        return 1

    exe_path = executar_build()
    if not exe_path:
        print_header("BUILD FALHOU")
        return 1

    # Copiar arquivos necessários para dist/ automaticamente
    copiar_arquivos_pos_build(exe_path)

    testar_executavel(exe_path)

    print("\n" + "=" * 70)
    print("Build finalizado com sucesso!".center(70))
    print("=" * 70)
    print(f"\nArquivos em dist/:")
    print(f"  - {exe_path.name}")
    print(f"  - docx/  (python-docx)")
    print(f"  - config_caminhos.json  (configuração de caminhos)")
    print(f"\nDistribua TODOS estes itens para o cliente.\n")

    return 0

# ====================================================================
# PONTO DE ENTRADA
# ====================================================================

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nBuild cancelado (Ctrl+C)")
        sys.exit(1)
    except Exception as e:
        print_error(f"Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
