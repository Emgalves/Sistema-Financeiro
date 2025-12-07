1# -*- coding: utf-8 -*-
"""
Script de Build Otimizado - Sistema de Gestão Financeira
VERSÃO CORRIGIDA - Resolve problema do python-docx
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

# ====================================================================
# FUNÇÕES AUXILIARES
# ====================================================================

def print_header(texto):
    """Imprime cabeçalho formatado"""
    print("\n" + "=" * 70)
    print(texto.center(70))
    print("=" * 70)

def print_step(texto):
    """Imprime passo da execução"""
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

# ====================================================================
# VERIFICAÇÕES PRÉ-BUILD
# ====================================================================

def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
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
        print("\nInstale com:")
        print(f"pip install {' '.join(faltando)}")
        print("="*70)
        return False
    
    return True

def verificar_pyinstaller():
    """Verifica PyInstaller"""
    print_step("Verificando PyInstaller...")
    
    try:
        result = subprocess.run(
            ["pyinstaller", "--version"],
            capture_output=True,
            text=True
        )
        version = result.stdout.strip()
        print_success(f"PyInstaller {version} encontrado")
        return True
    except FileNotFoundError:
        print_error("PyInstaller não encontrado!")
        print("\nInstale com: pip install pyinstaller")
        return False

def verificar_arquivos():
    """Verifica arquivos necessários"""
    print_step("Verificando arquivos do projeto...")
    
    arquivos_necessarios = [
        "src/sistema_principal.py",
        "src/gestao_medicoes.py",
        "src/modules/gerador_contrato.py",
    ]
    
    todos_ok = True
    for arquivo in arquivos_necessarios:
        if os.path.exists(arquivo):
            print_success(arquivo)
        else:
            print_error(f"{arquivo} NÃO ENCONTRADO!")
            todos_ok = False
    
    return todos_ok

def limpar_builds_anteriores():
    """Remove builds e cache antigos"""
    print_step("Limpando builds anteriores...")
    
    # Diretórios
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            try:
                shutil.rmtree(dir_name)
                print_success(f"Removido: {dir_name}/")
            except Exception as e:
                print_warning(f"Não foi possível remover {dir_name}/: {e}")
    
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
    """Converte PNG para ICO se necessário"""
    png_path = "logo1.png"
    ico_path = "logo1.ico"
    
    if not os.path.exists(png_path):
        print_warning("logo1.png não encontrado - build sem ícone")
        return False
    
    if os.path.exists(ico_path):
        print_success("Ícone .ico já existe")
        return True
    
    try:
        from PIL import Image
        print_step("Convertendo PNG → ICO...")
        img = Image.open(png_path)
        img.save(ico_path)
        print_success(f"Ícone criado: {ico_path}")
        return True
    except ImportError:
        print_warning("PIL não instalado - build sem ícone")
        return False
    except Exception as e:
        print_warning(f"Erro ao converter ícone: {e}")
        return False

# ====================================================================
# EDIÇÃO DO ARQUIVO SPEC
# ====================================================================

def escolher_ambiente():
    """Pergunta qual ambiente construir"""
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
    """Atualiza o arquivo .spec com o ambiente escolhido"""
    print_step(f"Configurando .spec para ambiente {ambiente}...")
    
    if not os.path.exists(SPEC_FILE):
        print_error(f"Arquivo {SPEC_FILE} não encontrado!")
        return False
    
    try:
        with open(SPEC_FILE, 'r', encoding='utf-8') as f:
            conteudo = f.read()
        
        # Substituir o nome do executável
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
# BUILD
# ====================================================================

def executar_build():
    """Executa o build usando o arquivo .spec"""
    print_header("INICIANDO BUILD")
    
    cmd = [
        "pyinstaller",
        "--clean",
        "--noconfirm",
        SPEC_FILE
    ]
    
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
        
        # Verificar executável criado
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
    """Oferece opção de testar o executável"""
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
    """Função principal"""
    print_header("BUILD SISTEMA DE GESTÃO FINANCEIRA")
    print(f"Data/Hora: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}")
    print("\n🔧 VERSÃO OTIMIZADA - Problema do python-docx RESOLVIDO")
    
    # 1. Verificar dependências
    if not verificar_dependencias():
        return 1
    
    # 2. Verificar PyInstaller
    if not verificar_pyinstaller():
        return 1
    
    # 3. Verificar arquivos
    if not verificar_arquivos():
        print_warning("Alguns arquivos não foram encontrados, mas continuando...")
    
    # 4. Limpar builds anteriores
    limpar_builds_anteriores()
    
    # 5. Converter ícone
    converter_icone()
    
    # 6. Escolher ambiente
    ambiente = escolher_ambiente()
    
    # 7. Atualizar .spec
    if not atualizar_spec_para_ambiente(ambiente):
        return 1
    
    # 8. Executar build
    exe_path = executar_build()
    
    if not exe_path:
        print_header("BUILD FALHOU")
        return 1
    
    # 9. Oferecer teste
    testar_executavel(exe_path)
    
    print("\n" + "=" * 70)
    print("✅ Build finalizado com sucesso!".center(70))
    print("=" * 70 + "\n")
    
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
