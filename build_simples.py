#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de build simplificado - sem emojis para compatibilidade com Windows
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def main():
    print("=" * 50)
    print("BUILD SISTEMA GESTAO FINANCEIRA")
    print("=" * 50)
    
    # Verificar diretório
    if not os.path.exists("src/sistema_principal.py"):
        print("ERRO: Execute no diretório raiz do projeto!")
        return
    
    print("Diretorio correto confirmado")
    
    # Limpar builds anteriores
    print("Limpando builds anteriores...")
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  Removido: {dir_name}")
    
    # Remover arquivos .spec antigos
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"  Removido: {spec_file}")
    
    # Criar comando PyInstaller direto
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",
        "--name=Sistema_Gestao_Financeira",
        "--add-data=logo.png;.",
        "--add-data=src;src",
        "--hidden-import=tkinter",
        "--hidden-import=tkinter.ttk",
        "--hidden-import=PIL",
        "--hidden-import=pandas",
        "--hidden-import=openpyxl",
        "--hidden-import=reportlab",
        "--hidden-import=matplotlib",
        "--hidden-import=xlwings",
        "--hidden-import=src.Sistema_Entrada_Dados",
        "--hidden-import=src.relatorios_interface",
        "--hidden-import=src.relatorio_despesas_aprimorado",
        "--hidden-import=src.relatorio_despesas_service",
        "--hidden-import=src.despesas_rateadas",
        "--hidden-import=src.gestao_medicoes",
        "--hidden-import=src.controle_pagamentos_taxas",
        "--hidden-import=src.controle_pagamentos",
        "--hidden-import=src.relatorio_tipo_despesa",
        "--hidden-import=src.verificador_sistema",
        "--hidden-import=src.gestao_taxas",
        "--hidden-import=src.pagamentos_eventos",
        "--hidden-import=src.relatorio_categoria",
        "--hidden-import=src.relatorio_fornecedores",
        "--hidden-import=src.relatorio_contratos_medicoes",
        "--hidden-import=src.corrigir_imports_sistema",
        "--hidden-import=src.finalizacao_quinzena",
        "--hidden-import=src.correcao_monetaria",
        "--hidden-import=src.configuracoes_sistema",
        "--hidden-import=src.version_control",
        "src/sistema_principal.py"
    ]
    
    # Adicionar ícone se existir
    if os.path.exists("logo1.png"):
        # Tentar converter para ICO# Adicionar ícone se existir
        try:
            from PIL import Image
            img = Image.open("logo1.png")
            img.save("logo1.ico")
            cmd.insert(-1, "--icon=logo1.ico")
            print("Icone convertido e adicionado")
        except:
            print("Nao foi possivel converter icone")
    
    print("Executando PyInstaller...")
    print("Comando:", " ".join(cmd[:5]) + "... (comando completo)")
    
    try:
        result = subprocess.run(cmd, check=True, capture_output=True, text=True)
        
        # Verificar se executável foi criado
        exe_path = Path("dist/Sistema_Gestao_Financeira.exe")
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024*1024)
            print(f"BUILD CONCLUIDO COM SUCESSO!")
            print(f"Executavel: {exe_path.absolute()}")
            print(f"Tamanho: {size_mb:.1f} MB")
            
            # Testar execução
            test = input("\nDeseja testar o executavel? (s/n): ").lower()
            if test == 's':
                print("Iniciando teste...")
                subprocess.Popen([str(exe_path)])
        else:
            print("ERRO: Executavel nao foi criado!")
            
    except subprocess.CalledProcessError as e:
        print("ERRO durante o build:")
        if e.stdout:
            print("Saida:", e.stdout[-500:])  # Últimas 500 chars
        if e.stderr:
            print("Erro:", e.stderr[-500:])   # Últimas 500 chars
            
    except FileNotFoundError:
        print("ERRO: PyInstaller nao encontrado!")
        print("Instale com: pip install pyinstaller")
    
    print("\n" + "=" * 50)

if __name__ == "__main__":
    main()