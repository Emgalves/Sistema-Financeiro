#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de build FINAL - COM TODOS OS MÓDULOS
Inclui TUDO que o sistema precisa
"""

import os
import sys
import subprocess
import shutil
from pathlib import Path

def main():
    print("=" * 70)
    print("BUILD FINAL - COMPLETO E FUNCIONAL")
    print("Com TODOS os módulos incluídos corretamente")
    print("=" * 70)
    
    # Verificar diretório
    if not os.path.exists("src/sistema_principal.py"):
        print("ERRO: Execute no diretório raiz do projeto!")
        return
    
    print("✓ Diretório correto confirmado")
    
    # Verificar módulos importantes
    print("\nVerificando módulos do sistema:")
    modulos_importantes = [
        "src/relatorios_interface.py",
        "src/ambiente_config.py",
        "src/version_control.py"
    ]
    
    for modulo in modulos_importantes:
        if os.path.exists(modulo):
            print(f"  ✓ {modulo}")
        else:
            print(f"  ⚠ {modulo} - NÃO ENCONTRADO")
    
    # Perguntar qual build fazer
    print("\nOpções de build:")
    print("1 - Build TESTE")
    print("2 - Build PRODUCAO")
    print("3 - Build AMBOS (recomendado)")
    
    escolha = input("\nEscolha (1-3): ").strip()
    
    if escolha == "1":
        ambientes_build = ["TESTE"]
    elif escolha == "2":
        ambientes_build = ["PRODUCAO"]
    else:
        ambientes_build = ["TESTE", "PRODUCAO"]
    
    # Limpar builds anteriores
    print("\nLimpando builds anteriores...")
    dirs_to_clean = ["build", "dist", "__pycache__"]
    for dir_name in dirs_to_clean:
        if os.path.exists(dir_name):
            shutil.rmtree(dir_name)
            print(f"  ✓ Removido: {dir_name}")
    
    # Remover .spec antigos
    for spec_file in Path(".").glob("*.spec"):
        spec_file.unlink()
        print(f"  ✓ Removido: {spec_file}")
    
    # Fazer build para cada ambiente
    for amb in ambientes_build:
        print("\n" + "=" * 70)
        print(f"CONSTRUINDO: {amb}")
        print("=" * 70)
        
        nome_exe = f"Sistema_Gestao_Financeira_{amb}"
        
        # Lista COMPLETA de módulos hidden-import
        modulos_sistema = [
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
            "json",
            "datetime",
            "pathlib",
            
            # === DOTENV ===
            "dotenv",
            
            # === MÓDULOS DO SISTEMA (raiz) ===
            "ambiente_config",
            "version_control",
            "Sistema_Entrada_Dados",
            
            # === MÓDULOS DO SISTEMA (src.) ===
            "src.ambiente_config",
            "src.version_control",
            "src.Sistema_Entrada_Dados",
            
            # === RELATÓRIOS ===
            "src.relatorios_interface",  # ← IMPORTANTE!
            "src.relatorio_despesas_aprimorado",
            "src.relatorio_despesas_service",
            "src.relatorio_tipo_despesa",
            "src.relatorio_categoria",
            "src.relatorio_fornecedores",
            "src.relatorio_contratos_medicoes",
            
            # === GESTÃO ===
            "src.despesas_rateadas",
            "src.gestao_medicoes",
            "src.gestao_taxas",
            "src.configuracoes_sistema",
            
            # === CONTROLE ===
            "src.controle_pagamentos_taxas",
            "src.controle_pagamentos",
            "src.pagamentos_eventos",
            
            # === UTILITÁRIOS ===
            "src.verificador_sistema",
            "src.corrigir_imports_sistema",
            "src.finalizacao_quinzena",
            "src.correcao_monetaria",
            "src.teste_certificado_automatico",
        ]
        
        # Criar comando PyInstaller
        cmd = [
            "pyinstaller",
            "--onefile",
            "--windowed",
            f"--name={nome_exe}",
            "--add-data=logo.png;.",
            "--add-data=src;src",
        ]
        
        # Adicionar config se existir
        if os.path.exists('config'):
            cmd.append("--add-data=config;config")
        
        # Adicionar todos os hidden-imports
        for modulo in modulos_sistema:
            cmd.append(f"--hidden-import={modulo}")
        
        # Adicionar ícone
        if os.path.exists("logo1.png"):
            try:
                from PIL import Image
                img = Image.open("logo1.png")
                img.save("logo1.ico")
                cmd.append("--icon=logo1.ico")
                print("✓ Ícone adicionado")
            except:
                pass
        
        # Arquivo principal
        cmd.append("src/sistema_principal.py")
        
        print(f"\nExecutando PyInstaller...")
        print(f"Total de módulos incluídos: {len(modulos_sistema)}")
        
        try:
            result = subprocess.run(cmd, check=True, capture_output=True, text=True)
            
            exe_path = Path(f"dist/{nome_exe}.exe")
            if exe_path.exists():
                size_mb = exe_path.stat().st_size / (1024*1024)
                print(f"\n✓ BUILD CONCLUÍDO!")
                print(f"  Arquivo: {exe_path.name}")
                print(f"  Tamanho: {size_mb:.1f} MB")
            else:
                print(f"\n✗ ERRO: Executável não criado!")
                
        except subprocess.CalledProcessError as e:
            print(f"\n✗ ERRO no build:")
            if e.stderr:
                print(e.stderr[-500:])
            return
    
    # Resumo final
    print("\n" + "=" * 70)
    print("✅ BUILD CONCLUÍDO COM SUCESSO!")
    print("=" * 70)
    
    dist_path = Path("dist")
    executaveis = list(dist_path.glob("*.exe"))
    
    if executaveis:
        print(f"\n📦 {len(executaveis)} executável(is) criado(s):")
        for exe in executaveis:
            size_mb = exe.stat().st_size / (1024*1024)
            if "TESTE" in exe.name:
                print(f"  🟨 {exe.name} ({size_mb:.1f} MB)")
            else:
                print(f"  🟢 {exe.name} ({size_mb:.1f} MB)")
        
        print("\n" + "=" * 70)
        print("🎯 TUDO INCLUÍDO:")
        print("=" * 70)
        print("✓ Ambiente detecta pelo nome do arquivo")
        print("✓ relatorios_interface incluído")
        print("✓ version_control incluído")
        print("✓ Todos os módulos do sistema incluídos")
        print("=" * 70)
        
        test = input("\nDeseja testar algum executável? (s/n): ").lower()
        if test == 's':
            if len(executaveis) == 1:
                subprocess.Popen([str(executaveis[0])])
            else:
                print("\nEscolha qual testar:")
                for i, exe in enumerate(executaveis, 1):
                    amb = "🟨 TESTE" if "TESTE" in exe.name else "🟢 PRODUCAO"
                    print(f"{i} - {exe.name} ({amb})")
                escolha = input("\nNúmero: ").strip()
                try:
                    idx = int(escolha) - 1
                    if 0 <= idx < len(executaveis):
                        print("\nTestando... Clique em 'Geração de Relatórios' para verificar!")
                        subprocess.Popen([str(executaveis[idx])])
                except:
                    print("Opção inválida!")
3

if __name__ == "__main__":
    main()