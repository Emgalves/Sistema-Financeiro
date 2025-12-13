#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de diagnóstico para problema de import do relatorios_interface
Execute este script para descobrir o problema exato
"""

import os
import sys
from pathlib import Path

print("=" * 70)
print("DIAGNÓSTICO - relatorios_interface")
print("=" * 70)

# 1. Verificar estrutura de pastas
print("\n1. ESTRUTURA DE PASTAS:")
print("-" * 70)

src_path = Path("src")
if src_path.exists():
    print(f"✅ Pasta src/ existe")
    
    # Listar arquivos importantes
    arquivos_importantes = [
        "src/__init__.py",
        "src/relatorios_interface.py",
        "src/sistema_principal.py",
        "src/ambiente_config.py",
        "src/version_control.py",
    ]
    
    for arquivo in arquivos_importantes:
        if Path(arquivo).exists():
            size = Path(arquivo).stat().st_size
            print(f"   ✅ {arquivo} ({size:,} bytes)")
        else:
            print(f"   ❌ {arquivo} - NÃO ENCONTRADO")
else:
    print(f"❌ Pasta src/ NÃO EXISTE!")

# 2. Testar imports
print("\n2. TESTE DE IMPORTS:")
print("-" * 70)

# Adicionar src ao path se necessário
if str(src_path.absolute()) not in sys.path:
    sys.path.insert(0, str(src_path.absolute()))
    print(f"➕ Adicionado ao path: {src_path.absolute()}")

# Testar imports diferentes
imports_testar = [
    ("relatorios_interface", "import relatorios_interface"),
    ("src.relatorios_interface", "from src import relatorios_interface"),
    ("ambiente_config", "import ambiente_config"),
    ("src.ambiente_config", "from src import ambiente_config"),
]

for nome, codigo in imports_testar:
    try:
        exec(codigo)
        print(f"   ✅ {nome} - OK")
    except ImportError as e:
        print(f"   ❌ {nome} - FALHOU: {e}")
    except Exception as e:
        print(f"   ⚠️ {nome} - ERRO: {e}")

# 3. Verificar conteúdo de relatorios_interface.py
print("\n3. CONTEÚDO DE relatorios_interface.py:")
print("-" * 70)

relatorios_path = Path("src/relatorios_interface.py")
if relatorios_path.exists():
    try:
        with open(relatorios_path, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
        
        print(f"   Total de linhas: {len(linhas)}")
        
        # Procurar pela classe SistemaRelatorios
        tem_classe = False
        for i, linha in enumerate(linhas, 1):
            if "class SistemaRelatorios" in linha:
                tem_classe = True
                print(f"   ✅ Classe SistemaRelatorios encontrada na linha {i}")
                print(f"      {linha.strip()}")
                break
        
        if not tem_classe:
            print(f"   ❌ Classe SistemaRelatorios NÃO ENCONTRADA!")
            
            # Listar classes que existem
            print("\n   Classes encontradas:")
            for i, linha in enumerate(linhas, 1):
                if linha.strip().startswith("class "):
                    print(f"      Linha {i}: {linha.strip()}")
        
        # Verificar imports
        print("\n   Primeiros imports do arquivo:")
        for i, linha in enumerate(linhas[:30], 1):
            if linha.strip().startswith("import ") or linha.strip().startswith("from "):
                print(f"      Linha {i}: {linha.strip()}")
                
    except Exception as e:
        print(f"   ❌ Erro ao ler arquivo: {e}")
else:
    print(f"   ❌ Arquivo NÃO EXISTE!")

# 4. Verificar __init__.py
print("\n4. CONTEÚDO DE src/__init__.py:")
print("-" * 70)

init_path = Path("src/__init__.py")
if init_path.exists():
    try:
        with open(init_path, 'r', encoding='utf-8') as f:
            conteudo = f.read()
        
        print(f"   Tamanho: {len(conteudo)} caracteres")
        
        if "relatorios_interface" in conteudo:
            print(f"   ✅ Referência a 'relatorios_interface' encontrada")
        else:
            print(f"   ⚠️ NÃO há referência a 'relatorios_interface'")
        
        print("\n   Conteúdo:")
        for i, linha in enumerate(conteudo.split('\n'), 1):
            if linha.strip() and not linha.strip().startswith('#'):
                print(f"      {i}: {linha}")
                
    except Exception as e:
        print(f"   ❌ Erro ao ler arquivo: {e}")
else:
    print(f"   ⚠️ Arquivo __init__.py NÃO EXISTE!")

# 5. Verificar se arquivo está corrompido
print("\n5. VERIFICAÇÃO DE INTEGRIDADE:")
print("-" * 70)

if relatorios_path.exists():
    try:
        # Tentar compilar o arquivo
        import py_compile
        py_compile.compile(str(relatorios_path), doraise=True)
        print(f"   ✅ Arquivo relatorios_interface.py é válido (compila sem erros)")
    except py_compile.PyCompileError as e:
        print(f"   ❌ ERRO DE SINTAXE no arquivo!")
        print(f"      {e}")
    except Exception as e:
        print(f"   ⚠️ Erro ao compilar: {e}")

# 6. Sugestões
print("\n6. RECOMENDAÇÕES:")
print("-" * 70)

problemas_encontrados = []

if not Path("src/__init__.py").exists():
    problemas_encontrados.append("Criar arquivo src/__init__.py")

if not Path("src/relatorios_interface.py").exists():
    problemas_encontrados.append("Arquivo relatorios_interface.py não existe!")
else:
    # Testar import real
    try:
        sys.path.insert(0, str(Path("src").absolute()))
        import relatorios_interface
        if not hasattr(relatorios_interface, 'SistemaRelatorios'):
            problemas_encontrados.append(
                "Arquivo existe mas não tem classe SistemaRelatorios"
            )
    except ImportError as e:
        problemas_encontrados.append(f"Import falha: {e}")

if problemas_encontrados:
    print("   ❌ PROBLEMAS ENCONTRADOS:")
    for i, problema in enumerate(problemas_encontrados, 1):
        print(f"      {i}. {problema}")
else:
    print("   ✅ Nenhum problema óbvio encontrado!")
    print("      O problema pode estar no PyInstaller.")

print("\n" + "=" * 70)
print("DIAGNÓSTICO CONCLUÍDO")
print("=" * 70)

input("\nPressione ENTER para fechar...")
