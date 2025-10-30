#!/usr/bin/env python3
"""
Verificação rápida antes de cada build
"""

import os
import re

def verificar_novos_modulos():
    """Verifica se há novos módulos em src/"""
    
    if not os.path.exists("src"):
        return []
    
    # Módulos conhecidos que já estão no build_simples.py
    modulos_conhecidos = {
        'sistema_principal',
        'relatorios_interface', 
        'relatorio_despesas_aprimorado',
        'despesas_rateadas',
        'gestao_medicoes',
        'configuracoes_sistema',
        'Sistema_Entrada_Dados',
        'controle_pagamentos_taxas',
        'version_control'
    }
    
    # Encontrar todos os módulos em src/
    modulos_encontrados = set()
    for arquivo in os.listdir("src"):
        if arquivo.endswith('.py') and not arquivo.startswith('__'):
            modulo = arquivo[:-3]
            modulos_encontrados.add(modulo)
    
    # Módulos novos
    novos_modulos = modulos_encontrados - modulos_conhecidos
    
    return list(novos_modulos)

def verificar_imports_sistema_principal():
    """Verifica se todos os imports no sistema_principal.py usam src."""
    
    arquivo = "src/sistema_principal.py"
    
    if not os.path.exists(arquivo):
        return []
    
    with open(arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    # Procurar por reload_module sem src.
    problemas = []
    
    # Padrão para encontrar reload_module
    pattern = r"self\.reload_module\('([^']+)'\)"
    matches = re.findall(pattern, conteudo)
    
    for match in matches:
        if not match.startswith('src.') and match not in ['tkinter', 'os', 'sys']:
            # Verificar se o módulo existe em src/
            if os.path.exists(f"src/{match}.py"):
                problemas.append(f"reload_module('{match}') deveria ser reload_module('src.{match}')")
    
    # Procurar por imports diretos sem src.
    pattern_import = r"from ([a-zA-Z_][a-zA-Z0-9_]*) import"
    imports = re.findall(pattern_import, conteudo)
    
    for imp in imports:
        if not imp.startswith('src.') and imp not in ['tkinter', 'os', 'sys', 'datetime', 'pathlib']:
            if os.path.exists(f"src/{imp}.py"):
                problemas.append(f"from {imp} import deveria ser from src.{imp} import")
    
    return problemas

def main():
    print("🔍 VERIFICAÇÃO PRÉ-BUILD")
    print("=" * 40)
    
    # Verificar novos módulos
    novos_modulos = verificar_novos_modulos()
    
    if novos_modulos:
        print(f"⚠️  NOVOS MÓDULOS ENCONTRADOS:")
        for modulo in novos_modulos:
            print(f"   - {modulo}")
        print(f"\n💡 AÇÃO NECESSÁRIA:")
        print(f"   Adicione estes módulos ao build_simples.py:")
        for modulo in novos_modulos:
            print(f"   --hidden-import={modulo}")
            print(f"   --hidden-import=src.{modulo}")
    else:
        print("✅ Nenhum módulo novo encontrado")
    
    # Verificar imports
    problemas_import = verificar_imports_sistema_principal()
    
    if problemas_import:
        print(f"\n⚠️  PROBLEMAS DE IMPORT ENCONTRADOS:")
        for problema in problemas_import:
            print(f"   - {problema}")
        print(f"\n💡 AÇÃO NECESSÁRIA:")
        print(f"   Corrija os imports no src/sistema_principal.py")
    else:
        print("✅ Todos os imports estão corretos")
    
    # Conclusão
    if not novos_modulos and not problemas_import:
        print(f"\n🚀 TUDO OK! Pode executar:")
        print(f"   python build_simples.py")
    else:
        print(f"\n⚠️  CORRIJA OS PROBLEMAS ANTES DO BUILD")
    
    print("=" * 40)

if __name__ == "__main__":
    main()