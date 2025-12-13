
import sys
import os

print("=== TESTE DE IMPORTS ===")

# Aplicar fix
from fix_imports import fix_imports
fix_imports()

print("\n=== TESTANDO IMPORTS ===")

# Testar módulos problemáticos
modulos_teste = [
    'relatorios_interface',
    'relatorio_despesas_aprimorado',
    'despesas_rateadas', 
    'gestao_medicoes',
    'configuracoes_sistema'
]

for modulo in modulos_teste:
    try:
        exec(f"import {modulo}")
        print(f"✅ {modulo}")
    except ImportError as e:
        print(f"❌ {modulo}: {e}")
        
        # Tentar versão com src
        try:
            exec(f"from src import {modulo}")
            print(f"✅ src.{modulo}")
        except ImportError as e2:
            print(f"❌ src.{modulo}: {e2}")

print("\n=== FIM DO TESTE ===")
