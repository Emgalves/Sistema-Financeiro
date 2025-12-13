
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagnóstico - Como o sistema determina os caminhos de dados
"""

import os
import sys
from pathlib import Path

print("=" * 80)
print("DIAGNÓSTICO - CAMINHOS DE DADOS")
print("=" * 80)

# 1. Verificar ambiente detectado
print("\n1. AMBIENTE DETECTADO:")
print("-" * 80)
try:
    from src.ambiente_config import config_ambiente
    print(f"   Ambiente: {config_ambiente.get_nome_ambiente()}")
    print(f"   É produção? {config_ambiente.eh_producao()}")
    print(f"   É teste? {config_ambiente.eh_teste()}")
except Exception as e:
    print(f"   ❌ Erro ao importar ambiente_config: {e}")

# 2. Verificar variáveis de ambiente
print("\n2. VARIÁVEIS DE AMBIENTE (.env):")
print("-" * 80)
from dotenv import load_dotenv
load_dotenv()

variaveis_importantes = [
    'AMBIENTE_SISTEMA',
    'SISTEMA_AMBIENTE',
    'PASTA_DADOS',
    'PASTA_DADOS_TESTE',
    'PASTA_DADOS_PRODUCAO',
    'BASE_DIR',
    'DATA_DIR',
    'DATABASE_PATH',
]

for var in variaveis_importantes:
    valor = os.getenv(var)
    if valor:
        print(f"   {var} = {valor}")

# 3. Procurar arquivos de configuração
print("\n3. ARQUIVOS DE CONFIGURAÇÃO:")
print("-" * 80)

arquivos_config = [
    'config/database.py',
    'config/paths.py',
    'config/settings.py',
    'src/config/database.py',
    'src/config/paths.py',
    'database_config.py',
    'paths_config.py',
]

for arquivo in arquivos_config:
    if Path(arquivo).exists():
        print(f"   ✅ {arquivo} - EXISTE")
        
        # Ler conteúdo
        try:
            with open(arquivo, 'r', encoding='utf-8') as f:
                conteudo = f.read()
            
            # Procurar por caminhos
            if 'sistema_gestao_testes' in conteudo.lower():
                print(f"      🔍 Contém referência a 'sistema_gestao_testes'")
            if 'shortcut-targets' in conteudo.lower():
                print(f"      🔍 Contém referência a 'shortcut-targets'")
            if 'C:\\Users\\Obras' in conteudo or 'C:/Users/Obras' in conteudo:
                print(f"      🔍 Contém caminho C:\\Users\\Obras")
            if 'H:' in conteudo:
                print(f"      🔍 Contém caminho H:")
                
        except Exception as e:
            print(f"      ⚠️ Erro ao ler: {e}")
    else:
        print(f"   ❌ {arquivo} - NÃO EXISTE")

# 4. Procurar em arquivos Python principais
print("\n4. CAMINHOS HARDCODED EM ARQUIVOS .PY:")
print("-" * 80)

arquivos_principais = [
    'src/sistema_principal.py',
    'src/Sistema_Entrada_Dados.py',
    'src/relatorios_interface.py',
    'src/ambiente_config.py',
]

for arquivo in arquivos_principais:
    if not Path(arquivo).exists():
        continue
    
    try:
        with open(arquivo, 'r', encoding='utf-8') as f:
            linhas = f.readlines()
        
        # Procurar por caminhos
        caminhos_encontrados = []
        for i, linha in enumerate(linhas, 1):
            linha_lower = linha.lower()
            
            if 'sistema_gestao_testes' in linha_lower:
                caminhos_encontrados.append(f"   Linha {i}: {linha.strip()}")
            elif 'shortcut-targets' in linha_lower:
                caminhos_encontrados.append(f"   Linha {i}: {linha.strip()}")
            elif ('c:\\users\\obras' in linha_lower or 'c:/users/obras' in linha_lower):
                caminhos_encontrados.append(f"   Linha {i}: {linha.strip()}")
            elif 'h:' in linha_lower and ('financeiro' in linha_lower or 'relatorio' in linha_lower):
                caminhos_encontrados.append(f"   Linha {i}: {linha.strip()}")
        
        if caminhos_encontrados:
            print(f"\n   📁 {arquivo}:")
            for caminho in caminhos_encontrados[:5]:  # Primeiros 5
                print(caminho)
                
    except Exception as e:
        print(f"   ⚠️ Erro ao ler {arquivo}: {e}")

# 5. Verificar se há função get_pasta_dados ou similar
print("\n5. FUNÇÕES DE CONFIGURAÇÃO DE CAMINHOS:")
print("-" * 80)

funcoes_procurar = [
    'get_pasta_dados',
    'get_data_path',
    'get_database_path',
    'obter_caminho',
    'configurar_paths',
]

for arquivo in ['src/ambiente_config.py', 'src/sistema_principal.py']:
    if not Path(arquivo).exists():
        continue
        
    try:
        with open(arquivo, 'r', encoding='utf-8') as f:
            conteudo = f.read()
        
        for funcao in funcoes_procurar:
            if funcao in conteudo:
                print(f"   ✅ Função '{funcao}' encontrada em {arquivo}")
                
                # Tentar encontrar a definição
                linhas = conteudo.split('\n')
                for i, linha in enumerate(linhas):
                    if f"def {funcao}" in linha:
                        # Mostrar definição + próximas 10 linhas
                        print(f"\n   Definição (linha {i+1}):")
                        for j in range(i, min(i+10, len(linhas))):
                            print(f"      {linhas[j]}")
                        break
                        
    except Exception as e:
        print(f"   ⚠️ Erro ao processar {arquivo}: {e}")

# 6. Verificar banco de dados SQLite
print("\n6. BANCOS DE DADOS ENCONTRADOS:")
print("-" * 80)

# Procurar arquivos .db
import glob
bancos = glob.glob("**/*.db", recursive=True)
for banco in bancos[:10]:  # Primeiros 10
    tamanho = Path(banco).stat().st_size / 1024  # KB
    print(f"   📊 {banco} ({tamanho:.1f} KB)")

print("\n" + "=" * 80)
print("DIAGNÓSTICO CONCLUÍDO")
print("=" * 80)
print("\n⚠️ IMPORTANTE:")
print("   Para resolver o problema, precisamos:")
print("   1. Identificar ONDE os caminhos são definidos")
print("   2. Fazer esses caminhos mudarem baseado no ambiente")
print("=" * 80)

input("\nPressione ENTER para fechar...")
