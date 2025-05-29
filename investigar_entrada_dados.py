#!/usr/bin/env python3
"""
Investigar o que o módulo "Entrada de Dados" faz que permite os outros funcionarem
"""

import os
import sys
import importlib
import subprocess

def analisar_sistema_entrada_dados():
    """Analisa o módulo Sistema_Entrada_Dados para entender o que ele faz"""
    
    print("=" * 70)
    print("INVESTIGAÇÃO: POR QUE ENTRADA DE DADOS FAZ OS OUTROS FUNCIONAREM?")
    print("=" * 70)
    
    # Procurar o arquivo Sistema_Entrada_Dados
    arquivos_possveis = [
        "Sistema_Entrada_Dados.py",
        "src/Sistema_Entrada_Dados.py"
    ]
    
    arquivo_encontrado = None
    for arquivo in arquivos_possveis:
        if os.path.exists(arquivo):
            arquivo_encontrado = arquivo
            break
    
    if not arquivo_encontrado:
        print("❌ Arquivo Sistema_Entrada_Dados.py não encontrado!")
        return
    
    print(f"📄 Analisando: {arquivo_encontrado}")
    
    # Ler o conteúdo do arquivo
    with open(arquivo_encontrado, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    print(f"\n🔍 ANÁLISE DO CÓDIGO:")
    print("-" * 50)
    
    # Procurar por modificações no sys.path
    if 'sys.path' in conteudo:
        print("✅ ENCONTRADO: Modificações em sys.path")
        linhas_syspath = [linha.strip() for linha in conteudo.split('\n') if 'sys.path' in linha]
        for linha in linhas_syspath:
            print(f"   {linha}")
    else:
        print("❌ Não encontrou modificações em sys.path")
    
    # Procurar por add_project_root ou similar
    if 'add_project_root' in conteudo or 'project_root' in conteudo:
        print("✅ ENCONTRADO: Configuração de project_root")
        linhas_root = []
        linhas = conteudo.split('\n')
        for i, linha in enumerate(linhas):
            if 'project_root' in linha.lower() or 'add_project_root' in linha:
                # Pegar linha atual e algumas ao redor para contexto
                start = max(0, i-2)
                end = min(len(linhas), i+3)
                for j in range(start, end):
                    linhas_root.append(f"   {j+1:3d}: {linhas[j]}")
        
        for linha in linhas_root[:10]:  # Mostrar no máximo 10 linhas
            print(linha)
    else:
        print("❌ Não encontrou configuração de project_root")
    
    # Procurar por imports específicos
    imports_importantes = [
        'importlib',
        'Path',
        'pathlib',
        '__file__',
        'resolve',
        'parent'
    ]
    
    print(f"\n🔍 IMPORTS E CONFIGURAÇÕES IMPORTANTES:")
    print("-" * 50)
    
    for termo in imports_importantes:
        if termo in conteudo:
            print(f"✅ Contém: {termo}")
            # Encontrar linhas com esse termo
            linhas_termo = [linha.strip() for linha in conteudo.split('\n') 
                           if termo in linha and not linha.strip().startswith('#')]
            for linha in linhas_termo[:3]:  # Máximo 3 linhas por termo
                if linha:
                    print(f"   {linha}")
        else:
            print(f"❌ Não contém: {termo}")
    
    return conteudo

def extrair_funcao_path():
    """Extrai a função que configura os paths do Sistema_Entrada_Dados"""
    
    arquivo_entrada = None
    for arquivo in ["Sistema_Entrada_Dados.py", "src/Sistema_Entrada_Dados.py"]:
        if os.path.exists(arquivo):
            arquivo_entrada = arquivo
            break
    
    if not arquivo_entrada:
        return None
    
    with open(arquivo_entrada, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    # Procurar pela função add_project_root ou similar
    import re
    
    # Padrão para encontrar funções relacionadas a path
    pattern = r'def\s+(add_project_root|setup_path|configure_path|add_.*path).*?(?=\ndef|\nclass|\n\n\S|\Z)'
    
    funcoes_path = re.findall(pattern, conteudo, re.DOTALL | re.IGNORECASE)
    
    if funcoes_path:
        print(f"\n📋 FUNÇÕES DE CONFIGURAÇÃO DE PATH ENCONTRADAS:")
        print("-" * 50)
        for funcao in funcoes_path:
            print(f"Função: {funcao}")
    
    # Também procurar por qualquer código que modifica sys.path
    pattern_syspath = r'(.*sys\.path.*)'
    linhas_syspath = re.findall(pattern_syspath, conteudo)
    
    if linhas_syspath:
        print(f"\n📋 MODIFICAÇÕES EM SYS.PATH:")
        print("-" * 50)
        for linha in linhas_syspath:
            print(f"   {linha.strip()}")
    
    return conteudo

def criar_fix_temporario():
    """Cria um fix temporário baseado no que encontramos"""
    
    print(f"\n🔧 CRIANDO FIX TEMPORÁRIO:")
    print("-" * 50)
    
    # Código básico para adicionar paths
    fix_code = '''
def fix_imports():
    """Fix temporário para resolver imports - baseado no que Sistema_Entrada_Dados faz"""
    import sys
    import os
    from pathlib import Path
    
    print("Aplicando fix de imports...")
    
    # Adicionar diretório atual
    current_dir = Path(__file__).resolve().parent
    if str(current_dir) not in sys.path:
        sys.path.insert(0, str(current_dir))
        print(f"Adicionado ao path: {current_dir}")
    
    # Adicionar diretório pai (raiz do projeto)
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
        print(f"Adicionado ao path: {project_root}")
    
    # Adicionar src especificamente
    src_dir = project_root / "src"
    if src_dir.exists() and str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
        print(f"Adicionado ao path: {src_dir}")
    
    # Forçar reload de módulos problemáticos se já estiverem carregados
    modulos_problematicos = [
        'relatorios_interface',
        'relatorio_despesas_aprimorado', 
        'despesas_rateadas',
        'gestao_medicoes',
        'configuracoes_sistema'
    ]
    
    for modulo in modulos_problematicos:
        if modulo in sys.modules:
            print(f"Removendo {modulo} do cache para forçar reload")
            del sys.modules[modulo]
        
        # Também versões com src
        modulo_src = f"src.{modulo}"
        if modulo_src in sys.modules:
            print(f"Removendo {modulo_src} do cache para forçar reload")
            del sys.modules[modulo_src]

if __name__ == "__main__":
    fix_imports()
'''
    
    with open("fix_imports.py", "w", encoding="utf-8") as f:
        f.write(fix_code)
    
    print("📄 Arquivo fix_imports.py criado")
    
    return "fix_imports.py"

def testar_fix():
    """Testa se o fix resolve os imports"""
    
    print(f"\n🧪 TESTANDO FIX:")
    print("-" * 50)
    
    # Criar script de teste
    teste_code = '''
import sys
import os

print("=== TESTE DE IMPORTS ===")

# Aplicar fix
from fix_imports import fix_imports
fix_imports()

print("\\n=== TESTANDO IMPORTS ===")

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

print("\\n=== FIM DO TESTE ===")
'''
    
    with open("teste_fix.py", "w", encoding="utf-8") as f:
        f.write(teste_code)
    
    print("📄 Arquivo teste_fix.py criado")
    print("🚀 Execute: python teste_fix.py")

def main():
    print("Investigando por que 'Entrada de Dados' faz os outros módulos funcionarem...")
    
    # Analisar Sistema_Entrada_Dados
    conteudo = analisar_sistema_entrada_dados()
    
    if conteudo:
        # Extrair configurações de path
        extrair_funcao_path()
        
        # Criar fix temporário
        fix_file = criar_fix_temporario()
        
        # Criar teste
        testar_fix()
        
        print(f"\n" + "=" * 70)
        print("🔍 PRÓXIMOS PASSOS:")
        print("-" * 70)
        print("1. Execute: python teste_fix.py")
        print("2. Se funcionar, podemos aplicar o fix ao sistema principal")
        print("3. Ou modificar o sistema_principal.py para incluir o mesmo código")
        print(f"   que está em Sistema_Entrada_Dados.py")
        
    else:
        print("❌ Não foi possível analisar Sistema_Entrada_Dados.py")

if __name__ == "__main__":
    main()