#!/usr/bin/env python3
"""
Correção rápida para o problema de módulos de relatório
Este script atualiza automaticamente o build com os módulos corretos
"""

import os
import sys
import subprocess
from pathlib import Path

def find_report_modules():
    """Encontra todos os módulos relacionados a relatórios"""
    report_modules = []
    
    # Procurar na pasta src
    src_path = Path("src")
    if src_path.exists():
        for py_file in src_path.glob("*.py"):
            filename = py_file.stem.lower()
            if any(term in filename for term in ['relatorio', 'relatório', 'report']):
                # Adicionar múltiplas variações
                base_name = py_file.stem
                report_modules.extend([
                    base_name,
                    f"src.{base_name}",
                ])
                print(f"✅ Encontrado módulo de relatório: {base_name}")
    
    # Procurar na raiz
    for py_file in Path(".").glob("*.py"):
        filename = py_file.stem.lower()
        if any(term in filename for term in ['relatorio', 'relatório', 'report']):
            base_name = py_file.stem
            report_modules.append(base_name)
            print(f"✅ Encontrado módulo de relatório na raiz: {base_name}")
    
    return report_modules

def update_build_script():
    """Atualiza o build_script.py com os módulos corretos"""
    
    print("🔍 Procurando módulos de relatório...")
    report_modules = find_report_modules()
    
    if not report_modules:
        print("⚠️  Nenhum módulo de relatório encontrado!")
        return False
    
    print(f"📊 Encontrados {len(report_modules)} módulos de relatório")
    
    # Ler o build_script atual
    build_script_path = "build_script.py"
    if not os.path.exists(build_script_path):
        print(f"❌ {build_script_path} não encontrado!")
        return False
    
    with open(build_script_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    # Procurar pela função get_hidden_imports
    start_marker = "def get_hidden_imports():"
    end_marker = "return hidden_imports"
    
    start_idx = content.find(start_marker)
    end_idx = content.find(end_marker, start_idx)
    
    if start_idx == -1 or end_idx == -1:
        print("❌ Não foi possível encontrar a função get_hidden_imports()")
        return False
    
    # Construir nova lista de módulos
    new_modules = []
    
    # Módulos básicos (manter os existentes)
    basic_modules = [
        # GUI
        "'tkinter'",
        "'tkinter.ttk'", 
        "'tkinter.messagebox'",
        "'tkinter.filedialog'",
        
        # Dados e relatórios
        "'pandas'",
        "'numpy'",
        "'openpyxl'",
        "'xlwings'",
        "'reportlab'",
        "'reportlab.pdfgen'",
        "'reportlab.lib'",
        "'matplotlib'",
        "'matplotlib.pyplot'",
        "'PIL'",
        "'PIL.Image'",
        
        # Específicos do sistema
        "'Sistema_Entrada_Dados'",
        "'src.Sistema_Entrada_Dados'",
        "'controle_pagamentos_taxas'",
        "'src.controle_pagamentos_taxas'",
        "'despesas_rateadas'",
        "'src.despesas_rateadas'",
        "'gestao_medicoes'",
        "'src.gestao_medicoes'",
    ]
    
    new_modules.extend(basic_modules)
    
    # Adicionar módulos de relatório encontrados
    for module in report_modules:
        new_modules.append(f"'{module}'")
    
    # Criar nova função
    new_function = f'''def get_hidden_imports():
    """Retorna lista de imports que podem não ser detectados automaticamente"""
    hidden_imports = [
        {chr(10).join(f"        {module}," for module in new_modules)}
    ]
    
    return hidden_imports'''
    
    # Substituir no conteúdo
    before = content[:start_idx]
    after = content[end_idx + len(end_marker):]
    new_content = before + new_function + after
    
    # Salvar arquivo atualizado
    with open(build_script_path, 'w', encoding='utf-8') as f:
        f.write(new_content)
    
    print("✅ build_script.py atualizado com sucesso!")
    return True

def rebuild_executable():
    """Reconstrói o executável com as correções"""
    print("\n🔨 Reconstruindo executável...")
    
    try:
        # Limpar build anterior
        if os.path.exists("build"):
            import shutil
            shutil.rmtree("build")
            print("🧹 Build anterior limpo")
        
        # Executar build
        result = subprocess.run([sys.executable, "build_script.py"], 
                              capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ Build concluído com sucesso!")
            
            # Verificar se executável foi criado
            exe_path = Path("dist/Sistema_Gestao_Financeira.exe")
            if exe_path.exists():
                print(f"📦 Executável atualizado: {exe_path.absolute()}")
                return True
            else:
                print("⚠️  Executável não encontrado após build")
                
        else:
            print("❌ Erro durante o build:")
            print(result.stderr)
            
    except Exception as e:
        print(f"❌ Erro ao executar build: {e}")
    
    return False

def main():
    print("=" * 60)
    print("🔧 CORREÇÃO RÁPIDA - Módulos de Relatório")
    print("=" * 60)
    
    # Verificar se estamos no diretório correto
    if not os.path.exists("src/sistema_principal.py"):
        print("❌ Execute este script no diretório raiz do projeto!")
        return
    
    print("📁 Diretório correto confirmado")
    
    # Passo 1: Atualizar build script
    if not update_build_script():
        print("❌ Falha ao atualizar build script")
        return
    
    # Passo 2: Perguntar se quer rebuildar
    response = input("\n🤔 Deseja rebuildar o executável agora? (s/n): ").lower()
    
    if response == 's':
        if rebuild_executable():
            print("\n🎉 CORREÇÃO CONCLUÍDA COM SUCESSO!")
            print("📋 Teste o executável abrindo diretamente 'Geração de Relatórios'")
        else:
            print("\n⚠️  Build falhou. Execute manualmente: python build_script.py")
    else:
        print("\n💡 Para rebuildar depois, execute: python build_script.py")
    
    print("\n" + "=" * 60)

if __name__ == "__main__":
    main()