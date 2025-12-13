#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para diagnosticar por que os módulos não estão sendo encontrados no executável
"""

import os
import sys
import importlib
import traceback

def testar_imports():
    """Testa se os módulos podem ser importados normalmente"""
    
    print("=" * 60)
    print("DIAGNÓSTICO DE IMPORTS")
    print("=" * 60)
    
    # Lista de módulos para testar
    modulos_teste = [
        'relatorios_interface',
        'src.relatorios_interface',
        'relatorio_despesas_aprimorado',
        'src.relatorio_despesas_aprimorado',
        'Sistema_Entrada_Dados',
        'src.Sistema_Entrada_Dados'
    ]
    
    resultados = {}
    
    for modulo in modulos_teste:
        print(f"\nTestando: {modulo}")
        try:
            # Tentar importar
            mod = importlib.import_module(modulo)
            print(f"✅ SUCESSO: {modulo}")
            
            # Verificar se tem as classes esperadas
            if 'relatorios_interface' in modulo:
                if hasattr(mod, 'SistemaRelatorios'):
                    print(f"  ✅ Classe SistemaRelatorios encontrada")
                else:
                    print(f"  ⚠️  Classe SistemaRelatorios NÃO encontrada")
            
            if 'relatorio_despesas_aprimorado' in modulo:
                if hasattr(mod, 'RelatorioUI'):
                    print(f"  ✅ Classe RelatorioUI encontrada")
                else:
                    print(f"  ⚠️  Classe RelatorioUI NÃO encontrada")
            
            resultados[modulo] = {'status': 'OK', 'modulo': mod}
            
        except ImportError as e:
            print(f"❌ ERRO: {modulo} - {str(e)}")
            resultados[modulo] = {'status': 'ERRO', 'erro': str(e)}
        except Exception as e:
            print(f"❌ ERRO INESPERADO: {modulo} - {str(e)}")
            resultados[modulo] = {'status': 'ERRO', 'erro': str(e)}
    
    return resultados

def verificar_caminhos():
    """Verifica os caminhos do Python"""
    
    print(f"\n" + "=" * 60)
    print("DIAGNÓSTICO DE CAMINHOS")
    print("=" * 60)
    
    print(f"Diretório atual: {os.getcwd()}")
    print(f"Python executável: {sys.executable}")
    print(f"Versão Python: {sys.version}")
    
    print(f"\nSys.path:")
    for i, caminho in enumerate(sys.path):
        print(f"  {i}: {caminho}")
    
    # Verificar se diretórios importantes existem
    print(f"\nDiretórios importantes:")
    dirs_verificar = ['.', 'src', 'dist', 'build']
    
    for dir_name in dirs_verificar:
        if os.path.exists(dir_name):
            arquivos = [f for f in os.listdir(dir_name) if f.endswith('.py')]
            print(f"✅ {dir_name}/ - {len(arquivos)} arquivos .py")
        else:
            print(f"❌ {dir_name}/ - NÃO EXISTE")

def simular_reload_module():
    """Simula o método reload_module do sistema principal"""
    
    print(f"\n" + "=" * 60)
    print("SIMULAÇÃO DO RELOAD_MODULE")
    print("=" * 60)
    
    def reload_module(module_name):
        """Simulação do método reload_module"""
        try:
            print(f"Tentando carregar: {module_name}")
            
            # Remover módulo se já estiver carregado
            for key in list(sys.modules.keys()):
                if key == module_name or key.startswith(f"{module_name}."):
                    del sys.modules[key]
                    print(f"  Removido do cache: {key}")
            
            # Tentar importar
            module = importlib.import_module(module_name)
            print(f"  ✅ Sucesso: {module_name}")
            return module
            
        except Exception as e:
            print(f"  ❌ Erro: {module_name} - {str(e)}")
            return None
    
    # Testar os módulos problemáticos
    modulos_testar = ['relatorios_interface', 'relatorio_despesas_aprimorado']
    
    for modulo in modulos_testar:
        print(f"\n--- Testando {modulo} ---")
        resultado = reload_module(modulo)
        
        if resultado:
            print(f"Módulo carregado: {resultado}")
            print(f"Arquivo: {getattr(resultado, '__file__', 'N/A')}")
            
            # Verificar classes específicas
            if modulo == 'relatorios_interface':
                if hasattr(resultado, 'SistemaRelatorios'):
                    print(f"Classe SistemaRelatorios: OK")
                else:
                    print(f"Classe SistemaRelatorios: NÃO ENCONTRADA")
                    print(f"Atributos disponíveis: {[attr for attr in dir(resultado) if not attr.startswith('_')]}")
            
            if modulo == 'relatorio_despesas_aprimorado':
                if hasattr(resultado, 'RelatorioUI'):
                    print(f"Classe RelatorioUI: OK")
                else:
                    print(f"Classe RelatorioUI: NÃO ENCONTRADA")
                    print(f"Atributos disponíveis: {[attr for attr in dir(resultado) if not attr.startswith('_')]}")

def verificar_executavel():
    """Verifica se estamos rodando de um executável PyInstaller"""
    
    print(f"\n" + "=" * 60)
    print("VERIFICAÇÃO DE AMBIENTE")
    print("=" * 60)
    
    # Verificar se é PyInstaller
    if hasattr(sys, '_MEIPASS'):
        print(f"✅ Executando de PyInstaller")
        print(f"Diretório temporário: {sys._MEIPASS}")
        
        # Listar arquivos no diretório temporário
        try:
            temp_files = os.listdir(sys._MEIPASS)
            print(f"Arquivos no diretório temporário: {len(temp_files)}")
            
            # Procurar por nossos módulos
            modulos_procurar = ['relatorios_interface', 'relatorio_despesas_aprimorado']
            for modulo in modulos_procurar:
                arquivos_modulo = [f for f in temp_files if modulo in f]
                if arquivos_modulo:
                    print(f"  ✅ {modulo}: {arquivos_modulo}")
                else:
                    print(f"  ❌ {modulo}: NÃO ENCONTRADO")
                    
        except Exception as e:
            print(f"Erro ao listar arquivos temporários: {str(e)}")
            
    else:
        print(f"⚠️  Executando em modo desenvolvimento")
        print(f"Diretório de execução: {os.getcwd()}")

def gerar_relatorio_diagnostico():
    """Gera um relatório completo do diagnóstico"""
    
    print(f"\n" + "=" * 60)
    print("RELATÓRIO DE DIAGNÓSTICO")
    print("=" * 60)
    
    # Executar todos os diagnósticos
    resultados_import = testar_imports()
    verificar_caminhos()
    simular_reload_module()
    verificar_executavel()
    
    # Gerar arquivo de log
    with open("diagnostico_executavel.log", "w", encoding="utf-8") as f:
        f.write("DIAGNÓSTICO DO EXECUTÁVEL\n")
        f.write("=" * 60 + "\n\n")
        
        f.write("RESULTADOS DOS IMPORTS:\n")
        for modulo, resultado in resultados_import.items():
            f.write(f"{modulo}: {resultado['status']}\n")
            if resultado['status'] == 'ERRO':
                f.write(f"  Erro: {resultado['erro']}\n")
        
        f.write(f"\nDiretório atual: {os.getcwd()}\n")
        f.write(f"Python: {sys.executable}\n")
        f.write(f"Versão: {sys.version}\n")
        
        if hasattr(sys, '_MEIPASS'):
            f.write(f"PyInstaller: SIM\n")
            f.write(f"Diretório temporário: {sys._MEIPASS}\n")
        else:
            f.write(f"PyInstaller: NÃO\n")
    
    print(f"Relatório salvo em: diagnostico_executavel.log")
    
    # Sugestões
    print(f"\n" + "=" * 60)
    print("SUGESTÕES DE CORREÇÃO")
    print("=" * 60)
    
    problemas_encontrados = [mod for mod, res in resultados_import.items() 
                           if res['status'] == 'ERRO' and 'relatorio' in mod]
    
    if problemas_encontrados:
        print(f"❌ Módulos com problema: {problemas_encontrados}")
        print(f"\nSOLUÇÕES:")
        print(f"1. Execute: python build_final.py")
        print(f"2. O novo build incluirá hooks específicos para estes módulos")
        print(f"3. Teste novamente o executável")
    else:
        print(f"✅ Todos os módulos estão sendo importados corretamente")
        print(f"O problema pode estar em outro lugar.")

def main():
    try:
        gerar_relatorio_diagnostico()
    except Exception as e:
        print(f"Erro durante diagnóstico: {str(e)}")
        traceback.print_exc()

if __name__ == "__main__":
    main()