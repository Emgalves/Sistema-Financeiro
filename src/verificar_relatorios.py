#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Verificação - Sistema de Relatórios
Executa diagnóstico completo para identificar problemas com relatorios_interface.py
"""

import sys
import os
from pathlib import Path
import importlib

def adicionar_paths():
    """Adiciona todos os paths possíveis ao sys.path"""
    # Obter diretório base
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys._MEIPASS)
        print(f"🏗️ Executável PyInstaller detectado: {base_dir}")
    else:
        base_dir = Path(__file__).resolve().parent
        print(f"🐍 Execução Python normal: {base_dir}")
    
    # Paths para adicionar
    paths_adicionar = [
        str(base_dir),
        str(base_dir / "src"),
        str(base_dir.parent),
        str(base_dir.parent / "src")
    ]
    
    print("\n📁 Adicionando paths ao sys.path:")
    for path in paths_adicionar:
        if os.path.exists(path) and path not in sys.path:
            sys.path.insert(0, path)
            print(f"  ✅ {path}")
        elif os.path.exists(path):
            print(f"  ⚠️ {path} (já existe)")
        else:
            print(f"  ❌ {path} (não existe)")

def buscar_arquivos():
    """Busca por arquivos relevantes do sistema de relatórios"""
    print("\n🔍 Buscando arquivos relevantes...")
    
    # Obter diretório base
    if getattr(sys, 'frozen', False):
        base_dir = Path(sys._MEIPASS)
    else:
        base_dir = Path(__file__).resolve().parent
    
    arquivos_procurados = [
        'relatorios_interface.py',
        'relatorio_despesas_service.py', 
        'relatorio_despesas_aprimorado.py'
    ]
    
    diretorios_busca = [
        base_dir,
        base_dir / "src",
        base_dir.parent,
        base_dir.parent / "src"
    ]
    
    arquivos_encontrados = {}
    
    for diretorio in diretorios_busca:
        if diretorio.exists():
            print(f"\n📂 Verificando: {diretorio}")
            for arquivo in arquivos_procurados:
                caminho_arquivo = diretorio / arquivo
                if caminho_arquivo.exists():
                    print(f"  ✅ {arquivo}")
                    if arquivo not in arquivos_encontrados:
                        arquivos_encontrados[arquivo] = []
                    arquivos_encontrados[arquivo].append(str(caminho_arquivo))
                else:
                    print(f"  ❌ {arquivo}")
        else:
            print(f"\n📂 ❌ NÃO EXISTE: {diretorio}")
    
    return arquivos_encontrados

def verificar_importacoes(arquivos_encontrados):
    """Tenta importar os módulos encontrados"""
    print("\n🔄 Testando importações...")
    
    # Limpar cache primeiro
    modulos_limpar = [
        'relatorios_interface',
        'src.relatorios_interface',
        'relatorio_despesas_service',
        'src.relatorio_despesas_service'
    ]
    
    for modulo in modulos_limpar:
        if modulo in sys.modules:
            del sys.modules[modulo]
            print(f"🗑️ Cache limpo: {modulo}")
    
    # Invalidar caches
    importlib.invalidate_caches()
    print("🔄 Cache do importlib invalidado")
    
    # Tentar importações
    resultados = {}
    
    # Testar relatorios_interface
    if 'relatorios_interface.py' in arquivos_encontrados:
        print("\n🎯 Testando relatorios_interface:")
        
        for caminho_tentativa in ['relatorios_interface', 'src.relatorios_interface']:
            try:
                print(f"  🔄 Tentando: {caminho_tentativa}")
                modulo = importlib.import_module(caminho_tentativa)
                print(f"  ✅ SUCESSO: {caminho_tentativa} importado!")
                
                # Verificar classes disponíveis
                classes = [name for name in dir(modulo) if not name.startswith('_') and name[0].isupper()]
                print(f"  📋 Classes encontradas: {classes}")
                
                # Verificar especificamente SistemaRelatorios
                if hasattr(modulo, 'SistemaRelatorios'):
                    print(f"  🎉 CLASSE SistemaRelatorios ENCONTRADA!")
                    resultados['relatorios_interface'] = {
                        'modulo': caminho_tentativa,
                        'sucesso': True,
                        'tem_sistema_relatorios': True
                    }
                else:
                    print(f"  ⚠️ Módulo importado mas SEM classe SistemaRelatorios")
                    resultados['relatorios_interface'] = {
                        'modulo': caminho_tentativa,
                        'sucesso': True,
                        'tem_sistema_relatorios': False
                    }
                break
                
            except ImportError as e:
                print(f"  ❌ Falha: {caminho_tentativa} - {str(e)}")
                continue
            except Exception as e:
                print(f"  💥 Erro inesperado: {caminho_tentativa} - {str(e)}")
                continue
    else:
        print("\n❌ relatorios_interface.py NÃO ENCONTRADO!")
        resultados['relatorios_interface'] = {
            'modulo': None,
            'sucesso': False,
            'tem_sistema_relatorios': False
        }
    
    # Testar relatorio_despesas_service
    if 'relatorio_despesas_service.py' in arquivos_encontrados:
        print("\n🎯 Testando relatorio_despesas_service:")
        
        for caminho_tentativa in ['relatorio_despesas_service', 'src.relatorio_despesas_service']:
            try:
                print(f"  🔄 Tentando: {caminho_tentativa}")
                modulo = importlib.import_module(caminho_tentativa)
                print(f"  ✅ SUCESSO: {caminho_tentativa} importado!")
                
                # Verificar classes
                classes = [name for name in dir(modulo) if not name.startswith('_') and name[0].isupper()]
                print(f"  📋 Classes encontradas: {classes}")
                
                if hasattr(modulo, 'RelatoriosDespesasService'):
                    print(f"  🎉 CLASSE RelatoriosDespesasService ENCONTRADA!")
                
                resultados['relatorio_despesas_service'] = {
                    'modulo': caminho_tentativa,
                    'sucesso': True
                }
                break
                
            except ImportError as e:
                print(f"  ❌ Falha: {caminho_tentativa} - {str(e)}")
                continue
            except Exception as e:
                print(f"  💥 Erro inesperado: {caminho_tentativa} - {str(e)}")
                continue
    
    return resultados

def gerar_relatorio_final(arquivos_encontrados, resultados_importacao):
    """Gera relatório final com conclusões"""
    print("\n" + "="*80)
    print("📊 RELATÓRIO FINAL")
    print("="*80)
    
    # Status dos arquivos
    print("\n📄 STATUS DOS ARQUIVOS:")
    arquivos_importantes = ['relatorios_interface.py', 'relatorio_despesas_service.py', 'relatorio_despesas_aprimorado.py']
    
    for arquivo in arquivos_importantes:
        if arquivo in arquivos_encontrados:
            print(f"  ✅ {arquivo}: {len(arquivos_encontrados[arquivo])} localização(ões)")
            for local in arquivos_encontrados[arquivo]:
                print(f"     📍 {local}")
        else:
            print(f"  ❌ {arquivo}: NÃO ENCONTRADO")
    
    # Status das importações
    print("\n🔄 STATUS DAS IMPORTAÇÕES:")
    for modulo, resultado in resultados_importacao.items():
        if resultado['sucesso']:
            print(f"  ✅ {modulo}: SUCESSO ({resultado['modulo']})")
            if 'tem_sistema_relatorios' in resultado:
                if resultado['tem_sistema_relatorios']:
                    print(f"     🎉 Classe SistemaRelatorios: DISPONÍVEL")
                else:
                    print(f"     ⚠️ Classe SistemaRelatorios: NÃO ENCONTRADA")
        else:
            print(f"  ❌ {modulo}: FALHA")
    
    # Diagnóstico e recomendações
    print("\n🔍 DIAGNÓSTICO:")
    
    # Verificar situação do relatorios_interface
    if 'relatorios_interface' in resultados_importacao:
        resultado_ri = resultados_importacao['relatorios_interface']
        
        if resultado_ri['sucesso'] and resultado_ri['tem_sistema_relatorios']:
            print("  🎉 SISTEMA INTEGRADO DISPONÍVEL!")
            print("     O relatorios_interface.py está funcionando corretamente.")
            
        elif resultado_ri['sucesso'] and not resultado_ri['tem_sistema_relatorios']:
            print("  ⚠️ PROBLEMA PARCIAL:")
            print("     O relatorios_interface.py pode ser importado, mas não tem a classe SistemaRelatorios.")
            print("     Possível problema no código do arquivo.")
            
        else:
            print("  ❌ PROBLEMA CRÍTICO:")
            print("     O relatorios_interface.py não pode ser importado.")
    
    # Verificar fallback
    if 'relatorio_despesas_aprimorado.py' in arquivos_encontrados:
        print("  📋 FALLBACK DISPONÍVEL:")
        print("     O sistema básico de despesas está disponível como alternativa.")
    else:
        print("  ❌ SEM FALLBACK:")
        print("     Nem o sistema básico de despesas está disponível.")
    
    print("\n🔧 RECOMENDAÇÕES:")
    
    # Recomendações baseadas no diagnóstico
    if 'relatorios_interface' not in resultados_importacao or not resultados_importacao['relatorios_interface']['sucesso']:
        print("  1. ❗ CRÍTICO: relatorios_interface.py não encontrado ou não importável")
        print("     - Verificar se o arquivo existe no projeto")
        print("     - Verificar se há erros de sintaxe no arquivo")
        print("     - Recompilar executável incluindo todos os arquivos")
        
    elif not resultados_importacao['relatorios_interface']['tem_sistema_relatorios']:
        print("  1. ⚠️ IMPORTANTE: Classe SistemaRelatorios não encontrada")
        print("     - Verificar se a classe está definida no arquivo")
        print("     - Verificar se há erros na definição da classe")
        
    else:
        print("  1. ✅ Sistema integrado está funcionando!")
        print("     - O problema pode estar na lógica de carregamento do sistema_principal.py")
    
    if 'relatorio_despesas_service.py' not in arquivos_encontrados:
        print("  2. ⚠️ relatorio_despesas_service.py não encontrado")
        print("     - Este arquivo é necessário para o sistema integrado funcionar")
    
    print("\n🚀 PRÓXIMOS PASSOS:")
    if 'relatorios_interface' in resultados_importacao and resultados_importacao['relatorios_interface']['sucesso'] and resultados_importacao['relatorios_interface']['tem_sistema_relatorios']:
        print("  ✅ O sistema integrado deve funcionar!")
        print("  💡 Se ainda há erro, o problema está no sistema_principal.py")
        print("  🔧 Aplicar a correção fornecida no código do sistema_principal.py")
    else:
        print("  ❌ Sistema integrado não funcionará")
        print("  🔧 Verificar/corrigir os arquivos conforme recomendações acima")
        print("  📦 Recompilar executável após correções")

def main():
    """Função principal do script de verificação"""
    print("🔍 SCRIPT DE VERIFICAÇÃO - SISTEMA DE RELATÓRIOS")
    print("="*60)
    print("Este script diagnostica problemas com o relatorios_interface.py")
    print()
    
    try:
        # Passo 1: Adicionar paths
        adicionar_paths()
        
        # Passo 2: Buscar arquivos
        arquivos_encontrados = buscar_arquivos()
        
        # Passo 3: Verificar importações
        resultados_importacao = verificar_importacoes(arquivos_encontrados)
        
        # Passo 4: Gerar relatório final
        gerar_relatorio_final(arquivos_encontrados, resultados_importacao)
        
        print("\n" + "="*80)
        print("✅ VERIFICAÇÃO CONCLUÍDA")
        print("="*80)
        
    except Exception as e:
        print(f"\n💥 ERRO DURANTE VERIFICAÇÃO: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()