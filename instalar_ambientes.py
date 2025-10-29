#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script de Instalação - Sistema de Diferenciação de Ambientes
Configura automaticamente o sistema para trabalhar com TESTE e PRODUÇÃO
"""

import os
import shutil
from pathlib import Path

def exibir_banner():
    """Exibe banner de boas-vindas"""
    print("=" * 70)
    print("  🎨 INSTALADOR - SISTEMA DE DIFERENCIAÇÃO DE AMBIENTES")
    print("  Sistema de Gestão Financeira")
    print("=" * 70)
    print()

def verificar_estrutura():
    """Verifica se está no diretório correto"""
    print("📁 Verificando estrutura do projeto...")
    
    if not os.path.exists("src"):
        print("❌ ERRO: Pasta 'src' não encontrada!")
        print("   Execute este script no diretório raiz do projeto.")
        return False
    
    if not os.path.exists("src/sistema_principal.py"):
        print("❌ ERRO: Arquivo 'src/sistema_principal.py' não encontrado!")
        return False
    
    print("✓ Estrutura do projeto OK")
    return True

def criar_backup():
    """Cria backup do sistema_principal.py original"""
    print("\n💾 Criando backup...")
    
    origem = "src/sistema_principal.py"
    destino = "src/sistema_principal.py.backup"
    
    if os.path.exists(destino):
        resposta = input("   Backup já existe. Sobrescrever? (s/n): ")
        if resposta.lower() != 's':
            print("   Backup mantido.")
            return True
    
    try:
        shutil.copy2(origem, destino)
        print(f"✓ Backup criado: {destino}")
        return True
    except Exception as e:
        print(f"❌ Erro ao criar backup: {e}")
        return False

def copiar_ambiente_config():
    """Copia ambiente_config.py para as pastas necessárias"""
    print("\n📄 Instalando módulo ambiente_config.py...")
    
    # Verificar se arquivo existe
    if not os.path.exists("ambiente_config.py"):
        print("❌ ERRO: ambiente_config.py não encontrado!")
        print("   Certifique-se de ter o arquivo na raiz do projeto.")
        return False
    
    # Copiar para src/
    try:
        shutil.copy2("ambiente_config.py", "src/ambiente_config.py")
        print("✓ Copiado para src/ambiente_config.py")
    except Exception as e:
        print(f"❌ Erro ao copiar: {e}")
        return False
    
    return True

def criar_arquivo_env():
    """Cria arquivo .env com configuração padrão"""
    print("\n⚙️  Configurando arquivo .env...")
    
    if os.path.exists(".env"):
        print("   Arquivo .env já existe.")
        resposta = input("   Deseja recriar? (s/n): ")
        if resposta.lower() != 's':
            print("   .env mantido.")
            return True
    
    conteudo_env = """# Configuração de Ambiente - Sistema de Gestão Financeira
# =========================================================

# AMBIENTE_SISTEMA: Define o ambiente de execução
# Valores possíveis: TESTE ou PRODUCAO
#
# TESTE: Interface amarela com avisos (seguro para experimentar)
# PRODUCAO: Interface padrão (dados reais)

AMBIENTE_SISTEMA=TESTE

# Outras configurações podem ser adicionadas abaixo
"""
    
    try:
        with open(".env", "w", encoding="utf-8") as f:
            f.write(conteudo_env)
        print("✓ Arquivo .env criado (Ambiente: TESTE)")
        return True
    except Exception as e:
        print(f"❌ Erro ao criar .env: {e}")
        return False

def atualizar_sistema_principal():
    """Atualiza sistema_principal.py com suporte a ambientes"""
    print("\n🔄 Atualizando sistema_principal.py...")
    
    resposta = input("   Deseja substituir sistema_principal.py? (s/n): ")
    if resposta.lower() != 's':
        print("   Sistema principal não foi alterado.")
        print("   ⚠️  Você precisará integrar manualmente as mudanças!")
        return True
    
    if not os.path.exists("sistema_principal_com_ambiente.py"):
        print("❌ Arquivo sistema_principal_com_ambiente.py não encontrado!")
        return False
    
    try:
        shutil.copy2("sistema_principal_com_ambiente.py", 
                     "src/sistema_principal.py")
        print("✓ sistema_principal.py atualizado")
        return True
    except Exception as e:
        print(f"❌ Erro ao atualizar: {e}")
        return False

def verificar_dependencias():
    """Verifica se as dependências necessárias estão instaladas"""
    print("\n📦 Verificando dependências...")
    
    dependencias = {
        'dotenv': 'python-dotenv',
        'tkinter': 'tk (geralmente já vem com Python)',
        'PIL': 'Pillow'
    }
    
    faltando = []
    
    for modulo, pacote in dependencias.items():
        try:
            if modulo == 'dotenv':
                __import__('dotenv')
            else:
                __import__(modulo)
            print(f"✓ {modulo} instalado")
        except ImportError:
            print(f"❌ {modulo} não encontrado")
            faltando.append(pacote)
    
    if faltando:
        print(f"\n⚠️  Dependências faltando: {', '.join(faltando)}")
        print("\nInstale com:")
        for pacote in faltando:
            if pacote != 'tk (geralmente já vem com Python)':
                print(f"  pip install {pacote}")
        return False
    
    print("✓ Todas as dependências instaladas")
    return True

def criar_atalhos_exemplo():
    """Cria exemplos de atalhos para Windows"""
    print("\n🔗 Criando exemplos de atalhos...")
    
    conteudo_bat_teste = """@echo off
REM Atalho para Ambiente de TESTE
echo Iniciando Sistema em modo TESTE...
set AMBIENTE_SISTEMA=TESTE
python sistema_principal.py
pause
"""
    
    conteudo_bat_prod = """@echo off
REM Atalho para Ambiente de PRODUCAO
echo Iniciando Sistema em modo PRODUCAO...
set AMBIENTE_SISTEMA=PRODUCAO
python sistema_principal.py
pause
"""
    
    try:
        with open("executar_TESTE.bat", "w") as f:
            f.write(conteudo_bat_teste)
        print("✓ executar_TESTE.bat criado")
        
        with open("executar_PRODUCAO.bat", "w") as f:
            f.write(conteudo_bat_prod)
        print("✓ executar_PRODUCAO.bat criado")
        
        print("\n   💡 Dica: Use estes arquivos .bat para iniciar rapidamente")
        print("   cada ambiente no Windows!")
        return True
    except Exception as e:
        print(f"❌ Erro ao criar atalhos: {e}")
        return False

def exibir_resumo():
    """Exibe resumo da instalação"""
    print("\n" + "=" * 70)
    print("  ✅ INSTALAÇÃO CONCLUÍDA!")
    print("=" * 70)
    print("\n📋 PRÓXIMOS PASSOS:\n")
    print("1. Verifique o arquivo .env na raiz do projeto")
    print("   • Atualmente configurado para: TESTE")
    print("   • Para produção, altere para: AMBIENTE_SISTEMA=PRODUCAO")
    print()
    print("2. Execute o sistema:")
    print("   • Windows: executar_TESTE.bat ou executar_PRODUCAO.bat")
    print("   • Ou: python src/sistema_principal.py")
    print()
    print("3. Teste a diferenciação visual:")
    print("   • TESTE: Tela amarela com banner laranja")
    print("   • PRODUÇÃO: Tela cinza padrão")
    print()
    print("4. Consulte a documentação:")
    print("   • GUIA_AMBIENTES.md - Guia completo")
    print("   • COMPARACAO_VISUAL_AMBIENTES.md - Diferenças visuais")
    print()
    print("5. Para gerar executáveis:")
    print("   • Use: python build_com_ambientes.py")
    print("   • Escolha gerar para TESTE, PRODUÇÃO ou ambos")
    print()
    print("=" * 70)
    print("\n🎉 Sistema pronto para uso com diferenciação de ambientes!")
    print("   Consulte GUIA_AMBIENTES.md para mais informações.\n")

def main():
    """Função principal do instalador"""
    exibir_banner()
    
    # Verificações e instalações
    etapas = [
        ("Verificar estrutura", verificar_estrutura),
        ("Criar backup", criar_backup),
        ("Copiar módulo", copiar_ambiente_config),
        ("Criar .env", criar_arquivo_env),
        ("Atualizar sistema", atualizar_sistema_principal),
        ("Verificar dependências", verificar_dependencias),
        ("Criar atalhos", criar_atalhos_exemplo)
    ]
    
    for nome, funcao in etapas:
        resultado = funcao()
        if not resultado:
            print(f"\n❌ Falha em: {nome}")
            print("   Corrija o erro e execute novamente.")
            return False
    
    # Sucesso!
    exibir_resumo()
    return True

if __name__ == "__main__":
    try:
        sucesso = main()
        if sucesso:
            input("\nPressione ENTER para sair...")
    except KeyboardInterrupt:
        print("\n\n❌ Instalação cancelada pelo usuário.")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
