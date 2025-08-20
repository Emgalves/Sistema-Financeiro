# -*- coding: utf-8 -*-
"""
CORREÇÃO SIMPLES PARA VSCODE
Salve como: corrigir_nfe_vscode.py (na raiz do projeto)
Execute: python corrigir_nfe_vscode.py
"""

import sys
import os
from pathlib import Path

def setup_system():
    """Configura o sistema para correção"""
    print("🔧 CORREÇÃO NFe/CERTIFICADO A1 - VSCODE")
    print("=" * 50)
    
    # 1. Configurar paths
    current_dir = Path(__file__).resolve().parent
    paths = [
        str(current_dir),
        str(current_dir / "src"),
        str(current_dir / "src" / "nfe")
    ]
    
    for path in paths:
        if os.path.exists(path) and path not in sys.path:
            sys.path.insert(0, path)
            print(f"➕ Path: {path}")
    
    # 2. Verificar arquivos essenciais
    arquivos_necessarios = [
        "src/nfe/correcao_certificado_a1.py",
        "src/Sistema_Entrada_Dados.py"
    ]
    
    print("\n📁 Verificando arquivos...")
    for arquivo in arquivos_necessarios:
        if os.path.exists(arquivo):
            print(f"✅ {arquivo}")
        else:
            print(f"❌ {arquivo} - NÃO ENCONTRADO!")
            return False
    
    return True

def aplicar_correcao():
    """Aplica a correção de certificado A1"""
    try:
        print("\n🚀 Aplicando correção...")
        
        # Importar Sistema_Entrada_Dados
        try:
            from src.Sistema_Entrada_Dados import SistemaEntradaDados
        except ImportError:
            from Sistema_Entrada_Dados import SistemaEntradaDados
        
        # Criar instância temporária (sem interface)
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Ocultar
        
        print("⚙️ Criando sistema temporário...")
        sistema = SistemaEntradaDados(parent=root)
        
        # Aplicar correção
        print("🔑 Carregando correção de certificado...")
        
        # Ler arquivo com codificação UTF-8
        with open('src/nfe/correcao_certificado_a1.py', 'r', encoding='utf-8') as f:
            codigo_correcao = f.read()
        
        # Executar código no namespace local para ter acesso às funções
        local_namespace = {}
        exec(codigo_correcao, globals(), local_namespace)
        
        # Agora as funções estão disponíveis
        if 'aplicar_correcao_automatica' in local_namespace:
            aplicar_correcao_automatica = local_namespace['aplicar_correcao_automatica']
            sucesso = aplicar_correcao_automatica(sistema)
        elif 'corrigir_sistema_certificado_a1' in local_namespace:
            corrigir_sistema_certificado_a1 = local_namespace['corrigir_sistema_certificado_a1']
            sucesso = corrigir_sistema_certificado_a1(sistema)
        else:
            # Tentar importação direta como fallback
            print("⚠️ Tentando importação direta...")
            try:
                from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
                sucesso = corrigir_sistema_certificado_a1(sistema)
            except ImportError:
                print("❌ Não foi possível encontrar função de correção")
                return False
        
        if sucesso:
            print("\n✅ CORREÇÃO APLICADA COM SUCESSO!")
            
            # Tornar disponível globalmente
            import builtins
            builtins.sistema_principal = sistema
            
            print("\n🎯 AGORA NO CONSOLE PYTHON DO VSCODE:")
            print(">>> sistema_principal.configurar_certificado_rapido()")
            print("\n🔍 Para diagnóstico:")
            print(">>> sistema_principal.diagnosticar_nfe()")
            
            # Instruções para VSCode
            print(f"""
📋 INSTRUÇÕES PARA VSCODE:

1. ABRA CONSOLE PYTHON:
   - Ctrl+Shift+P
   - Digite: "Python: Start REPL"
   - Pressione Enter

2. NO CONSOLE PYTHON QUE ABRIR, EXECUTE:
   >>> sistema_principal.configurar_certificado_rapido()

3. CONFIGURAR CERTIFICADO:
   - Na janela que abrir, clique "🔍 Procurar"
   - Selecione seu arquivo .pfx
   - Digite a senha/PIN (6 dígitos)
   - Clique "Configurar"

4. TESTAR:
   >>> sistema_principal.diagnosticar_nfe()
   >>> chave = "sua_chave_44_digitos"
   >>> dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)
            """)
            
            return True
        else:
            print("❌ Falha na aplicação da correção")
            return False
            
    except Exception as e:
        print(f"❌ Erro: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Função principal"""
    try:
        # Setup inicial
        if not setup_system():
            print("\n❌ Falha no setup inicial")
            return
        
        # Aplicar correção
        if aplicar_correcao():
            print("\n🎉 SUCESSO! Siga as instruções acima.")
            
            # Manter script ativo
            print("\n⏳ Script ativo (Ctrl+C para sair)")
            import time
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n👋 Finalizando...")
        else:
            print("\n❌ Falha na aplicação")
            
    except Exception as e:
        print(f"❌ Erro geral: {e}")

if __name__ == "__main__":
    main()
