# -*- coding: utf-8 -*-
"""
CORREÇÃO ROBUSTA PARA WINDOWS - NFe/CERTIFICADO A1
Salve como: corrigir_nfe_windows.py (na raiz do projeto)
Trata problemas de codificação do Windows
"""

import sys
import os
from pathlib import Path

def ler_arquivo_com_encoding(caminho_arquivo):
    """Lê arquivo tentando diferentes codificações"""
    encodings = ['utf-8', 'utf-8-sig', 'latin-1', 'cp1252', 'iso-8859-1']
    
    for encoding in encodings:
        try:
            with open(caminho_arquivo, 'r', encoding=encoding) as f:
                content = f.read()
            print(f"✅ Arquivo lido com encoding: {encoding}")
            return content
        except UnicodeDecodeError:
            continue
        except Exception as e:
            print(f"❌ Erro com {encoding}: {e}")
            continue
    
    raise Exception("Não foi possível ler o arquivo com nenhuma codificação")

def setup_system():
    """Configura o sistema para correção"""
    print("🔧 CORREÇÃO NFe/CERTIFICADO A1 - WINDOWS")
    print("=" * 50)
    
    # Configurar encoding do console para Windows
    if sys.platform.startswith('win'):
        try:
            # Tentar configurar UTF-8 no Windows
            os.system('chcp 65001 >nul 2>&1')
        except:
            pass
    
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
            print(f"+ Path: {path}")
    
    # 2. Verificar arquivos essenciais
    arquivos_necessarios = [
        "src/nfe/correcao_certificado_a1.py",
        "src/Sistema_Entrada_Dados.py"
    ]
    
    print("\nVerificando arquivos...")
    for arquivo in arquivos_necessarios:
        if os.path.exists(arquivo):
            print(f"OK {arquivo}")
        else:
            print(f"ERRO {arquivo} - NAO ENCONTRADO!")
            return False
    
    return True

def aplicar_correcao():
    """Aplica a correção de certificado A1"""
    try:
        print("\nAplicando correcao...")
        
        # Importar Sistema_Entrada_Dados
        try:
            from src.Sistema_Entrada_Dados import SistemaEntradaDados
        except ImportError:
            from Sistema_Entrada_Dados import SistemaEntradaDados
        
        # Criar instância temporária (sem interface)
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Ocultar
        
        print("Criando sistema temporario...")
        sistema = SistemaEntradaDados(parent=root)
        
        # Ler e executar correção com encoding correto
        print("Carregando correcao de certificado...")
        
        try:
            codigo_correcao = ler_arquivo_com_encoding('src/nfe/correcao_certificado_a1.py')
            
            # Executar código
            exec(codigo_correcao)
            
            # A função aplicar_correcao_automatica agora está disponível
            print("Aplicando correcao automatica...")
            sucesso = aplicar_correcao_automatica(sistema)
            
            if sucesso:
                print("\nSUCESSO! CORRECAO APLICADA!")
                
                # Tornar disponível globalmente
                import builtins
                builtins.sistema_principal = sistema
                
                print("\nAGORA NO CONSOLE PYTHON DO VSCODE:")
                print(">>> sistema_principal.configurar_certificado_rapido()")
                print("\nPara diagnostico:")
                print(">>> sistema_principal.diagnosticar_nfe()")
                
                # Instruções detalhadas
                print(f"""
INSTRUCOES PARA VSCODE:

1. ABRA CONSOLE PYTHON:
   - Ctrl+Shift+P
   - Digite: "Python: Start REPL"
   - Pressione Enter

2. NO CONSOLE PYTHON, EXECUTE:
   >>> sistema_principal.configurar_certificado_rapido()

3. CONFIGURAR CERTIFICADO:
   - Clique "Procurar" e selecione arquivo .pfx
   - Digite senha/PIN (normalmente 6 digitos)
   - Clique "Configurar"
   - Aguarde validacao automatica

4. TESTAR CONSULTA:
   >>> sistema_principal.diagnosticar_nfe()
   >>> chave = "sua_chave_44_digitos"
   >>> dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)
   >>> print(dados.get('razao_social_emitente', 'Erro'))

5. IMPORTAR NFE COMPLETA:
   - Use a interface grafica normal
   - Botao "Processar NFe" aparecera apos importar dados
                """)
                
                return True
            else:
                print("ERRO: Falha na aplicacao da correcao")
                return False
                
        except Exception as e:
            print(f"ERRO ao ler arquivo de correcao: {e}")
            print("\nTentando metodo alternativo...")
            return aplicar_correcao_alternativa(sistema)
            
    except Exception as e:
        print(f"ERRO geral: {e}")
        import traceback
        traceback.print_exc()
        return False

def aplicar_correcao_alternativa(sistema):
    """Método alternativo sem usar arquivo externo"""
    try:
        print("Aplicando correcao alternativa (codigo direto)...")
        
        # Importar e aplicar correção diretamente
        try:
            from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
            sucesso = corrigir_sistema_certificado_a1(sistema)
            
            if sucesso:
                print("Correcao alternativa aplicada com sucesso!")
                
                # Tornar disponível globalmente
                import builtins
                builtins.sistema_principal = sistema
                
                return True
            else:
                print("Falha na correcao alternativa")
                return False
                
        except ImportError as e:
            print(f"Erro de importacao: {e}")
            return False
            
    except Exception as e:
        print(f"Erro na correcao alternativa: {e}")
        return False

def verificar_sistema():
    """Verifica se o sistema está funcionando"""
    try:
        # Verificar se sistema_principal está disponível
        if 'sistema_principal' in dir(__builtins__) or hasattr(__builtins__, 'sistema_principal'):
            print("Sistema principal: DISPONIVEL")
            
            sistema = getattr(__builtins__, 'sistema_principal', None)
            if sistema:
                print(f"Tipo: {type(sistema)}")
                print(f"Tem processador NFe: {hasattr(sistema, 'processador_nfe')}")
                
                if hasattr(sistema, 'processador_nfe'):
                    print("Metodos disponveis:")
                    metodos = [m for m in dir(sistema.processador_nfe) if not m.startswith('_')]
                    for metodo in metodos[:10]:  # Primeiros 10
                        print(f"  - {metodo}")
                
                return True
        else:
            print("Sistema principal: NAO DISPONIVEL")
            return False
            
    except Exception as e:
        print(f"Erro na verificacao: {e}")
        return False

def main():
    """Função principal"""
    try:
        # Setup inicial
        if not setup_system():
            print("\nERRO no setup inicial")
            return
        
        # Aplicar correção
        if aplicar_correcao():
            print("\nSUCESSO! Siga as instrucoes acima.")
            
            # Verificar se funcionou
            if verificar_sistema():
                print("\nVerificacao: SISTEMA OK")
            
            # Manter script ativo
            print("\nScript ativo (Ctrl+C para sair)")
            import time
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\nFinalizando...")
        else:
            print("\nERRO na aplicacao")
            
    except Exception as e:
        print(f"ERRO geral: {e}")

def menu_simples():
    """Menu simplificado"""
    print("""
OPCOES:
1. Aplicar correcao automaticamente
2. Verificar sistema
3. Mostrar instrucoes
0. Sair
    """)
    
    try:
        opcao = input("Opcao: ").strip()
        
        if opcao == "1":
            main()
        elif opcao == "2":
            verificar_sistema()
        elif opcao == "3":
            print("""
INSTRUCOES COMPLETAS:

1. Execute este script: python corrigir_nfe_windows.py
2. Escolha opcao 1
3. Aguarde mensagem de sucesso
4. Abra Console Python no VSCode (Ctrl+Shift+P -> Python: Start REPL)
5. Execute: sistema_principal.configurar_certificado_rapido()
6. Configure seu certificado .pfx
7. Teste: sistema_principal.diagnosticar_nfe()
            """)
        elif opcao == "0":
            print("Saindo...")
        else:
            print("Opcao invalida")
            
    except Exception as e:
        print(f"Erro: {e}")

if __name__ == "__main__":
    print("""
CORRECAO NFe/CERTIFICADO A1 - VERSAO WINDOWS
Resolve problemas de codificacao e encoding
    """)
    
    try:
        main()
    except KeyboardInterrupt:
        print("\nInterrompido pelo usuario")
    except Exception as e:
        print(f"Erro critico: {e}")
        print("\nTente executar com menu:")
        print("python corrigir_nfe_windows.py")
        menu_simples()
