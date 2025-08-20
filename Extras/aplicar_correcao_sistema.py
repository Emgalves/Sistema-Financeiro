# -*- coding: utf-8 -*-
"""
CORREÇÃO ESPECÍFICA PARA SEU SISTEMA PRINCIPAL
Salve como: aplicar_correcao_sistema.py (na raiz do projeto)
"""

import sys
import os
from pathlib import Path

def encontrar_sistema_principal():
    """Encontra a instância do SistemaPrincipal em execução"""
    try:
        print("🔍 Procurando sistema principal em execução...")
        
        # Procurar em todos os módulos carregados
        for nome_modulo, modulo in sys.modules.items():
            if hasattr(modulo, '__dict__'):
                for attr_nome, attr_valor in modulo.__dict__.items():
                    # Procurar por SistemaPrincipal
                    if (hasattr(attr_valor, '__class__') and 
                        'SistemaPrincipal' in str(type(attr_valor))):
                        print(f"✅ Sistema encontrado: {nome_modulo}.{attr_nome}")
                        return attr_valor
                    
                    # Procurar por SistemaEntradaDados (que é referenciado no seu código)
                    if (hasattr(attr_valor, '__class__') and 
                        'SistemaEntradaDados' in str(type(attr_valor))):
                        print(f"✅ SistemaEntradaDados encontrado: {nome_modulo}.{attr_nome}")
                        return attr_valor
        
        print("❌ Sistema não encontrado em execução")
        return None
        
    except Exception as e:
        print(f"❌ Erro ao procurar sistema: {e}")
        return None

def configurar_paths():
    """Configura os paths necessários"""
    try:
        print("📁 Configurando paths...")
        
        # Obter diretório atual
        current_dir = Path(__file__).resolve().parent
        
        # Adicionar paths relevantes
        paths_para_adicionar = [
            str(current_dir),              # Raiz
            str(current_dir / "src"),      # src/
            str(current_dir / "src" / "nfe")  # src/nfe/
        ]
        
        for path in paths_para_adicionar:
            if os.path.exists(path) and path not in sys.path:
                sys.path.insert(0, path)
                print(f"  ➕ {path}")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao configurar paths: {e}")
        return False

def criar_sistema_entrada_dados():
    """Cria uma instância do SistemaEntradaDados se não existir"""
    try:
        print("🔧 Tentando criar instância do SistemaEntradaDados...")
        
        # Tentar importar
        try:
            from src.Sistema_Entrada_Dados import SistemaEntradaDados
        except ImportError:
            try:
                from Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                print("❌ Não foi possível importar SistemaEntradaDados")
                return None
        
        # Criar instância (sem abrir interface)
        print("⚙️ Criando instância...")
        sistema = SistemaEntradaDados(headless=True)  # Modo sem interface
        
        print("✅ SistemaEntradaDados criado com sucesso!")
        return sistema
        
    except Exception as e:
        print(f"❌ Erro ao criar SistemaEntradaDados: {e}")
        return None

def aplicar_correcao_nfe(sistema):
    """Aplica a correção específica de NFe/certificado"""
    try:
        print("\n🔧 APLICANDO CORREÇÃO DE CERTIFICADO A1")
        print("=" * 50)
        
        # Verificar se sistema híbrido NFe está inicializado
        if not hasattr(sistema, 'processador_nfe'):
            print("⚠️ Sistema híbrido NFe não inicializado")
            print("🔄 Inicializando sistema híbrido...")
            
            try:
                from src.nfe.extensao_sistema_hibrido import inicializar_sistema_nfe_estendido
                resultado = inicializar_sistema_nfe_estendido(sistema)
                
                if resultado:
                    print("✅ Sistema híbrido NFe inicializado!")
                else:
                    print("❌ Falha na inicialização do sistema híbrido")
                    return False
                    
            except ImportError:
                print("⚠️ Extensão não encontrada, tentando sistema básico...")
                try:
                    from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
                    inicializar_sistema_nfe_hibrido(sistema)
                    print("✅ Sistema híbrido básico inicializado!")
                except Exception as e:
                    print(f"❌ Erro ao inicializar sistema básico: {e}")
                    return False
        
def aplicar_correcao_nfe(sistema):
    """Aplica a correção específica de NFe/certificado"""
    try:
        print("\n🔧 APLICANDO CORREÇÃO DE CERTIFICADO A1")
        print("=" * 50)
        
        # Verificar se sistema híbrido NFe está inicializado
        if not hasattr(sistema, 'processador_nfe'):
            print("⚠️ Sistema híbrido NFe não inicializado")
            print("🔄 Inicializando sistema híbrido...")
            
            try:
                from src.nfe.extensao_sistema_hibrido import inicializar_sistema_nfe_estendido
                resultado = inicializar_sistema_nfe_estendido(sistema)
                
                if resultado:
                    print("✅ Sistema híbrido NFe inicializado!")
                else:
                    print("❌ Falha na inicialização do sistema híbrido")
                    return False
                    
            except ImportError:
                print("⚠️ Extensão não encontrada, tentando sistema básico...")
                try:
                    from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
                    inicializar_sistema_nfe_hibrido(sistema)
                    print("✅ Sistema híbrido básico inicializado!")
                except Exception as e:
                    print(f"❌ Erro ao inicializar sistema básico: {e}")
                    return False
        
        # Aplicar correção de certificado
        print("\n🔑 Aplicando correção de certificado A1...")
        
        try:
            from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
            sucesso = corrigir_sistema_certificado_a1(sistema)
            
            if sucesso:
                print("✅ Correção de certificado A1 aplicada com sucesso!")
                
                # Tornar sistema disponível globalmente para uso no console
                import builtins
                builtins.sistema_principal = sistema
                
                print("\n🎯 PRÓXIMOS PASSOS:")
                print("1. No console Python do VSCode, execute:")
                print("   sistema_principal.configurar_certificado_rapido()")
                print("\n2. Para diagnóstico:")
                print("   sistema_principal.diagnosticar_nfe()")
                print("\n3. Para testar consulta:")
                print("   chave = 'sua_chave_44_digitos'")
                print("   dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)")
                
                return True
            else:
                print("❌ Falha na aplicação da correção")
                return False
                
        except ImportError:
            print("❌ Arquivo de correção não encontrado!")
            print("💡 Certifique-se de que correcao_certificado_a1.py está em src/nfe/")
            return False
        except Exception as e:
            print(f"❌ Erro durante aplicação: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        return False

def main():
    """Função principal para aplicar correções"""
    print("\n🚀 CORREÇÃO DE CERTIFICADO A1 - SISTEMA ESPECÍFICO")
    print("=" * 60)
    
    # 1. Configurar paths
    if not configurar_paths():
        print("❌ Falha na configuração de paths")
        return False
    
    # 2. Procurar sistema em execução
    sistema = encontrar_sistema_principal()
    
    # 3. Se não encontrou, tentar criar
    if not sistema:
        print("\n🔧 Sistema não encontrado em execução")
        print("⚙️ Tentando criar instância do SistemaEntradaDados...")
        
        try:
            # Importar SistemaEntradaDados
            try:
                from src.Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                from Sistema_Entrada_Dados import SistemaEntradaDados
            
            # Criar instância sem interface gráfica
            import tkinter as tk
            root = tk.Tk()
            root.withdraw()  # Ocultar janela
            
            sistema = SistemaEntradaDados(parent=root)
            print("✅ SistemaEntradaDados criado!")
            
        except Exception as e:
            print(f"❌ Erro ao criar SistemaEntradaDados: {e}")
            print("\n💡 SOLUÇÃO ALTERNATIVA:")
            print("1. Execute seu sistema normalmente (sistema_principal.py)")
            print("2. Quando estiver funcionando, execute este script novamente")
            return False
    
    # 4. Aplicar correção
    if sistema:
        sucesso = aplicar_correcao_nfe(sistema)
        
        if sucesso:
            print("\n🎉 CORREÇÃO APLICADA COM SUCESSO!")
            print("\n📋 O QUE FAZER AGORA:")
            print("1. Abra o Console Python no VSCode (Ctrl+Shift+P → 'Python: Start REPL')")
            print("2. Execute os comandos mostrados acima")
            print("3. Configure seu certificado A1")
            
            # Manter script ativo para permitir uso do console
            print("\n⏳ Mantendo script ativo para uso do console...")
            print("   (Pressione Ctrl+C para sair)")
            
            try:
                import time
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n👋 Finalizando...")
                return True
        else:
            print("\n❌ FALHA NA APLICAÇÃO DA CORREÇÃO")
            return False
    
    return False

def menu_interativo():
    """Menu interativo para diferentes opções"""
    while True:
        print("\n" + "=" * 50)
        print("🔧 CORREÇÃO DE CERTIFICADO A1")
        print("=" * 50)
        print("1. 🚀 Aplicar correção automaticamente")
        print("2. 🔍 Verificar sistema em execução")
        print("3. 📋 Mostrar instruções manuais")
        print("4. 🧪 Testar conexão com sistema")
        print("0. ❌ Sair")
        print("=" * 50)
        
        try:
            opcao = input("Digite sua opção: ").strip()
            
            if opcao == "1":
                main()
                break
            
            elif opcao == "2":
                sistema = encontrar_sistema_principal()
                if sistema:
                    print(f"✅ Sistema encontrado: {type(sistema)}")
                    print(f"📋 Tem processador NFe: {hasattr(sistema, 'processador_nfe')}")
                    if hasattr(sistema, 'cliente_atual'):
                        print(f"👤 Cliente atual: {getattr(sistema, 'cliente_atual', 'Nenhum')}")
                else:
                    print("❌ Sistema não encontrado em execução")
            
            elif opcao == "3":
                mostrar_instrucoes_manuais()
            
            elif opcao == "4":
                testar_conexao_sistema()
            
            elif opcao == "0":
                print("👋 Saindo...")
                break
            
            else:
                print("❌ Opção inválida")
                
        except KeyboardInterrupt:
            print("\n👋 Saindo...")
            break
        except Exception as e:
            print(f"❌ Erro: {e}")

def mostrar_instrucoes_manuais():
    """Mostra instruções detalhadas para correção manual"""
    print("""
📋 INSTRUÇÕES MANUAIS PARA SEU SISTEMA:

1. EXECUTAR SEU SISTEMA:
   - Execute: python sistema_principal.py
   - Clique em "Entrada de Dados" para abrir o Sistema_Entrada_Dados
   
2. ABRIR CONSOLE PYTHON NO VSCODE:
   - Ctrl+Shift+P → "Python: Start REPL"
   
3. NO CONSOLE PYTHON, EXECUTE:
   
   # Configurar paths
   import sys
   sys.path.append('src/nfe')
   
   # Importar e executar correção
   exec(open('src/nfe/correcao_certificado_a1.py').read())
   
   # Encontrar sistema (substituir pela variável correta)
   # Opção A: Se você souber o nome da variável:
   # sistema_entrada = sua_variavel_do_sistema_entrada_dados
   
   # Opção B: Busca automática:
   import sys
   sistema_entrada = None
   for nome, modulo in sys.modules.items():
       if hasattr(modulo, '__dict__'):
           for attr, valor in modulo.__dict__.items():
               if 'SistemaEntradaDados' in str(type(valor)):
                   sistema_entrada = valor
                   print(f"Sistema encontrado: {attr}")
                   break
   
   # Aplicar correção
   if sistema_entrada:
       aplicar_correcao_automatica(sistema_entrada)
       
       # Configurar certificado
       sistema_entrada.configurar_certificado_rapido()
   else:
       print("Sistema não encontrado!")

4. CONFIGURAR CERTIFICADO:
   - Seguir as instruções da interface que abrir
   - Selecionar arquivo .pfx
   - Digitar senha/PIN
   - Aguardar validação
    """)

def testar_conexao_sistema():
    """Testa conexão com o sistema"""
    try:
        print("🧪 Testando conexão com sistema...")
        
        # Configurar paths
        configurar_paths()
        
        # Tentar importar módulos essenciais
        modulos_teste = [
            'src.Sistema_Entrada_Dados',
            'Sistema_Entrada_Dados',
            'src.nfe.correcao_certificado_a1',
            'src.nfe.sistema_hibrido_nfe'
        ]
        
        for modulo in modulos_teste:
            try:
                __import__(modulo)
                print(f"✅ {modulo}: OK")
            except ImportError:
                print(f"❌ {modulo}: NÃO ENCONTRADO")
            except Exception as e:
                print(f"⚠️ {modulo}: ERRO - {e}")
        
        # Verificar arquivos essenciais
        arquivos_teste = [
            'sistema_principal.py',
            'src/Sistema_Entrada_Dados.py',
            'src/nfe/correcao_certificado_a1.py'
        ]
        
        print("\n📁 Verificando arquivos:")
        for arquivo in arquivos_teste:
            if os.path.exists(arquivo):
                print(f"✅ {arquivo}: EXISTE")
            else:
                print(f"❌ {arquivo}: NÃO ENCONTRADO")
        
    except Exception as e:
        print(f"❌ Erro no teste: {e}")

if __name__ == "__main__":
    print("""
🎯 CORREÇÃO ESPECÍFICA PARA SEU SISTEMA

EXECUÇÃO RECOMENDADA:
1. Certifique-se de que sistema_principal.py NÃO está rodando
2. Execute este script: python aplicar_correcao_sistema.py
3. Escolha opção 1 para aplicar correção automaticamente
4. Depois configure o certificado conforme instruído

EXECUÇÃO ALTERNATIVA:
1. Execute sistema_principal.py normalmente
2. Execute este script em outra instância
3. O script encontrará o sistema em execução
    """)
    
    menu_interativo()
