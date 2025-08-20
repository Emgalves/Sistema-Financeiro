# -*- coding: utf-8 -*-
"""
CORREÇÃO COM IMPORTAÇÃO DIRETA - NFe/CERTIFICADO A1
Salve como: corrigir_nfe_direto.py (na raiz do projeto)
Usa importação direta em vez de exec()
"""

import sys
import os
from pathlib import Path

def setup_system():
    """Configura o sistema para correção"""
    print("🔧 CORREÇÃO NFe/CERTIFICADO A1 - IMPORTAÇÃO DIRETA")
    print("=" * 60)
    
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

def criar_sistema():
    """Cria instância do SistemaEntradaDados"""
    try:
        print("\n⚙️ Criando sistema...")
        
        # Importar Sistema_Entrada_Dados
        try:
            from src.Sistema_Entrada_Dados import SistemaEntradaDados
            print("✅ Importação: src.Sistema_Entrada_Dados")
        except ImportError:
            try:
                from Sistema_Entrada_Dados import SistemaEntradaDados
                print("✅ Importação: Sistema_Entrada_Dados")
            except ImportError:
                print("❌ Erro: Não foi possível importar SistemaEntradaDados")
                return None
        
        # Criar instância temporária (sem interface)
        import tkinter as tk
        root = tk.Tk()
        root.withdraw()  # Ocultar
        
        sistema = SistemaEntradaDados(parent=root)
        print("✅ Sistema criado com sucesso!")
        
        return sistema
        
    except Exception as e:
        print(f"❌ Erro ao criar sistema: {e}")
        return None

def aplicar_correcao_metodo1(sistema):
    """Método 1: Importação direta da função"""
    try:
        print("\n🔧 Método 1: Importação direta...")
        
        from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
        print("✅ Função importada com sucesso!")
        
        sucesso = corrigir_sistema_certificado_a1(sistema)
        
        if sucesso:
            print("✅ Correção aplicada via Método 1!")
            return True
        else:
            print("❌ Falha no Método 1")
            return False
            
    except ImportError as e:
        print(f"❌ Erro de importação no Método 1: {e}")
        return False
    except Exception as e:
        print(f"❌ Erro geral no Método 1: {e}")
        return False

def aplicar_correcao_metodo2(sistema):
    """Método 2: Importação do módulo completo"""
    try:
        print("\n🔧 Método 2: Importação do módulo...")
        
        import importlib
        
        # Limpar cache se existir
        modulo_nome = 'src.nfe.correcao_certificado_a1'
        if modulo_nome in sys.modules:
            del sys.modules[modulo_nome]
            
        # Importar módulo
        modulo = importlib.import_module(modulo_nome)
        print("✅ Módulo importado com sucesso!")
        
        # Procurar função correta
        funcoes_disponiveis = [name for name in dir(modulo) if 'corr' in name.lower() and callable(getattr(modulo, name))]
        print(f"🔍 Funções encontradas: {funcoes_disponiveis}")
        
        # Tentar diferentes nomes de função
        nomes_funcao = [
            'corrigir_sistema_certificado_a1',
            'aplicar_correcao_automatica',
            'aplicar_correcoes'
        ]
        
        for nome_funcao in nomes_funcao:
            if hasattr(modulo, nome_funcao):
                funcao = getattr(modulo, nome_funcao)
                print(f"✅ Usando função: {nome_funcao}")
                
                sucesso = funcao(sistema)
                
                if sucesso:
                    print("✅ Correção aplicada via Método 2!")
                    return True
                
        print("❌ Nenhuma função funcionou no Método 2")
        return False
        
    except Exception as e:
        print(f"❌ Erro no Método 2: {e}")
        return False

def aplicar_correcao_metodo3(sistema):
    """Método 3: Inicialização manual do sistema NFe"""
    try:
        print("\n🔧 Método 3: Inicialização manual...")
        
        # Verificar se já tem sistema NFe
        if hasattr(sistema, 'processador_nfe'):
            print("✅ Sistema NFe já inicializado")
        else:
            print("⚙️ Inicializando sistema NFe...")
            try:
                from src.nfe.extensao_sistema_hibrido import inicializar_sistema_nfe_estendido
                resultado = inicializar_sistema_nfe_estendido(sistema)
                if resultado:
                    print("✅ Sistema NFe estendido inicializado!")
                else:
                    print("⚠️ Falha no sistema estendido, tentando básico...")
                    raise ImportError("Fallback para básico")
            except ImportError:
                from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
                inicializar_sistema_nfe_hibrido(sistema)
                print("✅ Sistema NFe básico inicializado!")
        
        # Agora aplicar correção específica de certificado
        try:
            from src.nfe.correcao_certificado_a1 import ConsultorSefazA1Corrigido
            
            # Criar consultor corrigido
            consultor = ConsultorSefazA1Corrigido()
            
            # Substituir consultor antigo
            sistema.consultor_sefaz_a1 = consultor
            
            # Adicionar métodos ao processador
            if hasattr(sistema, 'processador_nfe'):
                sistema.processador_nfe.configurar_certificado_a1 = consultor.configurar_certificado
                sistema.processador_nfe.testar_certificado_a1 = consultor.testar_conectividade
                
                # Método de configuração rápida
                def config_rapida():
                    import tkinter as tk
                    from tkinter import filedialog, simpledialog, messagebox
                    
                    root = tk.Tk()
                    root.withdraw()
                    
                    # Selecionar arquivo
                    cert_path = filedialog.askopenfilename(
                        title="Selecionar Certificado A1",
                        filetypes=[("Certificado A1", "*.pfx *.p12"), ("Todos", "*.*")]
                    )
                    
                    if not cert_path:
                        print("❌ Arquivo não selecionado")
                        root.destroy()
                        return False
                    
                    # Solicitar PIN
                    cert_password = simpledialog.askstring(
                        "PIN do Certificado A1",
                        "Digite o PIN do certificado (6 dígitos):",
                        show='*'
                    )
                    
                    root.destroy()
                    
                    if not cert_password:
                        print("❌ PIN não informado")
                        return False
                    
                    # Configurar certificado
                    sucesso, msg = consultor.configurar_certificado(cert_path, cert_password)
                    
                    if sucesso:
                        print(f"✅ {msg}")
                        
                        # Testar conectividade
                        teste_ok, teste_msg = consultor.testar_conectividade()
                        print(f"🌐 {teste_msg}")
                        
                        messagebox.showinfo("Sucesso", f"✅ {msg}\n🌐 {teste_msg}")
                        return True
                    else:
                        print(f"❌ {msg}")
                        messagebox.showerror("Erro", f"❌ {msg}")
                        return False
                
                # Adicionar método ao sistema
                sistema.configurar_certificado_rapido = config_rapida
                
                # Método de diagnóstico
                def diagnostico():
                    print("\n🔍 DIAGNÓSTICO DO SISTEMA NFe")
                    print("=" * 40)
                    print(f"Sistema híbrido: {'✅' if hasattr(sistema, 'processador_nfe') else '❌'}")
                    print(f"Consultor A1: {'✅' if hasattr(sistema, 'consultor_sefaz_a1') else '❌'}")
                    
                    if hasattr(sistema, 'consultor_sefaz_a1'):
                        cert_info = sistema.consultor_sefaz_a1.obter_info_certificado()
                        if cert_info.get('is_valid'):
                            print(f"✅ Certificado: CONFIGURADO")
                            print(f"   📅 Válido até: {cert_info['not_valid_after'].strftime('%d/%m/%Y')}")
                        else:
                            print("⚠️ Certificado: NÃO CONFIGURADO")
                    
                    print("=" * 40)
                
                sistema.diagnosticar_nfe = diagnostico
                
                print("✅ Correção manual aplicada com sucesso!")
                return True
            else:
                print("❌ Sistema NFe não disponível")
                return False
                
        except Exception as e:
            print(f"❌ Erro na aplicação manual: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Erro no Método 3: {e}")
        return False

def aplicar_correcao(sistema):
    """Aplica correção usando múltiplos métodos"""
    metodos = [
        aplicar_correcao_metodo1,
        aplicar_correcao_metodo2, 
        aplicar_correcao_metodo3
    ]
    
    for i, metodo in enumerate(metodos, 1):
        try:
            if metodo(sistema):
                print(f"\n🎉 SUCESSO COM MÉTODO {i}!")
                return True
        except Exception as e:
            print(f"❌ Método {i} falhou: {e}")
            continue
    
    print("\n❌ TODOS OS MÉTODOS FALHARAM")
    return False

def main():
    """Função principal"""
    try:
        # Setup inicial
        if not setup_system():
            print("\n❌ Falha no setup inicial")
            return
        
        # Criar sistema
        sistema = criar_sistema()
        if not sistema:
            print("\n❌ Falha ao criar sistema")
            return
        
        # Aplicar correção
        if aplicar_correcao(sistema):
            print("\n🎉 CORREÇÃO APLICADA COM SUCESSO!")
            
            # Tornar disponível globalmente
            import builtins
            builtins.sistema_principal = sistema
            
            print("\n🎯 PRÓXIMOS PASSOS NO CONSOLE PYTHON:")
            print(">>> sistema_principal.configurar_certificado_rapido()")
            print(">>> sistema_principal.diagnosticar_nfe()")
            
            print(f"""
📋 INSTRUÇÕES COMPLETAS:

1. ABRA CONSOLE PYTHON NO VSCODE:
   - Ctrl+Shift+P → "Python: Start REPL"

2. CONFIGURE CERTIFICADO:
   >>> sistema_principal.configurar_certificado_rapido()

3. FAÇA DIAGNÓSTICO:
   >>> sistema_principal.diagnosticar_nfe()

4. TESTE CONSULTA:
   >>> chave = "sua_chave_44_digitos"
   >>> dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)
   >>> print(dados.get('razao_social_emitente', 'Erro'))

5. USAR INTERFACE GRÁFICA:
   - Abra seu sistema normalmente
   - Entre em "Entrada de Dados"
   - Use menu NFe → Importar NFe
   - Botão "🚀 Processar NFe" estará disponível
            """)
            
            # Manter script ativo
            print("\n⏳ Script ativo para uso do console (Ctrl+C para sair)")
            import time
            try:
                while True:
                    time.sleep(1)
            except KeyboardInterrupt:
                print("\n👋 Finalizando...")
        else:
            print("\n❌ FALHA NA APLICAÇÃO DA CORREÇÃO")
            
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
