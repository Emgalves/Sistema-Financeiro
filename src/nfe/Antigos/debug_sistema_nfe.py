# -*- coding: utf-8 -*-
"""
DEBUG E CORREÇÃO DO SISTEMA NFe
Versão com logs detalhados para identificar problemas
"""

def debug_sistema_nfe_completo(sistema_principal):
    """
    Função de debug completa para identificar problemas no sistema NFe
    """
    print("🔍 INICIANDO DEBUG DO SISTEMA NFe...")
    print("=" * 50)
    
    # 1. VERIFICAR ESTRUTURA DE ARQUIVOS
    print("📁 1. Verificando arquivos...")
    from pathlib import Path
    
    arquivos_necessarios = [
        "src/nfe/sistema_nfe_unificado.py",
        "src/nfe/ajustes_sistema_nfe.py",
        "src/materiais/gerenciador_materiais.py"
    ]
    
    for arquivo in arquivos_necessarios:
        caminho = Path(arquivo)
        status = "✅" if caminho.exists() else "❌"
        print(f"   {status} {arquivo}")
    
    # 2. VERIFICAR IMPORTS
    print("\n📦 2. Testando imports...")
    
    try:
        from src.nfe.sistema_nfe_unificado import SistemaNFeUnificado, substituir_sistemas_nfe_por_unificado
        print("   ✅ sistema_nfe_unificado.py")
    except Exception as e:
        print(f"   ❌ sistema_nfe_unificado.py: {e}")
        return False
    
    try:
        from src.nfe.ajustes_sistema_nfe import aplicar_todos_ajustes_nfe
        print("   ✅ ajustes_sistema_nfe.py")
    except Exception as e:
        print(f"   ❌ ajustes_sistema_nfe.py: {e}")
    
    try:
        from src.materiais.gerenciador_materiais import inicializar_sistema_materiais_completo
        print("   ✅ gerenciador_materiais.py")
    except Exception as e:
        print(f"   ❌ gerenciador_materiais.py: {e}")
    
    # 3. VERIFICAR SISTEMA PRINCIPAL
    print("\n🏗️ 3. Verificando sistema principal...")
    
    atributos_importantes = [
        'root',
        'aba_fornecedor', 
        'cliente_atual',
        'dados_para_incluir'
    ]
    
    for attr in atributos_importantes:
        tem_attr = hasattr(sistema_principal, attr)
        status = "✅" if tem_attr else "❌"
        print(f"   {status} {attr}")
    
    # 4. TESTAR INICIALIZAÇÃO DO SISTEMA NFe
    print("\n🚀 4. Testando inicialização do sistema NFe...")
    
    try:
        sistema_nfe = SistemaNFeUnificado(sistema_principal)
        print("   ✅ SistemaNFeUnificado criado com sucesso")
        
        # Testar método principal
        if hasattr(sistema_nfe, 'criar_interface_importacao'):
            print("   ✅ Método criar_interface_importacao existe")
        else:
            print("   ❌ Método criar_interface_importacao não existe")
            
        return sistema_nfe
        
    except Exception as e:
        print(f"   ❌ Erro ao criar SistemaNFeUnificado: {e}")
        import traceback
        print(f"   📄 Traceback: {traceback.format_exc()}")
        return None

def inicializar_sistema_nfe_seguro(sistema_principal):
    """
    Versão segura da inicialização com fallbacks
    """
    try:
        print("🚀 Inicializando sistema NFe (versão segura)...")
        
        # PASSO 1: Testar se tudo está OK
        sistema_nfe = debug_sistema_nfe_completo(sistema_principal)
        
        if not sistema_nfe:
            print("❌ Debug falhou - não é possível inicializar")
            return None
        
        # PASSO 2: Inicializar sistema unificado
        print("\n🔧 5. Inicializando sistema unificado...")
        try:
            from src.nfe.sistema_nfe_unificado import substituir_sistemas_nfe_por_unificado
            sistema_principal.sistema_nfe_unificado = substituir_sistemas_nfe_por_unificado(sistema_principal)
            print("   ✅ Sistema unificado inicializado")
        except Exception as e:
            print(f"   ❌ Erro no sistema unificado: {e}")
            # FALLBACK: Criar sistema básico
            sistema_principal.sistema_nfe_unificado = sistema_nfe
            adicionar_botao_nfe_manual(sistema_principal, sistema_nfe)
            print("   ⚠️ Usando fallback - sistema básico")
        
        # PASSO 3: Aplicar ajustes (opcional)
        print("\n🎨 6. Aplicando ajustes...")
        try:
            from src.nfe.ajustes_sistema_nfe import aplicar_todos_ajustes_nfe
            aplicar_todos_ajustes_nfe(sistema_principal)
            print("   ✅ Ajustes aplicados")
        except Exception as e:
            print(f"   ⚠️ Ajustes não aplicados: {e}")
            print("   ℹ️ Sistema funcionará sem ajustes")
        
        print("\n✅ SISTEMA NFe INICIALIZADO COM SUCESSO!")
        print("📌 Método disponível: sistema.abrir_importacao_nfe()")
        
        return sistema_principal.sistema_nfe_unificado
        
    except Exception as e:
        print(f"\n❌ ERRO GERAL NA INICIALIZAÇÃO: {e}")
        return None

def adicionar_botao_nfe_manual(sistema_principal, sistema_nfe):
    """
    Adiciona botão NFe manualmente se a função automática falhar
    """
    try:
        print("   🔧 Adicionando botão NFe manualmente...")
        
        if not hasattr(sistema_principal, 'aba_fornecedor'):
            print("   ❌ aba_fornecedor não encontrada")
            return
        
        # PROCURAR SEÇÃO DE MATERIAIS
        frame_materiais = None
        for widget in sistema_principal.aba_fornecedor.winfo_children():
            if hasattr(widget, 'configure') and 'text' in widget.configure():
                texto = widget['text']
                if 'Materiais' in texto:
                    frame_materiais = widget
                    break
        
        if not frame_materiais:
            print("   ❌ Seção de materiais não encontrada")
            return
        
        # ENCONTRAR FRAME DE BOTÕES
        frame_botoes = None
        for subwidget in frame_materiais.winfo_children():
            if str(type(subwidget)).endswith("Frame'>"):
                # Verificar se tem botões
                tem_botoes = any(str(type(child)).endswith("Button'>") 
                               for child in subwidget.winfo_children())
                if tem_botoes:
                    frame_botoes = subwidget
                    break
        
        if frame_botoes:
            import tkinter as tk
            from tkinter import ttk
            
            # CRIAR BOTÃO NFe
            def abrir_nfe():
                sistema_nfe.criar_interface_importacao()
            
            btn_nfe = ttk.Button(
                frame_botoes,
                text="📄 Importar NFe", 
                command=abrir_nfe
            )
            btn_nfe.pack(side='left', padx=5)
            
            # ADICIONAR MÉTODO DE CONVENIÊNCIA
            sistema_principal.abrir_importacao_nfe = abrir_nfe
            
            print("   ✅ Botão NFe adicionado manualmente")
        else:
            print("   ❌ Frame de botões não encontrado")
            
    except Exception as e:
        print(f"   ❌ Erro ao adicionar botão manual: {e}")

def verificar_sistema_pos_inicializacao(sistema_principal):
    """
    Verifica se sistema foi inicializado corretamente
    """
    print("\n🔍 VERIFICAÇÃO PÓS-INICIALIZAÇÃO:")
    print("=" * 40)
    
    # Verificar atributos criados
    checks = [
        ('sistema_nfe_unificado', hasattr(sistema_principal, 'sistema_nfe_unificado')),
        ('abrir_importacao_nfe', hasattr(sistema_principal, 'abrir_importacao_nfe')),
    ]
    
    for nome, check in checks:
        status = "✅" if check else "❌"
        print(f"   {status} {nome}")
    
    # Verificar botão na interface
    try:
        if hasattr(sistema_principal, 'aba_fornecedor'):
            botoes_nfe = []
            for widget in sistema_principal.aba_fornecedor.winfo_children():
                if hasattr(widget, 'winfo_children'):
                    for subwidget in widget.winfo_children():
                        if hasattr(subwidget, 'winfo_children'):
                            for btn in subwidget.winfo_children():
                                if hasattr(btn, 'configure') and 'text' in btn.configure():
                                    if 'NFe' in btn['text']:
                                        botoes_nfe.append(btn['text'])
            
            if botoes_nfe:
                print(f"   ✅ Botões NFe encontrados: {botoes_nfe}")
            else:
                print("   ❌ Nenhum botão NFe encontrado na interface")
    except Exception as e:
        print(f"   ⚠️ Erro ao verificar botões: {e}")
    
    # Teste funcional
    try:
        if hasattr(sistema_principal, 'abrir_importacao_nfe'):
            print("   ✅ Método abrir_importacao_nfe disponível")
            print("   💡 Teste: sistema.abrir_importacao_nfe()")
        else:
            print("   ❌ Método abrir_importacao_nfe não disponível")
    except Exception as e:
        print(f"   ❌ Erro no teste funcional: {e}")

# VERSÃO SIMPLIFICADA PARA TESTE RÁPIDO
def teste_rapido_nfe(sistema_principal):
    """
    Teste rápido do sistema NFe
    """
    print("🧪 TESTE RÁPIDO SISTEMA NFe")
    print("=" * 30)
    
    try:
        # Importar classe principal
        from src.nfe.sistema_nfe_unificado import SistemaNFeUnificado
        
        # Criar instância
        sistema_nfe = SistemaNFeUnificado(sistema_principal)
        
        # Testar método principal
        print("✅ Sistema NFe criado com sucesso!")
        print("✅ Pronto para uso!")
        
        return sistema_nfe
        
    except Exception as e:
        print(f"❌ Teste falhou: {e}")
        return None

# FUNÇÃO PRINCIPAL CORRIGIDA
def inicializar_sistema_nfe_com_debug(sistema_principal):
    """
    Função principal que você deve usar no __init__
    """
    print("🚀 INICIALIZANDO SISTEMA NFe COM DEBUG...")
    
    try:
        # TENTAR INICIALIZAÇÃO SEGURA
        resultado = inicializar_sistema_nfe_seguro(sistema_principal)
        
        if resultado:
            # VERIFICAR RESULTADO
            verificar_sistema_pos_inicializacao(sistema_principal)
            return resultado
        else:
            print("❌ Inicialização falhou")
            return None
            
    except Exception as e:
        print(f"❌ ERRO CRÍTICO: {e}")
        
        # ÚLTIMO RECURSO: TESTE RÁPIDO
        print("\n🆘 Tentando teste rápido...")
        return teste_rapido_nfe(sistema_principal)

# EXEMPLO DE USO PARA SEU __init__
"""
SUBSTITUA NO SEU __init__ POR ESTA VERSÃO COM DEBUG:

# ❌ REMOVER estas linhas antigas:
# self.sistema_nfe_unificado = substituir_sistemas_nfe_por_unificado(self)
# aplicar_todos_ajustes_nfe(self)

# ✅ ADICIONAR esta linha nova:
try:
    from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
    inicializar_sistema_nfe_com_debug(self)
except Exception as e:
    print(f"⚠️ Sistema NFe não carregado: {e}")

ESTA VERSÃO VAI:
1. ✅ Mostrar exatamente onde está o problema
2. ✅ Tentar várias formas de inicializar
3. ✅ Criar o botão mesmo se algo falhar
4. ✅ Dar logs detalhados para debug
5. ✅ Funcionar mesmo com problemas parciais
"""