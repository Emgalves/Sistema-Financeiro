# -*- coding: utf-8 -*-
"""
INICIALIZAÇÃO NFe FINAL - COM BOTÃO GARANTIDO
Solução definitiva que combina sistema simplificado + criação robusta do botão
"""

def inicializar_nfe_sistema_definitivo(sistema_principal):
    """
    Função definitiva que garante tanto o sistema quanto o botão NFe
    """
    try:
        print("🚀 Inicializando Sistema NFe Definitivo...")
        
        # IMPORTAR E INICIALIZAR SISTEMA COM BOTÃO GARANTIDO
        from src.nfe.sistema_nfe_com_botao_garantido import inicializar_sistema_nfe_com_botao_garantido
        resultado = inicializar_sistema_nfe_com_botao_garantido(sistema_principal)
        
        if resultado:
            # VERIFICAR SE TUDO ESTÁ FUNCIONANDO
            from src.nfe.sistema_nfe_com_botao_garantido import debug_sistema_nfe_com_botao
            sucesso = debug_sistema_nfe_com_botao(sistema_principal)
            
            if sucesso:
                print("✅ Sistema NFe Definitivo inicializado com SUCESSO!")
                print("📄 Fluxo: Selecionar XML → Configurar → Importar")
                print("🔌 Método: sistema.abrir_importacao_nfe()")
                print("🎯 Botão NFe disponível na interface!")
                return True
            else:
                print("⚠️ Sistema inicializado mas com problemas no botão")
                return False
        else:
            print("❌ Falha na inicialização")
            return False
            
    except ImportError as e:
        print(f"❌ Erro de importação: {e}")
        print("💡 Verifique se o arquivo sistema_nfe_com_botao_garantido.py existe")
        
        # FALLBACK: Tentar usar o sistema de debug que funcionava
        try:
            print("🆘 Tentando fallback com sistema de debug...")
            from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
            inicializar_sistema_nfe_com_debug(sistema_principal)
            print("✅ Fallback aplicado - botão deve estar disponível")
            return True
        except Exception as e2:
            print(f"❌ Fallback também falhou: {e2}")
            return False
        
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        
        # FALLBACK: Tentar usar o sistema de debug que funcionava
        try:
            print("🆘 Tentando fallback com sistema de debug...")
            from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
            inicializar_sistema_nfe_com_debug(sistema_principal)
            print("✅ Fallback aplicado - botão deve estar disponível")
            return True
        except Exception as e2:
            print(f"❌ Fallback também falhou: {e2}")
            return False


# EXEMPLO DE USO NO __init__ DO SistemaEntradaDados:
"""
SUBSTITUA TODO O CÓDIGO NFe ATUAL POR APENAS ESTAS LINHAS:

# ❌ REMOVER TODAS ESTAS LINHAS ANTIGAS:
# try:
#     from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
#     inicializar_sistema_nfe_com_debug(self)
#     from src.nfe.patch_final_otimizado import aplicar_patch_final_otimizado
#     aplicar_patch_final_otimizado(self)
#     print("✅ Patch funcional NFe aplicado!")
# except Exception as e:
#     print(f"⚠️ Patch não aplicado: {e}")

# ✅ ADICIONAR APENAS ESTAS LINHAS NOVAS:
try:
    from src.nfe.inicializacao_nfe_com_botao_garantido import inicializar_nfe_sistema_definitivo
    inicializar_nfe_sistema_definitivo(self)
except Exception as e:
    print(f"⚠️ Sistema NFe não carregado: {e}")

VANTAGENS DESTA SOLUÇÃO:
✅ Sistema simplificado (XML → Configuração → Importação)
✅ Criação robusta do botão (baseada no código do debug que funcionava)
✅ Fallback automático para o sistema de debug se algo falhar
✅ Verificação automática se tudo está funcionando
✅ Logs detalhados para debug
✅ Elimina telas intermediárias desnecessárias
✅ Data do período atual calculada automaticamente

RESULTADO ESPERADO:
🎯 Botão "📄 Importar NFe" aparece na seção de Materiais
🎯 Fluxo direto: Selecionar arquivo → Configurar → Importar
🎯 Funcionalidade completa mantida
🎯 Interface limpa e simplificada
"""