# -*- coding: utf-8 -*-
"""
INICIALIZAÇÃO NFe FINAL
Código para substituir no __init__ do SistemaEntradaDados
"""

def inicializar_nfe_sistema_final(sistema_principal):
    """
    Função única que substitui todo o código NFe anterior
    """
    try:
        print("🚀 Inicializando Sistema NFe Final...")
        
        # IMPORTAR E INICIALIZAR SISTEMA SIMPLIFICADO
        from src.nfe.sistema_nfe_simplificado import inicializar_sistema_nfe_simplificado
        resultado = inicializar_sistema_nfe_simplificado(sistema_principal)
        
        if resultado:
            print("✅ Sistema NFe Final inicializado com sucesso!")
            print("📄 Fluxo: Selecionar XML → Configurar → Importar")
            print("🔌 Método: sistema.abrir_importacao_nfe()")
            return True
        else:
            print("❌ Falha na inicialização")
            return False
            
    except ImportError as e:
        print(f"❌ Erro de importação: {e}")
        print("💡 Verifique se o arquivo sistema_nfe_simplificado.py existe")
        return False
        
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        return False


# EXEMPLO DE USO NO __init__ DO SistemaEntradaDados:
"""
SUBSTITUA TODO O CÓDIGO NFE ATUAL POR APENAS ESTAS LINHAS:

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
    from src.nfe.inicializacao_nfe_final import inicializar_nfe_sistema_final
    inicializar_nfe_sistema_final(self)
except Exception as e:
    print(f"⚠️ Sistema NFe não carregado: {e}")

RESULTADO:
✅ Código limpo e simples
✅ Apenas uma função de inicialização
✅ Fluxo direto: XML → Configuração → Importação
✅ Eliminação de todas as telas intermediárias
✅ Data do período atual calculada automaticamente
✅ Interface unificada e intuitiva
"""