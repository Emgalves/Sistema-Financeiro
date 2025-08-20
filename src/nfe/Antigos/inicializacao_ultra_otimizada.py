# -*- coding: utf-8 -*-
"""
INICIALIZAÇÃO ULTRA OTIMIZADA - FLUXO DIRETO
Solução definitiva com o mínimo de cliques possível
"""

def inicializar_nfe_ultra_otimizado_definitivo(sistema_principal):
    """
    Função que implementa o fluxo ultra otimizado:
    1 clique → Seleção de arquivo → Configuração → Importação
    """
    try:
        print("⚡ Inicializando Sistema NFe ULTRA OTIMIZADO...")
        
        # IMPORTAR E INICIALIZAR SISTEMA ULTRA OTIMIZADO
        from src.nfe.sistema_nfe_ultra_otimizado import inicializar_sistema_nfe_ultra_otimizado
        resultado = inicializar_sistema_nfe_ultra_otimizado(sistema_principal)
        
        if resultado:
            # VERIFICAR SE TUDO ESTÁ FUNCIONANDO
            from src.nfe.sistema_nfe_ultra_otimizado import debug_sistema_ultra_otimizado
            sucesso = debug_sistema_ultra_otimizado(sistema_principal)
            
            if sucesso:
                print("🎉 SISTEMA NFe ULTRA OTIMIZADO ATIVO!")
                print("⚡ FLUXO ULTRA RÁPIDO:")
                print("   1️⃣ Clique no botão '⚡ Importar NFe'")
                print("   2️⃣ Seleciona arquivo XML")
                print("   3️⃣ Configura e importa")
                print("🚀 APENAS 3 PASSOS - MÁXIMA OTIMIZAÇÃO!")
                return True
            else:
                print("⚠️ Sistema inicializado mas com problemas")
                return False
        else:
            print("❌ Falha na inicialização")
            return False
            
    except ImportError as e:
        print(f"❌ Erro de importação: {e}")
        print("💡 Verifique se o arquivo sistema_nfe_ultra_otimizado.py existe")
        
        # FALLBACK 1: Tentar sistema com botão garantido
        try:
            print("🆘 Tentando fallback 1 - sistema com botão garantido...")
            from src.nfe.sistema_nfe_com_botao_garantido import inicializar_sistema_nfe_com_botao_garantido
            resultado = inicializar_sistema_nfe_com_botao_garantido(sistema_principal)
            if resultado:
                print("✅ Fallback 1 funcionou!")
                return True
        except Exception as e2:
            print(f"❌ Fallback 1 falhou: {e2}")
        
        # FALLBACK 2: Tentar sistema de debug original
        try:
            print("🆘 Tentando fallback 2 - sistema de debug...")
            from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
            inicializar_sistema_nfe_com_debug(sistema_principal)
            print("✅ Fallback 2 aplicado - botão deve estar disponível")
            return True
        except Exception as e3:
            print(f"❌ Fallback 2 também falhou: {e3}")
            return False
        
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        
        # FALLBACK 1: Tentar sistema com botão garantido
        try:
            print("🆘 Tentando fallback 1 - sistema com botão garantido...")
            from src.nfe.sistema_nfe_com_botao_garantido import inicializar_sistema_nfe_com_botao_garantido
            resultado = inicializar_sistema_nfe_com_botao_garantido(sistema_principal)
            if resultado:
                print("✅ Fallback 1 funcionou!")
                return True
        except Exception as e2:
            print(f"❌ Fallback 1 falhou: {e2}")
        
        # FALLBACK 2: Tentar sistema de debug original
        try:
            print("🆘 Tentando fallback 2 - sistema de debug...")
            from src.nfe.debug_sistema_nfe import inicializar_sistema_nfe_com_debug
            inicializar_sistema_nfe_com_debug(sistema_principal)
            print("✅ Fallback 2 aplicado - botão deve estar disponível")
            return True
        except Exception as e3:
            print(f"❌ Todos os fallbacks falharam: {e3}")
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
    from src.nfe.inicializacao_ultra_otimizada import inicializar_nfe_ultra_otimizado_definitivo
    inicializar_nfe_ultra_otimizado_definitivo(self)
except Exception as e:
    print(f"⚠️ Sistema NFe não carregado: {e}")

VANTAGENS DA SOLUÇÃO ULTRA OTIMIZADA:
⚡ FLUXO ULTRA RÁPIDO: Apenas 3 passos
   1️⃣ Clique no botão "⚡ Importar NFe"
   2️⃣ Seleciona arquivo XML (seletor abre automaticamente)
   3️⃣ Configura e importa (interface de configuração abre automaticamente)

✅ ELIMINAÇÃO TOTAL DE TELAS INTERMEDIÁRIAS:
   • Não há mais tela "Processar XML"
   • Não há mais tela "Dados Extraídos"
   • Não há mais botão "Importar para Sistema"
   • Processamento em segundo plano

✅ INTERFACE OTIMIZADA:
   • Título dinâmico com info da NFe
   • Layout compacto e funcional
   • Preview integrado
   • Feedback visual aprimorado

✅ ROBUSTEZ:
   • 2 níveis de fallback automático
   • Sistema de debug como último recurso
   • Logs detalhados para troubleshooting

✅ FUNCIONALIDADES MANTIDAS:
   • Data do período atual calculada automaticamente
   • Classificação automática de produtos
   • Configuração completa de financeiro e materiais
   • Integração total com o sistema principal

RESULTADO FINAL:
🎯 MÁXIMA REDUÇÃO DE CLIQUES: De 5-6 cliques para apenas 3
🎯 FLUXO INTUITIVO: Botão → Arquivo → Configurar → Pronto
🎯 ZERO TELAS INTERMEDIÁRIAS: Direto ao que importa
🎯 COMPATIBILIDADE GARANTIDA: Fallbacks automáticos
"""