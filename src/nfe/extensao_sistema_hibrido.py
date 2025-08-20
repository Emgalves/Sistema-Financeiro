# -*- coding: utf-8 -*-
"""
Extensão do Sistema Híbrido NFe
Adiciona botão "Processar NFe Completa" após importar dados
Mantém todo sistema original funcionando
"""

def estender_sistema_hibrido_nfe(sistema_principal):
    """
    Estende o sistema híbrido existente adicionando funcionalidade completa
    """
    try:
        print("🔧 Estendendo Sistema Híbrido NFe...")
        
        # Verificar se sistema híbrido já está inicializado
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("⚠️ Sistema híbrido NFe não encontrado. Inicializando primeiro...")
            from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
            inicializar_sistema_nfe_hibrido(sistema_principal)
        
        # Adicionar integrador completo
        from src.nfe.integrador_nfe_sistema import IntegradorNFeFinanceiroMateriais
        sistema_principal.integrador_nfe_completo = IntegradorNFeFinanceiroMateriais(sistema_principal)
        
        # Estender métodos do processador existente
        estender_metodos_processador(sistema_principal)
        
        print("✅ Sistema híbrido NFe estendido com sucesso!")
        print("📄 Nova funcionalidade: Botão 'Processar NFe Completa' nas interfaces de importação")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao estender sistema híbrido: {e}")
        return False


def estender_metodos_processador(sistema_principal):
    """
    Estende os métodos do processador existente para ter 2 fluxos diferentes:
    1. Original (simples): Importa tudo automaticamente
    2. Avançado: Interface completa para configuração
    """
    try:
        processador = sistema_principal.processador_nfe
        
        # Salvar métodos originais
        processador.exibir_dados_extraidos_original = processador.exibir_dados_extraidos
        processador.criar_opcoes_importacao_original = processador.criar_opcoes_importacao
        processador.importar_dados_financeiro_original = processador.importar_dados_financeiro
        processador.importar_dados_material_original = processador.importar_dados_material
        
        # MODIFICAR método original para ser mais automático
        def importar_dados_financeiro_simples(dados_nfe):
            """Versão SIMPLES - importa automaticamente sem perguntar"""
            try:
                # Verificar se cliente está selecionado
                if not hasattr(sistema_principal, 'cliente_atual') or not sistema_principal.cliente_atual:
                    from tkinter import messagebox
                    messagebox.showerror("Erro", "Selecione um cliente antes de processar NFe!")
                    return "Erro: Cliente não selecionado"
                
                # Calcular data de referência automaticamente
                hoje = datetime.now()
                if hoje.day >= 21 or hoje.day <= 5:
                    data_ref = hoje.replace(day=5).strftime('%d/%m/%Y')
                else:
                    data_ref = hoje.replace(day=20).strftime('%d/%m/%Y')
                
                # Criar lançamento financeiro automaticamente
                dados_financeiro = {
                    'data': data_ref,
                    'cnpj_cpf': dados_nfe.get('cnpj_emitente', ''),
                    'nome': dados_nfe.get('razao_social_emitente', ''),
                    'categoria': 'MAT',
                    'tp_desp': '3',  # Fixo: Materiais
                    'referencia': 'MATERIAL VIA NFE',  # Fixo
                    'etapa_obra': '',
                    'nf': dados_nfe.get('numero_nf', ''),
                    'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}",
                    'dias': 1,
                    'valor': f"{dados_nfe.get('valor_total', 0):.2f}",
                    'dt_vencto': dados_nfe.get('data_emissao', ''),  # Data da NFe
                    'dados_bancarios': '',
                    'observacao': f"MATERIAL OBRA - NFE {dados_nfe.get('numero_nf', '')}",
                    'forma_pagamento': ''
                }
                
                # Adicionar aos dados do sistema
                sistema_principal.dados_para_incluir = [dados_financeiro]
                
                return f"Lançamento financeiro criado automaticamente - R$ {dados_nfe.get('valor_total', 0):,.2f}"
                
            except Exception as e:
                return f"Erro ao criar lançamento: {str(e)}"
        
        def importar_dados_material_simples(dados_nfe):
            """Versão SIMPLES - importa todos os materiais automaticamente"""
            try:
                if not hasattr(sistema_principal, 'gerenciador_materiais'):
                    from src.materiais.gerenciador_materiais import GerenciadorMateriais
                    sistema_principal.gerenciador_materiais = GerenciadorMateriais(sistema_principal)
                
                produtos = dados_nfe.get('produtos', [])
                if not produtos:
                    return "Nenhum produto encontrado na NFe"
                
                salvos = 0
                for produto in produtos:
                    material = {
                        'Cliente': sistema_principal.cliente_atual,
                        'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                        'Codigo_Produto': produto.get('codigo', ''),
                        'Descricao_Completa': produto.get('descricao', ''),
                        'Ambiente_Aplicacao': 'DEPÓSITO DA OBRA',  # Fixo
                        'Status_Instalacao': 'PENDENTE',
                        'Tem_Dados_Compra': True,
                        'Nome_Fornecedor': dados_nfe.get('razao_social_emitente', ''),
                        'CNPJ_Fornecedor': dados_nfe.get('cnpj_emitente', ''),
                        'Data_Compra': dados_nfe.get('data_emissao', ''),
                        'Quantidade': produto.get('quantidade', 0),
                        'Unidade': produto.get('unidade', 'UN'),
                        'Valor_Unitario': produto.get('valor_unitario', 0),
                        'Valor_Total': produto.get('valor_total', 0),
                        'Numero_NF': dados_nfe.get('numero_nf', ''),
                        'Observacoes': f"Importado automaticamente da NFe {dados_nfe.get('numero_nf', '')}"
                    }
                    
                    try:
                        sistema_principal.gerenciador_materiais.salvar_material(material)
                        salvos += 1
                    except:
                        continue
                
                return f"{salvos} materiais importados automaticamente"
                
            except Exception as e:
                return f"Erro ao importar materiais: {str(e)}"
        
        # Substituir métodos por versões estendidas
        processador.exibir_dados_extraidos = lambda dados, frame: exibir_dados_extraidos_estendido(
            processador, dados, frame, sistema_principal
        )
        processador.criar_opcoes_importacao = lambda frame, origem: criar_opcoes_importacao_estendido(
            processador, frame, origem, sistema_principal
        )
        
        # SUBSTITUIR métodos de importação por versões SIMPLES
        processador.importar_dados_financeiro = importar_dados_financeiro_simples
        processador.importar_dados_material = importar_dados_material_simples
        
        print("✅ Métodos do processador estendidos com 2 fluxos diferentes!")
        
    except Exception as e:
        print(f"❌ Erro ao estender métodos: {e}")


def exibir_dados_extraidos_estendido(processador, dados, frame_container, sistema_principal):
    """
    Versão estendida que exibe dados + AMBAS as opções (simples e avançada)
    """
    import tkinter as tk
    from tkinter import ttk
    
    try:
        # Executar método original primeiro
        processador.exibir_dados_extraidos_original(dados, frame_container)
        
        # Adicionar seção com DUAS opções distintas
        frame_opcoes = ttk.LabelFrame(frame_container, text="")  #, padding=10
        frame_opcoes.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(
            frame_opcoes,
            text="🚀 Processar NFe",
            command=lambda: abrir_processamento_completo(dados, sistema_principal)
        ).pack(pady=5)
        
        # Armazenar dados para uso posterior
        processador.dados_nfe_atual = dados
        
    except Exception as e:
        print(f"❌ Erro ao estender exibição: {e}")
        # Fallback: usar método original
        processador.exibir_dados_extraidos_original(dados, frame_container)

def criar_opcoes_importacao_estendido(processador, frame_container, origem, sistema_principal):
    """
    Versão estendida das opções de importação
    """
    import tkinter as tk
    from tkinter import ttk
    
    try:
        # Executar método original primeiro
        processador.criar_opcoes_importacao_original(frame_container, origem)
        
        # Verificar se há dados NFe carregados
        if hasattr(processador, 'dados_nfe_atual') and processador.dados_nfe_atual:
            # Adicionar opção de processamento completo
            frame_completo = ttk.Frame(frame_container)
            frame_completo.pack(fill='x', pady=10)
            
    except Exception as e:
        print(f"❌ Erro ao estender opções: {e}")
        # Fallback: usar método original
        processador.criar_opcoes_importacao_original(frame_container, origem)


def abrir_processamento_completo(dados_nfe, sistema_principal):
    """
    Abre o integrador completo com os dados da NFe
    """
    try:
        # Verificar se cliente está selecionado
        if not hasattr(sistema_principal, 'cliente_atual') or not sistema_principal.cliente_atual:
            from tkinter import messagebox
            messagebox.showerror("Erro", "Selecione um cliente antes de processar NFe!")
            return
        
        # Verificar se integrador completo está disponível
        if not hasattr(sistema_principal, 'integrador_nfe_completo'):
            from tkinter import messagebox
            messagebox.showerror("Erro", "Integrador completo não inicializado!")
            return
        
        # Abrir interface completa
        sistema_principal.integrador_nfe_completo.criar_interface_integracao_nfe(dados_nfe)
        
        print(f"✅ Processamento completo aberto para NFe {dados_nfe.get('numero_nf', '')}")
        
    except Exception as e:
        from tkinter import messagebox
        messagebox.showerror("Erro", f"Erro ao abrir processamento completo:\n{str(e)}")


def inicializar_sistema_nfe_estendido(sistema_principal):
    """
    Inicializa sistema NFe mantendo original + adicionando extensões
    """
    try:
        print("🚀 Inicializando Sistema NFe Estendido...")
        
        # 1. Inicializar sistema híbrido original (que já funciona)
        from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
        resultado_original = inicializar_sistema_nfe_hibrido(sistema_principal)
        
        if resultado_original:
            print("✅ Sistema híbrido original inicializado")
            
            # 2. Estender com funcionalidades completas
            estender_sistema_hibrido_nfe(sistema_principal)
            
            # 3. APLICAR MELHORIAS DE CERTIFICADO A1 (NOVO)
            try:
                from src.nfe.aplicar_melhorias_certificado import aplicar_melhorias_ao_sistema_existente
                sucesso_cert = aplicar_melhorias_ao_sistema_existente(sistema_principal)
                
                if sucesso_cert:
                    print("✅ Melhorias de certificado A1 aplicadas!")
                else:
                    print("⚠️ Melhorias de certificado A1 parcialmente aplicadas")
                    
            except ImportError:
                print("⚠️ Módulo de certificado A1 não encontrado - funcionalidades básicas mantidas")
            except Exception as e:
                print(f"⚠️ Erro ao aplicar melhorias A1: {e}")
            
            print("✅ Sistema NFe Estendido inicializado com sucesso!")
            print("🔄 Funcionalidades disponíveis:")
            print("   • Importar NFe (original) + botão 'Processar Completo'")
            print("   • Interface integrada financeiro + materiais")
            print("   • Consulta SEFAZ com certificado A1 (se configurado)")
            
            return sistema_principal.integrador_nfe_completo if hasattr(sistema_principal, 'integrador_nfe_completo') else True
        else:
            print("❌ Falha ao inicializar sistema híbrido original")
            return None
            
    except Exception as e:
        print(f"❌ Erro na inicialização estendida: {e}")
        # Fallback: tentar apenas sistema original
        try:
            from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
            return inicializar_sistema_nfe_hibrido(sistema_principal)
        except:
            return None


def testar_extensao_manualmente(sistema_principal):
    """
    Função para testar a extensão sem alterar o __init__
    Execute esta função no console após inicializar o sistema
    """
    try:
        print("🧪 Testando extensão NFe manualmente...")
        
        resultado = estender_sistema_hibrido_nfe(sistema_principal)
        
        if resultado:
            print("✅ Extensão aplicada com sucesso!")
            print("📄 Agora importe um XML para ver o botão 'Processar NFe Completa'")
            return True
        else:
            print("❌ Falha ao aplicar extensão")
            return False
            
    except Exception as e:
        print(f"❌ Erro no teste manual: {e}")
        return False


def diagnosticar_sistema_nfe(sistema_principal):
    """
    Diagnostica o estado atual do sistema NFe
    """
    print("\n🔍 DIAGNÓSTICO DO SISTEMA NFe:")
    print("=" * 40)
    
    # Verificar sistema híbrido
    if hasattr(sistema_principal, 'processador_nfe'):
        print("✅ Sistema híbrido: INICIALIZADO")
        
        # Verificar se tem integrador NFe
        if hasattr(sistema_principal, 'integrador_nfe'):
            print("✅ Integrador NFe: PRESENTE")
        else:
            print("❌ Integrador NFe: AUSENTE")
        
        # Verificar se tem integrador completo
        if hasattr(sistema_principal, 'integrador_nfe_completo'):
            print("✅ Integrador completo: PRESENTE")
        else:
            print("❌ Integrador completo: AUSENTE")
        
        # Verificar se métodos foram estendidos
        if hasattr(sistema_principal.processador_nfe, 'exibir_dados_extraidos_original'):
            print("✅ Métodos estendidos: SIM")
        else:
            print("❌ Métodos estendidos: NÃO")
            
    else:
        print("❌ Sistema híbrido: NÃO INICIALIZADO")
    
    print("=" * 40)