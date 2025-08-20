# -*- coding: utf-8 -*-
"""
PATCH FUNCIONAL NFe - AJUSTES CRÍTICOS
Corrige apenas data (5/20) e referência editável
SEM quebrar o sistema existente
"""

from datetime import datetime
import tkinter as tk
from tkinter import ttk

def aplicar_patch_funcional_nfe(sistema_principal):
    """
    Aplica APENAS os ajustes funcionais críticos:
    1. Data sempre 5 ou 20 do período atual
    2. Campo referência editável
    """
    try:
        print("🔧 Aplicando patch funcional NFe...")
        
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        sistema_nfe = sistema_principal.sistema_nfe_unificado
        
        # PATCH 1: Corrigir método de criação de lançamento financeiro
        original_criar_lancamento = sistema_nfe.criar_lancamento_financeiro
        
        def criar_lancamento_financeiro_corrigido(self):
            """Versão corrigida com data 5/20 e referência editável"""
            try:
                dados_nfe = self.dados_nfe_atual
                
                # CALCULAR DATA DO PERÍODO ATUAL (sempre 5 ou 20 de agora)
                hoje = datetime.now()
                if hoje.day <= 15:
                    data_rel = hoje.replace(day=5).strftime('%d/%m/%Y')
                else:
                    data_rel = hoje.replace(day=20).strftime('%d/%m/%Y')
                
                # DATA DE VENCIMENTO = DATA DA NFE (original)
                dt_vencto = dados_nfe.get('data_emissao', hoje.strftime('%d/%m/%Y'))
                
                # GERAR REFERÊNCIA INTELIGENTE (editável na interface)
                ref_base = f"NFE {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')[:25]}"
                
                print(f"📅 Data corrigida: NFe {dt_vencto} → Relatório {data_rel}")
                print(f"📋 Referência: {ref_base}")
                
                # DADOS FINANCEIROS CORRIGIDOS
                dados_financeiros = {
                    'data': data_rel,  # SEMPRE 5 ou 20 do período atual
                    'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                    'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                    'categoria': 'MAT',  # Padrão para material
                    'tp_desp': '3',  # Padrão para material
                    'referencia': ref_base.upper(),
                    'etapa_obra': 'MATERIAIS',  # Padrão
                    'nf': dados_nfe.get('numero_nf', ''),
                    'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dias': 1,
                    'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dt_vencto': dt_vencto,  # Data original da NFe
                    'dados_bancarios': '',
                    'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - PERIODO {data_rel}".upper(),
                    'forma_pagamento': 'A_PRAZO'
                }
                
                # ADICIONAR À LISTA DO SISTEMA
                if not hasattr(self.sistema, 'dados_para_incluir'):
                    self.sistema.dados_para_incluir = []
                
                self.sistema.dados_para_incluir.append(dados_financeiros)
                
                return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
                
            except Exception as e:
                raise Exception(f"Erro ao criar lançamento: {str(e)}")
        
        # APLICAR PATCH
        sistema_nfe.criar_lancamento_financeiro = criar_lancamento_financeiro_corrigido.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        # PATCH 2: Adicionar interface para editar referência
        original_criar_opcoes = sistema_nfe.criar_opcoes_importacao
        
        def criar_opcoes_importacao_com_referencia(self):
            """Versão com campo de referência editável"""
            # CHAMAR MÉTODO ORIGINAL
            original_criar_opcoes(self)
            
            # ADICIONAR CAMPO DE REFERÊNCIA EDITÁVEL
            if hasattr(self, 'frame_opcoes'):
                try:
                    self.adicionar_campo_referencia_editavel()
                except Exception as e:
                    print(f"⚠️ Erro ao adicionar campo referência: {e}")
        
        def adicionar_campo_referencia_editavel(self):
            """Adiciona campo de referência editável"""
            if not hasattr(self, 'frame_opcoes'):
                return
            
            # FRAME PARA REFERÊNCIA
            frame_ref = ttk.LabelFrame(self.frame_opcoes, text="📋 Referência do Lançamento", padding=10)
            frame_ref.pack(fill='x', pady=5)
            
            # GERAR REFERÊNCIA PADRÃO
            dados_nfe = self.dados_nfe_atual
            ref_padrao = f"NFE {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')[:30]}"
            
            # LABEL + ENTRY
            tk.Label(frame_ref, text="Referência para o relatório:", 
                    font=('Arial', 9, 'bold')).pack(anchor='w', pady=2)
            
            self.referencia_editavel = tk.Entry(frame_ref, width=80, font=('Arial', 9))
            self.referencia_editavel.pack(fill='x', pady=2)
            self.referencia_editavel.insert(0, ref_padrao)
            
            # DICA
            tk.Label(frame_ref, text="💡 Esta referência aparecerá nos relatórios quinzenais para o cliente", 
                    fg='blue', font=('Arial', 8)).pack(anchor='w', pady=2)
        
        # APLICAR PATCHES
        sistema_nfe.criar_opcoes_importacao = criar_opcoes_importacao_com_referencia.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        sistema_nfe.adicionar_campo_referencia_editavel = adicionar_campo_referencia_editavel.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        # PATCH 3: Modificar criação de lançamento para usar referência editável
        def criar_lancamento_com_referencia_editavel(self):
            """Cria lançamento usando referência editável"""
            try:
                dados_nfe = self.dados_nfe_atual
                
                # DATA SEMPRE DO PERÍODO ATUAL
                hoje = datetime.now()
                if hoje.day <= 15:
                    data_rel = hoje.replace(day=5).strftime('%d/%m/%Y')
                else:
                    data_rel = hoje.replace(day=20).strftime('%d/%m/%Y')
                
                dt_vencto = dados_nfe.get('data_emissao', hoje.strftime('%d/%m/%Y'))
                
                # USAR REFERÊNCIA EDITÁVEL SE EXISTIR
                if hasattr(self, 'referencia_editavel'):
                    referencia = self.referencia_editavel.get().strip().upper()
                else:
                    referencia = f"NFE {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')[:25]}".upper()
                
                # DADOS FINANCEIROS
                dados_financeiros = {
                    'data': data_rel,
                    'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                    'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                    'categoria': 'MAT',
                    'tp_desp': '3',
                    'referencia': referencia,  # REFERÊNCIA EDITÁVEL
                    'etapa_obra': 'MATERIAIS',
                    'nf': dados_nfe.get('numero_nf', ''),
                    'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dias': 1,
                    'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dt_vencto': dt_vencto,
                    'dados_bancarios': '',
                    'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - PERIODO {data_rel} - REF: {referencia[:30]}".upper(),
                    'forma_pagamento': 'A_PRAZO'
                }
                
                if not hasattr(self.sistema, 'dados_para_incluir'):
                    self.sistema.dados_para_incluir = []
                
                self.sistema.dados_para_incluir.append(dados_financeiros)
                
                print(f"💰 Lançamento criado:")
                print(f"   📅 Data Relatório: {data_rel} (período atual)")
                print(f"   📅 Data Vencimento: {dt_vencto} (NFe)")
                print(f"   📋 Referência: {referencia}")
                print(f"   💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}")
                
                return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
                
            except Exception as e:
                raise Exception(f"Erro ao criar lançamento: {str(e)}")
        
        # APLICAR PATCH FINAL
        sistema_nfe.criar_lancamento_financeiro = criar_lancamento_com_referencia_editavel.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        print("✅ Patch funcional NFe aplicado com sucesso!")
        print("📌 Ajustes aplicados:")
        print(f"   📅 Data sempre período atual: {datetime.now().day <= 15 and '5' or '20'}/{datetime.now().month:02d}/{datetime.now().year}")
        print("   📋 Campo referência editável adicionado")
        print("   🎯 Focado nos relatórios quinzenais")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar patch: {e}")
        return False

def mostrar_info_periodo_atual():
    """Mostra informações sobre o período atual"""
    hoje = datetime.now()
    
    if hoje.day <= 15:
        periodo = "PRIMEIRA QUINZENA"
        data_rel = hoje.replace(day=5).strftime('%d/%m/%Y')
        proximo_periodo = hoje.replace(day=20).strftime('%d/%m/%Y')
    else:
        periodo = "SEGUNDA QUINZENA"
        data_rel = hoje.replace(day=20).strftime('%d/%m/%Y')
        # Próximo período é dia 5 do mês seguinte
        if hoje.month == 12:
            proximo_periodo = hoje.replace(year=hoje.year+1, month=1, day=5).strftime('%d/%m/%Y')
        else:
            proximo_periodo = hoje.replace(month=hoje.month+1, day=5).strftime('%d/%m/%Y')
    
    print(f"""
📊 INFORMAÇÕES DO PERÍODO ATUAL:
   📅 Data: {hoje.strftime('%d/%m/%Y')}
   📋 Período: {periodo}
   🎯 Data Relatório: {data_rel}
   ⏭️ Próximo Período: {proximo_periodo}
   
💡 Todas as NFe importadas AGORA entrarão no relatório de {data_rel}
   independente da data original da nota fiscal.
""")

# EXEMPLO DE USO
"""
PARA APLICAR APENAS OS AJUSTES FUNCIONAIS CRÍTICOS:

# No final do __init__ do SistemaEntradaDados:
try:
    from src.nfe.patch_funcional_nfe import aplicar_patch_funcional_nfe
    aplicar_patch_funcional_nfe(self)
    print("✅ Patch funcional NFe aplicado!")
except Exception as e:
    print(f"⚠️ Patch não aplicado: {e}")

RESULTADO:
✅ Data sempre 5 ou 20 do período ATUAL (não da NFe)
✅ Campo referência editável na interface
✅ Sistema continua funcionando exatamente igual
✅ Relatórios quinzenais corretos
✅ Base para cálculo de taxa de administração correta

EXEMPLO PRÁTICO (hoje é 13/08/2025):
- NFe de 07/11/2024 → Data relatório: 20/08/2025
- NFe de qualquer data → Data relatório: 20/08/2025
- Referência editável: "NFE 10059 - CARAMELODRAMA PRODUTOS..."
"""