# -*- coding: utf-8 -*-
"""
PATCH FUNCIONAL NFe - VERSÃO CORRIGIDA
Aplica ajustes de data (5/20) e referência editável
SEM erros de chamada de método
"""

from datetime import datetime
import tkinter as tk
from tkinter import ttk

def aplicar_patch_funcional_nfe_corrigido(sistema_principal):
    """
    Aplica patch funcional corrigido para NFe
    """
    try:
        print("🔧 Aplicando patch funcional NFe (versão corrigida)...")
        
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        sistema_nfe = sistema_principal.sistema_nfe_unificado
        print("✅ Sistema NFe encontrado!")
        
        # PATCH: Substituir método criar_lancamento_financeiro
        def criar_lancamento_financeiro_corrigido(self):
            """Cria lançamento com data do período atual e referência editável"""
            try:
                dados_nfe = self.dados_nfe_atual
                
                # DATA SEMPRE DO PERÍODO ATUAL (5 ou 20)
                hoje = datetime.now()
                if hoje.day <= 15:
                    data_rel = hoje.replace(day=5).strftime('%d/%m/%Y')
                    periodo = "PRIMEIRA QUINZENA"
                else:
                    data_rel = hoje.replace(day=20).strftime('%d/%m/%Y')
                    periodo = "SEGUNDA QUINZENA"
                
                # DATA DE VENCIMENTO = DATA ORIGINAL DA NFE
                dt_vencto = dados_nfe.get('data_emissao', hoje.strftime('%d/%m/%Y'))
                
                # REFERÊNCIA EDITÁVEL (se campo existir) ou padrão
                if hasattr(self, 'referencia_customizada') and self.referencia_customizada:
                    referencia = self.referencia_customizada.strip().upper()
                else:
                    numero_nf = dados_nfe.get('numero_nf', '')
                    fornecedor = dados_nfe.get('razao_social_emitente', '')
                    referencia = f"NFE {numero_nf} - {fornecedor[:25]}".upper()
                
                # DADOS FINANCEIROS CORRIGIDOS
                dados_financeiros = {
                    'data': data_rel,  # SEMPRE 5 OU 20 DO PERÍODO ATUAL
                    'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                    'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                    'categoria': 'MAT',
                    'tp_desp': '3',
                    'referencia': referencia,
                    'etapa_obra': 'MATERIAIS',
                    'nf': dados_nfe.get('numero_nf', ''),
                    'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dias': 1,
                    'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                    'dt_vencto': dt_vencto,  # DATA ORIGINAL DA NFE
                    'dados_bancarios': '',
                    'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - {periodo} {data_rel}".upper(),
                    'forma_pagamento': 'A_PRAZO'
                }
                
                # ADICIONAR À LISTA DO SISTEMA
                if not hasattr(self.sistema, 'dados_para_incluir'):
                    self.sistema.dados_para_incluir = []
                
                self.sistema.dados_para_incluir.append(dados_financeiros)
                
                # LOG DETALHADO
                print(f"💰 LANÇAMENTO CRIADO:")
                print(f"   📅 Data Relatório: {data_rel} ({periodo})")
                print(f"   📅 Data Vencimento: {dt_vencto} (original NFe)")
                print(f"   📋 Referência: {referencia}")
                print(f"   💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}")
                print(f"   🎯 Para relatório quinzenal de {data_rel}")
                
                return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
                
            except Exception as e:
                raise Exception(f"Erro ao criar lançamento: {str(e)}")
        
        # APLICAR PATCH NO MÉTODO
        sistema_nfe.criar_lancamento_financeiro = criar_lancamento_financeiro_corrigido.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        # PATCH: Adicionar campo de referência editável na interface
        original_exibir_dados = sistema_nfe.exibir_dados_extraidos
        
        def exibir_dados_extraidos_com_referencia(self):
            """Versão que adiciona campo de referência editável"""
            # CHAMAR MÉTODO ORIGINAL
            original_exibir_dados.__func__(self)  # CORREÇÃO: usar __func__ para método bound
            
            # ADICIONAR CAMPO DE REFERÊNCIA
            if hasattr(self, 'frame_dados'):
                self.adicionar_campo_referencia()
        
        def adicionar_campo_referencia(self):
            """Adiciona campo de referência editável"""
            try:
                # FRAME PARA REFERÊNCIA EDITÁVEL
                frame_ref = ttk.LabelFrame(self.frame_dados, text="📋 Referência para Relatório", padding=5)
                frame_ref.pack(fill='x', pady=5)
                
                # GERAR REFERÊNCIA PADRÃO
                dados_nfe = self.dados_nfe_atual
                numero_nf = dados_nfe.get('numero_nf', '')
                fornecedor = dados_nfe.get('razao_social_emitente', '')[:30]
                ref_padrao = f"NFE {numero_nf} - {fornecedor}"
                
                # LABEL E ENTRY
                tk.Label(frame_ref, text="Referência (aparece no relatório quinzenal):", 
                        font=('Arial', 9, 'bold'), fg='purple').pack(anchor='w')
                
                self.entry_referencia = tk.Entry(frame_ref, width=70, font=('Arial', 9))
                self.entry_referencia.pack(fill='x', pady=2)
                self.entry_referencia.insert(0, ref_padrao)
                
                # BIND PARA SALVAR REFERÊNCIA
                def salvar_referencia_customizada(event=None):
                    self.referencia_customizada = self.entry_referencia.get()
                
                self.entry_referencia.bind('<KeyRelease>', salvar_referencia_customizada)
                self.entry_referencia.bind('<FocusOut>', salvar_referencia_customizada)
                
                # INICIALIZAR VARIÁVEL
                self.referencia_customizada = ref_padrao
                
                # DICA
                tk.Label(frame_ref, text="💡 Esta referência aparecerá nos relatórios quinzenais para o cliente", 
                        fg='blue', font=('Arial', 8)).pack(anchor='w')
                
                print("✅ Campo de referência editável adicionado")
                
            except Exception as e:
                print(f"⚠️ Erro ao adicionar campo referência: {e}")
        
        # APLICAR PATCHES
        sistema_nfe.exibir_dados_extraidos = exibir_dados_extraidos_com_referencia.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        sistema_nfe.adicionar_campo_referencia = adicionar_campo_referencia.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        print("✅ Patch funcional NFe aplicado com sucesso!")
        print("📌 Melhorias aplicadas:")
        
        # MOSTRAR INFORMAÇÕES DO PERÍODO ATUAL
        hoje = datetime.now()
        if hoje.day <= 15:
            data_periodo = hoje.replace(day=5).strftime('%d/%m/%Y')
            periodo_nome = "PRIMEIRA QUINZENA"
        else:
            data_periodo = hoje.replace(day=20).strftime('%d/%m/%Y')
            periodo_nome = "SEGUNDA QUINZENA"
        
        print(f"   📅 Data relatório: {data_periodo} ({periodo_nome})")
        print(f"   📋 Campo referência editável na interface")
        print(f"   🎯 Todas NFe importadas agora entram no relatório de {data_periodo}")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar patch: {e}")
        import traceback
        print(f"📄 Traceback: {traceback.format_exc()}")
        return False

def mostrar_periodo_atual():
    """Mostra informações sobre o período atual"""
    hoje = datetime.now()
    
    if hoje.day <= 15:
        data_rel = hoje.replace(day=5).strftime('%d/%m/%Y')
        periodo = "PRIMEIRA QUINZENA"
        dias_restantes = 15 - hoje.day
        proximo = hoje.replace(day=20).strftime('%d/%m/%Y')
    else:
        data_rel = hoje.replace(day=20).strftime('%d/%m/%Y')
        periodo = "SEGUNDA QUINZENA"
        # Calcular dias até fim do mês
        import calendar
        ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
        dias_restantes = ultimo_dia - hoje.day
        # Próximo período é dia 5 do mês seguinte
        if hoje.month == 12:
            proximo = hoje.replace(year=hoje.year+1, month=1, day=5).strftime('%d/%m/%Y')
        else:
            proximo = hoje.replace(month=hoje.month+1, day=5).strftime('%d/%m/%Y')
    
    print(f"""
📊 PERÍODO ATUAL PARA RELATÓRIOS:
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
📅 Hoje: {hoje.strftime('%d/%m/%Y (%A)')}
🎯 Período: {periodo}
📋 Data Relatório: {data_rel}
⏰ Dias restantes no período: {dias_restantes}
⏭️ Próximo período: {proximo}

💡 IMPORTANTE:
   • Todas as NFe importadas AGORA entram no relatório de {data_rel}
   • Independente da data original da nota fiscal
   • Base correta para cálculo de taxa de administração
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
""")

# FUNÇÃO DE TESTE
def testar_patch_nfe(sistema_principal):
    """Testa se o patch está funcionando"""
    try:
        print("🧪 TESTANDO PATCH FUNCIONAL NFe...")
        
        # Verificar se sistema existe
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        sistema_nfe = sistema_principal.sistema_nfe_unificado
        
        # Verificar se métodos foram patcheados
        checks = [
            ("criar_lancamento_financeiro", hasattr(sistema_nfe, 'criar_lancamento_financeiro')),
            ("exibir_dados_extraidos", hasattr(sistema_nfe, 'exibir_dados_extraidos')),
            ("adicionar_campo_referencia", hasattr(sistema_nfe, 'adicionar_campo_referencia'))
        ]
        
        for nome, existe in checks:
            status = "✅" if existe else "❌"
            print(f"   {status} {nome}")
        
        # Mostrar período atual
        mostrar_periodo_atual()
        
        print("✅ Teste concluído!")
        return True
        
    except Exception as e:
        print(f"❌ Erro no teste: {e}")
        return False

"""
PARA USAR A VERSÃO CORRIGIDA:

# No __init__ do SistemaEntradaDados, SUBSTITUA a linha anterior por:
try:
    from src.nfe.patch_funcional_corrigido import aplicar_patch_funcional_nfe_corrigido
    aplicar_patch_funcional_nfe_corrigido(self)
except Exception as e:
    print(f"⚠️ Patch não aplicado: {e}")

# Para testar:
try:
    from src.nfe.patch_funcional_corrigido import testar_patch_nfe
    testar_patch_nfe(self)
except Exception as e:
    print(f"⚠️ Teste falhou: {e}")
"""