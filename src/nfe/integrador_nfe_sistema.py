# -*- coding: utf-8 -*-
"""
Integrador NFe com Sistema Financeiro e Materiais - VERSÃO LIMPA
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from dateutil.relativedelta import relativedelta

class IntegradorNFeFinanceiroMateriais:
    """Integra dados da NFe com o sistema financeiro e de materiais"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.dados_nfe_atual = None
        
    def criar_interface_integracao_nfe(self, dados_nfe):
        """Cria interface para integração da NFe"""
        self.dados_nfe_atual = dados_nfe
        
        # Criar janela principal
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("Integração NFe - Financeiro e Materiais")
        self.janela.geometry("900x700")
        self.janela.grab_set()
        
        # Título
        titulo = tk.Label(
            self.janela,
            text="🚀 PROCESSAMENTO COMPLETO DE NFe",
            font=('Arial', 16, 'bold'),
            fg='#0056b3'
        )
        titulo.pack(pady=10)
        
        # Informações da NFe
        self.criar_info_nfe(dados_nfe)
        
        # Configurações
        self.criar_configuracoes()
        
        # Botões
        self.criar_botoes()
        
    def criar_info_nfe(self, dados_nfe):
        """Cria seção de informações da NFe"""
        frame_info = ttk.LabelFrame(self.janela, text="📋 Informações da NFe", padding=10)
        frame_info.pack(fill='x', padx=10, pady=5)
        
        # Dados principais
        info_texto = f"""
🏢 Fornecedor: {dados_nfe.get('razao_social_emitente', '')}
📋 CNPJ: {self.formatar_cnpj(dados_nfe.get('cnpj_emitente', ''))}
📄 NFe: {dados_nfe.get('numero_nf', '')}
📅 Data: {dados_nfe.get('data_emissao', '')}
💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}
📦 Produtos: {len(dados_nfe.get('produtos', []))} itens
        """
        
        tk.Label(frame_info, text=info_texto.strip(), 
                font=('Arial', 10), justify='left').pack(anchor='w')
    
    def criar_configuracoes(self):
        """Cria seção de configurações"""
        frame_config = ttk.LabelFrame(self.janela, text="⚙️ Configurações", padding=10)
        frame_config.pack(fill='both', expand=True, padx=10, pady=5)
        
        # === FINANCEIRO ===
        self.incluir_financeiro_var = tk.BooleanVar(value=True)
        cb_financeiro = tk.Checkbutton(
            frame_config,
            text="💰 Incluir lançamento financeiro",
            variable=self.incluir_financeiro_var,
            font=('Arial', 11, 'bold')
        )
        cb_financeiro.pack(anchor='w', pady=5)
        
        # Frame financeiro
        self.frame_financeiro = ttk.LabelFrame(frame_config, text="Dados Financeiros", padding=10)
        self.frame_financeiro.pack(fill='x', pady=5)
        
        # Data de referência
        tk.Label(self.frame_financeiro, text="Data Referência:", 
                font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        
        data_ref = self.calcular_data_referencia()
        tk.Label(self.frame_financeiro, text=data_ref, 
                fg='blue').grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        # Tipo de despesa
        tk.Label(self.frame_financeiro, text="Tipo Despesa:", 
                font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        
        self.tipo_despesa_var = tk.StringVar(value="3")
        tipo_combo = ttk.Combobox(
            self.frame_financeiro,
            textvariable=self.tipo_despesa_var,
            values=["2", "3", "5", "6"],
            state="readonly",
            width=5
        )
        tipo_combo.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # Referência
        tk.Label(self.frame_financeiro, text="Referência:", 
                font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        
        self.referencia_var = tk.StringVar(value="MATERIAL VIA NFE")
        ref_entry = tk.Entry(self.frame_financeiro, textvariable=self.referencia_var, width=30)
        ref_entry.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        # Data vencimento (usar data da NFe, não hoje)
        tk.Label(self.frame_financeiro, text="Vencimento:", 
                font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky='w', padx=5, pady=5)
        
        from tkcalendar import DateEntry
        self.data_vencimento = DateEntry(
            self.frame_financeiro,
            format='dd/mm/yyyy',
            locale='pt_BR'
        )
        self.data_vencimento.grid(row=3, column=1, sticky='w', padx=5, pady=5)
        
        # CORREÇÃO: Definir data vencimento como data da NFe
        try:
            data_nfe = datetime.strptime(self.dados_nfe_atual.get('data_emissao', ''), '%d/%m/%Y')
            self.data_vencimento.set_date(data_nfe.date())
        except:
            # Fallback: usar data de hoje se não conseguir ler data da NFe
            self.data_vencimento.set_date(datetime.now().date())
        
        # === MATERIAIS ===
        self.incluir_materiais_var = tk.BooleanVar(value=True)
        cb_materiais = tk.Checkbutton(
            frame_config,
            text="📦 Incluir materiais no controle de obra",
            variable=self.incluir_materiais_var,
            font=('Arial', 11, 'bold')
        )
        cb_materiais.pack(anchor='w', pady=(20, 5))
        
        # Frame materiais
        frame_materiais = ttk.LabelFrame(frame_config, text="Configuração dos Materiais", padding=10)
        frame_materiais.pack(fill='x', pady=5)
        
        tk.Label(frame_materiais, text="Ambiente:", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        self.ambiente_var = tk.StringVar(value="DEPÓSITO DA OBRA")
        ambiente_combo = ttk.Combobox(
            frame_materiais,
            textvariable=self.ambiente_var,
            values=[
                "DEPÓSITO DA OBRA", "SALA", "COZINHA", "BANHEIRO SUITE",
                "QUARTO CASAL", "ÁREA EXTERNA", "TODOS AMBIENTES"
            ],
            state="readonly"
        )
        ambiente_combo.pack(anchor='w', pady=5)
        
        # Produtos
        produtos = self.dados_nfe_atual.get('produtos', []) if self.dados_nfe_atual else []
        if produtos:
            tk.Label(frame_materiais, 
                    text=f"📦 {len(produtos)} produtos serão importados",
                    font=('Arial', 10)).pack(anchor='w', pady=5)
        
    def criar_botoes(self):
        """Cria botões principais"""
        frame_botoes = ttk.Frame(self.janela)
        frame_botoes.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(frame_botoes, text="❌ Cancelar",
                  command=self.janela.destroy).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, text="💾 Processar e Salvar",
                  command=self.processar_e_salvar).pack(side='right', padx=5)
    
    def calcular_data_referencia(self):
        """
        Calcula data de referência seguindo a regra 5/20:
        - Se hoje está entre 21 e 5: data é 5 do mês corrente
        - Caso contrário: dia 20 do mês corrente
        """
        try:
            # Usar data de HOJE para calcular (não da NFe)
            hoje = datetime.now()
            
            # REGRA CORRIGIDA: Entre 21 e 5 → dia 5, senão → dia 20
            if hoje.day >= 21 or hoje.day <= 5:
                # Entre 21 e 5 → dia 5 do mês corrente
                data_rel = hoje.replace(day=5)
            else:
                # Entre 6 e 20 → dia 20 do mês corrente  
                data_rel = hoje.replace(day=20)
            
            return data_rel.strftime('%d/%m/%Y')
            
        except:
            return datetime.now().strftime('%d/%m/%Y')
    
    def formatar_cnpj(self, cnpj):
        """Formata CNPJ"""
        if not cnpj or len(cnpj) != 14:
            return cnpj
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    
    def processar_e_salvar(self):
        """Processa e salva os dados"""
        try:
            if not hasattr(self.sistema, 'cliente_atual') or not self.sistema.cliente_atual:
                messagebox.showerror("Erro", "Selecione um cliente antes de processar!")
                return
            
            if not self.incluir_financeiro_var.get() and not self.incluir_materiais_var.get():
                messagebox.showwarning("Aviso", "Selecione pelo menos uma opção!")
                return
            
            # Confirmar
            msg = f"Processar NFe {self.dados_nfe_atual.get('numero_nf', '')}?\n\n"
            msg += f"Cliente: {self.sistema.cliente_atual}\n"
            msg += f"Valor: R$ {self.dados_nfe_atual.get('valor_total', 0):,.2f}"
            
            if not messagebox.askyesno("Confirmar", msg):
                return
            
            resultados = []
            
            # Processar financeiro
            if self.incluir_financeiro_var.get():
                resultado_fin = self.salvar_financeiro()
                resultados.append(f"💰 Financeiro: {resultado_fin}")
            
            # Processar materiais
            if self.incluir_materiais_var.get():
                resultado_mat = self.salvar_materiais()
                resultados.append(f"📦 Materiais: {resultado_mat}")
            
            # Mostrar resultado
            mensagem = "✅ PROCESSAMENTO CONCLUÍDO!\n\n"
            for resultado in resultados:
                mensagem += f"{resultado}\n"
            
            messagebox.showinfo("Sucesso", mensagem)
            self.janela.destroy()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")
    
    def salvar_financeiro(self):
        """Salva dados financeiros"""
        try:
            dados_financeiro = {
                'data': self.calcular_data_referencia(),
                'cnpj_cpf': self.dados_nfe_atual.get('cnpj_emitente', ''),
                'nome': self.dados_nfe_atual.get('razao_social_emitente', ''),
                'categoria': 'MAT',
                'tp_desp': self.tipo_despesa_var.get(),
                'referencia': self.referencia_var.get().upper(),
                'etapa_obra': '',
                'nf': self.dados_nfe_atual.get('numero_nf', ''),
                'vr_unit': f"{self.dados_nfe_atual.get('valor_total', 0):.2f}",
                'dias': 1,
                'valor': f"{self.dados_nfe_atual.get('valor_total', 0):.2f}",
                'dt_vencto': self.data_vencimento.get(),
                'dados_bancarios': '',
                'observacao': f"NFE {self.dados_nfe_atual.get('numero_nf', '')}",
                'forma_pagamento': ''
            }
            
            # Adicionar aos dados do sistema
            self.sistema.dados_para_incluir = [dados_financeiro]
            
            # Chamar método de envio
            self.sistema.enviar_dados()
            
            return "Salvo com sucesso"
            
        except Exception as e:
            return f"Erro: {str(e)}"
    
    def salvar_materiais(self):
        """Salva materiais"""
        try:
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                from src.materiais.gerenciador_materiais import GerenciadorMateriais
                self.sistema.gerenciador_materiais = GerenciadorMateriais(self.sistema)
            
            produtos = self.dados_nfe_atual.get('produtos', [])
            salvos = 0
            
            for produto in produtos:
                material = {
                    'Cliente': self.sistema.cliente_atual,
                    'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                    'Codigo_Produto': produto.get('codigo', ''),
                    'Descricao_Completa': produto.get('descricao', ''),
                    'Ambiente_Aplicacao': self.ambiente_var.get(),
                    'Status_Instalacao': 'PENDENTE',
                    'Tem_Dados_Compra': True,
                    'Nome_Fornecedor': self.dados_nfe_atual.get('razao_social_emitente', ''),
                    'CNPJ_Fornecedor': self.dados_nfe_atual.get('cnpj_emitente', ''),
                    'Data_Compra': self.dados_nfe_atual.get('data_emissao', ''),
                    'Quantidade': produto.get('quantidade', 0),
                    'Unidade': produto.get('unidade', 'UN'),
                    'Valor_Unitario': produto.get('valor_unitario', 0),
                    'Valor_Total': produto.get('valor_total', 0),
                    'Numero_NF': self.dados_nfe_atual.get('numero_nf', ''),
                    'Observacoes': f"Importado NFe {self.dados_nfe_atual.get('numero_nf', '')}"
                }
                
                try:
                    self.sistema.gerenciador_materiais.salvar_material(material)
                    salvos += 1
                except:
                    continue
            
            return f"{salvos} itens salvos"
            
        except Exception as e:
            return f"Erro: {str(e)}"