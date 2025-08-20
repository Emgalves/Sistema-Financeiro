# -*- coding: utf-8 -*-
"""
Integrador NFe com Sistema Financeiro e Materiais - VERSÃO CORRIGIDA
Integra dados da NFe com o sistema financeiro e de materiais
"""

import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pandas as pd
from pathlib import Path
import json

class IntegradorNFeFinanceiroMateriais:
    """
    Integra dados da NFe com o sistema financeiro e de materiais
    Respeita todas as regras e lógicas existentes
    """
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.dados_nfe_atual = None
        
    def criar_interface_integracao_nfe(self, dados_nfe):
        """
        Cria interface para integração da NFe com escolhas do usuário
        """
        self.dados_nfe_atual = dados_nfe
        
        # Criar janela principal
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("Integração NFe - Financeiro e Materiais")
        self.janela.geometry("900x800")
        self.janela.grab_set()
        
        # Criar notebook para organizar
        notebook = ttk.Notebook(self.janela)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Aba 1: Resumo da NFe
        self.criar_aba_resumo_nfe(notebook, dados_nfe)
        
        # Aba 2: Configuração Financeira
        self.criar_aba_configuracao_financeira(notebook, dados_nfe)
        
        # Aba 3: Seleção de Materiais
        self.criar_aba_selecao_materiais(notebook, dados_nfe)
        
        # Botões principais
        self.criar_botoes_principais()
        
    def criar_aba_resumo_nfe(self, notebook, dados_nfe):
        """Aba com resumo da NFe"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📄 Resumo NFe")
        
        # Frame scrollável
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Informações da NFe
        info_frame = ttk.LabelFrame(scrollable_frame, text="📋 Informações da NFe", padding=10)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        informacoes = [
            ("🏢 Emitente:", dados_nfe.get('razao_social_emitente', '')),
            ("📋 CNPJ:", self.formatar_cnpj(dados_nfe.get('cnpj_emitente', ''))),
            ("📄 Número NFe:", dados_nfe.get('numero_nf', '')),
            ("📅 Data Emissão:", dados_nfe.get('data_emissao', '')),
            ("💰 Valor Total:", f"R$ {dados_nfe.get('valor_total', 0):,.2f}"),
            ("📦 Produtos:", f"{len(dados_nfe.get('produtos', []))} itens"),
            ("🔗 Fonte:", dados_nfe.get('fonte_dados', ''))
        ]
        
        for i, (label, valor) in enumerate(informacoes):
            row = i // 2
            col = (i % 2) * 2
            
            tk.Label(info_frame, text=label, font=('Arial', 9, 'bold')).grid(
                row=row, column=col, sticky='w', padx=5, pady=3)
            tk.Label(info_frame, text=str(valor)[:60]).grid(
                row=row, column=col+1, sticky='w', padx=10, pady=3)
        
        # Lista de produtos (resumida)
        if dados_nfe.get('produtos'):
            produtos_frame = ttk.LabelFrame(scrollable_frame, text="📦 Produtos (Prévia)", padding=10)
            produtos_frame.pack(fill='both', expand=True, padx=10, pady=5)
            
            # TreeView simples
            tree = ttk.Treeview(produtos_frame, columns=('desc', 'qtd', 'valor'), show='headings', height=6)
            tree.heading('desc', text='Descrição')
            tree.heading('qtd', text='Qtd')
            tree.heading('valor', text='Valor')
            
            tree.column('desc', width=400)
            tree.column('qtd', width=80)
            tree.column('valor', width=100)
            
            for produto in dados_nfe['produtos'][:10]:  # Mostrar apenas primeiros 10
                tree.insert('', 'end', values=(
                    produto.get('descricao', '')[:50],
                    produto.get('quantidade', ''),
                    f"R$ {produto.get('valor_total', 0):.2f}"
                ))
            
            tree.pack(fill='both', expand=True)
            
            if len(dados_nfe['produtos']) > 10:
                tk.Label(produtos_frame, text=f"... e mais {len(dados_nfe['produtos']) - 10} produtos", 
                        font=('Arial', 8), fg='gray').pack(pady=5)
        
        # Configurar scroll
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def criar_aba_configuracao_financeira(self, notebook, dados_nfe):
        """Aba para configurar lançamento financeiro"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="💰 Configuração Financeira")
        
        # Checkbox para incluir no financeiro
        self.incluir_financeiro_var = tk.BooleanVar(value=True)
        cb_financeiro = tk.Checkbutton(
            frame, 
            text="✅ Incluir lançamento financeiro desta NFe",
            variable=self.incluir_financeiro_var,
            command=self.toggle_campos_financeiros,
            font=('Arial', 11, 'bold')
        )
        cb_financeiro.pack(anchor='w', padx=10, pady=10)
        
        # Frame principal para configurações financeiras
        self.frame_config_financeiro = ttk.LabelFrame(frame, text="Configurações do Lançamento", padding=10)
        self.frame_config_financeiro.pack(fill='x', padx=10, pady=5)
        
        # === SEÇÃO 1: DATAS ===
        datas_frame = ttk.LabelFrame(self.frame_config_financeiro, text="📅 Datas", padding=10)
        datas_frame.pack(fill='x', pady=5)
        
        # Data de referência (calculada automaticamente)
        tk.Label(datas_frame, text="Data Referência/Relatório:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        
        data_ref_calculada = self.calcular_data_referencia_nfe(dados_nfe.get('data_emissao', ''))
        self.data_referencia_var = tk.StringVar(value=data_ref_calculada)
        
        tk.Label(datas_frame, textvariable=self.data_referencia_var, 
                fg='blue', font=('Arial', 10)).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(datas_frame, text="(Calculada automaticamente conforme regra dos dias 5 e 20)", 
                font=('Arial', 8), fg='gray').grid(row=0, column=2, sticky='w', padx=5, pady=5)
        
        # Data de vencimento (usa data da NFe como padrão)
        tk.Label(datas_frame, text="Data Vencimento:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='w', padx=5, pady=5)
        
        from tkcalendar import DateEntry
        self.data_vencimento_entry = DateEntry(
            datas_frame,
            format='dd/mm/yyyy',
            locale='pt_BR',
            background='darkblue',
            foreground='white',
            borderwidth=2,
            font=('Arial', 10)
        )
        self.data_vencimento_entry.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        # Definir data de vencimento como data da NFe
        try:
            data_nfe = datetime.strptime(dados_nfe.get('data_emissao', ''), '%d/%m/%Y')
            self.data_vencimento_entry.set_date(data_nfe.date())
        except:
            self.data_vencimento_entry.set_date(datetime.now().date())
        
        # === SEÇÃO 2: CLASSIFICAÇÃO ===
        classif_frame = ttk.LabelFrame(self.frame_config_financeiro, text="🏷️ Classificação", padding=10)
        classif_frame.pack(fill='x', pady=5)
        
        # Tipo de despesa
        tk.Label(classif_frame, text="Tipo Despesa:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        
        self.tipo_despesa_var = tk.StringVar(value="3")  # Padrão material
        tipo_despesa_combo = ttk.Combobox(
            classif_frame,
            textvariable=self.tipo_despesa_var,
            values=["1", "2", "3", "4", "5", "6", "7"],
            state="readonly",
            width=5
        )
        tipo_despesa_combo.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        # Label explicativo dos tipos
        tipos_info = "1=Mão de Obra, 2=Equipamentos, 3=Materiais, 4=Terceiros, 5=Impostos, 6=Financeiro, 7=Outros"
        tk.Label(classif_frame, text=tipos_info, font=('Arial', 8), fg='gray').grid(
            row=0, column=2, sticky='w', padx=10, pady=5)
        
        # Referência (editável)
        tk.Label(classif_frame, text="Referência:", font=('Arial', 10, 'bold')).grid(
            row=1, column=0, sticky='w', padx=5, pady=5)
        
        referencia_sugerida = self.sugerir_referencia_nfe(dados_nfe)
        self.referencia_var = tk.StringVar(value=referencia_sugerida)
        self.referencia_entry = tk.Entry(classif_frame, textvariable=self.referencia_var, 
                                       width=40, font=('Arial', 10))
        self.referencia_entry.grid(row=1, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        
        # Etapa da obra
        tk.Label(classif_frame, text="Etapa da Obra:", font=('Arial', 10, 'bold')).grid(
            row=2, column=0, sticky='w', padx=5, pady=5)
        
        # Importar etapas da obra do sistema
        try:
            from src.configuracoes_sistema import GerenciadorConfiguracoes
            etapas_obra = GerenciadorConfiguracoes.get_etapas_obra()
        except:
            etapas_obra = ["ESTRUTURA", "ALVENARIA", "COBERTURA", "INSTALAÇÕES", "ACABAMENTOS", "LIMPEZA"]
        
        self.etapa_obra_var = tk.StringVar()
        etapa_combo = ttk.Combobox(
            classif_frame,
            textvariable=self.etapa_obra_var,
            values=etapas_obra,
            state="readonly",
            width=25
        )
        etapa_combo.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        # === SEÇÃO 3: VALORES ===
        valores_frame = ttk.LabelFrame(self.frame_config_financeiro, text="💰 Valores", padding=10)
        valores_frame.pack(fill='x', pady=5)
        
        # Valor da NFe (readonly)
        tk.Label(valores_frame, text="Valor Total NFe:", font=('Arial', 10, 'bold')).grid(
            row=0, column=0, sticky='w', padx=5, pady=5)
        
        valor_nfe_formatado = f"R$ {dados_nfe.get('valor_total', 0):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        tk.Label(valores_frame, text=valor_nfe_formatado, fg='green', 
                font=('Arial', 12, 'bold')).grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        # === SEÇÃO 4: OBSERVAÇÕES ===
        obs_frame = ttk.LabelFrame(self.frame_config_financeiro, text="📝 Informações Adicionais", padding=10)
        obs_frame.pack(fill='both', expand=True, pady=5)
        
        tk.Label(obs_frame, text="Observação:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=5)
        
        obs_sugerida = f"MATERIAL OBRA - NFE {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')}"
        self.observacao_var = tk.StringVar(value=obs_sugerida)
        self.observacao_entry = tk.Entry(obs_frame, textvariable=self.observacao_var, 
                                       width=80, font=('Arial', 10))
        self.observacao_entry.pack(fill='x', pady=5)
        
        # NF (preenchida automaticamente)
        tk.Label(obs_frame, text="Número NF:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(10,5))
        self.nf_var = tk.StringVar(value=dados_nfe.get('numero_nf', ''))
        tk.Entry(obs_frame, textvariable=self.nf_var, width=20, 
                font=('Arial', 10), state='readonly').pack(anchor='w')
        
        # Configurar grid weights
        classif_frame.columnconfigure(2, weight=1)
    
    def criar_aba_selecao_materiais(self, notebook, dados_nfe):
        """Aba para seleção de materiais"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📦 Materiais")
        
        # Checkbox para incluir materiais
        self.incluir_materiais_var = tk.BooleanVar(value=True)
        cb_materiais = tk.Checkbutton(
            frame,
            text="✅ Incluir materiais no controle de obra",
            variable=self.incluir_materiais_var,
            command=self.toggle_campos_materiais,
            font=('Arial', 11, 'bold')
        )
        cb_materiais.pack(anchor='w', padx=10, pady=10)
        
        # Frame para seleção de materiais
        self.frame_materiais = ttk.LabelFrame(frame, text="Seleção de Materiais", padding=10)
        self.frame_materiais.pack(fill='both', expand=True, padx=10, pady=5)
        
        produtos = dados_nfe.get('produtos', [])
        if produtos:
            # Frame para botões de seleção
            botoes_frame = ttk.Frame(self.frame_materiais)
            botoes_frame.pack(fill='x', pady=5)
            
            ttk.Button(botoes_frame, text="✅ Selecionar Todos",
                      command=self.selecionar_todos_materiais).pack(side='left', padx=5)
            ttk.Button(botoes_frame, text="❌ Desmarcar Todos",
                      command=self.desmarcar_todos_materiais).pack(side='left', padx=5)
            ttk.Button(botoes_frame, text="🔍 Apenas Materiais de Construção",
                      command=self.selecionar_apenas_construcao).pack(side='left', padx=5)
            
            # TreeView para seleção de produtos
            tree_frame = ttk.Frame(self.frame_materiais)
            tree_frame.pack(fill='both', expand=True, pady=5)
            
            self.tree_materiais = ttk.Treeview(
                tree_frame,
                columns=('sel', 'codigo', 'descricao', 'categoria', 'qtd', 'un', 'vl_unit', 'vl_total'),
                show='headings',
                height=12
            )
            
            # Configurar colunas
            colunas_config = {
                'sel': ('✓', 30),
                'codigo': ('Código', 80),
                'descricao': ('Descrição', 300),
                'categoria': ('Categoria', 120),
                'qtd': ('Qtd', 60),
                'un': ('Un', 40),
                'vl_unit': ('Vl Unit', 80),
                'vl_total': ('Vl Total', 80)
            }
            
            for col, (titulo, largura) in colunas_config.items():
                self.tree_materiais.heading(col, text=titulo)
                self.tree_materiais.column(col, width=largura)
            
            # Scrollbar
            scrollbar_mat = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_materiais.yview)
            self.tree_materiais.configure(yscrollcommand=scrollbar_mat.set)
            
            # Preencher dados
            self.produtos_selecionados = {}
            for i, produto in enumerate(produtos):
                item_id = self.tree_materiais.insert('', 'end', values=(
                    '✓',  # Selecionado por padrão
                    produto.get('codigo', ''),
                    produto.get('descricao', '')[:50],
                    produto.get('categoria_sugerida', 'OUTROS'),
                    produto.get('quantidade', ''),
                    produto.get('unidade', ''),
                    f"R$ {produto.get('valor_unitario', 0):.2f}",
                    f"R$ {produto.get('valor_total', 0):.2f}"
                ))
                self.produtos_selecionados[item_id] = True
            
            # Bind para clique
            self.tree_materiais.bind('<Button-1>', self.toggle_selecao_material)
            
            self.tree_materiais.pack(side='left', fill='both', expand=True)
            scrollbar_mat.pack(side='right', fill='y')
            
            # Configurações de ambiente para materiais selecionados
            config_frame = ttk.LabelFrame(self.frame_materiais, text="Configurações dos Materiais", padding=10)
            config_frame.pack(fill='x', pady=5)
            
            tk.Label(config_frame, text="Ambiente de Aplicação:", font=('Arial', 10, 'bold')).pack(anchor='w')
            
            # Ambientes padrão
            ambientes = [
                "DEPÓSITO DA OBRA", "SALA", "COZINHA", "BANHEIRO SUITE", "BANHEIRO SOCIAL",
                "QUARTO CASAL", "QUARTO SOLTEIRO", "VARANDA", "ÁREA EXTERNA", "TODOS AMBIENTES"
            ]
            
            self.ambiente_materiais_var = tk.StringVar(value="DEPÓSITO DA OBRA")
            ambiente_combo = ttk.Combobox(
                config_frame,
                textvariable=self.ambiente_materiais_var,
                values=ambientes,
                state="readonly",
                width=30
            )
            ambiente_combo.pack(anchor='w', pady=5)
        else:
            tk.Label(self.frame_materiais, 
                    text="📦 Esta NFe não possui produtos/materiais cadastrados.",
                    font=('Arial', 12), fg='gray').pack(pady=50)
    
    def criar_botoes_principais(self):
        """Cria botões principais da janela"""
        frame_botoes = ttk.Frame(self.janela)
        frame_botoes.pack(fill='x', padx=10, pady=10)
        
        # Botões à esquerda
        ttk.Button(frame_botoes, text="❌ Cancelar",
                  command=self.janela.destroy).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, text="👁️ Prévia",
                  command=self.visualizar_previa).pack(side='left', padx=5)
        
        # Botão principal à direita
        ttk.Button(frame_botoes, text="💾 Processar e Salvar",
                  command=self.processar_e_salvar,
                  style='Medium.TButton').pack(side='right', padx=5)
    
    # === MÉTODOS AUXILIARES ===
    
    def calcular_data_referencia_nfe(self, data_emissao_str):
        """
        Calcula data de referência usando a mesma lógica do sistema
        """
        try:
            # Usar a data da NFe como base ao invés de hoje
            data_base = datetime.strptime(data_emissao_str, '%d/%m/%Y')
        except:
            data_base = datetime.now()
        
        # Aplicar a mesma regra do sistema
        if 6 <= data_base.day <= 20:
            data_rel = data_base.replace(day=20)
        else:
            if data_base.day > 20:
                data_rel = (data_base + relativedelta(months=1)).replace(day=5)
            else:
                data_rel = data_base.replace(day=5)
        
        return data_rel.strftime('%d/%m/%Y')
    
    def sugerir_referencia_nfe(self, dados_nfe):
        """Sugere referência baseada nos dados da NFe"""
        produtos = dados_nfe.get('produtos', [])
        if not produtos:
            return "MATERIAL OBRA"
        
        # Analisar tipos de produtos mais comuns
        categorias = {}
        for produto in produtos:
            categoria = produto.get('categoria_sugerida', 'OUTROS')
            categorias[categoria] = categorias.get(categoria, 0) + 1
        
        if categorias:
            categoria_principal = max(categorias.items(), key=lambda x: x[1])[0]
            
            if categoria_principal == 'ACABAMENTOS':
                return "MATERIAL ACABAMENTO"
            elif categoria_principal == 'HIDRAULICO':
                return "MATERIAL HIDRÁULICO"
            elif categoria_principal == 'ELETRICO':
                return "MATERIAL ELÉTRICO"
            elif categoria_principal == 'ESTRUTURAL':
                return "MATERIAL ESTRUTURAL"
            else:
                return "MATERIAL OBRA"
        
        return "MATERIAL OBRA"
    
    def formatar_cnpj(self, cnpj):
        """Formata CNPJ"""
        if not cnpj or len(cnpj) != 14:
            return cnpj
        return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    
    def toggle_campos_financeiros(self):
        """Habilita/desabilita campos financeiros"""
        estado = 'normal' if self.incluir_financeiro_var.get() else 'disabled'
        
        def toggle_widget(widget):
            try:
                if isinstance(widget, (ttk.Entry, tk.Entry)):
                    widget.config(state=estado)
                elif isinstance(widget, ttk.Combobox):
                    widget.config(state='readonly' if estado == 'normal' else 'disabled')
                elif hasattr(widget, 'config'):
                    widget.config(state=estado)
            except:
                pass
        
        def toggle_frame_recursivo(frame):
            for child in frame.winfo_children():
                if isinstance(child, (ttk.Frame, tk.Frame, ttk.LabelFrame)):
                    toggle_frame_recursivo(child)
                else:
                    toggle_widget(child)
        
        if hasattr(self, 'frame_config_financeiro'):
            toggle_frame_recursivo(self.frame_config_financeiro)
    
    def toggle_campos_materiais(self):
        """Habilita/desabilita campos de materiais"""
        estado = 'normal' if self.incluir_materiais_var.get() else 'disabled'
        
        def toggle_frame_recursivo(frame):
            for child in frame.winfo_children():
                if isinstance(child, (ttk.Frame, tk.Frame, ttk.LabelFrame)):
                    toggle_frame_recursivo(child)
                else:
                    try:
                        if hasattr(child, 'config'):
                            child.config(state=estado)
                    except:
                        pass
        
        if hasattr(self, 'frame_materiais'):
            toggle_frame_recursivo(self.frame_materiais)
    
    def toggle_selecao_material(self, event):
        """Toggle seleção de material no tree"""
        item = self.tree_materiais.selection()[0] if self.tree_materiais.selection() else None
        if item:
            # Determinar coluna clicada
            region = self.tree_materiais.identify_region(event.x, event.y)
            if region == "cell":
                column = self.tree_materiais.identify_column(event.x, event.y)
                if column == '#1':  # Coluna de seleção
                    # Toggle seleção
                    atual = self.produtos_selecionados.get(item, False)
                    self.produtos_selecionados[item] = not atual
                    
                    # Atualizar visual
                    valores = list(self.tree_materiais.item(item, 'values'))
                    valores[0] = '✓' if not atual else '❌'
                    self.tree_materiais.item(item, values=valores)
    
    def selecionar_todos_materiais(self):
        """Seleciona todos os materiais"""
        for item in self.tree_materiais.get_children():
            self.produtos_selecionados[item] = True
            valores = list(self.tree_materiais.item(item, 'values'))
            valores[0] = '✓'
            self.tree_materiais.item(item, values=valores)
    
    def desmarcar_todos_materiais(self):
        """Desmarca todos os materiais"""
        for item in self.tree_materiais.get_children():
            self.produtos_selecionados[item] = False
            valores = list(self.tree_materiais.item(item, 'values'))
            valores[0] = '❌'
            self.tree_materiais.item(item, values=valores)
    
    def selecionar_apenas_construcao(self):
        """Seleciona apenas materiais relacionados à construção"""
        categorias_construcao = ['ESTRUTURAL', 'HIDRAULICO', 'ELETRICO', 'ACABAMENTOS', 'ESQUADRIAS']
        
        for item in self.tree_materiais.get_children():
            valores = list(self.tree_materiais.item(item, 'values'))
            categoria = valores[3]  # Coluna categoria
            
            if categoria in categorias_construcao:
                self.produtos_selecionados[item] = True
                valores[0] = '✓'
            else:
                self.produtos_selecionados[item] = False
                valores[0] = '❌'
            
            self.tree_materiais.item(item, values=valores)
    
    def visualizar_previa(self):
        """Mostra prévia do que será salvo"""
        try:
            dados_previa = self.preparar_dados_para_salvar()
            
            # Janela de prévia
            janela_previa = tk.Toplevel(self.janela)
            janela_previa.title("👁️ Prévia dos Dados")
            janela_previa.geometry("700x500")
            janela_previa.grab_set()
            
            # Texto com dados
            text_widget = tk.Text(janela_previa, wrap='word', font=('Courier', 10))
            scrollbar_previa = ttk.Scrollbar(janela_previa, orient="vertical", command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar_previa.set)
            
            # Preparar texto
            texto_previa = "🔍 PRÉVIA DOS DADOS QUE SERÃO SALVOS\n"
            texto_previa += "=" * 50 + "\n\n"
            
            if dados_previa['incluir_financeiro'] and dados_previa['lancamento_financeiro']:
                texto_previa += "💰 LANÇAMENTO FINANCEIRO:\n"
                lanc = dados_previa['lancamento_financeiro']
                texto_previa += f"📅 Data Referência: {lanc['data']}\n"
                texto_previa += f"🏢 Fornecedor: {lanc['nome']}\n"
                texto_previa += f"📋 CNPJ: {lanc['cnpj_cpf']}\n"
                texto_previa += f"🏷️ Tipo Despesa: {lanc['tp_desp']}\n"
                texto_previa += f"📝 Referência: {lanc['referencia']}\n"
                texto_previa += f"🏗️ Etapa: {lanc['etapa_obra']}\n"
                texto_previa += f"💰 Valor: R$ {float(lanc['valor']):,.2f}\n"
                texto_previa += f"📅 Vencimento: {lanc['dt_vencto']}\n"
                texto_previa += f"📄 NF: {lanc['nf']}\n"
                texto_previa += f"📝 Obs: {lanc['observacao']}\n\n"
            
            if dados_previa['incluir_materiais'] and dados_previa['materiais']:
                texto_previa += f"📦 MATERIAIS ({len(dados_previa['materiais'])} itens):\n"
                for i, material in enumerate(dados_previa['materiais'], 1):
                    texto_previa += f"\n{i}. {material['Descricao_Completa'][:50]}\n"
                    texto_previa += f"   🏷️ Categoria: {material['Categoria']}\n"
                    texto_previa += f"   📊 Qtd: {material['Quantidade']} {material['Unidade']}\n"
                    texto_previa += f"   💰 Valor: R$ {material['Valor_Total']:.2f}\n"
                    texto_previa += f"   🏠 Ambiente: {material['Ambiente_Aplicacao']}\n"
            
            if not dados_previa['incluir_financeiro'] and not dados_previa['incluir_materiais']:
                texto_previa += "⚠️ NENHUM DADO SERÁ SALVO!\n"
                texto_previa += "Marque pelo menos uma opção (Financeiro ou Materiais)\n"
            
            text_widget.insert('1.0', texto_previa)
            text_widget.config(state='disabled')
            
            text_widget.pack(side='left', fill='both', expand=True)
            scrollbar_previa.pack(side='right', fill='y')
            
            # Botão fechar
            ttk.Button(janela_previa, text="Fechar", 
                      command=janela_previa.destroy).pack(pady=10)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar prévia:\n{str(e)}")
    
    def preparar_dados_para_salvar(self):
        """Prepara todos os dados para salvamento"""
        dados = {
            'incluir_financeiro': self.incluir_financeiro_var.get(),
            'incluir_materiais': self.incluir_materiais_var.get(),
            'lancamento_financeiro': None,
            'materiais': []
        }
        
        # Preparar lançamento financeiro
        if dados['incluir_financeiro']:
            dados['lancamento_financeiro'] = {
                'data': self.data_referencia_var.get(),
                'cnpj_cpf': self.dados_nfe_atual.get('cnpj_emitente', ''),
                'nome': self.dados_nfe_atual.get('razao_social_emitente', ''),
                'categoria': 'MAT',  # Categoria padrão para materiais
                'tp_desp': self.tipo_despesa_var.get(),
                'referencia': self.referencia_var.get().upper(),
                'etapa_obra': self.etapa_obra_var.get(),
                'nf': self.nf_var.get().upper(),
                'vr_unit': f"{self.dados_nfe_atual.get('valor_total', 0):.2f}",
                'dias': 1,
                'valor': f"{self.dados_nfe_atual.get('valor_total', 0):.2f}",
                'dt_vencto': self.data_vencimento_entry.get(),
                'dados_bancarios': '',  # Deixar em branco conforme solicitado
                'observacao': self.observacao_var.get().upper(),
                'forma_pagamento': ''  # Deixar em branco conforme solicitado
            }
        
        # Preparar materiais selecionados
        if dados['incluir_materiais'] and hasattr(self, 'tree_materiais'):
            produtos_nfe = self.dados_nfe_atual.get('produtos', [])
            ambiente_aplicacao = self.ambiente_materiais_var.get()
            
            for i, item in enumerate(self.tree_materiais.get_children()):
                if self.produtos_selecionados.get(item, False):
                    if i < len(produtos_nfe):
                        produto = produtos_nfe[i]
                        
                        material = {
                            'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                            'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                            'Subcategoria': produto.get('subcategoria_sugerida', ''),
                            'Codigo_Produto': produto.get('codigo', ''),
                            'Descricao_Completa': produto.get('descricao', ''),
                            'Marca': '',  # Não disponível na NFe
                            'Modelo': '',
                            'Cor_Acabamento': '',
                            'Dimensoes': '',
                            'Especificacoes_Tecnicas': '',
                            'Ambiente_Aplicacao': ambiente_aplicacao,
                            'Localizacao_Especifica': '',
                            'Data_Instalacao': '',
                            'Instalador': '',
                            'Status_Instalacao': 'PENDENTE',
                            'Garantia_Meses': 12,  # Padrão
                            'Observacoes': f"Importado da NF-e {self.dados_nfe_atual.get('numero_nf', '')} - {self.dados_nfe_atual.get('fonte_dados', '')}",
                            'Tem_Dados_Compra': True,
                            'Nome_Fornecedor': self.dados_nfe_atual.get('razao_social_emitente', ''),
                            'CNPJ_Fornecedor': self.dados_nfe_atual.get('cnpj_emitente', ''),
                            'Data_Compra': self.dados_nfe_atual.get('data_emissao', ''),
                            'Quantidade': produto.get('quantidade', 0),
                            'Unidade': produto.get('unidade', 'UN'),
                            'Valor_Unitario': produto.get('valor_unitario', 0),
                            'Valor_Total': produto.get('valor_total', 0),
                            'Numero_NF': self.dados_nfe_atual.get('numero_nf', ''),
                            'Item_NF': produto.get('numero_item', ''),
                            'Origem_Dados': f"NFE_IMPORTADA_{self.dados_nfe_atual.get('fonte_dados', '')}"
                        }
                        
                        dados['materiais'].append(material)
        
        return dados
    
    def processar_e_salvar(self):
        """Processa e salva os dados no sistema"""
        try:
            # Verificar se cliente está selecionado
            if not hasattr(self.sistema, 'cliente_atual') or not self.sistema.cliente_atual:
                messagebox.showerror("Erro", "Nenhum cliente selecionado!\nSelecione um cliente antes de processar a NFe.")
                return
            
            # Preparar dados
            dados = self.preparar_dados_para_salvar()
            
            # Validar se pelo menos uma opção foi selecionada
            if not dados['incluir_financeiro'] and not dados['incluir_materiais']:
                messagebox.showwarning("Aviso", "Selecione pelo menos uma opção:\n• Incluir lançamento financeiro\n• Incluir materiais")
                return
            
            # Validações específicas
            if dados['incluir_financeiro']:
                if not self.tipo_despesa_var.get():
                    messagebox.showerror("Erro", "Tipo de despesa é obrigatório!")
                    return
                if not self.referencia_var.get().strip():
                    messagebox.showerror("Erro", "Referência é obrigatória!")
                    return
            
            # Confirmar operação
            msg_confirmacao = f"🔄 CONFIRMAR PROCESSAMENTO\n\n"
            msg_confirmacao += f"👤 Cliente: {self.sistema.cliente_atual}\n"
            msg_confirmacao += f"🏢 Fornecedor: {self.dados_nfe_atual.get('razao_social_emitente', '')}\n"
            msg_confirmacao += f"📄 NFe: {self.dados_nfe_atual.get('numero_nf', '')}\n"
            msg_confirmacao += f"💰 Valor: R$ {self.dados_nfe_atual.get('valor_total', 0):,.2f}\n\n"
            
            if dados['incluir_financeiro']:
                msg_confirmacao += f"✅ Lançamento financeiro será incluído\n"
            if dados['incluir_materiais']:
                msg_confirmacao += f"✅ {len(dados['materiais'])} materiais serão incluídos\n"
            
            msg_confirmacao += f"\nDeseja continuar?"
            
            if not messagebox.askyesno("Confirmar", msg_confirmacao):
                return
            
            # Desabilitar botão durante processamento
            for widget in self.janela.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for btn in widget.winfo_children():
                        if isinstance(btn, ttk.Button) and "Processar" in btn['text']:
                            btn.config(state='disabled', text="🔄 Processando...")
                            break
            
            self.janela.update()
            
            resultados = []
            
            # 1. SALVAR LANÇAMENTO FINANCEIRO
            if dados['incluir_financeiro']:
                try:
                    # Adicionar à lista de dados para incluir do sistema
                    self.sistema.dados_para_incluir = [dados['lancamento_financeiro']]
                    
                    # Simular o processo de envio usando o método existente
                    sucesso_financeiro = self.sistema.enviar_dados()
                    
                    if sucesso_financeiro is not False:  # Se não retornou False explicitamente
                        resultados.append("✅ Lançamento financeiro salvo")
                    else:
                        resultados.append("❌ Erro ao salvar lançamento financeiro")
                        
                except Exception as e:
                    resultados.append(f"❌ Erro financeiro: {str(e)}")
            
            # 2. SALVAR MATERIAIS
            if dados['incluir_materiais'] and dados['materiais']:
                try:
                    # Verificar se gerenciador de materiais existe
                    if not hasattr(self.sistema, 'gerenciador_materiais'):
                        from src.materiais.gerenciador_materiais import GerenciadorMateriais
                        self.sistema.gerenciador_materiais = GerenciadorMateriais(self.sistema)
                    
                    materiais_salvos = 0
                    materiais_erro = 0
                    
                    for material in dados['materiais']:
                        try:
                            material_id = self.sistema.gerenciador_materiais.salvar_material(material)
                            if material_id:
                                materiais_salvos += 1
                            else:
                                materiais_erro += 1
                        except Exception as e:
                            print(f"Erro ao salvar material: {e}")
                            materiais_erro += 1
                    
                    if materiais_salvos > 0:
                        resultados.append(f"✅ {materiais_salvos} materiais salvos")
                    if materiais_erro > 0:
                        resultados.append(f"⚠️ {materiais_erro} materiais com erro")
                        
                except Exception as e:
                    resultados.append(f"❌ Erro materiais: {str(e)}")
            
            # Mostrar resultados
            if resultados:
                msg_resultado = f"🎉 PROCESSAMENTO CONCLUÍDO!\n\n"
                msg_resultado += f"📋 Resumo:\n"
                for resultado in resultados:
                    msg_resultado += f"  {resultado}\n"
                
                msg_resultado += f"\n📄 NFe {self.dados_nfe_atual.get('numero_nf', '')} processada com sucesso!"
                
                messagebox.showinfo("Sucesso", msg_resultado)
                
                # Fechar janela
                self.janela.destroy()
            else:
                messagebox.showerror("Erro", "Nenhum dado foi processado!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro durante processamento:\n{str(e)}")
        
        finally:
            # Reabilitar botão
            for widget in self.janela.winfo_children():
                if isinstance(widget, ttk.Frame):
                    for btn in widget.winfo_children():
                        if isinstance(btn, ttk.Button) and "Processando" in btn['text']:
                            btn.config(state='normal', text="💾 Processar e Salvar")
                            break