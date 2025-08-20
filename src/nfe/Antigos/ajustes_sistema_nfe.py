# -*- coding: utf-8 -*-
"""
AJUSTES PARA O SISTEMA NFe UNIFICADO
Melhorias baseadas no feedback de uso
"""

import tkinter as tk
from tkinter import ttk
import json
from datetime import datetime
from pathlib import Path

class AjustesSistemaNFe:
    """
    Classe com ajustes e melhorias para o sistema NFe
    """
    
    @staticmethod
    def ajustar_data_para_sistema(data_emissao):
        """
        Ajusta data da NFe para o padrão do sistema (dia 5 ou 20)
        A data da NFe vai para dt_vencto, e data_rel fica 5 ou 20
        """
        try:
            if not data_emissao:
                return datetime.now().strftime('%d/%m/%Y'), datetime.now().strftime('%d/%m/%Y')
            
            # CONVERTER DATA DA NFE
            if isinstance(data_emissao, str):
                if '/' in data_emissao:
                    dt_nfe = datetime.strptime(data_emissao, '%d/%m/%Y')
                else:
                    dt_nfe = datetime.strptime(data_emissao, '%Y-%m-%d')
            else:
                dt_nfe = data_emissao
            
            # DATA DE VENCIMENTO = DATA DA NFE
            dt_vencto = dt_nfe.strftime('%d/%m/%Y')
            
            # DATA DE REFERÊNCIA = DIA 5 OU 20 DO MESMO MÊS
            dia_nfe = dt_nfe.day
            
            if dia_nfe <= 12:  # Primeira quinzena
                dia_ref = 5
            else:  # Segunda quinzena
                dia_ref = 20
            
            dt_ref = dt_nfe.replace(day=dia_ref)
            data_rel = dt_ref.strftime('%d/%m/%Y')
            
            print(f"📅 Data ajustada: NFe {dt_vencto} → Ref {data_rel}")
            
            return data_rel, dt_vencto
            
        except Exception as e:
            print(f"❌ Erro ao ajustar data: {e}")
            hoje = datetime.now()
            data_padrao = hoje.replace(day=5).strftime('%d/%m/%Y')
            return data_padrao, hoje.strftime('%d/%m/%Y')
    
    @staticmethod
    def carregar_parametros_sistema():
        """
        Carrega parâmetros do sistema principal
        """
        try:
            # TENTAR CARREGAR PARAMETROS_SISTEMA.JSON
            caminhos_possiveis = [
                Path.cwd() / "parametros_sistema.json",
                Path.cwd() / "data" / "parametros_sistema.json",
                Path.cwd() / "config" / "parametros_sistema.json"
            ]
            
            for caminho in caminhos_possiveis:
                if caminho.exists():
                    with open(caminho, 'r', encoding='utf-8') as f:
                        return json.load(f)
            
            # SE NÃO ENCONTRAR, RETORNAR PADRÕES
            return {
                "etapas_obra": [
                    "INSTALAÇÃO DA OBRA",
                    "FUNDAÇÃO",
                    "ESTRUTURA",
                    "ALVENARIA", 
                    "COBERTURA",
                    "INSTALAÇÕES",
                    "ACABAMENTOS",
                    "MATERIAIS",
                    "FINALIZAÇÃO"
                ]
            }
            
        except Exception as e:
            print(f"⚠️ Erro ao carregar parâmetros do sistema: {e}")
            return {"etapas_obra": ["MATERIAIS", "ACABAMENTOS", "INSTALAÇÕES"]}
    
    @staticmethod
    def carregar_parametros_materiais():
        """
        Carrega parâmetros específicos dos materiais
        """
        try:
            # TENTAR CARREGAR PARAMETROS_MATERIAIS.JSON
            caminhos_possiveis = [
                Path.cwd() / "data" / "materiais" / "parametros_materiais.json",
                Path.cwd() / "data" / "parametros_materiais.json",
                Path.cwd() / "parametros_materiais.json"
            ]
            
            for caminho in caminhos_possiveis:
                if caminho.exists():
                    with open(caminho, 'r', encoding='utf-8') as f:
                        return json.load(f)
            
            # PADRÕES SE NÃO ENCONTRAR
            return {
                "ambientes": [
                    "INSTALAÇÃO DA OBRA",
                    "SALA DE ESTAR", 
                    "SALA DE JANTAR",
                    "COZINHA",
                    "DORMITÓRIO SUITE",
                    "DORMITÓRIO 1",
                    "DORMITÓRIO 2", 
                    "DORMITÓRIO 3",
                    "BANHEIRO SUITE",
                    "BANHEIRO SOCIAL",
                    "LAVABO",
                    "ÁREA DE SERVIÇO",
                    "VARANDA",
                    "ÁREA EXTERNA",
                    "GARAGEM",
                    "DEPÓSITO",
                    "GERAL"
                ],
                "status_instalacao": [
                    "PENDENTE",
                    "INSTALADO", 
                    "EM_INSTALACAO",
                    "AGUARDANDO_MATERIAL",
                    "AGUARDANDO_INSTALADOR",
                    "DEFEITO",
                    "SUBSTITUIDO",
                    "CANCELADO"
                ],
                "unidades": [
                    "UN", "M", "M²", "M³", "KG", "G", "L", "ML", 
                    "PC", "CX", "SC", "PAR", "JG", "KIT", "ROL", "LT"
                ]
            }
            
        except Exception as e:
            print(f"⚠️ Erro ao carregar parâmetros de materiais: {e}")
            return {"ambientes": ["GERAL"], "status_instalacao": ["PENDENTE"], "unidades": ["UN"]}


class InterfaceNFeAprimorada:
    """
    Interface aprimorada para importação NFe com ajustes solicitados
    """
    
    def __init__(self, sistema_principal, dados_nfe):
        self.sistema = sistema_principal
        self.dados_nfe = dados_nfe
        self.parametros_sistema = AjustesSistemaNFe.carregar_parametros_sistema()
        self.parametros_materiais = AjustesSistemaNFe.carregar_parametros_materiais()
        
        # VARIÁVEIS DE CONTROLE
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        self.criar_interface()
    
    def criar_interface(self):
        """Cria interface aprimorada"""
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("Configuração de Importação NFe")
        self.janela.geometry("700x600")
        self.janela.grab_set()
        
        # FRAME PRINCIPAL COM SCROLL
        canvas = tk.Canvas(self.janela)
        scrollbar = ttk.Scrollbar(self.janela, orient="vertical", command=canvas.yview)
        self.frame_scrollable = ttk.Frame(canvas)
        
        # CONFIGURAR SCROLL
        self.frame_scrollable.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=self.frame_scrollable, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # CRIAR SEÇÕES
        self.criar_secao_resumo_nfe()
        self.criar_secao_opcoes_importacao()
        self.criar_secao_financeiro_aprimorada()
        self.criar_secao_materiais_aprimorada()
        self.criar_botoes_finais()
        
        # PACK SCROLL
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # CONFIGURAR ESTADOS INICIAIS
        self.toggle_opcoes_financeiro()
        self.toggle_opcoes_materiais()
    
    def criar_secao_resumo_nfe(self):
        """Seção com resumo da NFe"""
        frame_resumo = ttk.LabelFrame(self.frame_scrollable, text="📄 Dados da NFe", padding=10)
        frame_resumo.pack(fill='x', padx=10, pady=5)
        
        # INFORMAÇÕES EM GRID
        info_grid = ttk.Frame(frame_resumo)
        info_grid.pack(fill='x')
        
        infos = [
            ("📄 Número:", self.dados_nfe.get('numero_nf', '')),
            ("📅 Data Emissão:", self.dados_nfe.get('data_emissao', '')),
            ("🏢 Fornecedor:", self.dados_nfe.get('razao_social_emitente', '')[:40]),
            ("💰 Valor Total:", f"R$ {self.dados_nfe.get('valor_total', 0):,.2f}"),
            ("📦 Produtos:", str(len(self.dados_nfe.get('produtos', [])))),
            ("🔑 Chave:", self.dados_nfe.get('chave_acesso', '')[:20] + "...")
        ]
        
        for i, (label, valor) in enumerate(infos):
            row = i // 2
            col = (i % 2) * 2
            
            tk.Label(info_grid, text=label, font=('Arial', 9, 'bold')).grid(
                row=row, column=col, sticky='w', padx=5, pady=2)
            tk.Label(info_grid, text=valor).grid(
                row=row, column=col+1, sticky='w', padx=5, pady=2)
    
    def criar_secao_opcoes_importacao(self):
        """Seção de opções principais"""
        frame_opcoes = ttk.LabelFrame(self.frame_scrollable, text="⚙️ O que Importar", padding=10)
        frame_opcoes.pack(fill='x', padx=10, pady=5)
        
        tk.Checkbutton(
            frame_opcoes,
            text="💰 Dados Financeiros (lançamento no sistema)",
            variable=self.importar_financeiro,
            font=('Arial', 10),
            command=self.toggle_opcoes_financeiro
        ).pack(anchor='w', pady=2)
        
        tk.Checkbutton(
            frame_opcoes,
            text="📦 Materiais da Obra (banco de dados para manual)",
            variable=self.importar_materiais,
            font=('Arial', 10),
            command=self.toggle_opcoes_materiais
        ).pack(anchor='w', pady=2)
    
    def criar_secao_financeiro_aprimorada(self):
        """Seção financeiro com ajustes de data"""
        self.frame_financeiro = ttk.LabelFrame(self.frame_scrollable, 
                                              text="💰 Configurações Financeiras", 
                                              padding=10)
        self.frame_financeiro.pack(fill='x', padx=10, pady=5)
        
        # LINHA 1: DATAS
        linha_datas = ttk.LabelFrame(self.frame_financeiro, text="📅 Datas do Sistema", padding=5)
        linha_datas.pack(fill='x', pady=5)
        
        # CALCULAR DATAS AUTOMATICAMENTE
        data_rel, dt_vencto = AjustesSistemaNFe.ajustar_data_para_sistema(
            self.dados_nfe.get('data_emissao', '')
        )
        
        # DATA DE REFERÊNCIA (5 ou 20)
        frame_data_rel = ttk.Frame(linha_datas)
        frame_data_rel.pack(fill='x', pady=2)
        
        tk.Label(frame_data_rel, text="Data Referência (5 ou 20):", 
                font=('Arial', 9, 'bold')).pack(side='left')
        self.data_rel = tk.Entry(frame_data_rel, width=12)
        self.data_rel.pack(side='left', padx=5)
        self.data_rel.insert(0, data_rel)
        
        tk.Label(frame_data_rel, text="💡 Sistema só aceita dia 5 ou 20", 
                fg='blue', font=('Arial', 8)).pack(side='left', padx=10)
        
        # DATA DE VENCIMENTO (da NFe)
        frame_vencto = ttk.Frame(linha_datas)
        frame_vencto.pack(fill='x', pady=2)
        
        tk.Label(frame_vencto, text="Data Vencimento (da NFe):", 
                font=('Arial', 9, 'bold')).pack(side='left')
        self.dt_vencto = tk.Entry(frame_vencto, width=12)
        self.dt_vencto.pack(side='left', padx=5)
        self.dt_vencto.insert(0, dt_vencto)
        
        # LINHA 2: CLASSIFICAÇÃO
        linha_class = ttk.Frame(self.frame_financeiro)
        linha_class.pack(fill='x', pady=5)
        
        tk.Label(linha_class, text="Tipo Despesa:").grid(row=0, column=0, sticky='w', pady=2)
        self.tipo_despesa = ttk.Combobox(linha_class, width=15, state='readonly')
        self.tipo_despesa['values'] = ['1', '2', '3', '4', '5', '6', '7']
        self.tipo_despesa.set('3')  # Material
        self.tipo_despesa.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        tk.Label(linha_class, text="Categoria:").grid(row=0, column=2, sticky='w', padx=(20,0), pady=2)
        self.categoria_financeira = tk.Entry(linha_class, width=10)
        self.categoria_financeira.insert(0, 'MAT')
        self.categoria_financeira.grid(row=0, column=3, sticky='w', padx=5, pady=2)
        
        # LINHA 3: ETAPA DA OBRA (com parâmetros)
        linha_etapa = ttk.Frame(self.frame_financeiro)
        linha_etapa.pack(fill='x', pady=5)
        
        tk.Label(linha_etapa, text="Etapa da Obra:").grid(row=0, column=0, sticky='w', pady=2)
        self.etapa_obra = ttk.Combobox(linha_etapa, width=25, state='readonly')
        etapas = self.parametros_sistema.get('etapas_obra', ['MATERIAIS', 'ACABAMENTOS'])
        self.etapa_obra['values'] = etapas
        self.etapa_obra.set('MATERIAIS')
        self.etapa_obra.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        tk.Label(linha_etapa, text="Forma Pgto:").grid(row=0, column=2, sticky='w', padx=(20,0), pady=2)
        self.forma_pagamento = ttk.Combobox(linha_etapa, width=15, state='readonly')
        self.forma_pagamento['values'] = ['A_VISTA', 'A_PRAZO', 'CARTAO', 'PIX']
        self.forma_pagamento.set('A_PRAZO')
        self.forma_pagamento.grid(row=0, column=3, sticky='w', padx=5, pady=2)
        
        # LINHA 4: REFERÊNCIA EDITÁVEL
        linha_ref = ttk.Frame(self.frame_financeiro)
        linha_ref.pack(fill='x', pady=5)
        
        tk.Label(linha_ref, text="Referência:", font=('Arial', 9, 'bold')).pack(side='left')
        self.referencia_editavel = tk.Entry(linha_ref, width=50)
        self.referencia_editavel.pack(side='left', padx=5, fill='x', expand=True)
        
        # GERAR REFERÊNCIA INICIAL
        ref_inicial = f"NFE {self.dados_nfe.get('numero_nf', '')} - {self.dados_nfe.get('razao_social_emitente', '')[:25]}".upper()
        self.referencia_editavel.insert(0, ref_inicial)
    
    def criar_secao_materiais_aprimorada(self):
        """Seção materiais detalhada com parâmetros"""
        self.frame_materiais = ttk.LabelFrame(self.frame_scrollable, 
                                             text="📦 Configurações Materiais", 
                                             padding=10)
        self.frame_materiais.pack(fill='x', padx=10, pady=5)
        
        # LINHA 1: LOCALIZAÇÃO
        linha_local = ttk.LabelFrame(self.frame_materiais, text="🏠 Localização", padding=5)
        linha_local.pack(fill='x', pady=5)
        
        frame_ambiente = ttk.Frame(linha_local)
        frame_ambiente.pack(fill='x', pady=2)
        
        tk.Label(frame_ambiente, text="Ambiente Padrão:").pack(side='left')
        self.ambiente_padrao = ttk.Combobox(frame_ambiente, width=30, state='readonly')
        ambientes = self.parametros_materiais.get('ambientes', ['GERAL'])
        self.ambiente_padrao['values'] = ambientes
        self.ambiente_padrao.pack(side='left', padx=5)
        
        frame_localizacao = ttk.Frame(linha_local)
        frame_localizacao.pack(fill='x', pady=2)
        
        tk.Label(frame_localizacao, text="Localização Específica:").pack(side='left')
        self.localizacao_especifica = tk.Entry(frame_localizacao, width=40)
        self.localizacao_especifica.pack(side='left', padx=5)
        self.localizacao_especifica.insert(0, "Conforme projeto")
        
        # LINHA 2: STATUS E GARANTIA
        linha_status = ttk.LabelFrame(self.frame_materiais, text="⚙️ Status e Garantia", padding=5)
        linha_status.pack(fill='x', pady=5)
        
        frame_status = ttk.Frame(linha_status)
        frame_status.pack(fill='x', pady=2)
        
        tk.Label(frame_status, text="Status Instalação:").pack(side='left')
        self.status_instalacao = ttk.Combobox(frame_status, width=20, state='readonly')
        status_list = self.parametros_materiais.get('status_instalacao', ['PENDENTE'])
        self.status_instalacao['values'] = status_list
        self.status_instalacao.set('PENDENTE')
        self.status_instalacao.pack(side='left', padx=5)
        
        tk.Label(frame_status, text="Garantia:").pack(side='left', padx=(20,0))
        self.garantia_meses = tk.Entry(frame_status, width=5)
        self.garantia_meses.insert(0, '12')
        self.garantia_meses.pack(side='left', padx=5)
        tk.Label(frame_status, text="meses").pack(side='left')
        
        # LINHA 3: FORNECEDOR E MARCA
        linha_fornec = ttk.LabelFrame(self.frame_materiais, text="🏢 Fornecedor", padding=5)
        linha_fornec.pack(fill='x', pady=5)
        
        frame_marca = ttk.Frame(linha_fornec)
        frame_marca.pack(fill='x', pady=2)
        
        tk.Label(frame_marca, text="Marca/Fabricante:").pack(side='left')
        self.marca_fabricante = tk.Entry(frame_marca, width=30)
        marca_sugerida = self.dados_nfe.get('razao_social_emitente', '')[:25]
        self.marca_fabricante.insert(0, marca_sugerida)
        self.marca_fabricante.pack(side='left', padx=5)
        
        # LINHA 4: INSTALAÇÃO
        linha_instal = ttk.LabelFrame(self.frame_materiais, text="🔧 Instalação", padding=5)
        linha_instal.pack(fill='x', pady=5)
        
        frame_instal = ttk.Frame(linha_instal)
        frame_instal.pack(fill='x', pady=2)
        
        tk.Label(frame_instal, text="Data Instalação:").pack(side='left')
        self.data_instalacao = tk.Entry(frame_instal, width=12)
        self.data_instalacao.pack(side='left', padx=5)
        
        tk.Label(frame_instal, text="Instalador:").pack(side='left', padx=(20,0))
        self.instalador = tk.Entry(frame_instal, width=25)
        self.instalador.pack(side='left', padx=5)
        
        # INFORMAÇÕES ÚTEIS
        info_frame = ttk.Frame(self.frame_materiais)
        info_frame.pack(fill='x', pady=5)
        
        info_text = """💡 Dicas: Deixe campos vazios para preencher depois. Ambiente e status podem ser 
alterados individualmente em 'Consultar Materiais'."""
        
        tk.Label(info_frame, text=info_text, justify='left', 
                fg='blue', font=('Arial', 8)).pack(anchor='w')
    
    def criar_botoes_finais(self):
        """Botões finais da interface"""
        frame_botoes = ttk.Frame(self.frame_scrollable)
        frame_botoes.pack(fill='x', padx=10, pady=20)
        
        ttk.Button(frame_botoes, 
                  text="👁️ Preview Dados", 
                  command=self.preview_dados).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ Processar Importação", 
                  command=self.processar_importacao).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="❌ Cancelar", 
                  command=self.janela.destroy).pack(side='right', padx=5)
    
    def toggle_opcoes_financeiro(self):
        """Toggle opções financeiras"""
        estado = 'normal' if self.importar_financeiro.get() else 'disabled'
        
        def alterar_estado_widget(widget):
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.config(state=estado)
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    alterar_estado_widget(child)
        
        alterar_estado_widget(self.frame_financeiro)
    
    def toggle_opcoes_materiais(self):
        """Toggle opções materiais"""
        estado = 'normal' if self.importar_materiais.get() else 'disabled'
        
        def alterar_estado_widget(widget):
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.config(state=estado)
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    alterar_estado_widget(child)
        
        alterar_estado_widget(self.frame_materiais)
    
    def preview_dados(self):
        """Mostra preview dos dados que serão importados"""
        # COLETAR DADOS ATUAIS
        opcoes = self.coletar_opcoes()
        
        # CRIAR JANELA DE PREVIEW
        janela_preview = tk.Toplevel(self.janela)
        janela_preview.title("Preview da Importação")
        janela_preview.geometry("600x500")
        janela_preview.grab_set()
        
        # NOTEBOOK PARA ORGANIZAR
        notebook = ttk.Notebook(janela_preview)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # ABA FINANCEIRO
        if opcoes['importar_financeiro']:
            self.criar_preview_financeiro(notebook, opcoes)
        
        # ABA MATERIAIS
        if opcoes['importar_materiais']:
            self.criar_preview_materiais(notebook, opcoes)
        
        # BOTÃO FECHAR
        ttk.Button(janela_preview, text="Fechar", 
                  command=janela_preview.destroy).pack(pady=10)
    
    def criar_preview_financeiro(self, notebook, opcoes):
        """Preview dos dados financeiros"""
        frame_fin = ttk.Frame(notebook)
        notebook.add(frame_fin, text="💰 Dados Financeiros")
        
        preview_text = f"""
LANÇAMENTO FINANCEIRO:

📅 Data Referência: {opcoes['data_rel']} (padrão do sistema)
📅 Data Vencimento: {opcoes['dt_vencto']} (data da NFe)
🏢 Fornecedor: {self.dados_nfe.get('razao_social_emitente', '')}
📄 CNPJ: {self.dados_nfe.get('cnpj_emitente', '')}
🏷️ Categoria: {opcoes['categoria_financeira']}
🔧 Tipo Despesa: {opcoes['tipo_despesa']}
📋 Referência: {opcoes['referencia']}
🏗️ Etapa Obra: {opcoes['etapa_obra']}
📄 NF: {self.dados_nfe.get('numero_nf', '')}
💰 Valor: R$ {self.dados_nfe.get('valor_total', 0):,.2f}
💳 Forma Pgto: {opcoes['forma_pagamento']}

⚠️ Este lançamento será adicionado à lista de dados.
Use 'Enviar Dados' no sistema principal para salvar na planilha.
        """.strip()
        
        text_widget = tk.Text(frame_fin, wrap='word', font=('Courier', 10))
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', preview_text)
        text_widget.config(state='disabled')
    
    def criar_preview_materiais(self, notebook, opcoes):
        """Preview dos materiais"""
        frame_mat = ttk.Frame(notebook)
        notebook.add(frame_mat, text="📦 Materiais")
        
        # HEADER COM CONFIGURAÇÕES
        header_text = f"""
CONFIGURAÇÕES DOS MATERIAIS:
🏠 Ambiente: {opcoes['ambiente_padrao']} | 📍 Localização: {opcoes['localizacao_especifica']}
⚙️ Status: {opcoes['status_instalacao']} | 🛡️ Garantia: {opcoes['garantia_meses']} meses
🏢 Marca: {opcoes['marca_fabricante']} | 🔧 Instalador: {opcoes['instalador']}
📅 Data Instalação: {opcoes['data_instalacao']}

PRODUTOS QUE SERÃO IMPORTADOS:
        """.strip()
        
        tk.Label(frame_mat, text=header_text, justify='left', 
                font=('Arial', 9)).pack(anchor='w', padx=10, pady=5)
        
        # TREEVIEW COM PRODUTOS
        colunas = ('Item', 'Código', 'Descrição', 'Categoria', 'Qtd', 'Valor')
        tree = ttk.Treeview(frame_mat, columns=colunas, show='headings', height=12)
        
        # CONFIGURAR COLUNAS
        for col in colunas:
            tree.heading(col, text=col)
        
        tree.column('Item', width=40)
        tree.column('Código', width=80)
        tree.column('Descrição', width=200)
        tree.column('Categoria', width=100)
        tree.column('Qtd', width=60)
        tree.column('Valor', width=80)
        
        # PREENCHER PRODUTOS
        produtos = self.dados_nfe.get('produtos', [])
        for i, produto in enumerate(produtos, 1):
            tree.insert('', 'end', values=(
                i,
                produto.get('codigo', '')[:12],
                produto.get('descricao', '')[:30],
                produto.get('categoria_sugerida', ''),
                produto.get('quantidade', ''),
                f"R$ {produto.get('valor_total', 0):.2f}"
            ))
        
        # SCROLLBAR
        scrollbar = ttk.Scrollbar(frame_mat, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side='left', fill='both', expand=True, padx=10, pady=5)
        scrollbar.pack(side='right', fill='y', pady=5)
    
    def coletar_opcoes(self):
        """Coleta todas as opções da interface"""
        return {
            'importar_financeiro': self.importar_financeiro.get(),
            'importar_materiais': self.importar_materiais.get(),
            
            # FINANCEIRO
            'data_rel': self.data_rel.get(),
            'dt_vencto': self.dt_vencto.get(),
            'tipo_despesa': self.tipo_despesa.get(),
            'categoria_financeira': self.categoria_financeira.get(),
            'etapa_obra': self.etapa_obra.get(),
            'forma_pagamento': self.forma_pagamento.get(),
            'referencia': self.referencia_editavel.get(),
            
            # MATERIAIS
            'ambiente_padrao': self.ambiente_padrao.get(),
            'localizacao_especifica': self.localizacao_especifica.get(),
            'status_instalacao': self.status_instalacao.get(),
            'garantia_meses': int(self.garantia_meses.get() or 12),
            'marca_fabricante': self.marca_fabricante.get(),
            'data_instalacao': self.data_instalacao.get(),
            'instalador': self.instalador.get()
        }
    
    def processar_importacao(self):
        """Processa a importação com as opções configuradas"""
        try:
            # VALIDAR SELEÇÕES
            if not self.importar_financeiro.get() and not self.importar_materiais.get():
                tk.messagebox.showwarning("Aviso", "Selecione pelo menos uma opção!")
                return
            
            # VALIDAR DATAS
            data_rel = self.data_rel.get().strip()
            if self.importar_financeiro.get() and data_rel:
                try:
                    dt = datetime.strptime(data_rel, '%d/%m/%Y')
                    if dt.day not in [5, 20]:
                        resposta = tk.messagebox.askyesno(
                            "Data Inválida", 
                            f"Data de referência deve ser dia 5 ou 20.\n"
                            f"Atual: {data_rel}\n\n"
                            f"Deseja ajustar automaticamente?"
                        )
                        if resposta:
                            data_corrigida, _ = AjustesSistemaNFe.ajustar_data_para_sistema(data_rel)
                            self.data_rel.delete(0, tk.END)
                            self.data_rel.insert(0, data_corrigida)
                        else:
                            return
                except ValueError:
                    tk.messagebox.showerror("Erro", "Data de referência inválida!")
                    return
            
            # COLETAR OPÇÕES FINAIS
            opcoes = self.coletar_opcoes()
            
            # FECHAR INTERFACE
            self.janela.destroy()
            
            # PROCESSAR DADOS
            self.executar_importacao(opcoes)
            
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao processar importação: {str(e)}")
    
    def executar_importacao(self, opcoes):
        """Executa a importação propriamente dita"""
        try:
            resultados = []
            
            # IMPORTAR FINANCEIRO
            if opcoes['importar_financeiro']:
                resultado_fin = self.criar_lancamento_financeiro_aprimorado(opcoes)
                resultados.append(f"💰 Financeiro: {resultado_fin}")
            
            # IMPORTAR MATERIAIS
            if opcoes['importar_materiais']:
                resultado_mat = self.criar_materiais_aprimorados(opcoes)
                resultados.append(f"📦 Materiais: {resultado_mat}")
            
            # MOSTRAR RESULTADO
            self.mostrar_resultado_final(resultados, opcoes)
            
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro na importação: {str(e)}")
    
    def criar_lancamento_financeiro_aprimorado(self, opcoes):
        """Cria lançamento financeiro com ajustes"""
        try:
            dados_nfe = self.dados_nfe
            
            # DADOS FINANCEIROS APRIMORADOS
            dados_financeiros = {
                'data': opcoes['data_rel'],  # Data de referência (5 ou 20)
                'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                'categoria': opcoes['categoria_financeira'].upper(),
                'tp_desp': opcoes['tipo_despesa'],
                'referencia': opcoes['referencia'].upper(),
                'etapa_obra': opcoes['etapa_obra'].upper(),
                'nf': dados_nfe.get('numero_nf', ''),
                'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dias': 1,
                'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dt_vencto': opcoes['dt_vencto'],  # Data de vencimento (da NFe)
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - CHAVE: {dados_nfe.get('chave_acesso', '')[:20]}...".upper(),
                'forma_pagamento': opcoes['forma_pagamento']
            }
            
            # ADICIONAR À LISTA DO SISTEMA
            if not hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir = []
            
            self.sistema.dados_para_incluir.append(dados_financeiros)
            
            print(f"💰 Lançamento financeiro criado: R$ {dados_nfe.get('valor_total', 0):,.2f}")
            
            return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            
        except Exception as e:
            raise Exception(f"Erro ao criar lançamento: {str(e)}")
    
    def criar_materiais_aprimorados(self, opcoes):
        """Cria materiais com configurações detalhadas"""
        try:
            produtos = self.dados_nfe.get('produtos', [])
            if not produtos:
                return "Nenhum produto encontrado"
            
            # VERIFICAR SISTEMA DE MATERIAIS
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                return f"{len(produtos)} produtos (sistema materiais não inicializado)"
            
            materiais_criados = 0
            dados_nfe = self.dados_nfe
            
            for produto in produtos:
                try:
                    # DADOS COMPLETOS DO MATERIAL
                    dados_material = {
                        # IDENTIFICAÇÃO
                        'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                        'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                        'Subcategoria': produto.get('subcategoria_sugerida', ''),
                        'Codigo_Produto': produto.get('codigo', ''),
                        'Descricao_Completa': produto.get('descricao', ''),
                        
                        # ESPECIFICAÇÕES
                        'Marca': opcoes['marca_fabricante'],
                        'Modelo': '',
                        'Cor_Acabamento': '',
                        'Dimensoes': '',
                        'Especificacoes_Tecnicas': self.gerar_especificacoes_produto(produto),
                        
                        # LOCALIZAÇÃO
                        'Ambiente_Aplicacao': opcoes['ambiente_padrao'],
                        'Localizacao_Especifica': opcoes['localizacao_especifica'],
                        
                        # INSTALAÇÃO
                        'Data_Instalacao': opcoes['data_instalacao'],
                        'Instalador': opcoes['instalador'],
                        'Status_Instalacao': opcoes['status_instalacao'],
                        'Garantia_Meses': opcoes['garantia_meses'],
                        
                        # OBSERVAÇÕES
                        'Observacoes': f"Importado NFe {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')}",
                        
                        # DADOS DE COMPRA
                        'Tem_Dados_Compra': True,
                        'Data_Compra': dados_nfe.get('data_emissao', ''),
                        'CNPJ_Fornecedor': dados_nfe.get('cnpj_emitente', ''),
                        'Nome_Fornecedor': dados_nfe.get('razao_social_emitente', ''),
                        'Numero_NF': dados_nfe.get('numero_nf', ''),
                        'Item_NF': produto.get('numero_item', ''),
                        'Quantidade': produto.get('quantidade', 1),
                        'Unidade': produto.get('unidade', 'UN'),
                        'Valor_Unitario': produto.get('valor_unitario', 0),
                        'Valor_Total': produto.get('valor_total', 0),
                        'Origem_Dados': 'IMPORTACAO_NFE'
                    }
                    
                    # SALVAR MATERIAL
                    material_id = self.sistema.gerenciador_materiais.salvar_material(dados_material)
                    materiais_criados += 1
                    
                    print(f"📦 Material criado - ID: {material_id} - {produto.get('descricao', '')[:30]}")
                    
                except Exception as e:
                    print(f"⚠️ Erro ao criar material {produto.get('descricao', '')}: {e}")
                    continue
            
            return f"{materiais_criados} de {len(produtos)} produtos"
            
        except Exception as e:
            raise Exception(f"Erro ao criar materiais: {str(e)}")
    
    def gerar_especificacoes_produto(self, produto):
        """Gera especificações técnicas detalhadas"""
        specs = []
        
        if produto.get('ncm'):
            specs.append(f"NCM: {produto['ncm']}")
        
        if produto.get('cfop'):
            specs.append(f"CFOP: {produto['cfop']}")
        
        if produto.get('unidade') and produto.get('quantidade'):
            specs.append(f"Embalagem: {produto['quantidade']} {produto['unidade']}")
        
        if produto.get('valor_unitario'):
            specs.append(f"Valor Unit: R$ {produto['valor_unitario']:.2f}")
        
        return " | ".join(specs) if specs else "Conforme NFe"
    
    def mostrar_resultado_final(self, resultados, opcoes):
        """Mostra resultado final da importação"""
        janela_resultado = tk.Toplevel(self.sistema.root)
        janela_resultado.title("🎉 Importação Concluída")
        janela_resultado.geometry("600x500")
        janela_resultado.grab_set()
        
        # FRAME PRINCIPAL
        frame_main = ttk.Frame(janela_resultado)
        frame_main.pack(fill='both', expand=True, padx=20, pady=20)
        
        # TÍTULO
        titulo = tk.Label(frame_main, 
                         text="✅ IMPORTAÇÃO NFe CONCLUÍDA COM SUCESSO!", 
                         font=('Arial', 14, 'bold'),
                         fg='green')
        titulo.pack(pady=10)
        
        # RESUMO NFE
        frame_nfe = ttk.LabelFrame(frame_main, text="📄 NFe Processada", padding=10)
        frame_nfe.pack(fill='x', pady=5)
        
        dados_nfe = self.dados_nfe
        resumo_nfe = f"""📄 Número: {dados_nfe.get('numero_nf', '')}  |  📅 Data: {dados_nfe.get('data_emissao', '')}
🏢 Fornecedor: {dados_nfe.get('razao_social_emitente', '')}
💰 Valor Total: R$ {dados_nfe.get('valor_total', 0):,.2f}  |  📦 Produtos: {len(dados_nfe.get('produtos', []))}"""
        
        tk.Label(frame_nfe, text=resumo_nfe, justify='left').pack(anchor='w')
        
        # DADOS IMPORTADOS
        frame_importados = ttk.LabelFrame(frame_main, text="✅ Dados Importados", padding=10)
        frame_importados.pack(fill='x', pady=5)
        
        for resultado in resultados:
            tk.Label(frame_importados, text=resultado, fg='blue', 
                    font=('Arial', 10, 'bold')).pack(anchor='w', pady=1)
        
        # CONFIGURAÇÕES APLICADAS
        if opcoes['importar_financeiro']:
            frame_config_fin = ttk.LabelFrame(frame_main, text="💰 Configurações Financeiras", padding=10)
            frame_config_fin.pack(fill='x', pady=5)
            
            config_fin = f"""📅 Data Ref: {opcoes['data_rel']}  |  📅 Vencimento: {opcoes['dt_vencto']}
🏗️ Etapa: {opcoes['etapa_obra']}  |  📋 Referência: {opcoes['referencia'][:50]}"""
            
            tk.Label(frame_config_fin, text=config_fin, justify='left', 
                    font=('Arial', 9)).pack(anchor='w')
        
        if opcoes['importar_materiais']:
            frame_config_mat = ttk.LabelFrame(frame_main, text="📦 Configurações Materiais", padding=10)
            frame_config_mat.pack(fill='x', pady=5)
            
            config_mat = f"""🏠 Ambiente: {opcoes['ambiente_padrao']}  |  ⚙️ Status: {opcoes['status_instalacao']}
🛡️ Garantia: {opcoes['garantia_meses']} meses  |  🏢 Marca: {opcoes['marca_fabricante']}"""
            
            tk.Label(frame_config_mat, text=config_mat, justify='left', 
                    font=('Arial', 9)).pack(anchor='w')
        
        # PRÓXIMOS PASSOS
        frame_passos = ttk.LabelFrame(frame_main, text="🚀 Próximos Passos", padding=10)
        frame_passos.pack(fill='x', pady=5)
        
        passos_text = """1. ✅ Dados foram adicionados ao sistema
2. 📊 Use 'Enviar Dados' para salvar na planilha do cliente
3. 📦 Confira materiais em 'Consultar Materiais'
4. 🔧 Atualize status de instalação conforme obra avança
5. 📄 Gere 'Manual do Proprietário' quando todos materiais estiverem instalados"""
        
        tk.Label(frame_passos, text=passos_text, justify='left', 
                font=('Arial', 9)).pack(anchor='w')
        
        # BOTÕES
        frame_botoes = ttk.Frame(frame_main)
        frame_botoes.pack(fill='x', pady=15)
        
        ttk.Button(frame_botoes, 
                  text="📊 Processar no Sistema", 
                  command=lambda: self.processar_no_sistema_final(janela_resultado)).pack(side='left', padx=5)
        
        if opcoes['importar_materiais']:
            ttk.Button(frame_botoes, 
                      text="📦 Ver Materiais", 
                      command=self.abrir_consulta_materiais).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ Fechar", 
                  command=janela_resultado.destroy).pack(side='right', padx=5)
    
    def processar_no_sistema_final(self, janela_resultado):
        """Executa enviar_dados() do sistema principal"""
        try:
            janela_resultado.destroy()
            
            # VERIFICAR DADOS
            if not hasattr(self.sistema, 'dados_para_incluir') or not self.sistema.dados_para_incluir:
                tk.messagebox.showwarning("Aviso", "Não há dados financeiros para processar!")
                return
            
            print(f"📊 Processando {len(self.sistema.dados_para_incluir)} lançamentos...")
            
            # CHAMAR MÉTODO DO SISTEMA
            self.sistema.enviar_dados()
            
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")
    
    def abrir_consulta_materiais(self):
        """Abre consulta de materiais"""
        try:
            if hasattr(self.sistema, 'integrador_materiais'):
                self.sistema.integrador_materiais.abrir_consulta_materiais()
            else:
                tk.messagebox.showinfo("Info", "Sistema de materiais não disponível")
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro: {str(e)}")


# FUNÇÃO PARA INTEGRAR OS AJUSTES NO SISTEMA UNIFICADO
def aplicar_ajustes_sistema_nfe(sistema_nfe_unificado):
    """
    Aplica os ajustes no sistema NFe unificado existente
    """
    try:
        # SUBSTITUIR MÉTODO DE IMPORTAÇÃO
        def importar_para_sistema_aprimorado(self):
            """Versão aprimorada da importação"""
            try:
                if not self.dados_nfe_atual:
                    tk.messagebox.showerror("Erro", "Nenhum dado carregado!")
                    return
                
                # USAR INTERFACE APRIMORADA
                interface_aprimorada = InterfaceNFeAprimorada(
                    self.sistema, 
                    self.dados_nfe_atual
                )
                
            except Exception as e:
                tk.messagebox.showerror("Erro", f"Erro na importação: {str(e)}")
        
        # SUBSTITUIR MÉTODO
        sistema_nfe_unificado.importar_para_sistema = importar_para_sistema_aprimorado.__get__(
            sistema_nfe_unificado, 
            type(sistema_nfe_unificado)
        )
        
        print("✅ Ajustes aplicados ao sistema NFe!")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar ajustes: {e}")
        return False


# FUNÇÃO PRINCIPAL PARA APLICAR TODOS OS AJUSTES
def aplicar_todos_ajustes_nfe(sistema_principal):
    """
    Aplica todos os ajustes identificados no sistema NFe
    """
    try:
        print("🔧 Aplicando ajustes no sistema NFe...")
        
        # VERIFICAR SE SISTEMA NFE EXISTE
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe unificado não encontrado!")
            return False
        
        # APLICAR AJUSTES
        sucesso = aplicar_ajustes_sistema_nfe(sistema_principal.sistema_nfe_unificado)
        
        if sucesso:
            print("✅ Todos os ajustes aplicados com sucesso!")
            print("📌 Melhorias implementadas:")
            print("   - ✅ Datas ajustadas para padrão do sistema (5 ou 20)")
            print("   - ✅ Campo referência editável")
            print("   - ✅ Etapas da obra carregadas dos parâmetros")
            print("   - ✅ Configurações materiais detalhadas")
            print("   - ✅ Interface aprimorada com scroll")
            print("   - ✅ Preview completo dos dados")
            return True
        else:
            print("❌ Erro ao aplicar ajustes!")
            return False
        
    except Exception as e:
        print(f"❌ Erro geral nos ajustes: {e}")
        return False


# EXEMPLO DE USO DOS AJUSTES
"""
PARA APLICAR OS AJUSTES NO SEU SISTEMA:

# No final do __init__ do SistemaEntradaDados, adicione:
try:
    from src.nfe.ajustes_sistema_nfe import aplicar_todos_ajustes_nfe
    aplicar_todos_ajustes_nfe(self)
    print("✅ Ajustes NFe aplicados!")
except Exception as e:
    print(f"⚠️ Ajustes NFe não aplicados: {e}")

RESULTADO:
- ✅ Data da NFe vai para vencimento
- ✅ Data de referência fica dia 5 ou 20
- ✅ Campo referência editável pelo usuário  
- ✅ Etapas carregadas do parametros_sistema.json
- ✅ Configurações materiais detalhadas
- ✅ Interface com scroll e melhor organização
- ✅ Preview completo antes da importação
"""