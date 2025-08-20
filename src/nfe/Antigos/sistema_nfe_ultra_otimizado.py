# -*- coding: utf-8 -*-
"""
SISTEMA NFe ULTRA OTIMIZADO - FLUXO DIRETO
Elimina TODAS as telas intermediárias: Botão → Selecionar → Configurar
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path
import re
import calendar


class SistemaNFeUltraOtimizado:
    """
    Sistema ultra otimizado: clique único → seleção → configuração
    """
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.dados_nfe_atual = None
        print("🚀 Sistema NFe Ultra Otimizado inicializado")
    
    def abrir_importacao_nfe(self):
        """
        FLUXO ULTRA OTIMIZADO:
        1. Abre seletor de arquivo diretamente
        2. Processa XML em segundo plano
        3. Abre interface de configuração
        """
        try:
            # PASSO 1: SELECIONAR ARQUIVO DIRETAMENTE
            arquivo = filedialog.askopenfilename(
                title="📄 Selecionar XML da NFe para Importação",
                filetypes=[
                    ("Arquivos XML da NFe", "*.xml"),
                    ("Todos os arquivos", "*.*")
                ]
            )
            
            if not arquivo:
                return
            
            # PASSO 2: MOSTRAR PROGRESSO COM JANELA TEMPORÁRIA
            janela_progresso = self.criar_janela_progresso()
            
            try:
                # PASSO 3: PROCESSAR XML EM SEGUNDO PLANO
                self.dados_nfe_atual = self.processar_xml_nfe(arquivo)
                
                if self.dados_nfe_atual:
                    # PASSO 4: FECHAR PROGRESSO E ABRIR CONFIGURAÇÃO
                    janela_progresso.destroy()
                    InterfaceConfiguracaoNFeOtimizada(self.sistema, self.dados_nfe_atual)
                    print("✅ Fluxo ultra otimizado concluído!")
                else:
                    janela_progresso.destroy()
                    messagebox.showerror("Erro", "Não foi possível processar o arquivo XML selecionado.")
                    
            except Exception as e:
                janela_progresso.destroy()
                messagebox.showerror("Erro", f"Erro ao processar XML:\n{str(e)}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processo de importação:\n{str(e)}")
    
    def criar_janela_progresso(self):
        """Cria janela de progresso minimalista"""
        janela = tk.Toplevel(self.sistema.root)
        janela.title("Processando NFe...")
        janela.geometry("400x150")
        janela.grab_set()
        janela.resizable(False, False)
        
        # CENTRALIZAR NA TELA
        janela.transient(self.sistema.root)
        
        # CONTEÚDO
        frame = ttk.Frame(janela)
        frame.pack(expand=True, fill='both', padx=20, pady=20)
        
        # ÍCONE E TEXTO
        tk.Label(frame, text="🔄", font=('Arial', 24)).pack(pady=5)
        tk.Label(frame, text="Processando XML da NFe...", 
                font=('Arial', 12, 'bold')).pack(pady=5)
        tk.Label(frame, text="Extraindo dados e preparando configuração", 
                font=('Arial', 9), fg='gray').pack(pady=2)
        
        # BARRA DE PROGRESSO INDETERMINADA
        progress = ttk.Progressbar(frame, mode='indeterminate')
        progress.pack(fill='x', pady=10)
        progress.start(10)
        
        # ATUALIZAR INTERFACE
        janela.update()
        
        return janela
    
    def processar_xml_nfe(self, caminho_arquivo):
        """Processa arquivo XML da NFe (versão otimizada)"""
        try:
            print(f"📄 Processando XML: {caminho_arquivo}")
            
            # CARREGAR XML
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # NAMESPACE NFE
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # ESTRUTURA BÁSICA
            dados = {
                'fonte_dados': 'XML Local',
                'arquivo_origem': caminho_arquivo,
                'nome_arquivo': Path(caminho_arquivo).name,
                'chave_acesso': '',
                'numero_nf': '',
                'serie': '1',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'cnpj_emitente': '',
                'razao_social_emitente': '',
                'valor_total': 0.0,
                'valor_produtos': 0.0,
                'produtos': []
            }
            
            # BUSCAR DADOS NO XML
            inf_nfe = root.find('.//nfe:infNFe', ns)
            if inf_nfe is None:
                raise Exception("Arquivo XML inválido - estrutura NFe não encontrada")
            
            # CHAVE DE ACESSO
            dados['chave_acesso'] = inf_nfe.get('Id', '').replace('NFe', '')
            
            # DADOS IDENTIFICAÇÃO
            ide = inf_nfe.find('nfe:ide', ns)
            if ide is not None:
                dados['numero_nf'] = self.get_xml_text(ide.find('nfe:nNF', ns))
                dados['serie'] = self.get_xml_text(ide.find('nfe:serie', ns))
                
                # DATA DE EMISSÃO
                dh_emi = self.get_xml_text(ide.find('nfe:dhEmi', ns))
                if dh_emi:
                    dados['data_emissao'] = self.formatar_data_xml(dh_emi)
            
            # DADOS DO EMITENTE
            emit = inf_nfe.find('nfe:emit', ns)
            if emit is not None:
                dados['cnpj_emitente'] = self.get_xml_text(emit.find('nfe:CNPJ', ns))
                dados['razao_social_emitente'] = self.get_xml_text(emit.find('nfe:xNome', ns))
            
            # TOTAIS
            total = inf_nfe.find('.//nfe:total/nfe:ICMSTot', ns)
            if total is not None:
                dados['valor_total'] = float(self.get_xml_text(total.find('nfe:vNF', ns)) or 0)
                dados['valor_produtos'] = float(self.get_xml_text(total.find('nfe:vProd', ns)) or 0)
            
            # PRODUTOS
            dados['produtos'] = self.extrair_produtos_xml(inf_nfe, ns)
            
            print(f"✅ XML processado: NFe {dados['numero_nf']} - R$ {dados['valor_total']:,.2f}")
            
            return dados
            
        except Exception as e:
            print(f"❌ Erro ao processar XML: {e}")
            raise e
    
    def extrair_produtos_xml(self, inf_nfe, ns):
        """Extrai produtos do XML (versão otimizada)"""
        produtos = []
        
        try:
            itens = inf_nfe.findall('nfe:det', ns)
            
            for item in itens:
                prod = item.find('nfe:prod', ns)
                if prod is None:
                    continue
                
                produto = {
                    'numero_item': item.get('nItem', ''),
                    'codigo': self.get_xml_text(prod.find('nfe:cProd', ns)),
                    'descricao': self.get_xml_text(prod.find('nfe:xProd', ns)),
                    'ncm': self.get_xml_text(prod.find('nfe:NCM', ns)),
                    'cfop': self.get_xml_text(prod.find('nfe:CFOP', ns)),
                    'unidade': self.get_xml_text(prod.find('nfe:uCom', ns)),
                    'quantidade': float(self.get_xml_text(prod.find('nfe:qCom', ns)) or 0),
                    'valor_unitario': float(self.get_xml_text(prod.find('nfe:vUnCom', ns)) or 0),
                    'valor_total': float(self.get_xml_text(prod.find('nfe:vProd', ns)) or 0)
                }
                
                # CLASSIFICAÇÃO AUTOMÁTICA
                produto['categoria_sugerida'] = self.classificar_produto(produto['descricao'])
                
                produtos.append(produto)
                
        except Exception as e:
            print(f"⚠️ Erro ao extrair produtos: {e}")
        
        return produtos
    
    # MÉTODOS UTILITÁRIOS
    
    def get_xml_text(self, element):
        """Extrai texto de elemento XML"""
        return element.text if element is not None else ''
    
    def formatar_data_xml(self, data_str):
        """Formata data do XML para dd/mm/yyyy"""
        try:
            if 'T' in data_str:
                dt = datetime.fromisoformat(data_str.replace('Z', '+00:00'))
            else:
                dt = datetime.strptime(data_str, '%Y-%m-%d')
            return dt.strftime('%d/%m/%Y')
        except:
            return data_str
    
    def classificar_produto(self, descricao):
        """Classifica produto em categoria"""
        if not descricao:
            return 'OUTROS'
        
        desc_upper = descricao.upper()
        
        categorias = {
            'ACABAMENTOS': [
                'CERAMICA', 'PORCELANATO', 'AZULEJO', 'PASTILHA', 'REVESTIMENTO', 'PISO',
                'RODAPE', 'MOLDURA', 'REJUNTE', 'GESSO', 'FORRO', 'LAMINADO'
            ],
            'TINTAS': [
                'TINTA', 'VERNIZ', 'ESMALTE', 'PRIMER', 'SELADOR', 'MASSA CORRIDA',
                'TEXTURA', 'RESINA'
            ],
            'ELETRICO': [
                'FIO', 'CABO', 'TOMADA', 'INTERRUPTOR', 'LAMPADA', 'LED', 'ELETRICO',
                'DISJUNTOR', 'QUADRO', 'CONDULETE', 'ELETRODUTO'
            ],
            'HIDRAULICO': [
                'TUBO', 'CONEXAO', 'REGISTRO', 'TORNEIRA', 'VALVULA', 'HIDRAULICO',
                'CANO', 'CHUVEIRO', 'VASO SANITARIO', 'PIA', 'SIFAO'
            ],
            'FERRAGENS': [
                'PARAFUSO', 'PREGO', 'BUCHA', 'CHAVE', 'CADEADO', 'REBITE',
                'PORCA', 'ARRUELA', 'GANCHO'
            ]
        }
        
        for categoria, palavras_chave in categorias.items():
            if any(palavra in desc_upper for palavra in palavras_chave):
                return categoria
        
        return 'OUTROS'


class InterfaceConfiguracaoNFeOtimizada:
    """Interface de configuração otimizada com título melhorado"""
    
    def __init__(self, sistema_principal, dados_nfe):
        self.sistema = sistema_principal
        self.dados_nfe = dados_nfe
        
        # CALCULAR DATAS DO PERÍODO ATUAL
        self.calcular_datas_periodo_atual()
        
        # CRIAR INTERFACE
        self.criar_interface()
    
    def calcular_datas_periodo_atual(self):
        """Calcula datas do período ATUAL (não da NFe)"""
        hoje = datetime.now()
        
        # DATA DE REFERÊNCIA = SEMPRE PERÍODO ATUAL
        if hoje.day <= 5:
            self.data_referencia = hoje.replace(day=5).strftime('%d/%m/%Y')
            self.periodo_nome = "PRIMEIRA QUINZENA"
            # self.data_fim_periodo = hoje.replace(day=5).strftime('%d/%m/%Y')
        else:
            self.data_referencia = hoje.replace(day=20).strftime('%d/%m/%Y')
            self.periodo_nome = "SEGUNDA QUINZENA"
            # Último dia do mês
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            # self.data_fim_periodo = hoje.replace(day=ultimo_dia).strftime('%d/%m/%Y')
        
        # DATA DE VENCIMENTO = DATA ORIGINAL DA NFE
        self.data_vencimento = self.dados_nfe.get('data_emissao', hoje.strftime('%d/%m/%Y'))
    
    def criar_interface(self):
        """Cria interface otimizada de configuração"""
        self.janela = tk.Toplevel(self.sistema.root)
        
        # TÍTULO DINÂMICO COM INFO DA NFE
        nfe_info = f"NFe {self.dados_nfe.get('numero_nf', '')} - {self.dados_nfe.get('razao_social_emitente', '')[:30]}"
        self.janela.title(f"⚙️ Configurar Importação: {nfe_info}")
        self.janela.geometry("750x700")
        self.janela.grab_set()
        
        # FRAME PRINCIPAL COM SCROLL
        canvas = tk.Canvas(self.janela)
        scrollbar = ttk.Scrollbar(self.janela, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # PACK SCROLL
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        main_frame = scrollable_frame
        
        # CONTEÚDO
        self.criar_cabecalho_otimizado(main_frame)
        self.criar_resumo_nfe_compacto(main_frame)
        self.criar_periodo_atual_destaque(main_frame)
        self.criar_configuracoes_simplificadas(main_frame)
        self.criar_botoes_finais_otimizados(main_frame)
        
        # BIND SCROLL
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def criar_cabecalho_otimizado(self, parent):
        """Cria cabeçalho otimizado"""
        frame_header = ttk.Frame(parent)
        frame_header.pack(fill='x', padx=15, pady=10)
        
        # TÍTULO PRINCIPAL
        titulo_principal = f"📄 NFe {self.dados_nfe.get('numero_nf', '')} - PRONTA PARA IMPORTAÇÃO"
        titulo = tk.Label(frame_header, 
                         text=titulo_principal,
                         font=('Arial', 14, 'bold'),
                         fg='darkgreen')
        titulo.pack()
        
        # ARQUIVO ORIGEM
        arquivo_nome = self.dados_nfe.get('nome_arquivo', 'arquivo.xml')
        subtitulo = tk.Label(frame_header,
                           text=f"📁 Origem: {arquivo_nome}",
                           font=('Arial', 9),
                           fg='blue')
        subtitulo.pack()
    
    def criar_resumo_nfe_compacto(self, parent):
        """Cria resumo compacto da NFe"""
        frame_nfe = ttk.LabelFrame(parent, text="📊 Resumo da NFe", padding=8)
        frame_nfe.pack(fill='x', padx=15, pady=5)
        
        # LINHA ÚNICA COM INFO ESSENCIAL
        info_frame = ttk.Frame(frame_nfe)
        info_frame.pack(fill='x')
        
        # INFO COMPACTA
        fornecedor = self.dados_nfe.get('razao_social_emitente', '')[:35]
        valor = f"R$ {self.dados_nfe.get('valor_total', 0):,.2f}"
        produtos = len(self.dados_nfe.get('produtos', []))
        data_nfe = self.dados_nfe.get('data_emissao', '')
        
        info_text = f"🏢 {fornecedor}  |  💰 {valor}  |  📦 {produtos} produtos  |  📅 {data_nfe}"
        
        tk.Label(info_frame, text=info_text, font=('Arial', 10)).pack()
    
    def criar_periodo_atual_destaque(self, parent):
        """Cria seção do período atual em destaque"""
        frame_periodo = ttk.LabelFrame(parent, text="🎯 PERÍODO DO RELATÓRIO", padding=10)
        frame_periodo.pack(fill='x', padx=15, pady=8)
        
        # FRAME PARA LAYOUT EM COLUNAS
        cols_frame = ttk.Frame(frame_periodo)
        cols_frame.pack(fill='x')
        
        # COLUNA 1: PERÍODO ATUAL
        col1 = ttk.Frame(cols_frame)
        col1.pack(side='left', fill='x', expand=True)
        
        tk.Label(col1, text=f"📊 {self.periodo_nome}", 
                font=('Arial', 10, 'bold'), fg='darkgreen').pack(anchor='w')
        tk.Label(col1, text=f"Data do Relatório: {self.data_referencia}", 
                font=('Arial', 10)).pack(anchor='w')
        # tk.Label(col1, text=f"Período até: {self.data_fim_periodo}", 
        #         font=('Arial', 9), fg='gray').pack(anchor='w')
        
        # COLUNA 2: VENCIMENTO NFE
        col2 = ttk.Frame(cols_frame)
        col2.pack(side='left', fill='x', expand=True)
        
        tk.Label(col2, text="📄 Vencimento da NFe", 
                font=('Arial', 11, 'bold'), fg='darkblue').pack(anchor='w')
        tk.Label(col2, text=f"Data: {self.data_vencimento}", 
                font=('Arial', 10)).pack(anchor='w')
        tk.Label(col2, text="(mantém data original)", 
                font=('Arial', 9), fg='gray').pack(anchor='w')
        
        # AVISO IMPORTANTE
        aviso_frame = ttk.Frame(frame_periodo)
        aviso_frame.pack(fill='x', pady=(8,0))
        
        tk.Label(aviso_frame, text="💡", font=('Arial', 14)).pack(side='left')
        aviso_text = f"Esta NFe entrará no relatório de {self.data_referencia} ({self.periodo_nome}) para cálculo da taxa"
        tk.Label(aviso_frame, text=aviso_text, font=('Arial', 9, 'bold'), 
                fg='purple', wraplength=600).pack(side='left', padx=(5,0))
    
    def criar_configuracoes_simplificadas(self, parent):
        """Cria configurações simplificadas"""
        frame_config = ttk.LabelFrame(parent, text="⚙️ Configurações de Importação", padding=10)
        frame_config.pack(fill='x', padx=15, pady=5)
        
        # VARIÁVEIS DE CONTROLE
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        # CHECKBOXES PRINCIPAIS
        cb_frame = ttk.Frame(frame_config)
        cb_frame.pack(fill='x', pady=5)
        
        cb_financeiro = tk.Checkbutton(
            cb_frame,
            text="💰 Criar lançamento financeiro no sistema",
            variable=self.importar_financeiro,
            font=('Arial', 11, 'bold'),
            command=self.toggle_opcoes
        )
        cb_financeiro.pack(anchor='w', pady=2)
        
        cb_materiais = tk.Checkbutton(
            cb_frame,
            text="📦 Salvar materiais no banco de dados da obra",
            variable=self.importar_materiais,
            font=('Arial', 11, 'bold'),
            command=self.toggle_opcoes
        )
        cb_materiais.pack(anchor='w', pady=2)
        
        # CONFIGURAÇÕES ESPECÍFICAS (COMPACTAS)
        self.criar_config_financeiro_compacto(frame_config)
        self.criar_config_materiais_compacto(frame_config)
    
    def criar_config_financeiro_compacto(self, parent):
        """Configurações financeiras compactas"""
        self.frame_fin = ttk.LabelFrame(parent, text="💰 Configuração Financeira", padding=8)
        self.frame_fin.pack(fill='x', pady=5)
        
        # REFERÊNCIA (DESTAQUE)
        ref_frame = ttk.Frame(self.frame_fin)
        ref_frame.pack(fill='x', pady=3)
        
        tk.Label(ref_frame, text="📋 Referência:", font=('Arial', 9, 'bold')).pack(side='left')
        self.referencia_entry = tk.Entry(ref_frame, font=('Arial', 9))
        self.referencia_entry.pack(side='left', fill='x', expand=True, padx=(5,0))
        
        # REFERÊNCIA PADRÃO INTELIGENTE
        numero_nf = self.dados_nfe.get('numero_nf', '')
        fornecedor = self.dados_nfe.get('razao_social_emitente', '')[:25]
        ref_padrao = f"NFE {numero_nf} - {fornecedor}"
        self.referencia_entry.insert(0, ref_padrao)
        
        # CONFIGURAÇÕES EM LINHA COMPACTA
        config_frame = ttk.Frame(self.frame_fin)
        config_frame.pack(fill='x', pady=3)
        
        # CATEGORIA
        tk.Label(config_frame, text="Categoria:").pack(side='left')
        self.categoria_entry = tk.Entry(config_frame, width=6)
        self.categoria_entry.insert(0, 'MAT')
        self.categoria_entry.pack(side='left', padx=(2,10))
        
        # TIPO
        tk.Label(config_frame, text="Tipo:").pack(side='left')
        self.tipo_combo = ttk.Combobox(config_frame, width=4, state='readonly')
        self.tipo_combo['values'] = ['1', '2', '3', '4', '5', '6', '7']
        self.tipo_combo.set('3')
        self.tipo_combo.pack(side='left', padx=(2,10))
        
        # FORMA PAGAMENTO
        tk.Label(config_frame, text="Forma Pgto:").pack(side='left')
        self.forma_combo = ttk.Combobox(config_frame, width=10, state='readonly')
        self.forma_combo['values'] = ['A_VISTA', 'A_PRAZO', 'CARTAO', 'PIX']
        self.forma_combo.set('')
        self.forma_combo.pack(side='left', padx=(2,0))
    
    def criar_config_materiais_compacto(self, parent):
        """Configurações materiais compactas"""
        self.frame_mat = ttk.LabelFrame(parent, text="📦 Configuração Materiais", padding=8)
        self.frame_mat.pack(fill='x', pady=5)
        
        config_frame = ttk.Frame(self.frame_mat)
        config_frame.pack(fill='x', pady=3)
        
        # AMBIENTE
        tk.Label(config_frame, text="Ambiente:").pack(side='left')
        self.ambiente_combo = ttk.Combobox(config_frame, width=18, state='readonly')
        self.ambiente_combo['values'] = self.carregar_ambientes()
        self.ambiente_combo.pack(side='left', padx=(2,10))
        
        # STATUS
        tk.Label(config_frame, text="Status:").pack(side='left')
        self.status_combo = ttk.Combobox(config_frame, width=12, state='readonly')
        self.status_combo['values'] = ['PENDENTE', 'INSTALADO', 'EM_INSTALACAO']
        self.status_combo.set('PENDENTE')
        self.status_combo.pack(side='left', padx=(2,10))
        
        # GARANTIA
        tk.Label(config_frame, text="Garantia:").pack(side='left')
        self.garantia_entry = tk.Entry(config_frame, width=4)
        self.garantia_entry.insert(0, '12')
        self.garantia_entry.pack(side='left', padx=(2,2))
        tk.Label(config_frame, text="meses").pack(side='left')
    
    def carregar_ambientes(self):
        """Carrega ambientes dos parâmetros"""
        try:
            if hasattr(self.sistema, 'gerenciador_materiais'):
                return self.sistema.gerenciador_materiais.parametros.get('ambientes', ['GERAL'])
        except:
            pass
        return ['GERAL', 'INSTALAÇÃO DA OBRA', 'SALA DE ESTAR', 'COZINHA']
    
    def toggle_opcoes(self):
        """Habilita/desabilita opções baseado nas seleções"""
        # Financeiro
        estado_fin = 'normal' if self.importar_financeiro.get() else 'disabled'
        for widget in self.frame_fin.winfo_children():
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.config(state=estado_fin)
            elif isinstance(widget, ttk.Frame):
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, (tk.Entry, ttk.Combobox)):
                        subwidget.config(state=estado_fin)
        
        # Materiais
        estado_mat = 'normal' if self.importar_materiais.get() else 'disabled'
        for widget in self.frame_mat.winfo_children():
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.config(state=estado_mat)
            elif isinstance(widget, ttk.Frame):
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, (tk.Entry, ttk.Combobox)):
                        subwidget.config(state=estado_mat)
    
    def criar_botoes_finais_otimizados(self, parent):
        """Cria botões finais otimizados"""
        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill='x', padx=15, pady=15)
        
        # BOTÃO PRINCIPAL EM DESTAQUE
        btn_importar = ttk.Button(frame_botoes, 
                                 text="✅ IMPORTAR NFe PARA O SISTEMA", 
                                 command=self.executar_importacao_otimizada)
        btn_importar.pack(side='left', padx=5)
        
        # BOTÕES SECUNDÁRIOS
        ttk.Button(frame_botoes, 
                  text="👁️ Preview", 
                  command=self.mostrar_preview_otimizado).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="📦 Ver Produtos", 
                  command=self.mostrar_produtos_rapido).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="❌ Cancelar", 
                  command=self.janela.destroy).pack(side='right', padx=5)
    
    def mostrar_preview_otimizado(self):
        """Mostra preview otimizado"""
        opcoes = self.coletar_opcoes()
        
        # PREVIEW COMPACTO
        preview_text = f"""🎯 PREVIEW DA IMPORTAÇÃO NFe {self.dados_nfe.get('numero_nf', '')}
{'='*60}

💰 VALOR: R$ {self.dados_nfe.get('valor_total', 0):,.2f}
🏢 FORNECEDOR: {self.dados_nfe.get('razao_social_emitente', '')}
📦 PRODUTOS: {len(self.dados_nfe.get('produtos', []))} itens
"""
        
        if opcoes['importar_financeiro']:
            preview_text += f"""
💰 LANÇAMENTO FINANCEIRO:
   📅 Período: {self.periodo_nome} ({self.data_referencia})
   📄 Vencimento: {self.data_vencimento}
   📋 Referência: {opcoes['referencia']}
   🏷️ Categoria: {opcoes['categoria']} | Tipo: {opcoes['tipo']} | Forma: {opcoes['forma_pagamento']}
"""
        
        if opcoes['importar_materiais']:
            preview_text += f"""
📦 MATERIAIS DA OBRA:
   🏠 Ambiente: {opcoes['ambiente']}
   ⚙️ Status: {opcoes['status']}
   🛡️ Garantia: {opcoes['garantia']} meses
   📊 Quantidade: {len(self.dados_nfe.get('produtos', []))} produtos
"""
        
        preview_text += f"""
🎯 RESULTADO:
   Esta NFe será incluída no relatório de {self.data_referencia}
   Base para cálculo da taxa de administração do período atual
"""
        
        # JANELA DE PREVIEW COMPACTA
        janela_preview = tk.Toplevel(self.janela)
        janela_preview.title("👁️ Preview da Importação")
        janela_preview.geometry("550x400")
        janela_preview.grab_set()
        
        # TEXTO
        text_widget = tk.Text(janela_preview, wrap='word', font=('Consolas', 9))
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', preview_text.strip())
        text_widget.config(state='disabled')
        
        # BOTÃO
        ttk.Button(janela_preview, text="✅ Confirma - Importar", 
                  command=lambda: [janela_preview.destroy(), self.executar_importacao_otimizada()]).pack(pady=5)
        ttk.Button(janela_preview, text="❌ Fechar Preview", 
                  command=janela_preview.destroy).pack(pady=2)
    
    def mostrar_produtos_rapido(self):
        """Mostra lista rápida de produtos"""
        janela_produtos = tk.Toplevel(self.janela)
        janela_produtos.title(f"📦 Produtos da NFe {self.dados_nfe.get('numero_nf', '')}")
        janela_produtos.geometry("900x500")
        janela_produtos.grab_set()
        
        # FRAME PRINCIPAL
        frame_main = ttk.Frame(janela_produtos)
        frame_main.pack(fill='both', expand=True, padx=10, pady=10)
        
        # TREEVIEW COMPACTO
        colunas = ('Item', 'Descrição', 'Categoria', 'Qtd', 'Valor Unit', 'Total')
        tree = ttk.Treeview(frame_main, columns=colunas, show='headings', height=15)
        
        # CONFIGURAR COLUNAS
        tree.heading('Item', text='#')
        tree.heading('Descrição', text='Descrição')
        tree.heading('Categoria', text='Categoria')
        tree.heading('Qtd', text='Qtd')
        tree.heading('Valor Unit', text='Vl Unit')
        tree.heading('Total', text='Total')
        
        tree.column('Item', width=40)
        tree.column('Descrição', width=300)
        tree.column('Categoria', width=120)
        tree.column('Qtd', width=60)
        tree.column('Valor Unit', width=80)
        tree.column('Total', width=80)
        
        # PREENCHER DADOS
        produtos = self.dados_nfe.get('produtos', [])
        for i, produto in enumerate(produtos, 1):
            tree.insert('', 'end', values=(
                i,
                produto.get('descricao', '')[:40],
                produto.get('categoria_sugerida', ''),
                produto.get('quantidade', ''),
                f"R$ {produto.get('valor_unitario', 0):.2f}",
                f"R$ {produto.get('valor_total', 0):.2f}"
            ))
        
        # SCROLLBAR
        scrollbar = ttk.Scrollbar(frame_main, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # BOTÃO FECHAR
        ttk.Button(janela_produtos, text="✅ Fechar", 
                  command=janela_produtos.destroy).pack(pady=10)
    
    def coletar_opcoes(self):
        """Coleta todas as opções"""
        return {
            'importar_financeiro': self.importar_financeiro.get(),
            'importar_materiais': self.importar_materiais.get(),
            'referencia': self.referencia_entry.get().strip(),
            'categoria': self.categoria_entry.get().strip(),
            'tipo': self.tipo_combo.get(),
            'forma_pagamento': self.forma_combo.get(),
            'ambiente': self.ambiente_combo.get(),
            'status': self.status_combo.get(),
            'garantia': int(self.garantia_entry.get() or 12)
        }
    
    def executar_importacao_otimizada(self):
        """Executa importação com feedback otimizado"""
        try:
            opcoes = self.coletar_opcoes()
            
            if not opcoes['importar_financeiro'] and not opcoes['importar_materiais']:
                messagebox.showwarning("⚠️ Seleção Necessária", 
                    "Selecione pelo menos uma opção:\n• Lançamento financeiro\n• Materiais da obra")
                return
            
            # JANELA DE PROGRESSO
            janela_progresso = self.criar_janela_progresso_importacao()
            
            resultados = []
            
            # IMPORTAR FINANCEIRO
            if opcoes['importar_financeiro']:
                janela_progresso.update()
                resultado_fin = self.criar_lancamento_financeiro_otimizado(opcoes)
                resultados.append(f"💰 Financeiro: {resultado_fin}")
            
            # IMPORTAR MATERIAIS
            if opcoes['importar_materiais']:
                janela_progresso.update()
                resultado_mat = self.criar_materiais_otimizado(opcoes)
                resultados.append(f"📦 Materiais: {resultado_mat}")
            
            # FECHAR PROGRESSO
            janela_progresso.destroy()
            
            # FECHAR CONFIGURAÇÃO
            self.janela.destroy()
            
            # MOSTRAR RESULTADO
            self.mostrar_resultado_final_otimizado(resultados, opcoes)
            
        except Exception as e:
            messagebox.showerror("❌ Erro na Importação", f"Erro durante a importação:\n{str(e)}")
    
    def criar_janela_progresso_importacao(self):
        """Cria janela de progresso para importação"""
        janela = tk.Toplevel(self.sistema.root)
        janela.title("Importando NFe...")
        janela.geometry("350x120")
        janela.grab_set()
        janela.resizable(False, False)
        
        frame = ttk.Frame(janela)
        frame.pack(expand=True, fill='both', padx=15, pady=15)
        
        tk.Label(frame, text="⚙️", font=('Arial', 20)).pack()
        tk.Label(frame, text="Importando dados da NFe...", 
                font=('Arial', 11, 'bold')).pack()
        
        progress = ttk.Progressbar(frame, mode='indeterminate')
        progress.pack(fill='x', pady=8)
        progress.start(15)
        
        janela.update()
        return janela
    
    def criar_lancamento_financeiro_otimizado(self, opcoes):
        """Cria lançamento financeiro otimizado"""
        try:
            dados_nfe = self.dados_nfe
            
            dados_financeiros = {
                'data': self.data_referencia,  # PERÍODO ATUAL (5 ou 20)
                'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                'categoria': opcoes['categoria'].upper(),
                'tp_desp': opcoes['tipo'],
                'referencia': opcoes['referencia'].upper(),
                'etapa_obra': 'MATERIAIS',
                'nf': dados_nfe.get('numero_nf', ''),
                'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dias': 1,
                'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dt_vencto': self.data_vencimento,  # DATA ORIGINAL DA NFE
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - {self.periodo_nome} {self.data_referencia}".upper(),
                'forma_pagamento': opcoes['forma_pagamento']
            }
            
            if not hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir = []
            
            self.sistema.dados_para_incluir.append(dados_financeiros)
            
            print(f"💰 LANÇAMENTO CRIADO: R$ {dados_nfe.get('valor_total', 0):,.2f} - {self.periodo_nome}")
            
            return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            
        except Exception as e:
            raise Exception(f"Erro ao criar lançamento: {str(e)}")
    
    def criar_materiais_otimizado(self, opcoes):
        """Cria materiais otimizado"""
        try:
            produtos = self.dados_nfe.get('produtos', [])
            if not produtos:
                return "Nenhum produto"
            
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                return f"{len(produtos)} produtos (sistema materiais não inicializado)"
            
            materiais_criados = 0
            dados_nfe = self.dados_nfe
            
            for produto in produtos:
                try:
                    dados_material = {
                        'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                        'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                        'Subcategoria': '',
                        'Codigo_Produto': produto.get('codigo', ''),
                        'Descricao_Completa': produto.get('descricao', ''),
                        'Marca': dados_nfe.get('razao_social_emitente', '')[:20],
                        'Ambiente_Aplicacao': opcoes['ambiente'],
                        'Status_Instalacao': opcoes['status'],
                        'Garantia_Meses': opcoes['garantia'],
                        'Observacoes': f"Importado NFe {dados_nfe.get('numero_nf', '')} - {self.periodo_nome}",
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
                    
                    material_id = self.sistema.gerenciador_materiais.salvar_material(dados_material)
                    materiais_criados += 1
                    
                except Exception as e:
                    print(f"⚠️ Erro ao criar material: {e}")
                    continue
            
            return f"{materiais_criados} de {len(produtos)} produtos"
            
        except Exception as e:
            raise Exception(f"Erro ao criar materiais: {str(e)}")
    
    def mostrar_resultado_final_otimizado(self, resultados, opcoes):
        """Mostra resultado final otimizado"""
        janela_resultado = tk.Toplevel(self.sistema.root)
        janela_resultado.title("🎉 NFe Importada com Sucesso!")
        janela_resultado.geometry("600x350")
        janela_resultado.grab_set()
        
        frame_main = ttk.Frame(janela_resultado)
        frame_main.pack(fill='both', expand=True, padx=20, pady=15)
        
        # TÍTULO DE SUCESSO
        titulo_frame = ttk.Frame(frame_main)
        titulo_frame.pack(fill='x', pady=5)
        
        tk.Label(titulo_frame, text="🎉", font=('Arial', 24)).pack(side='left')
        titulo_text = f"NFe {self.dados_nfe.get('numero_nf', '')} IMPORTADA COM SUCESSO!"
        tk.Label(titulo_frame, text=titulo_text, 
                font=('Arial', 13, 'bold'), fg='darkgreen').pack(side='left', padx=(10,0))
        
        # RESUMO DOS RESULTADOS
        resumo_frame = ttk.LabelFrame(frame_main, text="📊 Dados Importados", padding=10)
        resumo_frame.pack(fill='x', pady=8)
        
        for resultado in resultados:
            tk.Label(resumo_frame, text=resultado, fg='blue', 
                    font=('Arial', 11, 'bold')).pack(anchor='w', pady=1)
        
        # PERÍODO E REFERÊNCIA
        periodo_frame = ttk.LabelFrame(frame_main, text="📅 Período do Relatório", padding=10)
        periodo_frame.pack(fill='x', pady=8)
        
        periodo_info = f"🎯 {self.periodo_nome} - Data: {self.data_referencia}\n📋 Referência: {opcoes.get('referencia', 'N/A')}"
        tk.Label(periodo_frame, text=periodo_info, justify='left', 
                font=('Arial', 10), fg='darkblue').pack(anchor='w')
        
        # BOTÕES DE AÇÃO
        botoes_frame = ttk.Frame(frame_main)
        botoes_frame.pack(fill='x', pady=15)
        
        # BOTÃO PRINCIPAL
        btn_processar = ttk.Button(botoes_frame, 
                                  text="📊 PROCESSAR NO SISTEMA AGORA", 
                                  command=lambda: self.processar_no_sistema_otimizado(janela_resultado))
        btn_processar.pack(side='left', padx=5)
        
        # BOTÃO SECUNDÁRIO
        ttk.Button(botoes_frame, 
                  text="✅ Concluído", 
                  command=janela_resultado.destroy).pack(side='right', padx=5)
        
        # DICA
        dica_text = "💡 Use 'Processar no Sistema' para salvar os dados na planilha imediatamente"
        tk.Label(frame_main, text=dica_text, font=('Arial', 8), 
                fg='gray').pack(pady=(5,0))
    
    def processar_no_sistema_otimizado(self, janela_resultado):
        """Processa no sistema com feedback otimizado"""
        try:
            janela_resultado.destroy()
            
            if hasattr(self.sistema, 'dados_para_incluir') and self.sistema.dados_para_incluir:
                # MOSTRAR PROGRESSO
                progresso = tk.Toplevel(self.sistema.root)
                progresso.title("Processando...")
                progresso.geometry("300x100")
                progresso.grab_set()
                
                tk.Label(progresso, text="📊 Salvando na planilha...", 
                        font=('Arial', 10, 'bold')).pack(pady=20)
                progresso.update()
                
                # PROCESSAR
                self.sistema.enviar_dados()
                
                # FECHAR PROGRESSO
                progresso.destroy()
                
                # CONFIRMAR SUCESSO
                messagebox.showinfo("✅ Processamento Concluído", 
                    f"✅ {len(self.sistema.dados_para_incluir)} lançamento(s) processado(s)\n"
                    f"📊 Dados salvos na planilha do sistema\n"
                    f"🎯 NFe incluída no relatório de {self.data_referencia}")
            else:
                messagebox.showwarning("⚠️ Aviso", "Não há dados para processar!")
                
        except Exception as e:
            messagebox.showerror("❌ Erro", f"Erro ao processar:\n{str(e)}")


# FUNÇÃO PRINCIPAL ULTRA OTIMIZADA
def inicializar_sistema_nfe_ultra_otimizado(sistema_principal):
    """
    Inicialização ultra otimizada: clique único → seleção → configuração
    """
    try:
        print("🚀 Inicializando Sistema NFe Ultra Otimizado...")
        
        # LIMPAR SISTEMAS ANTERIORES
        limpar_todos_sistemas_nfe(sistema_principal)
        
        # CRIAR SISTEMA ULTRA OTIMIZADO
        sistema_nfe = SistemaNFeUltraOtimizado(sistema_principal)
        sistema_principal.sistema_nfe_ultra = sistema_nfe
        
        # MÉTODO DE CONVENIÊNCIA
        sistema_principal.abrir_importacao_nfe = sistema_nfe.abrir_importacao_nfe
        
        # ADICIONAR BOTÃO OTIMIZADO
        adicionar_botao_nfe_ultra_otimizado(sistema_principal)
        
        print("✅ Sistema NFe Ultra Otimizado inicializado!")
        print("🎯 FLUXO: Clique no Botão → Selecionar XML → Configurar → Importar")
        print("⚡ OTIMIZAÇÃO: Eliminadas TODAS as telas intermediárias!")
        
        return sistema_nfe
        
    except Exception as e:
        print(f"❌ Erro ao inicializar sistema ultra otimizado: {e}")
        return None


def limpar_todos_sistemas_nfe(sistema_principal):
    """Remove TODOS os sistemas NFe anteriores"""
    try:
        print("🧹 Limpando TODOS os sistemas NFe anteriores...")
        
        # TODOS OS ATRIBUTOS NFE
        atributos_para_remover = [
            'sistema_nfe_unificado',
            'sistema_nfe_simplificado', 
            'sistema_nfe_com_botao',
            'sistema_nfe_ultra',
            'processador_nfe',
            'integrador_nfe',
            'importar_nfe_xml',
            'importar_nfe_com_interface'
        ]
        
        for atributo in atributos_para_remover:
            if hasattr(sistema_principal, atributo):
                delattr(sistema_principal, atributo)
                print(f"  ✅ Removido: {atributo}")
        
        print("✅ Limpeza total concluída!")
        
    except Exception as e:
        print(f"❌ Erro na limpeza: {e}")


def adicionar_botao_nfe_ultra_otimizado(sistema_principal):
    """Adiciona botão NFe ultra otimizado"""
    try:
        print("⚡ Adicionando botão NFe ultra otimizado...")
        
        if not hasattr(sistema_principal, 'aba_fornecedor'):
            print("❌ aba_fornecedor não encontrada")
            return
        
        # ENCONTRAR SEÇÃO DE MATERIAIS
        frame_materiais = None
        for widget in sistema_principal.aba_fornecedor.winfo_children():
            if hasattr(widget, 'configure') and 'text' in widget.configure():
                if 'Materiais' in widget['text']:
                    frame_materiais = widget
                    break
        
        if not frame_materiais:
            print("❌ Seção de materiais não encontrada")
            return
        
        # ENCONTRAR FRAME DE BOTÕES
        frame_botoes = None
        for subwidget in frame_materiais.winfo_children():
            if str(type(subwidget)).endswith("Frame'>"):
                tem_botoes = any(str(type(child)).endswith("Button'>") 
                               for child in subwidget.winfo_children())
                if tem_botoes:
                    frame_botoes = subwidget
                    break
        
        if frame_botoes:
            # REMOVER BOTÕES NFE ANTIGOS
            botoes_nfe = [w for w in frame_botoes.winfo_children() 
                         if hasattr(w, 'configure') and 'text' in w.configure() 
                         and 'NFe' in w['text']]
            
            for btn in botoes_nfe:
                btn.destroy()
            
            # CRIAR BOTÃO ULTRA OTIMIZADO
            btn_nfe = ttk.Button(
                frame_botoes,
                text="⚡ Importar NFe", 
                command=sistema_principal.abrir_importacao_nfe
            )
            btn_nfe.pack(side='left', padx=5)
            
            print("✅ Botão NFe ultra otimizado adicionado!")
            print("⚡ FLUXO: Clique → Selecionar → Configurar → Importar")
        else:
            print("❌ Frame de botões não encontrado")
            
    except Exception as e:
        print(f"❌ Erro ao adicionar botão: {e}")


def debug_sistema_ultra_otimizado(sistema_principal):
    """Debug do sistema ultra otimizado"""
    print("🔍 DEBUG SISTEMA NFe ULTRA OTIMIZADO")
    print("=" * 50)
    
    # VERIFICAÇÕES
    tem_sistema = hasattr(sistema_principal, 'sistema_nfe_ultra')
    tem_metodo = hasattr(sistema_principal, 'abrir_importacao_nfe')
    
    print(f"⚡ Sistema ultra inicializado: {tem_sistema}")
    print(f"⚡ Método disponível: {tem_metodo}")
    
    # VERIFICAR BOTÃO
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
            print(f"⚡ Botões encontrados: {botoes_nfe}")
        else:
            print("❌ Nenhum botão NFe encontrado")
    
    if tem_sistema and tem_metodo:
        print("🎯 SISTEMA ULTRA OTIMIZADO PRONTO!")
        print("⚡ FLUXO: Clique → Selecionar → Configurar → Importar")
        return True
    else:
        print("❌ Sistema não funcional")
        return False