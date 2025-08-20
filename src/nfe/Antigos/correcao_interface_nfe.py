# -*- coding: utf-8 -*-
"""
CORREÇÃO DA INTERFACE NFe
Aplica os ajustes visuais que não foram aplicados automaticamente
"""

import tkinter as tk
from tkinter import ttk
from datetime import datetime
import json
from pathlib import Path

def corrigir_interface_nfe_manualmente(sistema_principal):
    """
    Substitui o método importar_para_sistema por uma versão com interface aprimorada
    """
    try:
        print("🔧 Corrigindo interface NFe manualmente...")
        
        # VERIFICAR SE SISTEMA NFe EXISTE
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        sistema_nfe = sistema_principal.sistema_nfe_unificado
        
        # SUBSTITUIR MÉTODO importar_para_sistema
        def importar_para_sistema_aprimorado(self):
            """Versão aprimorada com interface corrigida"""
            try:
                if not self.dados_nfe_atual:
                    tk.messagebox.showerror("Erro", "Nenhum dado carregado!")
                    return
                
                # FECHAR JANELA ATUAL
                self.janela.withdraw()  # Esconder em vez de destruir
                
                # ABRIR INTERFACE APRIMORADA
                InterfaceNFeCorrigida(self.sistema, self.dados_nfe_atual, self.janela)
                
            except Exception as e:
                tk.messagebox.showerror("Erro", f"Erro na importação: {str(e)}")
        
        # APLICAR SUBSTITUIÇÃO
        sistema_nfe.importar_para_sistema = importar_para_sistema_aprimorado.__get__(sistema_nfe, type(sistema_nfe))
        
        print("✅ Interface NFe corrigida com sucesso!")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao corrigir interface: {e}")
        return False


class InterfaceNFeCorrigida:
    """Interface NFe com todos os ajustes visuais aplicados"""
    
    def __init__(self, sistema_principal, dados_nfe, janela_anterior):
        self.sistema = sistema_principal
        self.dados_nfe = dados_nfe
        self.janela_anterior = janela_anterior
        
        # CARREGAR PARÂMETROS
        self.carregar_parametros()
        
        # CRIAR INTERFACE
        self.criar_interface_aprimorada()
    
    def carregar_parametros(self):
        """Carrega parâmetros do sistema"""
        try:
            # PARÂMETROS DO SISTEMA
            self.parametros_sistema = self.carregar_parametros_sistema()
            self.parametros_materiais = self.carregar_parametros_materiais()
        except Exception as e:
            print(f"⚠️ Erro ao carregar parâmetros: {e}")
            self.parametros_sistema = {"etapas_obra": ["MATERIAIS"]}
            self.parametros_materiais = {"ambientes": ["GERAL"], "status_instalacao": ["PENDENTE"]}
    
    def carregar_parametros_sistema(self):
        """Carrega parametros_sistema.json"""
        caminhos = [
            Path("parametros_sistema.json"),
            Path("data/parametros_sistema.json"),
            Path("config/parametros_sistema.json")
        ]
        
        for caminho in caminhos:
            if caminho.exists():
                try:
                    with open(caminho, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except:
                    continue
        
        return {
            "etapas_obra": [
                "INSTALAÇÃO DA OBRA", "FUNDAÇÃO", "ESTRUTURA", "ALVENARIA",
                "COBERTURA", "INSTALAÇÕES", "ACABAMENTOS", "MATERIAIS", "FINALIZAÇÃO"
            ]
        }
    
    def carregar_parametros_materiais(self):
        """Carrega parametros_materiais.json"""
        caminhos = [
            Path("data/materiais/parametros_materiais.json"),
            Path("data/parametros_materiais.json"),
            Path("parametros_materiais.json")
        ]
        
        for caminho in caminhos:
            if caminho.exists():
                try:
                    with open(caminho, 'r', encoding='utf-8') as f:
                        return json.load(f)
                except:
                    continue
        
        return {
            "ambientes": [
                "INSTALAÇÃO DA OBRA", "SALA DE ESTAR", "COZINHA", "DORMITÓRIOS",
                "BANHEIROS", "ÁREA EXTERNA", "GERAL"
            ],
            "status_instalacao": ["PENDENTE", "INSTALADO", "EM_INSTALACAO", "DEFEITO"],
            "unidades": ["UN", "M", "M²", "KG", "L", "PC", "CX"]
        }
    
    def criar_interface_aprimorada(self):
        """Cria a interface aprimorada"""
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("⚙️ Configuração de Importação NFe - VERSÃO APRIMORADA")
        self.janela.geometry("750x700")
        self.janela.grab_set()
        
        # PROTOCOLO DE FECHAMENTO
        self.janela.protocol("WM_DELETE_WINDOW", self.fechar_janela)
        
        # FRAME PRINCIPAL COM SCROLL
        self.criar_frame_com_scroll()
        
        # SEÇÕES DA INTERFACE
        self.criar_cabecalho()
        self.criar_resumo_nfe()
        self.criar_opcoes_importacao()
        self.criar_configuracoes_financeiras()
        self.criar_configuracoes_materiais()
        self.criar_botoes_finais()
        
        # CONFIGURAR ESTADOS INICIAIS
        self.configurar_estados_iniciais()
    
    def criar_frame_com_scroll(self):
        """Cria frame principal com scroll"""
        # CANVAS E SCROLLBAR
        self.canvas = tk.Canvas(self.janela, bg='white')
        self.scrollbar = ttk.Scrollbar(self.janela, orient="vertical", command=self.canvas.yview)
        self.frame_scroll = ttk.Frame(self.canvas)
        
        # CONFIGURAR SCROLL
        self.frame_scroll.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )
        
        self.canvas.create_window((0, 0), window=self.frame_scroll, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        # BIND MOUSE WHEEL
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        self.canvas.bind("<MouseWheel>", _on_mousewheel)
        
        # PACK
        self.canvas.pack(side="left", fill="both", expand=True)
        self.scrollbar.pack(side="right", fill="y")
    
    def criar_cabecalho(self):
        """Cria cabeçalho da interface"""
        frame_header = ttk.Frame(self.frame_scroll)
        frame_header.pack(fill='x', padx=15, pady=10)
        
        # TÍTULO
        titulo = tk.Label(frame_header, 
                         text="⚙️ CONFIGURAÇÃO DE IMPORTAÇÃO NFe",
                         font=('Arial', 14, 'bold'),
                         fg='darkblue')
        titulo.pack(pady=5)
        
        # SUBTÍTULO
        subtitulo = tk.Label(frame_header,
                           text="Configure como os dados da NFe serão importados para o sistema",
                           font=('Arial', 9),
                           fg='gray')
        subtitulo.pack(pady=2)
    
    def criar_resumo_nfe(self):
        """Cria seção de resumo da NFe"""
        frame_resumo = ttk.LabelFrame(self.frame_scroll, text="📄 Dados da NFe", padding=15)
        frame_resumo.pack(fill='x', padx=15, pady=5)
        
        # GRID DE INFORMAÇÕES
        info_frame = ttk.Frame(frame_resumo)
        info_frame.pack(fill='x')
        
        dados = [
            ("📄 Número:", self.dados_nfe.get('numero_nf', '')),
            ("📅 Data Emissão:", self.dados_nfe.get('data_emissao', '')),
            ("🏢 Fornecedor:", self.dados_nfe.get('razao_social_emitente', '')[:45]),
            ("💰 Valor Total:", f"R$ {self.dados_nfe.get('valor_total', 0):,.2f}"),
            ("📦 Produtos:", str(len(self.dados_nfe.get('produtos', [])))),
            ("🔑 Chave:", self.dados_nfe.get('chave_acesso', '')[:25] + "...")
        ]
        
        for i, (label, valor) in enumerate(dados):
            row = i // 2
            col = (i % 2) * 2
            
            tk.Label(info_frame, text=label, font=('Arial', 9, 'bold')).grid(
                row=row, column=col, sticky='w', padx=(0,5), pady=3)
            tk.Label(info_frame, text=valor, font=('Arial', 9)).grid(
                row=row, column=col+1, sticky='w', padx=(0,20), pady=3)
    
    def criar_opcoes_importacao(self):
        """Cria seção de opções de importação"""
        frame_opcoes = ttk.LabelFrame(self.frame_scroll, text="⚙️ O que Importar", padding=15)
        frame_opcoes.pack(fill='x', padx=15, pady=5)
        
        # VARIÁVEIS DE CONTROLE
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        # CHECKBOXES
        cb_financeiro = tk.Checkbutton(
            frame_opcoes,
            text="💰 Dados Financeiros (lançamento de despesa no sistema)",
            variable=self.importar_financeiro,
            font=('Arial', 10, 'bold'),
            command=self.toggle_configuracoes_financeiras
        )
        cb_financeiro.pack(anchor='w', pady=3)
        
        cb_materiais = tk.Checkbutton(
            frame_opcoes,
            text="📦 Materiais da Obra (banco de dados para manual do proprietário)",
            variable=self.importar_materiais,
            font=('Arial', 10, 'bold'),
            command=self.toggle_configuracoes_materiais
        )
        cb_materiais.pack(anchor='w', pady=3)
    
    def criar_configuracoes_financeiras(self):
        """Cria seção de configurações financeiras com ajustes de data"""
        self.frame_financeiro = ttk.LabelFrame(self.frame_scroll, 
                                              text="💰 Configurações Financeiras", 
                                              padding=15)
        self.frame_financeiro.pack(fill='x', padx=15, pady=5)
        
        # ===== SEÇÃO DE DATAS (AJUSTE PRINCIPAL) =====
        frame_datas = ttk.LabelFrame(self.frame_financeiro, text="📅 Datas do Sistema", padding=10)
        frame_datas.pack(fill='x', pady=5)
        
        # CALCULAR DATAS AUTOMATICAMENTE
        data_rel, dt_vencto = self.ajustar_datas_sistema()
        
        # DATA DE REFERÊNCIA
        linha_ref = ttk.Frame(frame_datas)
        linha_ref.pack(fill='x', pady=3)
        
        tk.Label(linha_ref, text="Data Referência (5 ou 20):", 
                font=('Arial', 9, 'bold'), fg='darkgreen').pack(side='left')
        self.data_rel = tk.Entry(linha_ref, width=12, font=('Arial', 9, 'bold'))
        self.data_rel.pack(side='left', padx=5)
        self.data_rel.insert(0, data_rel)
        
        tk.Label(linha_ref, text="💡 Sistema só trabalha com dia 5 ou 20", 
                fg='blue', font=('Arial', 8)).pack(side='left', padx=10)
        
        # DATA DE VENCIMENTO
        linha_venc = ttk.Frame(frame_datas)
        linha_venc.pack(fill='x', pady=3)
        
        tk.Label(linha_venc, text="Data Vencimento (da NFe):", 
                font=('Arial', 9, 'bold'), fg='darkred').pack(side='left')
        self.dt_vencto = tk.Entry(linha_venc, width=12, font=('Arial', 9, 'bold'))
        self.dt_vencto.pack(side='left', padx=5)
        self.dt_vencto.insert(0, dt_vencto)
        
        tk.Label(linha_venc, text="📄 Data original da nota fiscal", 
                fg='gray', font=('Arial', 8)).pack(side='left', padx=10)
        
        # ===== CLASSIFICAÇÃO =====
        frame_class = ttk.LabelFrame(self.frame_financeiro, text="🏷️ Classificação", padding=10)
        frame_class.pack(fill='x', pady=5)
        
        linha_class = ttk.Frame(frame_class)
        linha_class.pack(fill='x', pady=3)
        
        # TIPO DESPESA
        tk.Label(linha_class, text="Tipo Despesa:").pack(side='left')
        self.tipo_despesa = ttk.Combobox(linha_class, width=12, state='readonly')
        self.tipo_despesa['values'] = ['1', '2', '3', '4', '5', '6', '7']
        self.tipo_despesa.set('3')
        self.tipo_despesa.pack(side='left', padx=5)
        
        # CATEGORIA
        tk.Label(linha_class, text="Categoria:").pack(side='left', padx=(20,0))
        self.categoria_financeira = tk.Entry(linha_class, width=8)
        self.categoria_financeira.insert(0, 'MAT')
        self.categoria_financeira.pack(side='left', padx=5)
        
        # FORMA PAGAMENTO
        tk.Label(linha_class, text="Forma Pgto:").pack(side='left', padx=(20,0))
        self.forma_pagamento = ttk.Combobox(linha_class, width=12, state='readonly')
        self.forma_pagamento['values'] = ['A_VISTA', 'A_PRAZO', 'CARTAO', 'PIX']
        self.forma_pagamento.set('A_PRAZO')
        self.forma_pagamento.pack(side='left', padx=5)
        
        # ===== ETAPA E REFERÊNCIA =====
        frame_etapa = ttk.LabelFrame(self.frame_financeiro, text="🏗️ Etapa e Referência", padding=10)
        frame_etapa.pack(fill='x', pady=5)
        
        # ETAPA DA OBRA (COM PARÂMETROS)
        linha_etapa = ttk.Frame(frame_etapa)
        linha_etapa.pack(fill='x', pady=3)
        
        tk.Label(linha_etapa, text="Etapa da Obra:").pack(side='left')
        self.etapa_obra = ttk.Combobox(linha_etapa, width=25, state='readonly')
        etapas = self.parametros_sistema.get('etapas_obra', ['MATERIAIS'])
        self.etapa_obra['values'] = etapas
        self.etapa_obra.set('MATERIAIS')
        self.etapa_obra.pack(side='left', padx=5)
        
        # REFERÊNCIA EDITÁVEL (AJUSTE PRINCIPAL)
        linha_ref = ttk.Frame(frame_etapa)
        linha_ref.pack(fill='x', pady=3)
        
        tk.Label(linha_ref, text="Referência:", 
                font=('Arial', 9, 'bold'), fg='purple').pack(side='left')
        self.referencia_editavel = tk.Entry(linha_ref, width=60, font=('Arial', 9))
        self.referencia_editavel.pack(side='left', padx=5, fill='x', expand=True)
        
        # REFERÊNCIA INICIAL INTELIGENTE
        ref_inicial = f"NFE {self.dados_nfe.get('numero_nf', '')} - {self.dados_nfe.get('razao_social_emitente', '')[:30]}".upper()
        self.referencia_editavel.insert(0, ref_inicial)
    
    def criar_configuracoes_materiais(self):
        """Cria seção de configurações materiais detalhada"""
        self.frame_materiais = ttk.LabelFrame(self.frame_scroll, 
                                             text="📦 Configurações Materiais", 
                                             padding=15)
        self.frame_materiais.pack(fill='x', padx=15, pady=5)
        
        # ===== LOCALIZAÇÃO =====
        frame_local = ttk.LabelFrame(self.frame_materiais, text="🏠 Localização na Obra", padding=10)
        frame_local.pack(fill='x', pady=5)
        
        linha_amb = ttk.Frame(frame_local)
        linha_amb.pack(fill='x', pady=3)
        
        tk.Label(linha_amb, text="Ambiente Padrão:").pack(side='left')
        self.ambiente_padrao = ttk.Combobox(linha_amb, width=25, state='readonly')
        ambientes = self.parametros_materiais.get('ambientes', ['GERAL'])
        self.ambiente_padrao['values'] = ambientes
        self.ambiente_padrao.pack(side='left', padx=5)
        
        linha_loc = ttk.Frame(frame_local)
        linha_loc.pack(fill='x', pady=3)
        
        tk.Label(linha_loc, text="Localização Específica:").pack(side='left')
        self.localizacao_especifica = tk.Entry(linha_loc, width=40)
        self.localizacao_especifica.insert(0, "Conforme projeto")
        self.localizacao_especifica.pack(side='left', padx=5)
        
        # ===== STATUS E GARANTIA =====
        frame_status = ttk.LabelFrame(self.frame_materiais, text="⚙️ Status e Garantia", padding=10)
        frame_status.pack(fill='x', pady=5)
        
        linha_st = ttk.Frame(frame_status)
        linha_st.pack(fill='x', pady=3)
        
        tk.Label(linha_st, text="Status:").pack(side='left')
        self.status_instalacao = ttk.Combobox(linha_st, width=20, state='readonly')
        status_list = self.parametros_materiais.get('status_instalacao', ['PENDENTE'])
        self.status_instalacao['values'] = status_list
        self.status_instalacao.set('PENDENTE')
        self.status_instalacao.pack(side='left', padx=5)
        
        tk.Label(linha_st, text="Garantia:").pack(side='left', padx=(20,0))
        self.garantia_meses = tk.Entry(linha_st, width=5)
        self.garantia_meses.insert(0, '12')
        self.garantia_meses.pack(side='left', padx=5)
        tk.Label(linha_st, text="meses").pack(side='left')
        
        # ===== FORNECEDOR =====
        frame_fornec = ttk.LabelFrame(self.frame_materiais, text="🏢 Dados do Fornecedor", padding=10)
        frame_fornec.pack(fill='x', pady=5)
        
        linha_marca = ttk.Frame(frame_fornec)
        linha_marca.pack(fill='x', pady=3)
        
        tk.Label(linha_marca, text="Marca/Fabricante:").pack(side='left')
        self.marca_fabricante = tk.Entry(linha_marca, width=35)
        marca_sugerida = self.dados_nfe.get('razao_social_emitente', '')[:30]
        self.marca_fabricante.insert(0, marca_sugerida)
        self.marca_fabricante.pack(side='left', padx=5)
    
    def criar_botoes_finais(self):
        """Cria botões finais"""
        frame_botoes = ttk.Frame(self.frame_scroll)
        frame_botoes.pack(fill='x', padx=15, pady=20)
        
        # BOTÕES PRINCIPAIS
        ttk.Button(frame_botoes, 
                  text="👁️ Preview dos Dados", 
                  command=self.preview_dados,
                  style='Accent.TButton').pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ PROCESSAR IMPORTAÇÃO", 
                  command=self.processar_importacao,
                  style='Accent.TButton').pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="❌ Cancelar", 
                  command=self.fechar_janela).pack(side='right', padx=5)
    
    def configurar_estados_iniciais(self):
        """Configura estados iniciais dos widgets"""
        self.toggle_configuracoes_financeiras()
        self.toggle_configuracoes_materiais()
    
    def ajustar_datas_sistema(self):
        """Ajusta datas para o padrão do sistema"""
        try:
            data_emissao = self.dados_nfe.get('data_emissao', '')
            
            if not data_emissao:
                hoje = datetime.now()
                return hoje.replace(day=5).strftime('%d/%m/%Y'), hoje.strftime('%d/%m/%Y')
            
            # CONVERTER DATA DA NFE
            if '/' in data_emissao:
                dt_nfe = datetime.strptime(data_emissao, '%d/%m/%Y')
            else:
                dt_nfe = datetime.strptime(data_emissao[:10], '%Y-%m-%d')
            
            # DATA DE VENCIMENTO = DATA DA NFE
            dt_vencto = dt_nfe.strftime('%d/%m/%Y')
            
            # DATA DE REFERÊNCIA = DIA 5 OU 20
            if dt_nfe.day <= 12:
                dia_ref = 5
            else:
                dia_ref = 20
            
            dt_ref = dt_nfe.replace(day=dia_ref)
            data_rel = dt_ref.strftime('%d/%m/%Y')
            
            print(f"📅 Datas ajustadas: NFe {dt_vencto} → Ref {data_rel}")
            
            return data_rel, dt_vencto
            
        except Exception as e:
            print(f"❌ Erro ao ajustar datas: {e}")
            hoje = datetime.now()
            return hoje.replace(day=5).strftime('%d/%m/%Y'), hoje.strftime('%d/%m/%Y')
    
    def toggle_configuracoes_financeiras(self):
        """Liga/desliga configurações financeiras"""
        estado = 'normal' if self.importar_financeiro.get() else 'disabled'
        self.alterar_estado_frame(self.frame_financeiro, estado)
    
    def toggle_configuracoes_materiais(self):
        """Liga/desliga configurações materiais"""
        estado = 'normal' if self.importar_materiais.get() else 'disabled'
        self.alterar_estado_frame(self.frame_materiais, estado)
    
    def alterar_estado_frame(self, frame, estado):
        """Altera estado de todos os widgets em um frame"""
        def alterar_recursivo(widget):
            if isinstance(widget, (tk.Entry, ttk.Combobox)):
                widget.config(state=estado)
            elif hasattr(widget, 'winfo_children'):
                for child in widget.winfo_children():
                    alterar_recursivo(child)
        
        alterar_recursivo(frame)
    
    def preview_dados(self):
        """Mostra preview dos dados"""
        opcoes = self.coletar_opcoes()
        
        # CRIAR JANELA DE PREVIEW
        janela_preview = tk.Toplevel(self.janela)
        janela_preview.title("👁️ Preview da Importação")
        janela_preview.geometry("700x500")
        janela_preview.grab_set()
        
        # TEXTO DE PREVIEW
        text_preview = tk.Text(janela_preview, wrap='word', font=('Courier', 10))
        scrollbar_preview = ttk.Scrollbar(janela_preview, orient='vertical', command=text_preview.yview)
        text_preview.configure(yscrollcommand=scrollbar_preview.set)
        
        # GERAR PREVIEW
        preview_content = self.gerar_preview_completo(opcoes)
        text_preview.insert('1.0', preview_content)
        text_preview.config(state='disabled')
        
        # PACK
        text_preview.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        scrollbar_preview.pack(side='right', fill='y', pady=10)
        
        # BOTÃO FECHAR
        ttk.Button(janela_preview, text="Fechar", 
                  command=janela_preview.destroy).pack(pady=10)
    
    def gerar_preview_completo(self, opcoes):
        """Gera preview completo dos dados"""
        dados_nfe = self.dados_nfe
        preview = f"""
🎯 PREVIEW DA IMPORTAÇÃO NFe
{'='*50}

📄 DADOS DA NFe:
   Número: {dados_nfe.get('numero_nf', '')}
   Fornecedor: {dados_nfe.get('razao_social_emitente', '')}
   Data Emissão: {dados_nfe.get('data_emissao', '')}
   Valor Total: R$ {dados_nfe.get('valor_total', 0):,.2f}
   Produtos: {len(dados_nfe.get('produtos', []))}

"""
        
        if opcoes['importar_financeiro']:
            preview += f"""💰 LANÇAMENTO FINANCEIRO:
   📅 Data Referência: {opcoes['data_rel']} (padrão sistema)
   📅 Data Vencimento: {opcoes['dt_vencto']} (data NFe)
   🏷️ Categoria: {opcoes['categoria_financeira']}
   🔧 Tipo Despesa: {opcoes['tipo_despesa']}
   🏗️ Etapa: {opcoes['etapa_obra']}
   📋 Referência: {opcoes['referencia']}
   💳 Forma Pgto: {opcoes['forma_pagamento']}

"""
        
        if opcoes['importar_materiais']:
            preview += f"""📦 CONFIGURAÇÕES MATERIAIS:
   🏠 Ambiente: {opcoes['ambiente_padrao']}
   📍 Localização: {opcoes['localizacao_especifica']}
   ⚙️ Status: {opcoes['status_instalacao']}
   🛡️ Garantia: {opcoes['garantia_meses']} meses
   🏢 Marca: {opcoes['marca_fabricante']}

   PRODUTOS QUE SERÃO IMPORTADOS:
"""
            
            for i, produto in enumerate(dados_nfe.get('produtos', []), 1):
                preview += f"   {i:2d}. {produto.get('descricao', '')[:50]}\n"
                preview += f"       Categoria: {produto.get('categoria_sugerida', 'OUTROS')}\n"
                preview += f"       Qtd: {produto.get('quantidade', '')} | Valor: R$ {produto.get('valor_total', 0):.2f}\n\n"
        
        preview += f"""
⚠️ PRÓXIMOS PASSOS:
1. Dados serão adicionados ao sistema
2. Use 'Enviar Dados' para salvar na planilha
3. Materiais ficarão disponíveis para consulta
4. Manual do proprietário poderá ser gerado
"""
        
        return preview.strip()
    
    def coletar_opcoes(self):
        """Coleta todas as opções configuradas"""
        return {
            'importar_financeiro': self.importar_financeiro.get(),
            'importar_materiais': self.importar_materiais.get(),
            'data_rel': self.data_rel.get(),
            'dt_vencto': self.dt_vencto.get(),
            'tipo_despesa': self.tipo_despesa.get(),
            'categoria_financeira': self.categoria_financeira.get(),
            'etapa_obra': self.etapa_obra.get(),
            'forma_pagamento': self.forma_pagamento.get(),
            'referencia': self.referencia_editavel.get(),
            'ambiente_padrao': self.ambiente_padrao.get(),
            'localizacao_especifica': self.localizacao_especifica.get(),
            'status_instalacao': self.status_instalacao.get(),
            'garantia_meses': int(self.garantia_meses.get() or 12),
            'marca_fabricante': self.marca_fabricante.get()
        }
    
    def processar_importacao(self):
        """Processa a importação com as configurações"""
        try:
            # VALIDAR SELEÇÕES
            if not self.importar_financeiro.get() and not self.importar_materiais.get():
                tk.messagebox.showwarning("Aviso", "Selecione pelo menos uma opção!")
                return
            
            # VALIDAR DATAS
            if self.importar_financeiro.get():
                data_rel = self.data_rel.get().strip()
                if data_rel:
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
                                data_corrigida, _ = self.ajustar_datas_sistema()
                                self.data_rel.delete(0, tk.END)
                                self.data_rel.insert(0, data_corrigida)
                            else:
                                return
                    except ValueError:
                        tk.messagebox.showerror("Erro", "Data de referência inválida!")
                        return
            
            # COLETAR OPÇÕES
            opcoes = self.coletar_opcoes()
            
            # EXECUTAR IMPORTAÇÃO
            self.executar_importacao_completa(opcoes)
            
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")
    
    def executar_importacao_completa(self, opcoes):
        """Executa a importação completa"""
        try:
            resultados = []
            
            # IMPORTAR FINANCEIRO
            if opcoes['importar_financeiro']:
                resultado_fin = self.criar_lancamento_financeiro_completo(opcoes)
                resultados.append(f"💰 Financeiro: {resultado_fin}")
            
            # IMPORTAR MATERIAIS
            if opcoes['importar_materiais']:
                resultado_mat = self.criar_materiais_completos(opcoes)
                resultados.append(f"📦 Materiais: {resultado_mat}")
            
            # FECHAR JANELAS
            self.fechar_janela()
            
            # MOSTRAR RESULTADO FINAL
            self.mostrar_resultado_importacao(resultados, opcoes)
            
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro na importação: {str(e)}")
    
    def criar_lancamento_financeiro_completo(self, opcoes):
        """Cria lançamento financeiro com todos os ajustes"""
        try:
            dados_nfe = self.dados_nfe
            
            # PREPARAR DADOS FINANCEIROS
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
                'dt_vencto': opcoes['dt_vencto'],  # Data da NFe
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - CHAVE: {dados_nfe.get('chave_acesso', '')[:20]}...".upper(),
                'forma_pagamento': opcoes['forma_pagamento']
            }
            
            # ADICIONAR À LISTA DO SISTEMA
            if not hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir = []
            
            self.sistema.dados_para_incluir.append(dados_financeiros)
            
            print(f"💰 Lançamento financeiro criado:")
            print(f"   📅 Data Ref: {opcoes['data_rel']} | Vencto: {opcoes['dt_vencto']}")
            print(f"   📋 Referência: {opcoes['referencia']}")
            print(f"   💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}")
            
            return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            
        except Exception as e:
            raise Exception(f"Erro ao criar lançamento: {str(e)}")
    
    def criar_materiais_completos(self, opcoes):
        """Cria materiais com todas as configurações"""
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
                        'Data_Instalacao': '',
                        'Instalador': '',
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
                    
                    print(f"📦 Material criado - ID: {material_id} - {produto.get('descricao', '')[:40]}")
                    
                except Exception as e:
                    print(f"⚠️ Erro ao criar material {produto.get('descricao', '')}: {e}")
                    continue
            
            return f"{materiais_criados} de {len(produtos)} produtos"
            
        except Exception as e:
            raise Exception(f"Erro ao criar materiais: {str(e)}")
    
    def gerar_especificacoes_produto(self, produto):
        """Gera especificações técnicas do produto"""
        specs = []
        
        if produto.get('ncm'):
            specs.append(f"NCM: {produto['ncm']}")
        
        if produto.get('cfop'):
            specs.append(f"CFOP: {produto['cfop']}")
        
        if produto.get('unidade') and produto.get('quantidade'):
            specs.append(f"Embalagem: {produto['quantidade']} {produto['unidade']}")
        
        return " | ".join(specs) if specs else "Conforme NFe"
    
    def mostrar_resultado_importacao(self, resultados, opcoes):
        """Mostra resultado final da importação"""
        janela_resultado = tk.Toplevel(self.sistema.root)
        janela_resultado.title("🎉 Importação NFe Concluída")
        janela_resultado.geometry("650x600")
        janela_resultado.grab_set()
        
        # FRAME PRINCIPAL
        frame_main = ttk.Frame(janela_resultado)
        frame_main.pack(fill='both', expand=True, padx=20, pady=20)
        
        # TÍTULO
        titulo = tk.Label(frame_main, 
                         text="🎉 IMPORTAÇÃO NFe CONCLUÍDA COM SUCESSO!", 
                         font=('Arial', 14, 'bold'),
                         fg='darkgreen')
        titulo.pack(pady=10)
        
        # RESUMO NFE
        frame_nfe = ttk.LabelFrame(frame_main, text="📄 NFe Processada", padding=10)
        frame_nfe.pack(fill='x', pady=5)
        
        dados_nfe = self.dados_nfe
        resumo_nfe = f"""📄 Número: {dados_nfe.get('numero_nf', '')}  |  📅 Data: {dados_nfe.get('data_emissao', '')}
🏢 Fornecedor: {dados_nfe.get('razao_social_emitente', '')}
💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}  |  📦 Produtos: {len(dados_nfe.get('produtos', []))}"""
        
        tk.Label(frame_nfe, text=resumo_nfe, justify='left', font=('Arial', 9)).pack(anchor='w')
        
        # DADOS IMPORTADOS
        frame_resultados = ttk.LabelFrame(frame_main, text="✅ Dados Importados", padding=10)
        frame_resultados.pack(fill='x', pady=5)
        
        for resultado in resultados:
            tk.Label(frame_resultados, text=resultado, fg='blue', 
                    font=('Arial', 10, 'bold')).pack(anchor='w', pady=2)
        
        # CONFIGURAÇÕES APLICADAS
        frame_config = ttk.LabelFrame(frame_main, text="⚙️ Configurações Aplicadas", padding=10)
        frame_config.pack(fill='x', pady=5)
        
        if opcoes['importar_financeiro']:
            config_text = f"""💰 Financeiro:
   📅 Data Ref: {opcoes['data_rel']} (padrão sistema) | Vencto: {opcoes['dt_vencto']} (NFe)
   🏗️ Etapa: {opcoes['etapa_obra']} | 📋 Ref: {opcoes['referencia'][:50]}"""
            tk.Label(frame_config, text=config_text, justify='left', font=('Arial', 9)).pack(anchor='w', pady=2)
        
        if opcoes['importar_materiais']:
            mat_text = f"""📦 Materiais:
   🏠 Ambiente: {opcoes['ambiente_padrao']} | ⚙️ Status: {opcoes['status_instalacao']}
   🛡️ Garantia: {opcoes['garantia_meses']} meses | 🏢 Marca: {opcoes['marca_fabricante']}"""
            tk.Label(frame_config, text=mat_text, justify='left', font=('Arial', 9)).pack(anchor='w', pady=2)
        
        # PRÓXIMOS PASSOS
        frame_passos = ttk.LabelFrame(frame_main, text="🚀 Próximos Passos", padding=10)
        frame_passos.pack(fill='x', pady=5)
        
        passos_text = """1. ✅ Dados foram adicionados ao sistema
2. 📊 Use 'Enviar Dados' no sistema principal para salvar na planilha
3. 📦 Confira materiais em 'Consultar Materiais' 
4. 🔧 Atualize status conforme instalação avança
5. 📄 Gere 'Manual do Proprietário' ao final da obra"""
        
        tk.Label(frame_passos, text=passos_text, justify='left', font=('Arial', 9)).pack(anchor='w')
        
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
                  text="✅ Concluído", 
                  command=janela_resultado.destroy).pack(side='right', padx=5)
    
    def processar_no_sistema_final(self, janela_resultado):
        """Chama enviar_dados() do sistema principal"""
        try:
            janela_resultado.destroy()
            
            # VERIFICAR DADOS
            if not hasattr(self.sistema, 'dados_para_incluir') or not self.sistema.dados_para_incluir:
                tk.messagebox.showwarning("Aviso", "Não há dados financeiros para processar!")
                return
            
            print(f"📊 Processando {len(self.sistema.dados_para_incluir)} lançamentos no sistema...")
            
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
    
    def fechar_janela(self):
        """Fecha janela e restaura anterior"""
        try:
            self.janela.destroy()
            if self.janela_anterior and self.janela_anterior.winfo_exists():
                self.janela_anterior.deiconify()  # Mostrar janela anterior
        except:
            pass


# FUNÇÃO PRINCIPAL PARA APLICAR A CORREÇÃO
def aplicar_correcao_interface_nfe(sistema_principal):
    """
    Aplica a correção da interface NFe ao sistema existente
    """
    try:
        print("🔧 Aplicando correção da interface NFe...")
        
        sucesso = corrigir_interface_nfe_manualmente(sistema_principal)
        
        if sucesso:
            print("✅ Correção da interface NFe aplicada com sucesso!")
            print("📌 A partir de agora, ao clicar 'Importar para Sistema' você verá:")
            print("   - ✅ Interface aprimorada com scroll")
            print("   - ✅ Datas ajustadas (5/20 + vencimento NFe)")
            print("   - ✅ Campo referência editável")
            print("   - ✅ Etapas carregadas dos parâmetros")
            print("   - ✅ Configurações materiais detalhadas")
            print("   - ✅ Preview completo dos dados")
            return True
        else:
            print("❌ Erro ao aplicar correção da interface!")
            return False
        
    except Exception as e:
        print(f"❌ Erro geral na correção: {e}")
        return False


# EXEMPLO DE USO
"""
PARA APLICAR A CORREÇÃO DA INTERFACE:

# Adicione no final do __init__ do SistemaEntradaDados:
try:
    from src.nfe.correcao_interface_nfe import aplicar_correcao_interface_nfe
    aplicar_correcao_interface_nfe(self)
    print("✅ Interface NFe corrigida!")
except Exception as e:
    print(f"⚠️ Correção interface não aplicada: {e}")

RESULTADO:
- ✅ Interface antiga substituída pela aprimorada
- ✅ Todos os ajustes visuais aplicados
- ✅ Datas, referência e configurações corretas
- ✅ Preview completo antes da importação
"""