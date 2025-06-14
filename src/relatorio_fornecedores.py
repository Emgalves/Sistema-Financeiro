import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime, timedelta
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
import matplotlib.cm as cm
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from collections import defaultdict
import numpy as np

# Adicionar diretório raiz ao path
def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Importar configurações
try:
    from src.config.config import (
        ARQUIVO_CLIENTES,
        PASTA_CLIENTES,
        BASE_PATH,
        ARQUIVO_FORNECEDORES
    )
    print("Configurações importadas com sucesso")
except ImportError as e:
    print(f"Erro ao importar configurações: {str(e)}")
    # Definir valores padrão em caso de falha
    BASE_PATH = Path(".")
    ARQUIVO_CLIENTES = BASE_PATH / "dados" / "clientes.xlsx"
    PASTA_CLIENTES = BASE_PATH / "dados" / "clientes"
    ARQUIVO_FORNECEDORES = BASE_PATH / "dados" / "fornecedores.xlsx"

# Importar o utils.py
from src.config.utils import atualizar_combobox_clientes, cliente_esta_ativo, obter_info_cliente

try:
    from src.config.window_config import configurar_janela
    print("window_config importado com sucesso")
except ImportError as e:
    print(f"Erro ao importar window_config: {str(e)}")
    def configurar_janela(janela, titulo="Janela", largura=800, altura=950):
        janela.title(titulo)
        janela.geometry(f"{largura}x{altura}")
        janela.resizable(True, True)
        
        # Centralizar na tela
        janela.update_idletasks()
        width = janela.winfo_width()
        height = janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (width // 2)
        y = (janela.winfo_screenheight() // 2) - (height // 2)
        janela.geometry(f'{width}x{height}+{x}+{y}')

# Função para formatação de moeda
def formatar_moeda_br(valor):
    """Formata um valor numérico como moeda brasileira"""
    try:
        valor_float = float(valor)
        return f"R$ {valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except (ValueError, TypeError):
        return f"R$ 0,00"

class RelatorioFornecedores:
    """Classe para geração de relatórios de principais fornecedores"""
    
    def __init__(self, parent=None):
        """Inicializa a interface do relatório de fornecedores"""
        self.parent = parent
        
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        configurar_janela(self.root, "Relatório de Fornecedores", 900, 1000)
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.data_referencia = datetime.now()
        self.periodo_inicio = None
        self.periodo_fim = None
        self.todos_clientes = False
        self.fornecedor_especifico = None  # NOVO: para o relatório inverso
        self.modo_relatorio = "fornecedores"  # "fornecedores" ou "por_fornecedor"
        self.dados_fornecedores = {}
        self.dados_por_fornecedor = {}  # NOVO: dados agrupados por cliente para um fornecedor
        self.top_fornecedores = []
        self.dados_carregados = False
        
        # Configurar interface
        self.setup_gui()
        
    def setup_gui(self):
        """Configuração da interface gráfica principal"""
        # Frame principal
        self.frame_principal = ttk.Frame(self.root, padding=10)
        self.frame_principal.pack(fill='both', expand=True)

        # Dividir o frame principal em três partes usando grid
        self.frame_principal.columnconfigure(0, weight=1)
        self.frame_principal.rowconfigure(0, weight=0)  # Seleção de modo
        self.frame_principal.rowconfigure(1, weight=0)  # Seleção
        self.frame_principal.rowconfigure(2, weight=1)  # Resultados
        self.frame_principal.rowconfigure(3, weight=0)  # Botões
        
        # NOVO: Frame para seleção do modo de relatório
        self.frame_modo = ttk.LabelFrame(self.frame_principal, text="Tipo de Relatório")
        self.frame_modo.grid(row=0, column=0, sticky='ew', pady=5)
        
        frame_modo_int = ttk.Frame(self.frame_modo)
        frame_modo_int.pack(fill='x', padx=10, pady=10)
        
        self.var_modo = tk.StringVar(value="fornecedores")
        
        ttk.Radiobutton(
            frame_modo_int,
            text="Principais Fornecedores (por cliente)",
            variable=self.var_modo,
            value="fornecedores",
            command=self.alterar_modo_relatorio
        ).pack(side='left', padx=20)
        
        ttk.Radiobutton(
            frame_modo_int,
            text="Clientes de um Fornecedor Específico",
            variable=self.var_modo,
            value="por_fornecedor",
            command=self.alterar_modo_relatorio
        ).pack(side='left', padx=20)
            
        # Frame para seleção
        self.frame_selecao = ttk.LabelFrame(self.frame_principal, text="Seleção de Cliente e Período")
        self.frame_selecao.grid(row=1, column=0, sticky='ew', pady=10)
        
        # Container para cliente
        self.frame_cliente = ttk.Frame(self.frame_selecao)
        self.frame_cliente.pack(fill='x', padx=10, pady=10)
        
        self.lbl_cliente = ttk.Label(self.frame_cliente, text="Selecione o Cliente:", font=('Arial', 11))
        self.lbl_cliente.pack(side='left', pady=5)
        
        self.cliente_combobox = ttk.Combobox(self.frame_cliente, width=40, font=('Arial', 11))
        self.cliente_combobox.pack(side='left', padx=5)
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.selecionar_cliente)
        
        # Checkbox para todos os clientes
        self.var_todos_clientes = tk.BooleanVar(value=False)
        self.cb_todos_clientes = ttk.Checkbutton(
            self.frame_cliente, 
            text="Analisar todos os clientes",
            variable=self.var_todos_clientes,
            command=self.alternar_todos_clientes
        )
        self.cb_todos_clientes.pack(side='left', padx=20)
        
        # NOVO: Container para fornecedor (inicialmente oculto)
        self.frame_fornecedor = ttk.Frame(self.frame_selecao)
        
        self.lbl_fornecedor = ttk.Label(self.frame_fornecedor, text="Selecione o Fornecedor:", font=('Arial', 11))
        self.lbl_fornecedor.pack(side='left', pady=5)
        
        self.fornecedor_especifico_combobox = ttk.Combobox(self.frame_fornecedor, width=40, font=('Arial', 11))
        self.fornecedor_especifico_combobox.pack(side='left', padx=5)
        self.fornecedor_especifico_combobox.bind('<<ComboboxSelected>>', self.selecionar_fornecedor_especifico)
        
        ttk.Button(
            self.frame_fornecedor,
            text="Buscar Fornecedores",
            command=self.carregar_fornecedores
        ).pack(side='left', padx=20)
        
        # Container para período
        frame_periodo = ttk.Frame(self.frame_selecao)
        frame_periodo.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_periodo, text="Período de Análise:", font=('Arial', 11)).pack(side='left', pady=5)
        
        # Combobox para seleção de período pré-definido
        self.periodo_combobox = ttk.Combobox(
            frame_periodo, 
            width=15, 
            font=('Arial', 11),
            values=[
                "Últimos 30 dias",
                "Últimos 90 dias",
                "Últimos 180 dias",
                "Último ano",
                "Todo o período",
                "Personalizado"
            ]
        )
        self.periodo_combobox.pack(side='left', padx=5)
        self.periodo_combobox.current(2)  # Padrão: Últimos 180 dias
        self.periodo_combobox.bind('<<ComboboxSelected>>', self.alterar_periodo)
        
        # Frame para datas personalizadas (inicialmente oculto)
        self.frame_datas_personalizadas = ttk.Frame(frame_periodo)
        
        ttk.Label(self.frame_datas_personalizadas, text="De:", font=('Arial', 11)).pack(side='left', padx=(20, 5))
        
        # Usar DateEntry se disponível
        try:
            from tkcalendar import DateEntry
            self.data_inicio = DateEntry(
                self.frame_datas_personalizadas, 
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR',
                font=('Arial', 11)
            )
            self.data_inicio.pack(side='left', padx=5)
            
            ttk.Label(self.frame_datas_personalizadas, text="Até:", font=('Arial', 11)).pack(side='left', padx=(20, 5))
            
            self.data_fim = DateEntry(
                self.frame_datas_personalizadas, 
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR',
                font=('Arial', 11)
            )
            self.data_fim.pack(side='left', padx=5)
            
        except ImportError:
            # Fallback para Entry
            self.data_inicio_var = tk.StringVar(value=(datetime.now() - timedelta(days=180)).strftime('%d/%m/%Y'))
            ttk.Entry(
                self.frame_datas_personalizadas,
                textvariable=self.data_inicio_var,
                width=12,
                font=('Arial', 11)
            ).pack(side='left', padx=5)
            
            ttk.Label(self.frame_datas_personalizadas, text="Até:", font=('Arial', 11)).pack(side='left', padx=(20, 5))
            
            self.data_fim_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
            ttk.Entry(
                self.frame_datas_personalizadas,
                textvariable=self.data_fim_var,
                width=12,
                font=('Arial', 11)
            ).pack(side='left', padx=5)
        
        # Botão de gerar relatório
        ttk.Button(
            frame_periodo,
            text="Gerar Relatório",
            command=self.gerar_relatorio,
            style='Big.TButton'
        ).pack(side='right', padx=20)
        
        # Configurar opções de análise
        self.frame_opcoes = ttk.LabelFrame(self.frame_principal, text="Opções de Análise")
        self.frame_opcoes.grid(row=2, column=0, sticky='ew', pady=10)
        
        frame_opcoes_int = ttk.Frame(self.frame_opcoes)
        frame_opcoes_int.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_opcoes_int, text="Quantidade a exibir:", font=('Arial', 11)).pack(side='left', pady=5)
        
        self.top_n_var = tk.StringVar(value="10")
        self.top_n_combobox = ttk.Combobox(
            frame_opcoes_int, 
            width=5, 
            font=('Arial', 11),
            textvariable=self.top_n_var,
            values=["5", "10", "15", "20", "25", "30", "50", "100"]
        )
        self.top_n_combobox.pack(side='left', padx=5)
        
        # Checkbox para mostrar fornecedores agrupados por tipo de despesa
        self.var_agrupar_tipo = tk.BooleanVar(value=True)
        self.cb_agrupar_tipo = ttk.Checkbutton(
            frame_opcoes_int, 
            text="Agrupar por tipo de despesa",
            variable=self.var_agrupar_tipo
        )
        self.cb_agrupar_tipo.pack(side='left', padx=20)
        
        # Frame para resultados - com notebook para separar visões
        self.frame_resultados = ttk.LabelFrame(self.frame_principal, text="Resultados")
        self.frame_resultados.grid(row=2, column=0, sticky='nsew', pady=10)
        
        # Notebook (abas)
        self.notebook = ttk.Notebook(self.frame_resultados)
        self.notebook.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Abas
        self.aba_resumo = ttk.Frame(self.notebook)
        self.aba_detalhes = ttk.Frame(self.notebook)
        self.aba_grafico = ttk.Frame(self.notebook)
        
        self.notebook.add(self.aba_resumo, text='Resumo')
        self.notebook.add(self.aba_detalhes, text='Detalhes')
        self.notebook.add(self.aba_grafico, text='Gráfico')
        
        # Configurar cada aba
        self.setup_aba_resumo()
        self.setup_aba_detalhes()
        self.setup_aba_grafico()
        
        # Botões na parte inferior
        frame_botoes = ttk.Frame(self.frame_principal)
        frame_botoes.grid(row=3, column=0, sticky='ew', pady=10)

        btn_excel = ttk.Button(
            frame_botoes,
            text="Exportar para Excel",
            command=self.exportar_excel
        )
        btn_excel.grid(row=0, column=0, padx=5, sticky='w')

        btn_pdf = ttk.Button(
            frame_botoes,
            text="Exportar para PDF",
            command=self.exportar_pdf
        )
        btn_pdf.grid(row=0, column=1, padx=5, sticky='w')

        btn_voltar = ttk.Button(
            frame_botoes,
            text="Voltar ao Menu",
            command=self.voltar_menu
        )
        btn_voltar.grid(row=0, column=2, padx=5, sticky='e')

        # Configure as colunas para posicionar corretamente
        frame_botoes.columnconfigure(0, weight=0)
        frame_botoes.columnconfigure(1, weight=0)
        frame_botoes.columnconfigure(2, weight=1)
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 11, 'bold'), padding=(10, 5))
        
        # Carregar lista de clientes
        self.atualizar_lista_clientes()
        
        # Configurar período inicial
        self.alterar_periodo()
        
        # Configurar modo inicial
        self.alterar_modo_relatorio()

    # NOVO: Método para alternar modo de relatório
    def alterar_modo_relatorio(self):
        """Alterna entre os modos de relatório"""
        self.modo_relatorio = self.var_modo.get()
        
        if self.modo_relatorio == "fornecedores":
            # Modo original: principais fornecedores
            self.frame_cliente.pack(fill='x', padx=10, pady=10)
            self.frame_fornecedor.pack_forget()
            
            # Atualizar labels das abas
            self.notebook.tab(0, text="Resumo")
            self.lbl_cliente.config(text="Selecione o Cliente:")
            self.frame_opcoes.config(text="Opções de Análise")
            
        else:
            # Novo modo: clientes de um fornecedor
            self.frame_cliente.pack_forget()
            self.frame_fornecedor.pack(fill='x', padx=10, pady=10)
            
            # Atualizar labels das abas
            self.notebook.tab(0, text="Clientes do Fornecedor")
            self.frame_opcoes.config(text="Opções de Análise")
            
            # Carregar fornecedores se ainda não foi feito
            if not self.fornecedor_especifico_combobox['values']:
                self.carregar_fornecedores()

    # NOVO: Método para carregar lista de fornecedores
    def carregar_fornecedores(self):
        """Carrega a lista de fornecedores do arquivo de fornecedores"""
        try:
            if not os.path.exists(ARQUIVO_FORNECEDORES):
                messagebox.showwarning("Aviso", "Arquivo de fornecedores não encontrado!")
                return
            
            df = pd.read_excel(ARQUIVO_FORNECEDORES, sheet_name='Fornecedores')
            
            # Ordenar por nome
            fornecedores = sorted(df['NOME'].dropna().unique())
            
            # Atualizar combobox
            self.fornecedor_especifico_combobox['values'] = fornecedores
            
            messagebox.showinfo("Sucesso", f"{len(fornecedores)} fornecedores carregados!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar fornecedores: {str(e)}")

    # NOVO: Método para selecionar fornecedor específico
    def selecionar_fornecedor_especifico(self, event=None):
        """Seleciona o fornecedor para análise"""
        self.fornecedor_especifico = self.fornecedor_especifico_combobox.get()

    def setup_aba_resumo(self):
        """Configura a aba de resumo do relatório"""
        # Frame para informações do cliente/período
        frame_info = ttk.Frame(self.aba_resumo, padding=5)
        frame_info.pack(fill='x', pady=5)
        
        self.lbl_cliente_resumo = ttk.Label(
            frame_info, 
            text="Cliente: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_cliente_resumo.pack(side='left', padx=10)
        
        self.lbl_periodo_resumo = ttk.Label(
            frame_info, 
            text="Período: Não definido", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_periodo_resumo.pack(side='left', padx=10)
        
        # Frame para tabela de resumo
        frame_tabela = ttk.Frame(self.aba_resumo, padding=5)
        frame_tabela.pack(fill='both', expand=True, pady=5)
        
        # Criar Treeview para listar (fornecedores ou clientes)
        self.colunas_resumo = ('posicao', 'nome', 'total_gasto', 'percentual', 'qtd_lancamentos', 'tipos_despesa')
        self.tree_resumo = ttk.Treeview(frame_tabela, columns=self.colunas_resumo, show='headings', height=20)
        
        # Configurar colunas (serão atualizadas conforme o modo)
        self.atualizar_colunas_resumo()
        
        self.tree_resumo.bind("<Double-1>", self.mostrar_detalhes_por_clique)
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tree_resumo.yview)
        scrollx = ttk.Scrollbar(frame_tabela, orient='horizontal', command=self.tree_resumo.xview)
        self.tree_resumo.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_resumo.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
        # Frame para totais
        frame_total = ttk.Frame(self.aba_resumo, padding=5)
        frame_total.pack(fill='x', pady=5)
        
        self.lbl_total_geral = ttk.Label(
            frame_total,
            text="Total Geral: R$ 0,00",
            font=('Arial', 12, 'bold')
        )
        self.lbl_total_geral.pack(side='left', padx=10)
        
        self.lbl_total_apresentado = ttk.Label(
            frame_total,
            text="Total Apresentado: R$ 0,00 (0%)",
            font=('Arial', 12)
        )
        self.lbl_total_apresentado.pack(side='left', padx=10)

    # NOVO: Método para atualizar colunas conforme o modo
    def atualizar_colunas_resumo(self):
        """Atualiza as colunas da tabela de resumo conforme o modo"""
        if self.modo_relatorio == "fornecedores":
            # Modo original
            self.tree_resumo.heading('posicao', text='#')
            self.tree_resumo.heading('nome', text='Fornecedor')
            self.tree_resumo.heading('total_gasto', text='Total Gasto')
            self.tree_resumo.heading('percentual', text='% do Total')
            self.tree_resumo.heading('qtd_lancamentos', text='Qtd. Lançamentos')
            self.tree_resumo.heading('tipos_despesa', text='Tipos de Despesa')
        else:
            # Novo modo: clientes do fornecedor
            self.tree_resumo.heading('posicao', text='#')
            self.tree_resumo.heading('nome', text='Cliente')
            self.tree_resumo.heading('total_gasto', text='Total Gasto')
            self.tree_resumo.heading('percentual', text='% do Total')
            self.tree_resumo.heading('qtd_lancamentos', text='Qtd. Lançamentos')
            self.tree_resumo.heading('tipos_despesa', text='Tipos de Despesa')
        
        # Definir larguras
        self.tree_resumo.column('posicao', width=50, anchor='center')
        self.tree_resumo.column('nome', width=250)
        self.tree_resumo.column('total_gasto', width=150, anchor='e')
        self.tree_resumo.column('percentual', width=100, anchor='center')
        self.tree_resumo.column('qtd_lancamentos', width=150, anchor='center')
        self.tree_resumo.column('tipos_despesa', width=150)

    def setup_aba_detalhes(self):
        """Configura a aba de detalhes do relatório"""
        # Frame para seleção
        frame_selecao = ttk.Frame(self.aba_detalhes, padding=5)
        frame_selecao.pack(fill='x', pady=5)
        
        self.lbl_selecao_detalhes = ttk.Label(frame_selecao, text="Selecione o Item:", font=('Arial', 11))
        self.lbl_selecao_detalhes.pack(side='left', pady=5)
        
        self.item_detalhes_combobox = ttk.Combobox(frame_selecao, width=40, font=('Arial', 11))
        self.item_detalhes_combobox.pack(side='left', padx=5)
        self.item_detalhes_combobox.bind('<<ComboboxSelected>>', self.carregar_detalhes_item)
        
        # Frame para informações do item
        frame_info_item = ttk.LabelFrame(self.aba_detalhes, text="Informações")
        frame_info_item.pack(fill='x', pady=5, padx=5)
        
        # Grid para informações
        frame_grid = ttk.Frame(frame_info_item, padding=10)
        frame_grid.pack(fill='x')
        
        # Primeira linha
        self.lbl_titulo_nome = ttk.Label(frame_grid, text="Nome:", font=('Arial', 10, 'bold'))
        self.lbl_titulo_nome.grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.lbl_nome_item = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_nome_item.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Total Gasto:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.lbl_total_item = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_total_item.grid(row=0, column=3, sticky='w', padx=5, pady=2)
        
        # Segunda linha
        ttk.Label(frame_grid, text="Quantidade de Lançamentos:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.lbl_qtd_lancamentos_item = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_qtd_lancamentos_item.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Média por Lançamento:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=2)
        self.lbl_media_lancamento_item = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_media_lancamento_item.grid(row=1, column=3, sticky='w', padx=5, pady=2)
        
        # Terceira linha
        ttk.Label(frame_grid, text="Tipos de Despesa:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.lbl_tipos_despesa_item = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_tipos_despesa_item.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        
        # Frame para tabela de lançamentos
        frame_lancamentos = ttk.LabelFrame(self.aba_detalhes, text="Lançamentos")
        frame_lancamentos.pack(fill='both', expand=True, pady=5, padx=5)
        
        # Tree para lançamentos
        colunas = ('data', 'cliente_ou_fornecedor', 'tipo_despesa', 'referencia', 'nf', 'dt_vencto', 'valor', 'observacao')
        self.tree_lancamentos = ttk.Treeview(frame_lancamentos, columns=colunas, show='headings', height=15)
        
        # Configurar colunas (serão atualizadas conforme o modo)
        self.atualizar_colunas_detalhes()
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_lancamentos, orient='vertical', command=self.tree_lancamentos.yview)
        scrollx = ttk.Scrollbar(frame_lancamentos, orient='horizontal', command=self.tree_lancamentos.xview)
        self.tree_lancamentos.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_lancamentos.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')

    # NOVO: Método para atualizar colunas da aba de detalhes
    def atualizar_colunas_detalhes(self):
        """Atualiza as colunas da tabela de detalhes conforme o modo"""
        if self.modo_relatorio == "fornecedores":
            # Modo original: detalhes de um fornecedor
            self.tree_lancamentos.heading('data', text='Data')
            self.tree_lancamentos.heading('cliente_ou_fornecedor', text='Cliente')
            self.tree_lancamentos.heading('tipo_despesa', text='Tipo')
            self.tree_lancamentos.heading('referencia', text='Referência')
            self.tree_lancamentos.heading('nf', text='NF')
            self.tree_lancamentos.heading('dt_vencto', text='Vencimento')
            self.tree_lancamentos.heading('valor', text='Valor')
            self.tree_lancamentos.heading('observacao', text='Observação')
        else:
            # Novo modo: detalhes de um cliente (compras de um fornecedor)
            self.tree_lancamentos.heading('data', text='Data')
            self.tree_lancamentos.heading('cliente_ou_fornecedor', text='Fornecedor')
            self.tree_lancamentos.heading('tipo_despesa', text='Tipo')
            self.tree_lancamentos.heading('referencia', text='Referência')
            self.tree_lancamentos.heading('nf', text='NF')
            self.tree_lancamentos.heading('dt_vencto', text='Vencimento')
            self.tree_lancamentos.heading('valor', text='Valor')
            self.tree_lancamentos.heading('observacao', text='Observação')
        
        # Ajustar larguras
        self.tree_lancamentos.column('data', width=80, anchor='center')
        self.tree_lancamentos.column('cliente_ou_fornecedor', width=200)
        self.tree_lancamentos.column('tipo_despesa', width=50, anchor='center')
        self.tree_lancamentos.column('referencia', width=250)
        self.tree_lancamentos.column('nf', width=100)
        self.tree_lancamentos.column('dt_vencto', width=80, anchor='center')
        self.tree_lancamentos.column('valor', width=100, anchor='e')
        self.tree_lancamentos.column('observacao', width=150)

    def setup_aba_grafico(self):
        """Configura a aba de gráficos"""
        # Frame para controles do gráfico
        frame_controles = ttk.Frame(self.aba_grafico, padding=5)
        frame_controles.pack(fill='x', pady=5)
        
        ttk.Label(frame_controles, text="Tipo de Gráfico:").pack(side='left', padx=5)
        self.combo_tipo_grafico = ttk.Combobox(frame_controles, values=[
            "Pizza - Total por Item",
            "Barras - Top Items",
            "Linhas - Evolução Mensal",
            "Barras Empilhadas - Por Tipo de Despesa"
        ], state='readonly', width=30)
        self.combo_tipo_grafico.pack(side='left', padx=5)
        self.combo_tipo_grafico.current(0)
        
        ttk.Button(frame_controles, text="Atualizar Gráfico", command=self.atualizar_grafico).pack(side='left', padx=20)
        
        # Frame para o gráfico
        self.frame_grafico = ttk.Frame(self.aba_grafico)
        self.frame_grafico.pack(fill='both', expand=True, pady=5)

    def mostrar_detalhes_por_clique(self, event):
        """Abre a aba de detalhes quando o usuário dá duplo clique em um item"""
        # Obter o item selecionado
        item = self.tree_resumo.identify('item', event.x, event.y)
        if not item:
            return
            
        # Obter a posição do item (valor da primeira coluna)
        posicao = self.tree_resumo.item(item, 'values')[0]
        
        # Selecionar o item correspondente no combobox
        if self.item_detalhes_combobox['values']:
            # Os valores do combobox começam com a posição (ex: "1. Nome do Item")
            for i, valor in enumerate(self.item_detalhes_combobox['values']):
                if valor.startswith(f"{posicao}. "):
                    self.item_detalhes_combobox.current(i)
                    self.carregar_detalhes_item()
                    # Mudar para a aba de detalhes
                    self.notebook.select(1)  # Índice 1 = aba de detalhes
                    break

    def alterar_periodo(self, event=None):
        """Atualiza o período de análise com base na seleção"""
        periodo_selecionado = self.periodo_combobox.get()
        
        # Ocultar frame de datas personalizadas por padrão
        self.frame_datas_personalizadas.pack_forget()
        
        # Data de hoje
        hoje = datetime.now()
        
        # Configurar período com base na seleção
        if periodo_selecionado == "Últimos 30 dias":
            self.periodo_inicio = hoje - timedelta(days=30)
            self.periodo_fim = hoje
        elif periodo_selecionado == "Últimos 90 dias":
            self.periodo_inicio = hoje - timedelta(days=90)
            self.periodo_fim = hoje
        elif periodo_selecionado == "Últimos 180 dias":
            self.periodo_inicio = hoje - timedelta(days=180)
            self.periodo_fim = hoje
        elif periodo_selecionado == "Último ano":
            self.periodo_inicio = hoje - timedelta(days=365)
            self.periodo_fim = hoje
        elif periodo_selecionado == "Todo o período":
            # Usar um período bem amplo
            self.periodo_inicio = datetime(2000, 1, 1)
            self.periodo_fim = hoje
        elif periodo_selecionado == "Personalizado":
            # Mostrar frame de datas personalizadas
            self.frame_datas_personalizadas.pack(side='left')
            
            # Configurar valores iniciais se estiver usando DateEntry
            if hasattr(self, 'data_inicio'):
                data_inicio = hoje - timedelta(days=180)
                self.data_inicio.set_date(data_inicio)
                self.data_fim.set_date(hoje)
            
            # Datas serão definidas quando o relatório for gerado
            return
        
        # Atualizar label do período (se já tiver sido gerado algum relatório)
        if hasattr(self, 'lbl_periodo_resumo'):
            self.lbl_periodo_resumo.config(
                text=f"Período: {self.periodo_inicio.strftime('%d/%m/%Y')} a {self.periodo_fim.strftime('%d/%m/%Y')}"
            )
    
    def alternar_todos_clientes(self):
        """Alterna entre análise de um cliente específico ou todos"""
        todos = self.var_todos_clientes.get()
        
        if todos:
            self.cliente_combobox.set("TODOS OS CLIENTES")
            self.cliente_combobox.config(state='disabled')
            self.cliente_atual = None
            self.todos_clientes = True
            
            # Atualizar label
            if hasattr(self, 'lbl_cliente_resumo'):
                self.lbl_cliente_resumo.config(text="Cliente: Todos os Clientes")
        else:
            self.cliente_combobox.config(state='normal')
            self.cliente_combobox.set("")
            self.cliente_atual = None
            self.todos_clientes = False
            
            # Atualizar label
            if hasattr(self, 'lbl_cliente_resumo'):
                self.lbl_cliente_resumo.config(text="Cliente: Nenhum selecionado")
    
    def atualizar_lista_clientes(self):
        """Atualiza a lista de clientes no combobox usando a função centralizada"""
        try:
            # Usar a função centralizada (apenas clientes ativos)
            self.info_clientes = atualizar_combobox_clientes(self.cliente_combobox, mostrar_inativos=False)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")

    def selecionar_cliente(self, event=None):
        """Atualiza o cliente selecionado"""
        self.cliente_atual = self.cliente_combobox.get()
        
        if self.cliente_atual:
            # Verificar se o cliente está ativo (extra proteção)
            if not cliente_esta_ativo(self.cliente_atual):
                messagebox.showwarning(
                    "Cliente Inativo", 
                    f"O cliente '{self.cliente_atual}' está inativo (contrato finalizado). " +
                    "Os dados serão mostrados somente para consulta."
                )
            
            # Obter informações do cliente
            info_cliente = obter_info_cliente(self.cliente_atual)
            
            # Atualizar label
            if hasattr(self, 'lbl_cliente_resumo'):
                texto_cliente = f"Cliente: {self.cliente_atual}"
                if info_cliente and not info_cliente['ativo']:
                    texto_cliente += " (INATIVO)"
                self.lbl_cliente_resumo.config(text=texto_cliente)
            
            # Definir o caminho do arquivo
            if info_cliente and 'arquivo' in info_cliente:
                self.arquivo_cliente = info_cliente['arquivo']
            else:
                self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
    
    def gerar_relatorio(self):
        """Gera o relatório com base nos dados selecionados"""
        # Validações específicas por modo
        if self.modo_relatorio == "fornecedores":
            if not self.cliente_atual and not self.todos_clientes:
                messagebox.showwarning("Aviso", "Selecione um cliente ou marque a opção 'Analisar todos os clientes'!")
                return
        else:
            if not self.fornecedor_especifico:
                messagebox.showwarning("Aviso", "Selecione um fornecedor para análise!")
                return
        
        # Se o período for personalizado, obter as datas selecionadas
        if self.periodo_combobox.get() == "Personalizado":
            try:
                if hasattr(self, 'data_inicio'):  # Usando DateEntry
                    self.periodo_inicio = datetime.strptime(self.data_inicio.get(), '%d/%m/%Y')
                    self.periodo_fim = datetime.strptime(self.data_fim.get(), '%d/%m/%Y')
                else:  # Usando Entry
                    self.periodo_inicio = datetime.strptime(self.data_inicio_var.get(), '%d/%m/%Y')
                    self.periodo_fim = datetime.strptime(self.data_fim_var.get(), '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida no período personalizado!")
                return
        
        # Atualizar label do período
        self.lbl_periodo_resumo.config(
            text=f"Período: {self.periodo_inicio.strftime('%d/%m/%Y')} a {self.periodo_fim.strftime('%d/%m/%Y')}"
        )
        
        # Carregar dados conforme o modo
        if self.modo_relatorio == "fornecedores":
            if not self.carregar_dados_fornecedores():
                return
        else:
            if not self.carregar_dados_por_fornecedor():
                return
        
        # Preencher resumo
        self.preencher_resumo()
        
        # Atualizar lista para a aba de detalhes
        self.atualizar_lista_detalhes()
        
        # Limpar detalhes
        self.limpar_detalhes()
        
        # Gerar gráfico inicial
        self.atualizar_grafico()
        
        # Marcar que os dados foram carregados
        self.dados_carregados = True
        
        # Selecionar aba de resumo
        self.notebook.select(0)

    # NOVO: Método para carregar dados por fornecedor específico
    def carregar_dados_por_fornecedor(self):
        """Carrega os dados de todos os clientes para um fornecedor específico"""
        try:
            # Dicionário para armazenar dados por cliente
            self.dados_por_fornecedor = defaultdict(lambda: {
                'total': 0.0,
                'lancamentos': [],
                'qtd_lancamentos': 0,
                'tipos_despesa': set(),
                'por_mes': defaultdict(float),
                'por_tipo': defaultdict(float)
            })
            
            # Variáveis para somatórios
            self.total_geral = 0.0
            
            # Processar todos os arquivos de clientes
            clientes_processados = []
            fornecedor_encontrado = False
            
            for arquivo in os.listdir(PASTA_CLIENTES):
                if arquivo.endswith('.xlsx'):
                    try:
                        caminho_arquivo = os.path.join(PASTA_CLIENTES, arquivo)
                        nome_cliente = os.path.splitext(arquivo)[0]
                        
                        # Processar arquivo do cliente para o fornecedor específico
                        if self.processar_arquivo_cliente_fornecedor(caminho_arquivo, nome_cliente):
                            clientes_processados.append(nome_cliente)
                            fornecedor_encontrado = True
                    except Exception as e:
                        print(f"Erro ao processar arquivo {arquivo}: {str(e)}")
            
            if not fornecedor_encontrado:
                messagebox.showwarning("Aviso", f"Fornecedor '{self.fornecedor_especifico}' não encontrado em nenhum cliente no período selecionado!")
                return False
            
            if self.total_geral == 0:
                messagebox.showinfo("Aviso", f"Nenhum lançamento encontrado para o fornecedor '{self.fornecedor_especifico}' no período selecionado.")
                return False
            
            # Ordenar clientes por valor total (decrescente)
            clientes_ordenados = sorted(
                self.dados_por_fornecedor.items(),
                key=lambda x: x[1]['total'],
                reverse=True
            )
            
            # Obter top N clientes
            try:
                top_n = int(self.top_n_var.get())
            except ValueError:
                top_n = 10
                
            self.top_fornecedores = clientes_ordenados[:top_n]  # Reutilizando a variável
            
            print(f"Clientes processados para {self.fornecedor_especifico}: {len(clientes_processados)}")
            return True
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            import traceback
            traceback.print_exc()
            return False
    
    def processar_arquivo_cliente(self, caminho_arquivo, nome_cliente):
        """Processa um arquivo de cliente (método original)"""
        try:
            # Carregar dados do Excel
            df = pd.read_excel(caminho_arquivo, sheet_name='Dados')
            
            # Verificar colunas necessárias
            colunas_necessarias = ['DATA_REL', 'TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR']
            if not all(coluna in df.columns for coluna in colunas_necessarias):
                print(f"Arquivo {nome_cliente} não contém todas as colunas necessárias.")
                return
            
            # Filtrar apenas lançamentos ativos
            if 'STATUS' in df.columns:
                # Filtrar apenas registros com STATUS = 'ATIVO'
                df_original_len = len(df)
                df = df[df['STATUS'].str.upper().str.strip() == 'ATIVO'].copy()
                df_filtrado_len = len(df)
                
                print(f"Cliente {nome_cliente}: {df_original_len} registros totais, {df_filtrado_len} ativos processados")
                
                # Se não há registros ativos, não processar
                if df.empty:
                    print(f"Nenhum lançamento ativo encontrado para {nome_cliente}")
                    return
            else:
                # Se não existe a coluna STATUS, processar todos (compatibilidade)
                print(f"Cliente {nome_cliente}: Coluna STATUS não encontrada, processando todos os registros")

            # Converter DATA_REL para datetime
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'])
            
            # Filtrar por período
            df_periodo = df[
                (df['DATA_REL'] >= self.periodo_inicio) & 
                (df['DATA_REL'] <= self.periodo_fim)
            ]
            
            # Verificar se há dados no período
            if df_periodo.empty:
                print(f"Nenhum lançamento encontrado para {nome_cliente} no período selecionado.")
                return
            
            # Processar cada lançamento
            for _, row in df_periodo.iterrows():
                try:
                    # Obter nome do fornecedor
                    fornecedor = row['NOME']
                    if not isinstance(fornecedor, str) or not fornecedor.strip():
                        continue
                    
                    fornecedor = fornecedor.strip().upper()
                    
                    # Obter valor do lançamento
                    valor = 0.0
                    if isinstance(row['VALOR'], (int, float)):
                        valor = float(row['VALOR'])
                    elif isinstance(row['VALOR'], str):
                        # Limpar string e converter para float
                        valor_str = row['VALOR'].replace('R

    # NOVO: Método para processar arquivo de cliente para um fornecedor específico
    def processar_arquivo_cliente_fornecedor(self, caminho_arquivo, nome_cliente):
        """Processa um arquivo de cliente buscando lançamentos de um fornecedor específico"""
        try:
            # Carregar dados do Excel
            df = pd.read_excel(caminho_arquivo, sheet_name='Dados')
            
            # Verificar colunas necessárias
            colunas_necessarias = ['DATA_REL', 'TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR']
            if not all(coluna in df.columns for coluna in colunas_necessarias):
                print(f"Arquivo {nome_cliente} não contém todas as colunas necessárias.")
                return False
            
            # Filtrar apenas lançamentos ativos
            if 'STATUS' in df.columns:
                df = df[df['STATUS'].str.upper().str.strip() == 'ATIVO'].copy()
                if df.empty:
                    return False
            
            # Converter DATA_REL para datetime
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'])
            
            # Filtrar por período
            df_periodo = df[
                (df['DATA_REL'] >= self.periodo_inicio) & 
                (df['DATA_REL'] <= self.periodo_fim)
            ]
            
            if df_periodo.empty:
                return False
            
            # Filtrar pelo fornecedor específico (busca flexível)
            fornecedor_mask = df_periodo['NOME'].str.upper().str.contains(
                self.fornecedor_especifico.upper(), 
                na=False, 
                regex=False
            )
            df_fornecedor = df_periodo[fornecedor_mask].copy()
            
            if df_fornecedor.empty:
                return False
            
            encontrou_lancamentos = False
            
            # Processar cada lançamento do fornecedor
            for _, row in df_fornecedor.iterrows():
                try:
                    # Obter valor do lançamento
                    valor = 0.0
                    if isinstance(row['VALOR'], (int, float)):
                        valor = float(row['VALOR'])
                    elif isinstance(row['VALOR'], str):
                        # Limpar string e converter para float
                        valor_str = row['VALOR'].replace('R, '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor = float(valor_str)
                        except ValueError:
                            valor = 0.0
                    
                    # Ignorar lançamentos com valor zero
                    if valor <= 0:
                        continue
                    
                    # Obter tipo de despesa
                    tipo_despesa = int(row['TP_DESP']) if pd.notnull(row['TP_DESP']) else 0
                    
                    # Obter referência e incluir NF se disponível
                    referencia = str(row['REFERÊNCIA']) if pd.notnull(row['REFERÊNCIA']) else ""
                    
                    # Verificar se existe coluna 'NF' e adicionar à referência se disponível
                    nf = ""
                    if 'NF' in df.columns and pd.notnull(row['NF']) and str(row['NF']).strip():
                        nf = str(row['NF']).strip()
                        if nf and nf.lower() != 'nan':
                            referencia = f"{referencia} (NF: {nf})"
                    
                    # Obter data
                    data = row['DATA_REL']
                    
                    # Obter data de vencimento se disponível
                    dt_vencto = None
                    if 'DT_VENCTO' in df.columns and pd.notnull(row['DT_VENCTO']):
                        try:
                            dt_vencto = pd.to_datetime(row['DT_VENCTO'])
                        except:
                            dt_vencto = None
                    
                    # Obter observação se disponível
                    observacao = ""
                    if 'OBSERVACAO' in df.columns and pd.notnull(row['OBSERVACAO']):
                        observacao = str(row['OBSERVAÇÃO'])
                    elif 'OBSERVAÇÃO' in df.columns and pd.notnull(row['OBSERVAÇÃO']):
                        observacao = str(row['OBSERVAÇÃO'])
                    
                    # Criar identificador do mês para análise mensal
                    mes_ano = f"{data.year}-{data.month:02d}"
                    
                    # Atualizar dados do cliente
                    self.dados_por_fornecedor[nome_cliente]['total'] += valor
                    self.dados_por_fornecedor[nome_cliente]['qtd_lancamentos'] += 1
                    self.dados_por_fornecedor[nome_cliente]['tipos_despesa'].add(tipo_despesa)
                    self.dados_por_fornecedor[nome_cliente]['por_mes'][mes_ano] += valor
                    self.dados_por_fornecedor[nome_cliente]['por_tipo'][tipo_despesa] += valor
                    
                    # Adicionar lançamento à lista de lançamentos do cliente
                    self.dados_por_fornecedor[nome_cliente]['lancamentos'].append({
                        'data': data,
                        'fornecedor': row['NOME'],  # Nome real do fornecedor
                        'tipo_despesa': tipo_despesa,
                        'referencia': referencia,
                        'nf': nf,
                        'dt_vencto': dt_vencto,
                        'valor': valor,
                        'observacao': observacao
                    })
                    
                    # Atualizar total geral
                    self.total_geral += valor
                    encontrou_lancamentos = True
                    
                except Exception as e:
                    print(f"Erro ao processar lançamento: {str(e)}")
                    continue
            
            if encontrou_lancamentos:
                print(f"Cliente {nome_cliente} processado: {len(df_fornecedor)} lançamentos do fornecedor {self.fornecedor_especifico}.")
            
            return encontrou_lancamentos
            
        except Exception as e:
            print(f"Erro ao processar arquivo {nome_cliente}: {str(e)}")
            return False
    
    def carregar_dados_fornecedores(self):
        """Carrega os dados para o relatório (modo original)"""
        try:
            # Dicionário para armazenar dados por fornecedor
            self.dados_fornecedores = defaultdict(lambda: {
                'total': 0.0,
                'lancamentos': [],
                'qtd_lancamentos': 0,
                'tipos_despesa': set(),
                'clientes': set(),
                'por_mes': defaultdict(float),
                'por_tipo': defaultdict(float)
            })
            
            # Variáveis para somatórios
            self.total_geral = 0.0
            
            if self.todos_clientes:
                # Processar todos os arquivos de clientes
                clientes_processados = []
                
                for arquivo in os.listdir(PASTA_CLIENTES):
                    if arquivo.endswith('.xlsx'):
                        try:
                            caminho_arquivo = os.path.join(PASTA_CLIENTES, arquivo)
                            nome_cliente = os.path.splitext(arquivo)[0]
                            
                            # Processar arquivo do cliente
                            self.processar_arquivo_cliente(caminho_arquivo, nome_cliente)
                            clientes_processados.append(nome_cliente)
                        except Exception as e:
                            print(f"Erro ao processar arquivo {arquivo}: {str(e)}")
                
                if not clientes_processados:
                    messagebox.showwarning("Aviso", "Nenhum arquivo de cliente encontrado!")
                    return False
                    
                print(f"Clientes processados: {len(clientes_processados)}")
                
            else:
                # Processar apenas o cliente selecionado
                if not os.path.exists(self.arquivo_cliente):
                    messagebox.showerror("Erro", f"Arquivo do cliente '{self.cliente_atual}' não encontrado!")
                    return False
                
                # Processar arquivo do cliente
                self.processar_arquivo_cliente(self.arquivo_cliente, self.cliente_atual)
            
            # Verificar se encontrou algum lançamento
            if self.total_geral == 0:
                messagebox.showinfo("Aviso", "Nenhum lançamento encontrado no período selecionado.")
                return False
            
            # Ordenar fornecedores por valor total (decrescente)
            fornecedores_ordenados = sorted(
                self.dados_fornecedores.items(),
                key=lambda x: x[1]['total'],
                reverse=True
            )
            
            # Obter top N fornecedores
            try:
                top_n = int(self.top_n_var.get())
            except ValueError:
                top_n = 10
                
            self.top_fornecedores = fornecedores_ordenados[:top_n]
            
            return True
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            import traceback
            traceback.print_exc()
            return False, '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor = float(valor_str)
                        except ValueError:
                            valor = 0.0
                    
                    # Ignorar lançamentos com valor zero
                    if valor <= 0:
                

    # NOVO: Método para processar arquivo de cliente para um fornecedor específico
    def processar_arquivo_cliente_fornecedor(self, caminho_arquivo, nome_cliente):
        """Processa um arquivo de cliente buscando lançamentos de um fornecedor específico"""
        try:
            # Carregar dados do Excel
            df = pd.read_excel(caminho_arquivo, sheet_name='Dados')
            
            # Verificar colunas necessárias
            colunas_necessarias = ['DATA_REL', 'TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR']
            if not all(coluna in df.columns for coluna in colunas_necessarias):
                print(f"Arquivo {nome_cliente} não contém todas as colunas necessárias.")
                return False
            
            # Filtrar apenas lançamentos ativos
            if 'STATUS' in df.columns:
                df = df[df['STATUS'].str.upper().str.strip() == 'ATIVO'].copy()
                if df.empty:
                    return False
            
            # Converter DATA_REL para datetime
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'])
            
            # Filtrar por período
            df_periodo = df[
                (df['DATA_REL'] >= self.periodo_inicio) & 
                (df['DATA_REL'] <= self.periodo_fim)
            ]
            
            if df_periodo.empty:
                return False
            
            # Filtrar pelo fornecedor específico (busca flexível)
            fornecedor_mask = df_periodo['NOME'].str.upper().str.contains(
                self.fornecedor_especifico.upper(), 
                na=False, 
                regex=False
            )
            df_fornecedor = df_periodo[fornecedor_mask].copy()
            
            if df_fornecedor.empty:
                return False
            
            encontrou_lancamentos = False
            
            # Processar cada lançamento do fornecedor
            for _, row in df_fornecedor.iterrows():
                try:
                    # Obter valor do lançamento
                    valor = 0.0
                    if isinstance(row['VALOR'], (int, float)):
                        valor = float(row['VALOR'])
                    elif isinstance(row['VALOR'], str):
                        # Limpar string e converter para float
                        valor_str = row['VALOR'].replace('R, '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor = float(valor_str)
                        except ValueError:
                            valor = 0.0
                    
                    # Ignorar lançamentos com valor zero
                    if valor <= 0:
                        continue
                    
                    # Obter tipo de despesa
                    tipo_despesa = int(row['TP_DESP']) if pd.notnull(row['TP_DESP']) else 0
                    
                    # Obter referência e incluir NF se disponível
                    referencia = str(row['REFERÊNCIA']) if pd.notnull(row['REFERÊNCIA']) else ""
                    
                    # Verificar se existe coluna 'NF' e adicionar à referência se disponível
                    nf = ""
                    if 'NF' in df.columns and pd.notnull(row['NF']) and str(row['NF']).strip():
                        nf = str(row['NF']).strip()
                        if nf and nf.lower() != 'nan':
                            referencia = f"{referencia} (NF: {nf})"
                    
                    # Obter data
                    data = row['DATA_REL']
                    
                    # Obter data de vencimento se disponível
                    dt_vencto = None
                    if 'DT_VENCTO' in df.columns and pd.notnull(row['DT_VENCTO']):
                        try:
                            dt_vencto = pd.to_datetime(row['DT_VENCTO'])
                        except:
                            dt_vencto = None
                    
                    # Obter observação se disponível
                    observacao = ""
                    if 'OBSERVACAO' in df.columns and pd.notnull(row['OBSERVACAO']):
                        observacao = str(row['OBSERVAÇÃO'])
                    elif 'OBSERVAÇÃO' in df.columns and pd.notnull(row['OBSERVAÇÃO']):
                        observacao = str(row['OBSERVAÇÃO'])
                    
                    # Criar identificador do mês para análise mensal
                    mes_ano = f"{data.year}-{data.month:02d}"
                    
                    # Atualizar dados do cliente
                    self.dados_por_fornecedor[nome_cliente]['total'] += valor
                    self.dados_por_fornecedor[nome_cliente]['qtd_lancamentos'] += 1
                    self.dados_por_fornecedor[nome_cliente]['tipos_despesa'].add(tipo_despesa)
                    self.dados_por_fornecedor[nome_cliente]['por_mes'][mes_ano] += valor
                    self.dados_por_fornecedor[nome_cliente]['por_tipo'][tipo_despesa] += valor
                    
                    # Adicionar lançamento à lista de lançamentos do cliente
                    self.dados_por_fornecedor[nome_cliente]['lancamentos'].append({
                        'data': data,
                        'fornecedor': row['NOME'],  # Nome real do fornecedor
                        'tipo_despesa': tipo_despesa,
                        'referencia': referencia,
                        'nf': nf,
                        'dt_vencto': dt_vencto,
                        'valor': valor,
                        'observacao': observacao
                    })
                    
                    # Atualizar total geral
                    self.total_geral += valor
                    encontrou_lancamentos = True
                    
                except Exception as e:
                    print(f"Erro ao processar lançamento: {str(e)}")
                    continue
            
            if encontrou_lancamentos:
                print(f"Cliente {nome_cliente} processado: {len(df_fornecedor)} lançamentos do fornecedor {self.fornecedor_especifico}.")
            
            return encontrou_lancamentos
            
        except Exception as e:
            print(f"Erro ao processar arquivo {nome_cliente}: {str(e)}")
            return False
    
    def carregar_dados_fornecedores(self):
        """Carrega os dados para o relatório (modo original)"""
        try:
            # Dicionário para armazenar dados por fornecedor
            self.dados_fornecedores = defaultdict(lambda: {
                'total': 0.0,
                'lancamentos': [],
                'qtd_lancamentos': 0,
                'tipos_despesa': set(),
                'clientes': set(),
                'por_mes': defaultdict(float),
                'por_tipo': defaultdict(float)
            })
            
            # Variáveis para somatórios
            self.total_geral = 0.0
            
            if self.todos_clientes:
                # Processar todos os arquivos de clientes
                clientes_processados = []
                
                for arquivo in os.listdir(PASTA_CLIENTES):
                    if arquivo.endswith('.xlsx'):
                        try:
                            caminho_arquivo = os.path.join(PASTA_CLIENTES, arquivo)
                            nome_cliente = os.path.splitext(arquivo)[0]
                            
                            # Processar arquivo do cliente
                            self.processar_arquivo_cliente(caminho_arquivo, nome_cliente)
                            clientes_processados.append(nome_cliente)
                        except Exception as e:
                            print(f"Erro ao processar arquivo {arquivo}: {str(e)}")
                
                if not clientes_processados:
                    messagebox.showwarning("Aviso", "Nenhum arquivo de cliente encontrado!")
                    return False
                    
                print(f"Clientes processados: {len(clientes_processados)}")
                
            else:
                # Processar apenas o cliente selecionado
                if not os.path.exists(self.arquivo_cliente):
                    messagebox.showerror("Erro", f"Arquivo do cliente '{self.cliente_atual}' não encontrado!")
                    return False
                
                # Processar arquivo do cliente
                self.processar_arquivo_cliente(self.arquivo_cliente, self.cliente_atual)
            
            # Verificar se encontrou algum lançamento
            if self.total_geral == 0:
                messagebox.showinfo("Aviso", "Nenhum lançamento encontrado no período selecionado.")
                return False
            
            # Ordenar fornecedores por valor total (decrescente)
            fornecedores_ordenados = sorted(
                self.dados_fornecedores.items(),
                key=lambda x: x[1]['total'],
                reverse=True
            )
            
            # Obter top N fornecedores
            try:
                top_n = int(self.top_n_var.get())
            except ValueError:
                top_n = 10
                
            self.top_fornecedores = fornecedores_ordenados[:top_n]
            
            return True
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            import traceback
            traceback.print_exc()
            return False

def main():
    """Função principal para executar o módulo de forma independente"""
    app = RelatorioFornecedores()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()