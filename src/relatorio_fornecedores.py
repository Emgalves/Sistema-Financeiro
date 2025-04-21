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
    from config.config import (
        ARQUIVO_CLIENTES,
        PASTA_CLIENTES,
        BASE_PATH
    )
    print("Configurações importadas com sucesso")
except ImportError as e:
    print(f"Erro ao importar configurações: {str(e)}")
    # Definir valores padrão em caso de falha
    BASE_PATH = Path(".")
    ARQUIVO_CLIENTES = BASE_PATH / "dados" / "clientes.xlsx"
    PASTA_CLIENTES = BASE_PATH / "dados" / "clientes"

try:
    from config.window_config import configurar_janela
    print("window_config importado com sucesso")
except ImportError as e:
    print(f"Erro ao importar window_config: {str(e)}")
    # Implementação simples de configurar_janela como fallback
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
            
        configurar_janela(self.root, "Relatório de Principais Fornecedores", 900, 1000)
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.data_referencia = datetime.now()
        self.periodo_inicio = None
        self.periodo_fim = None
        self.todos_clientes = False
        self.dados_fornecedores = {}
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
        self.frame_principal.columnconfigure(0, weight=1)  # Coluna única expande horizontalmente
        self.frame_principal.rowconfigure(0, weight=0)  # Linha de seleção não expande
        self.frame_principal.rowconfigure(1, weight=1)  # Linha de resultados expande
        self.frame_principal.rowconfigure(2, weight=0)  # Linha de botões não expande
            
        # Frame para seleção
        self.frame_selecao = ttk.LabelFrame(self.frame_principal, text="Seleção de Cliente e Período")
        self.frame_selecao.grid(row=0, column=0, sticky='ew', pady=10)
        
        # Container para cliente
        frame_cliente = ttk.Frame(self.frame_selecao)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Selecione o Cliente:", font=('Arial', 11)).pack(side='left', pady=5)
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=40, font=('Arial', 11))
        self.cliente_combobox.pack(side='left', padx=5)
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.selecionar_cliente)
        
        # Checkbox para todos os clientes
        self.var_todos_clientes = tk.BooleanVar(value=False)
        self.cb_todos_clientes = ttk.Checkbutton(
            frame_cliente, 
            text="Analisar todos os clientes",
            variable=self.var_todos_clientes,
            command=self.alternar_todos_clientes
        )
        self.cb_todos_clientes.pack(side='left', padx=20)
        
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
        self.frame_opcoes.grid(row=1, column=0, sticky='ew', pady=10)
        
        frame_opcoes_int = ttk.Frame(self.frame_opcoes)
        frame_opcoes_int.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_opcoes_int, text="Quantidade de fornecedores a exibir:", font=('Arial', 11)).pack(side='left', pady=5)
        
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
        self.frame_resultados.grid(row=1, column=0, sticky='nsew', pady=10)
        
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
        frame_botoes.grid(row=2, column=0, sticky='ew', pady=10)

        # Use grid para os botões também
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
        frame_botoes.columnconfigure(0, weight=0)  # Primeira coluna não expande
        frame_botoes.columnconfigure(1, weight=0)  # Segunda coluna não expande
        frame_botoes.columnconfigure(2, weight=1)  # Terceira coluna expande para empurrar o botão para a direita
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 11, 'bold'), padding=(10, 5))
        
        # Carregar lista de clientes
        self.atualizar_lista_clientes()
        
        # Configurar período inicial
        self.alterar_periodo()
        
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
        
        # Criar Treeview para listar os fornecedores
        colunas = ('posicao', 'fornecedor', 'total_gasto', 'percentual', 'qtd_lancamentos', 'tipos_despesa')
        self.tree_resumo = ttk.Treeview(frame_tabela, columns=colunas, show='headings', height=20)
        
        # Configurar colunas
        self.tree_resumo.heading('posicao', text='#')
        self.tree_resumo.heading('fornecedor', text='Fornecedor')
        self.tree_resumo.heading('total_gasto', text='Total Gasto')
        self.tree_resumo.heading('percentual', text='% do Total')
        self.tree_resumo.heading('qtd_lancamentos', text='Qtd. Lançamentos')
        self.tree_resumo.heading('tipos_despesa', text='Tipos de Despesa')
        
        # Definir larguras
        self.tree_resumo.column('posicao', width=50, anchor='center')
        self.tree_resumo.column('fornecedor', width=250)
        self.tree_resumo.column('total_gasto', width=150, anchor='e')
        self.tree_resumo.column('percentual', width=100, anchor='center')
        self.tree_resumo.column('qtd_lancamentos', width=150, anchor='center')
        self.tree_resumo.column('tipos_despesa', width=150)
        
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
            text="Total Top Fornecedores: R$ 0,00 (0%)",
            font=('Arial', 12)
        )
        self.lbl_total_apresentado.pack(side='left', padx=10)
        
    def setup_aba_detalhes(self):
        """Configura a aba de detalhes do relatório"""
        # Frame para seleção de fornecedor
        frame_selecao = ttk.Frame(self.aba_detalhes, padding=5)
        frame_selecao.pack(fill='x', pady=5)
        
        ttk.Label(frame_selecao, text="Selecione o Fornecedor:", font=('Arial', 11)).pack(side='left', pady=5)
        
        self.fornecedor_combobox = ttk.Combobox(frame_selecao, width=40, font=('Arial', 11))
        self.fornecedor_combobox.pack(side='left', padx=5)
        self.fornecedor_combobox.bind('<<ComboboxSelected>>', self.carregar_detalhes_fornecedor)
        
        # Frame para informações do fornecedor
        frame_info_fornecedor = ttk.LabelFrame(self.aba_detalhes, text="Informações do Fornecedor")
        frame_info_fornecedor.pack(fill='x', pady=5, padx=5)
        
        # Grid para informações
        frame_grid = ttk.Frame(frame_info_fornecedor, padding=10)
        frame_grid.pack(fill='x')
        
        # Primeira linha
        ttk.Label(frame_grid, text="Fornecedor:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.lbl_nome_fornecedor = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_nome_fornecedor.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Total Gasto:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.lbl_total_fornecedor = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_total_fornecedor.grid(row=0, column=3, sticky='w', padx=5, pady=2)
        
        # Segunda linha
        ttk.Label(frame_grid, text="Quantidade de Lançamentos:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.lbl_qtd_lancamentos = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_qtd_lancamentos.grid(row=1, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Média por Lançamento:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=2)
        self.lbl_media_lancamento = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_media_lancamento.grid(row=1, column=3, sticky='w', padx=5, pady=2)
        
        # Terceira linha
        ttk.Label(frame_grid, text="Tipos de Despesa:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.lbl_tipos_despesa = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_tipos_despesa.grid(row=2, column=1, columnspan=3, sticky='w', padx=5, pady=2)
        
        # Frame para tabela de lançamentos
        frame_lancamentos = ttk.LabelFrame(self.aba_detalhes, text="Lançamentos")
        frame_lancamentos.pack(fill='both', expand=True, pady=5, padx=5)
        
        # Tree para lançamentos
        colunas = ('data', 'cliente', 'tipo_despesa', 'referencia', 'valor')
        self.tree_lancamentos = ttk.Treeview(frame_lancamentos, columns=colunas, show='headings', height=15)
        
        # Configurar colunas
        self.tree_lancamentos.heading('data', text='Data')
        self.tree_lancamentos.heading('cliente', text='Cliente')
        self.tree_lancamentos.heading('tipo_despesa', text='Tipo')
        self.tree_lancamentos.heading('referencia', text='Referência')
        self.tree_lancamentos.heading('valor', text='Valor')
        
        # Ajustar larguras
        self.tree_lancamentos.column('data', width=100, anchor='center')
        self.tree_lancamentos.column('cliente', width=150)
        self.tree_lancamentos.column('tipo_despesa', width=50, anchor='center')
        self.tree_lancamentos.column('referencia', width=300)
        self.tree_lancamentos.column('valor', width=100, anchor='e')
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_lancamentos, orient='vertical', command=self.tree_lancamentos.yview)
        scrollx = ttk.Scrollbar(frame_lancamentos, orient='horizontal', command=self.tree_lancamentos.xview)
        self.tree_lancamentos.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_lancamentos.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
    def setup_aba_grafico(self):
        """Configura a aba de gráficos"""
        # Frame para controles do gráfico
        frame_controles = ttk.Frame(self.aba_grafico, padding=5)
        frame_controles.pack(fill='x', pady=5)
        
        ttk.Label(frame_controles, text="Tipo de Gráfico:").pack(side='left', padx=5)
        self.combo_tipo_grafico = ttk.Combobox(frame_controles, values=[
            "Pizza - Total por Fornecedor",
            "Barras - Top Fornecedores",
            "Linhas - Evolução Mensal",
            "Barras Empilhadas - Por Tipo de Despesa"
        ], state='readonly', width=30)
        self.combo_tipo_grafico.pack(side='left', padx=5)
        self.combo_tipo_grafico.current(0)
        
        ttk.Button(frame_controles, text="Atualizar Gráfico", command=self.atualizar_grafico).pack(side='left', padx=20)
        
        # Frame para o gráfico
        self.frame_grafico = ttk.Frame(self.aba_grafico)
        self.frame_grafico.pack(fill='both', expand=True, pady=5)
        
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
        """Atualiza a lista de clientes no combobox"""
        try:
            # Carregar arquivo de clientes
            workbook = load_workbook(ARQUIVO_CLIENTES)
            sheet = workbook['Clientes']  # Assumindo que existe uma aba chamada 'Clientes'
            
            # Limpar lista atual
            self.cliente_combobox['values'] = []
            
            # Pegar todos os clientes (pulando o cabeçalho)
            clientes = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Nome do cliente está na primeira coluna
                    clientes.append(row[0])
            
            # Atualizar combobox
            self.cliente_combobox['values'] = sorted(clientes)
            workbook.close()
            
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de clientes não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
    
    def selecionar_cliente(self, event=None):
        """Atualiza o cliente selecionado"""
        # Desmarcar checkbox de todos os clientes
        self.var_todos_clientes.set(False)
        self.todos_clientes = False
        
        self.cliente_atual = self.cliente_combobox.get()
        
        if self.cliente_atual:
            # Atualizar label
            self.lbl_cliente_resumo.config(text=f"Cliente: {self.cliente_atual}")
            
            # Definir o caminho do arquivo
            self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
    
    def gerar_relatorio(self):
        """Gera o relatório com base nos dados selecionados"""
        if not self.cliente_atual and not self.todos_clientes:
            messagebox.showwarning("Aviso", "Selecione um cliente ou marque a opção 'Analisar todos os clientes'!")
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
        
        # Carregar dados
        if not self.carregar_dados():
            return
        
        # Preencher resumo
        self.preencher_resumo()
        
        # Atualizar lista de fornecedores para a aba de detalhes
        self.atualizar_lista_fornecedores()
        
        # Limpar detalhes
        self.limpar_detalhes()
        
        # Gerar gráfico inicial
        self.atualizar_grafico()
        
        # Marcar que os dados foram carregados
        self.dados_carregados = True
        
        # Selecionar aba de resumo
        self.notebook.select(0)
    
    def carregar_dados(self):
        """Carrega os dados para o relatório"""
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
    
    def processar_arquivo_cliente(self, caminho_arquivo, nome_cliente):
        """Processa um arquivo de cliente"""
        try:
            # Carregar dados do Excel
            df = pd.read_excel(caminho_arquivo, sheet_name='Dados')
            
            # Verificar colunas necessárias
            colunas_necessarias = ['DATA_REL', 'TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR']
            if not all(coluna in df.columns for coluna in colunas_necessarias):
                print(f"Arquivo {nome_cliente} não contém todas as colunas necessárias.")
                return
            
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
                        valor_str = row['VALOR'].replace('R$', '').replace('.', '').replace(',', '.').strip()
                        try:
                            valor = float(valor_str)
                        except ValueError:
                            valor = 0.0
                    
                    # Ignorar lançamentos com valor zero
                    if valor <= 0:
                        continue
                    
                    # Obter tipo de despesa
                    tipo_despesa = int(row['TP_DESP']) if pd.notnull(row['TP_DESP']) else 0
                    
                    # Obter referência
                    referencia = str(row['REFERÊNCIA']) if pd.notnull(row['REFERÊNCIA']) else ""
                    
                    # Obter data
                    data = row['DATA_REL']
                    
                    # Criar identificador do mês para análise mensal
                    mes_ano = f"{data.year}-{data.month:02d}"
                    
                    # Atualizar dados do fornecedor
                    self.dados_fornecedores[fornecedor]['total'] += valor
                    self.dados_fornecedores[fornecedor]['qtd_lancamentos'] += 1
                    self.dados_fornecedores[fornecedor]['tipos_despesa'].add(tipo_despesa)
                    self.dados_fornecedores[fornecedor]['clientes'].add(nome_cliente)
                    self.dados_fornecedores[fornecedor]['por_mes'][mes_ano] += valor
                    self.dados_fornecedores[fornecedor]['por_tipo'][tipo_despesa] += valor
                    
                    # Adicionar lançamento à lista de lançamentos do fornecedor
                    self.dados_fornecedores[fornecedor]['lancamentos'].append({
                        'data': data,
                        'cliente': nome_cliente,
                        'tipo_despesa': tipo_despesa,
                        'referencia': referencia,
                        'valor': valor
                    })
                    
                    # Atualizar total geral
                    self.total_geral += valor
                    
                except Exception as e:
                    print(f"Erro ao processar lançamento: {str(e)}")
                    continue
            
            print(f"Cliente {nome_cliente} processado: {len(df_periodo)} lançamentos.")
            
        except Exception as e:
            print(f"Erro ao processar arquivo {nome_cliente}: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def preencher_resumo(self):
        """Preenche a tabela de resumo com os principais fornecedores"""
        # Limpar treeview
        for item in self.tree_resumo.get_children():
            self.tree_resumo.delete(item)
        
        # Verificar se há dados
        if not self.top_fornecedores:
            return
        
        # Total dos top fornecedores
        total_top = sum(dados['total'] for _, dados in self.top_fornecedores)
        
        # Atualizar labels de totais
        self.lbl_total_geral.config(text=f"Total Geral: {formatar_moeda_br(self.total_geral)}")
        self.lbl_total_apresentado.config(
            text=f"Total Top Fornecedores: {formatar_moeda_br(total_top)} ({total_top/self.total_geral*100:.1f}%)"
        )
        
        # Adicionar fornecedores à tabela
        for i, (fornecedor, dados) in enumerate(self.top_fornecedores, 1):
            # Formatar tipos de despesa
            tipos_str = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
            
            # Calcular percentual em relação ao total geral
            percentual = (dados['total'] / self.total_geral) * 100
            
            # Inserir linha na tabela
            self.tree_resumo.insert('', 'end', values=(
                i,  # Posição
                fornecedor,  # Nome do fornecedor
                formatar_moeda_br(dados['total']),  # Total gasto
                f"{percentual:.1f}%",  # Percentual
                dados['qtd_lancamentos'],  # Quantidade de lançamentos
                tipos_str  # Tipos de despesa
            ))
    
    def atualizar_lista_fornecedores(self):
        """Atualiza a lista de fornecedores no combobox da aba de detalhes"""
        # Limpar lista atual
        self.fornecedor_combobox['values'] = []
        
        # Verificar se há dados
        if not self.top_fornecedores:
            return
        
        # Obter nomes dos fornecedores na ordem do ranking
        fornecedores = [f"{i}. {fornecedor}" for i, (fornecedor, _) in enumerate(self.top_fornecedores, 1)]
        
        # Atualizar combobox
        self.fornecedor_combobox['values'] = fornecedores
        
        # Selecionar o primeiro
        if fornecedores:
            self.fornecedor_combobox.current(0)
            self.carregar_detalhes_fornecedor()
    
    def limpar_detalhes(self):
        """Limpa os detalhes do fornecedor"""
        self.lbl_nome_fornecedor.config(text="")
        self.lbl_total_fornecedor.config(text="")
        self.lbl_qtd_lancamentos.config(text="")
        self.lbl_media_lancamento.config(text="")
        self.lbl_tipos_despesa.config(text="")
        
        # Limpar treeview de lançamentos
        for item in self.tree_lancamentos.get_children():
            self.tree_lancamentos.delete(item)
    
    def carregar_detalhes_fornecedor(self, event=None):
        """Carrega os detalhes do fornecedor selecionado"""
        # Verificar se há dados carregados
        if not hasattr(self, 'top_fornecedores') or not self.top_fornecedores:
            return
        
        # Obter fornecedor selecionado
        selecao = self.fornecedor_combobox.get()
        if not selecao:
            return
        
        # Extrair nome do fornecedor (remover número do ranking)
        partes = selecao.split('. ', 1)
        if len(partes) < 2:
            return
            
        fornecedor = partes[1]
        
        # Verificar se o fornecedor existe nos dados
        if fornecedor not in self.dados_fornecedores:
            messagebox.showwarning("Aviso", f"Fornecedor {fornecedor} não encontrado nos dados!")
            return
        
        # Obter dados do fornecedor
        dados = self.dados_fornecedores[fornecedor]
        
        # Atualizar labels
        self.lbl_nome_fornecedor.config(text=fornecedor)
        self.lbl_total_fornecedor.config(text=formatar_moeda_br(dados['total']))
        self.lbl_qtd_lancamentos.config(text=str(dados['qtd_lancamentos']))
        
        # Calcular média por lançamento
        if dados['qtd_lancamentos'] > 0:
            media = dados['total'] / dados['qtd_lancamentos']
            self.lbl_media_lancamento.config(text=formatar_moeda_br(media))
        else:
            self.lbl_media_lancamento.config(text="R$ 0,00")
        
        # Formatar tipos de despesa
        tipos_str = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
        self.lbl_tipos_despesa.config(text=tipos_str)
        
        # Limpar treeview de lançamentos
        for item in self.tree_lancamentos.get_children():
            self.tree_lancamentos.delete(item)
        
        # Ordenar lançamentos por data (decrescente)
        lancamentos_ordenados = sorted(
            dados['lancamentos'],
            key=lambda x: x['data'],
            reverse=True
        )
        
        # Adicionar lançamentos à tabela
        for lancamento in lancamentos_ordenados:
            self.tree_lancamentos.insert('', 'end', values=(
                lancamento['data'].strftime('%d/%m/%Y'),
                lancamento['cliente'],
                lancamento['tipo_despesa'],
                lancamento['referencia'],
                formatar_moeda_br(lancamento['valor'])
            ))
    
    def atualizar_grafico(self):
        """Atualiza o gráfico com base no tipo selecionado"""
        # Verificar se há dados carregados
        if not hasattr(self, 'top_fornecedores') or not self.top_fornecedores:
            return
            
        tipo_grafico = self.combo_tipo_grafico.get()
        
        # Limpar frame do gráfico
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
            
        # Criar figura com tamanho adequado e mais espaço para títulos
        fig = plt.figure(figsize=(10, 6), constrained_layout=True)
        
        # Adicionar mais espaço na parte superior para títulos
        ax = fig.add_subplot(111)
        
        # Verificar qual gráfico criar
        if tipo_grafico == "Pizza - Total por Fornecedor":
            self.criar_grafico_pizza(fig, ax)
        elif tipo_grafico == "Barras - Top Fornecedores":
            self.criar_grafico_barras(fig, ax)
        elif tipo_grafico == "Linhas - Evolução Mensal":
            self.criar_grafico_linha(fig, ax)
        elif tipo_grafico == "Barras Empilhadas - Por Tipo de Despesa":
            self.criar_grafico_barras_empilhadas(fig, ax)
        
        # Adicionar título principal com espaçamento adequado
        # Usar suptitle com y=0.98 para posicionar acima do título do gráfico
        fig.suptitle(
            f"Análise de Fornecedores - {self.periodo_inicio.strftime('%d/%m/%Y')} a {self.periodo_fim.strftime('%d/%m/%Y')}",
            fontsize=14,
            fontweight='bold',
            y=0.98  # Posicionamento mais alto
        )
        
        # Ajuste de layout automático para evitar sobreposições
        fig.tight_layout(rect=[0, 0, 1, 0.95])  # Reservar espaço para o título principal
            
        # Exibir o gráfico
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
        
        # Adicionar barra de ferramentas de navegação (opcional)
        try:
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            toolbar = NavigationToolbar2Tk(canvas, self.frame_grafico)
            toolbar.update()
        except ImportError:
            pass  # Se não conseguir importar, prossegue sem a barra de ferramentas
    
    def criar_grafico_pizza(self, fig, ax):
        """Cria um gráfico de pizza com legenda reposicionada e melhor formatação de título"""
        # Limpar o eixo antes de desenhar
        ax.clear()
        
        # Ajustar o tamanho da figura e layout
        fig.subplots_adjust(left=0.05, right=0.65, top=0.9, bottom=0.1)
        
        # Preparar dados
        # Mostrar apenas os top 7 e agrupar o resto como "Outros"
        top_n = min(7, len(self.top_fornecedores))
        
        labels = []
        valores = []
        
        # Adicionar top fornecedores
        for i in range(top_n):
            fornecedor, dados = self.top_fornecedores[i]
            # Limitar tamanho do nome para o gráfico
            nome_curto = fornecedor[:15] + '...' if len(fornecedor) > 15 else fornecedor
            labels.append(f"{i+1}. {nome_curto}")
            valores.append(dados['total'])
        
        # Adicionar "Outros" se houver mais fornecedores
        if len(self.top_fornecedores) > top_n:
            valor_outros = sum(dados['total'] for _, dados in self.top_fornecedores[top_n:])
            labels.append("Outros")
            valores.append(valor_outros)
        
        # Criar gráfico com tamanho reduzido para acomodar legenda
        wedges, texts, autotexts = ax.pie(
            valores, 
            labels=None,
            autopct='%1.1f%%',
            startangle=90,
            shadow=False,
            colors=plt.cm.tab20.colors[:len(valores)],
            wedgeprops={'linewidth': 1, 'edgecolor': 'white'}
        )
        
        # Ajustar tamanho do texto das porcentagens
        for autotext in autotexts:
            autotext.set_fontsize(8)
            autotext.set_weight('bold')
            autotext.set_color('white')
        
        # Adicionar legenda com melhor posicionamento e tamanho
        legend = ax.legend(
            wedges, 
            labels, 
            loc="center left",
            bbox_to_anchor=(1.05, 0.5),  # Mover mais para a direita
            fontsize=8,                  # Fonte menor
            frameon=True,                # Adicionar borda
            framealpha=0.8,              # Tornar fundo semi-transparente
            title="Fornecedores",
            title_fontsize=9
        )
        
        # Adicionar título e subtítulo separados com espaço adequado
        ax.set_title('Distribuição do Valor Total por Fornecedor', 
                    fontsize=12, 
                    pad=20,            # Adicionar espaço entre título e gráfico
                    fontweight='bold')
        
        # Adicionar um círculo central (opcional, para criar efeito de "donut")
        # Isso pode ajudar a tornar o gráfico de pizza mais atraente
        centre_circle = plt.Circle((0, 0), 0.3, fc='white', ec='lightgray')
        ax.add_patch(centre_circle)
        
        # Garantir que o aspecto do gráfico seja igual (círculo perfeito)
        ax.set_aspect('equal')
        
        # Adicionarr totais no círculo central
        total_fmt = formatar_moeda_br(sum(valores))
        ax.text(0, 0, f"Total\n{total_fmt}", 
                ha='center', va='center', fontsize=9, fontweight='bold')
    
    def criar_grafico_barras(self, fig, ax):
        """Cria um gráfico de barras horizontais com melhor formatação"""
        # Limpar o eixo antes de desenhar
        ax.clear()
        
        # Ajustar margens
        fig.subplots_adjust(left=0.3, right=0.95, top=0.85, bottom=0.1)
        
        # Mostrar apenas os top 15 fornecedores
        top_n = min(15, len(self.top_fornecedores))
        
        # Preparar dados
        labels = []
        valores = []
        
        # Adicionar fornecedores em ordem inversa para que o maior apareça no topo
        for i in range(top_n - 1, -1, -1):
            fornecedor, dados = self.top_fornecedores[i]
            # Limitar tamanho do nome para o gráfico
            nome_curto = fornecedor[:20] + '...' if len(fornecedor) > 20 else fornecedor
            labels.append(f"{i+1}. {nome_curto}")
            valores.append(dados['total'])
        
        # Definir cores com gradiente baseado nos valores
        cores = plt.cm.Blues(np.linspace(0.4, 0.9, len(valores)))
        
        # Criar barras horizontais
        bars = ax.barh(
            labels, 
            valores,
            color=cores,
            height=0.7,  # Barras mais finas para melhor visualização
            edgecolor='grey',
            linewidth=0.5
        )
        
        # Adicionar rótulos de valor nas barras
        for i, bar in enumerate(bars):
            width = bar.get_width()
            label_x_pos = width * 1.01
            # Verificar se o valor é muito grande para caber na figura
            if label_x_pos > max(valores) * 0.95:
                label_x_pos = width * 0.95
                ha = 'right'
                color = 'white'
            else:
                ha = 'left'
                color = 'black'
                
            ax.text(
                label_x_pos,
                bar.get_y() + bar.get_height()/2,
                formatar_moeda_br(width),
                va='center',
                ha=ha,
                fontsize=8,
                color=color,
                fontweight='bold'
            )
        
        # Formatar eixo Y (nomes) para melhor legibilidade
        ax.tick_params(axis='y', labelsize=9)
        
        # Remover linhas de grade no eixo X (valores)
        ax.xaxis.grid(True, linestyle='--', alpha=0.7)
        ax.yaxis.grid(False)
        
        # Remover bordas desnecessárias
        for spine in ['top', 'right']:
            ax.spines[spine].set_visible(False)
        
        ax.set_title('Top Fornecedores por Valor Total', fontsize=12, pad=20, fontweight='bold')
        ax.set_xlabel('Valor Total', fontsize=10)
        
        # Formatação de números no eixo X
        from matplotlib.ticker import FuncFormatter
        def formato_eixo(x, pos):
            if x >= 1000000:
                return f'R$ {x/1000000:.1f}M'
            elif x >= 1000:
                return f'R$ {x/1000:.0f}K'
            else:
                return f'R$ {x:.0f}'
        
        ax.xaxis.set_major_formatter(FuncFormatter(formato_eixo))
    
    def criar_grafico_linha(self, fig, ax):
        """Cria um gráfico de linha com melhor formatação"""
        # Limpar o eixo antes de desenhar
        ax.clear()
        
        # Ajustar layout
        fig.subplots_adjust(left=0.1, right=0.75, top=0.85, bottom=0.15)
        
        # Mostrar apenas os top 5 fornecedores
        top_n = min(5, len(self.top_fornecedores))
        
        # Obter todos os meses do período
        todos_meses = set()
        for _, dados in self.top_fornecedores[:top_n]:
            todos_meses.update(dados['por_mes'].keys())
        
        # Ordenar meses
        meses_ordenados = sorted(todos_meses)
        
        # Verificar se há meses para plotar
        if not meses_ordenados:
            ax.text(0.5, 0.5, "Sem dados mensais para exibir", 
                    ha='center', va='center', fontsize=12)
            ax.set_title('Evolução Mensal dos Top Fornecedores', fontsize=12, pad=20)
            return
        
        # Formatação de meses para exibição
        meses_formatados = []
        for mes in meses_ordenados:
            ano, mes_num = mes.split('-')
            # Converter para nomes abreviados de meses
            nomes_meses = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 
                        'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez']
            mes_fmt = f"{nomes_meses[int(mes_num)-1]}/{ano[2:]}"
            meses_formatados.append(mes_fmt)
        
        # Cores para cada fornecedor
        cores = plt.cm.tab10.colors[:top_n]
        
        # Preparar dados por fornecedor
        for i in range(top_n):
            fornecedor, dados = self.top_fornecedores[i]
            
            # Limitar tamanho do nome para o gráfico
            nome_curto = fornecedor[:20] + '...' if len(fornecedor) > 20 else fornecedor
            label = f"{i+1}. {nome_curto}"
            
            # Preparar valores mensais
            valores = [dados['por_mes'].get(mes, 0) for mes in meses_ordenados]
            
            # Plotar linha
            linha = ax.plot(
                meses_formatados, 
                valores,
                marker='o',
                linewidth=2,
                label=label,
                color=cores[i],
                markersize=6,
                markeredgecolor='white',
                markeredgewidth=1
            )
            
            # Adicionar rótulos de valor no último ponto (opcional)
            if valores[-1] > 0:
                ax.annotate(
                    formatar_moeda_br(valores[-1]),
                    xy=(len(valores)-1, valores[-1]),
                    xytext=(10, 0),
                    textcoords="offset points",
                    fontsize=8,
                    color=cores[i],
                    fontweight='bold'
                )
        
        # Configurar eixo X
        ax.set_xticks(range(len(meses_formatados)))
        ax.set_xticklabels(meses_formatados, rotation=45, ha='right')
        
        # Melhorar formatação do eixo Y
        from matplotlib.ticker import FuncFormatter
        def formato_eixo(y, pos):
            if y >= 1000000:
                return f'R${y/1000000:.1f}M'
            elif y >= 1000:
                return f'R${y/1000:.0f}K'
            else:
                return f'R${y:.0f}'
        
        ax.yaxis.set_major_formatter(FuncFormatter(formato_eixo))
        
        ax.set_title('Evolução Mensal dos Top Fornecedores', fontsize=12, pad=20, fontweight='bold')
        ax.set_xlabel('Mês/Ano', fontsize=10)
        ax.set_ylabel('Valor Total', fontsize=10)
        
        # Melhorar posicionamento da legenda
        ax.legend(
            loc='center left', 
            bbox_to_anchor=(1.02, 0.5),
            fontsize=9,
            frameon=True,
            framealpha=0.8,
            title="Fornecedores",
            title_fontsize=10
        )
        
        # Adicionar grid
        ax.grid(True, linestyle='--', alpha=0.7, axis='both')
        
        # Remover bordas desnecessárias
        for spine in ['top', 'right']:
            ax.spines[spine].set_visible(False)
    
    def criar_grafico_barras_empilhadas(self, fig, ax):
        """Cria um gráfico de barras empilhadas com melhor formatação"""
        # Limpar o eixo antes de desenhar
        ax.clear()
        
        # Ajustar layout
        fig.subplots_adjust(left=0.1, right=0.75, top=0.85, bottom=0.2)
        
        # Mostrar apenas os top 8 fornecedores (para não ficar muito apertado)
        top_n = min(8, len(self.top_fornecedores))
        
        # Obter todos os tipos de despesa
        todos_tipos = set()
        for _, dados in self.top_fornecedores[:top_n]:
            todos_tipos.update(dados['por_tipo'].keys())
        
        # Ordenar tipos
        tipos_ordenados = sorted(todos_tipos)
        
        # Verificar se há tipos para mostrar
        if not tipos_ordenados:
            ax.text(0.5, 0.5, "Sem dados de tipos de despesa para exibir", 
                    ha='center', va='center', fontsize=12)
            ax.set_title('Fornecedores por Tipo de Despesa', fontsize=12, pad=20)
            return
        
        # Preparar dados
        fornecedores = []
        valores_por_tipo = {tipo: [] for tipo in tipos_ordenados}
        
        # Para cada fornecedor, obter valores por tipo
        for i in range(top_n):
            fornecedor, dados = self.top_fornecedores[i]
            
            # Adicionar nome do fornecedor
            nome_curto = fornecedor[:12] + '...' if len(fornecedor) > 12 else fornecedor
            fornecedores.append(f"{i+1}. {nome_curto}")
            
            # Adicionar valores por tipo
            for tipo in tipos_ordenados:
                valores_por_tipo[tipo].append(dados['por_tipo'].get(tipo, 0))
        
        # Criar barras empilhadas
        bottom = np.zeros(len(fornecedores))
        
        # Definir cores para os tipos de despesa (com paleta mais diferenciada)
        cores = plt.cm.tab20.colors[:len(tipos_ordenados)]
        
        # Criar barras para cada tipo
        barras = []
        for i, tipo in enumerate(tipos_ordenados):
            barra = ax.bar(
                fornecedores, 
                valores_por_tipo[tipo],
                bottom=bottom,
                label=f"Tipo {tipo}",
                color=cores[i % len(cores)],
                width=0.7,  # Barras mais finas
                edgecolor='white',
                linewidth=0.5
            )
            barras.append(barra)
            bottom += np.array(valores_por_tipo[tipo])
        
        # Adicionar rótulos de valor total no topo de cada barra
        for i in range(len(fornecedores)):
            total = sum(valores_por_tipo[tipo][i] for tipo in tipos_ordenados)
            if total > 0:
                ax.text(
                    i, 
                    total * 1.01, 
                    formatar_moeda_br(total),
                    ha='center',
                    va='bottom',
                    fontsize=8,
                    fontweight='bold',
                    rotation=0
                )
        
        ax.set_title('Fornecedores por Tipo de Despesa', 
                    fontsize=12, 
                    pad=20,
                    fontweight='bold')
        ax.set_xlabel('Fornecedor', fontsize=10)
        ax.set_ylabel('Valor Total', fontsize=10)
        
        # Melhorar formatação do eixo Y
        from matplotlib.ticker import FuncFormatter
        def formato_eixo(y, pos):
            if y >= 1000000:
                return f'R${y/1000000:.1f}M'
            elif y >= 1000:
                return f'R${y/1000:.0f}K'
            else:
                return f'R${y:.0f}'
        
        ax.yaxis.set_major_formatter(FuncFormatter(formato_eixo))
        
        # Melhorar posicionamento da legenda
        ax.legend(
            loc='center left', 
            bbox_to_anchor=(1.02, 0.5),
            fontsize=9,
            frameon=True,
            framealpha=0.8,
            title="Tipos de Despesa",
            title_fontsize=10
        )
        
        # Adicionar grid apenas no eixo Y
        ax.yaxis.grid(True, linestyle='--', alpha=0.5)
        ax.xaxis.grid(False)
        
        # Remover bordas desnecessárias
        for spine in ['top', 'right']:
            ax.spines[spine].set_visible(False)
        
        # Rotacionar labels do eixo X para melhor visualização
        plt.setp(ax.get_xticklabels(), rotation=45, ha='right', fontsize=9)
        
        # Garantir que há espaço suficiente para os rótulos
        plt.tight_layout(rect=[0, 0, 0.75, 0.9])
    
    def exportar_excel(self):
        """Exporta o relatório para um arquivo Excel"""
        if not self.dados_carregados:
            messagebox.showwarning("Aviso", "Não há dados para exportar!")
            return
            
        # Solicitar nome do arquivo ao usuário
        periodo_str = f"{self.periodo_inicio.strftime('%d-%m-%Y')}_{self.periodo_fim.strftime('%d-%m-%Y')}"
        nome_padrao = f"Relatorio_Fornecedores_{periodo_str}.xlsx"
        
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Arquivos Excel", "*.xlsx")],
            initialfile=nome_padrao
        )
        
        if not arquivo:
            return
            
        try:
            # Criar workbook
            wb = Workbook()
            
            # Estilos
            titulo_font = Font(name='Arial', size=12, bold=True)
            cabecalho_font = Font(name='Arial', size=11, bold=True, color="FFFFFF")
            cabecalho_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            borda = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            
            # Aba de resumo
            ws_resumo = wb.active
            ws_resumo.title = "Resumo Fornecedores"
            
            # Título
            if self.todos_clientes:
                cliente_str = "Todos os Clientes"
            else:
                cliente_str = self.cliente_atual
                
            ws_resumo['A1'] = f"Relatório de Principais Fornecedores - {cliente_str}"
            ws_resumo['A1'].font = titulo_font
            ws_resumo.merge_cells('A1:F1')
            
            ws_resumo['A2'] = f"Período: {self.periodo_inicio.strftime('%d/%m/%Y')} a {self.periodo_fim.strftime('%d/%m/%Y')}"
            ws_resumo.merge_cells('A2:F2')
            
            # Informações de totais
            ws_resumo['A4'] = "Total Geral:"
            ws_resumo['B4'] = self.total_geral
            ws_resumo['B4'].number_format = '#,##0.00'
            
            # Cabeçalho da tabela
            row = 6
            cabecalhos = ['Posição', 'Fornecedor', 'Total Gasto', '% do Total', 'Qtd. Lançamentos', 'Tipos de Despesa']
            for col, texto in enumerate(cabecalhos, 1):
                celula = ws_resumo.cell(row=row, column=col, value=texto)
                celula.font = cabecalho_font
                celula.fill = cabecalho_fill
                celula.border = borda
                celula.alignment = Alignment(horizontal='center')
            
            # Dados dos fornecedores
            for i, (fornecedor, dados) in enumerate(self.top_fornecedores, 1):
                row += 1
                
                # Calcular percentual
                percentual = dados['total'] / self.total_geral * 100
                
                # Formatar tipos de despesa
                tipos_str = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
                
                # Adicionar linha
                ws_resumo.cell(row=row, column=1, value=i)
                ws_resumo.cell(row=row, column=2, value=fornecedor)
                ws_resumo.cell(row=row, column=3, value=dados['total'])
                ws_resumo.cell(row=row, column=4, value=f"{percentual:.1f}%")
                ws_resumo.cell(row=row, column=5, value=dados['qtd_lancamentos'])
                ws_resumo.cell(row=row, column=6, value=tipos_str)
                
                # Formatar células de valor como moeda
                ws_resumo.cell(row=row, column=3).number_format = '#,##0.00'
            
            # Ajustar larguras das colunas
            ws_resumo.column_dimensions['A'].width = 10
            ws_resumo.column_dimensions['B'].width = 40
            ws_resumo.column_dimensions['C'].width = 15
            ws_resumo.column_dimensions['D'].width = 15
            ws_resumo.column_dimensions['E'].width = 18
            ws_resumo.column_dimensions['F'].width = 25
            
            # Adicionar aba para cada fornecedor do top 5
            for i, (fornecedor, dados) in enumerate(self.top_fornecedores[:5], 1):
                # Limitar o nome da aba para 31 caracteres (limite do Excel)
                nome_aba = f"{i}_{fornecedor}"[:31]
                ws_fornecedor = wb.create_sheet(nome_aba)
                
                # Título
                ws_fornecedor['A1'] = f"Detalhamento do Fornecedor: {fornecedor}"
                ws_fornecedor['A1'].font = titulo_font
                ws_fornecedor.merge_cells('A1:E1')
                
                # Informações do fornecedor
                ws_fornecedor['A3'] = "Total Gasto:"
                ws_fornecedor['B3'] = dados['total']
                ws_fornecedor['B3'].number_format = '#,##0.00'
                
                ws_fornecedor['A4'] = "Quantidade de Lançamentos:"
                ws_fornecedor['B4'] = dados['qtd_lancamentos']
                
                ws_fornecedor['A5'] = "Média por Lançamento:"
                if dados['qtd_lancamentos'] > 0:
                    ws_fornecedor['B5'] = dados['total'] / dados['qtd_lancamentos']
                else:
                    ws_fornecedor['B5'] = 0
                ws_fornecedor['B5'].number_format = '#,##0.00'
                
                ws_fornecedor['A6'] = "Tipos de Despesa:"
                ws_fornecedor['B6'] = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
                
                # Cabeçalho da tabela de lançamentos
                row = 8
                cabecalhos = ['Data', 'Cliente', 'Tipo', 'Referência', 'Valor']
                for col, texto in enumerate(cabecalhos, 1):
                    celula = ws_fornecedor.cell(row=row, column=col, value=texto)
                    celula.font = cabecalho_font
                    celula.fill = cabecalho_fill
                    celula.border = borda
                    celula.alignment = Alignment(horizontal='center')
                
                # Ordenar lançamentos por data
                lancamentos_ordenados = sorted(dados['lancamentos'], key=lambda x: x['data'])
                
                # Adicionar lançamentos
                for lancamento in lancamentos_ordenados:
                    row += 1
                    
                    ws_fornecedor.cell(row=row, column=1, value=lancamento['data'])
                    ws_fornecedor.cell(row=row, column=2, value=lancamento['cliente'])
                    ws_fornecedor.cell(row=row, column=3, value=lancamento['tipo_despesa'])
                    ws_fornecedor.cell(row=row, column=4, value=lancamento['referencia'])
                    ws_fornecedor.cell(row=row, column=5, value=lancamento['valor'])
                    
                    # Formatar data
                    ws_fornecedor.cell(row=row, column=1).number_format = 'dd/mm/yyyy'
                    
                    # Formatar valor como moeda
                    ws_fornecedor.cell(row=row, column=5).number_format = '#,##0.00'
                
                # Ajustar larguras
                ws_fornecedor.column_dimensions['A'].width = 15
                ws_fornecedor.column_dimensions['B'].width = 25
                ws_fornecedor.column_dimensions['C'].width = 10
                ws_fornecedor.column_dimensions['D'].width = 40
                ws_fornecedor.column_dimensions['E'].width = 15
            
            # Salvar o arquivo
            wb.save(arquivo)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {str(e)}")
    
    def exportar_pdf(self):
        """Exporta o relatório para um arquivo PDF com layout aprimorado"""
        if not self.dados_carregados:
            messagebox.showwarning("Aviso", "Não há dados para exportar!")
            return
            
        try:
            # Importar reportlab apenas quando necessário
            from reportlab.lib.pagesizes import landscape, A4
            from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
            from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
            from reportlab.lib import colors
            from reportlab.lib.units import mm
            from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
            from reportlab.platypus.flowables import KeepTogether
            
            # Definir nome do cliente para o nome do arquivo
            if self.todos_clientes:
                cliente_str = "Todos_Clientes"
            else:
                # Substituir espaços por underscores e remover caracteres inválidos
                cliente_str = self.cliente_atual.replace(' ', '_').replace('/', '-').replace('\\', '-')
            
            # Gerar nome do arquivo mais específico 
            data_geracao = datetime.now().strftime('%Y%m%d-%H%M')
            periodo_str = f"{self.periodo_inicio.strftime('%d-%m-%Y')}_{self.periodo_fim.strftime('%d-%m-%Y')}"
            nome_padrao = f"Relatorio_Fornecedores_{cliente_str}_{periodo_str}.pdf"
            
            arquivo = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Arquivos PDF", "*.pdf")],
                initialfile=nome_padrao
            )
            
            if not arquivo:
                return
                
            # Dimensões de página do documento
            largura, altura = landscape(A4)
            
            # Criar documento com margens adequadas
            doc = SimpleDocTemplate(
                arquivo,
                pagesize=landscape(A4),
                rightMargin=15*mm,
                leftMargin=15*mm,
                topMargin=15*mm,
                bottomMargin=20*mm
            )
            
            # Estilos aprimorados
            styles = getSampleStyleSheet()
            
            titulo_style = ParagraphStyle(
                'TituloStyle',
                parent=styles['Heading1'],
                fontSize=16,
                leading=20,
                alignment=TA_CENTER,
                spaceBefore=10,
                spaceAfter=20
            )
            
            subtitulo_style = ParagraphStyle(
                'SubtituloStyle',
                parent=styles['Heading2'],
                fontSize=14,
                leading=18,
                spaceBefore=10,
                spaceAfter=10,
                textColor=colors.navy
            )
            
            texto_style = ParagraphStyle(
                'TextoStyle',
                parent=styles['Normal'],
                fontSize=10,
                leading=12,
                spaceBefore=5,
                spaceAfter=5
            )
            
            info_style = ParagraphStyle(
                'InfoStyle',
                parent=styles['Normal'],
                fontSize=10,
                leading=14,
                leftIndent=5*mm,
                spaceBefore=2,
                spaceAfter=2
            )
            
            # Função para quebrar textos longos
            def quebrar_texto(texto, tamanho_max=70):
                """Quebra textos longos para evitar sobreposição nas células da tabela"""
                if not texto or len(texto) <= tamanho_max:
                    return texto
                    
                palavras = texto.split()
                linhas = []
                linha_atual = []
                
                for palavra in palavras:
                    if len(' '.join(linha_atual + [palavra])) <= tamanho_max:
                        linha_atual.append(palavra)
                    else:
                        if linha_atual:
                            linhas.append(' '.join(linha_atual))
                            linha_atual = [palavra]
                        else:
                            # Caso a palavra seja maior que o tamanho máximo
                            linhas.append(palavra)
                            linha_atual = []
                
                if linha_atual:
                    linhas.append(' '.join(linha_atual))
                    
                return '\n'.join(linhas)
            
            # Lista de elementos para o PDF
            elementos = []
            
            # Título e informações de cabeçalho
            titulo = f"Relatório de Principais Fornecedores - {self.cliente_atual if not self.todos_clientes else 'Todos os Clientes'}"
            elementos.append(Paragraph(titulo, titulo_style))
            
            # Informações do período e outras informações em negrito
            periodo_info = f"<b>Período:</b> {self.periodo_inicio.strftime('%d/%m/%Y')} a {self.periodo_fim.strftime('%d/%m/%Y')}"
            elementos.append(Paragraph(periodo_info, texto_style))
            
            # Data de geração
            data_atual = datetime.now().strftime('%d/%m/%Y %H:%M')
            # elementos.append(Paragraph(f"<b>Data de geração:</b> {data_atual}", texto_style))
            
            # Totais
            # elementos.append(Spacer(1, 10*mm))
            elementos.append(Paragraph(f"<b>Total Geral:</b> {formatar_moeda_br(self.total_geral)}", texto_style))
            
            # Calcular total dos top fornecedores
            total_top = sum(dados['total'] for _, dados in self.top_fornecedores)
            percentual_top = (total_top / self.total_geral) * 100
            elementos.append(Paragraph(
                f"<b>Total Top Fornecedores:</b> {formatar_moeda_br(total_top)} ({percentual_top:.1f}%)",
                texto_style
            ))
            
            # Tabela de resumo
            elementos.append(Spacer(1, 10*mm))
            elementos.append(Paragraph("RESUMO DOS PRINCIPAIS FORNECEDORES", subtitulo_style))
            elementos.append(Spacer(1, 3*mm))
            
            # Cabeçalho da tabela com larguras ajustadas
            dados_tabela = [
                ['#', 'Fornecedor', 'Total Gasto', '% do Total', 'Qtd.', 'Tipos']
            ]
            
            # Adicionar dados da tabela
            for i, (fornecedor, dados) in enumerate(self.top_fornecedores, 1):
                # Calcular percentual
                percentual = dados['total'] / self.total_geral * 100
                
                # Formatar tipos de despesa (limitar tamanho)
                tipos_str = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
                if len(tipos_str) > 20:  # Limitar para evitar tabelas muito largas
                    tipos_str = tipos_str[:17] + "..."
                
                # Adicionar linha com textos quebrados conforme necessário
                dados_tabela.append([
                    str(i),
                    quebrar_texto(fornecedor, 40),  # Quebrar nomes muito longos
                    formatar_moeda_br(dados['total']),
                    f"{percentual:.1f}%",
                    str(dados['qtd_lancamentos']),
                    tipos_str
                ])
            
            # Larguras de coluna ajustadas para caber na página landscape
            larguras_colunas = [20*mm, 110*mm, 40*mm, 30*mm, 25*mm, 20*mm]
            
            # Criar tabela com larguras específicas
            tabela = Table(dados_tabela, colWidths=larguras_colunas, repeatRows=1)
            
            # Estilo da tabela aprimorado
            estilo_tabela = TableStyle([
                # Cabeçalho
                ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                ('TOPPADDING', (0, 0), (-1, 0), 6),
                
                # Células de dados
                ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BOX', (0, 0), (-1, -1), 1, colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),  # Coluna de posição
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),  # Centralização vertical
                ('ALIGN', (2, 1), (2, -1), 'RIGHT'),   # Coluna de valor
                ('ALIGN', (3, 1), (3, -1), 'CENTER'),  # Coluna de percentual
                ('ALIGN', (4, 1), (4, -1), 'CENTER'),  # Coluna de quantidade
                
                # Espaçamento interno das células
                ('LEFTPADDING', (0, 0), (-1, -1), 4),
                ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                ('TOPPADDING', (0, 1), (-1, -1), 4),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
                
                # Linhas verticais internas - mais finas
                ('LINEABOVE', (0, 1), (-1, -1), 0.25, colors.grey),
            ])
            
            # Zebrar linhas para melhor leitura
            for i in range(1, len(dados_tabela)):
                if i % 2 == 0:
                    estilo_tabela.add('BACKGROUND', (0, i), (-1, i), colors.lightgrey)
            
            tabela.setStyle(estilo_tabela)
            
            # Manter tabela junta se possível
            elementos.append(KeepTogether([tabela, Spacer(1, 5*mm)]))
            
            # Adicionar detalhes dos top fornecedores em páginas separadas
            for i, (fornecedor, dados) in enumerate(self.top_fornecedores[:min(5, len(self.top_fornecedores))], 1):
                elementos.append(PageBreak())
                
                # Título da página do fornecedor
                elementos.append(Paragraph(f"Detalhamento do Fornecedor", subtitulo_style))
                elementos.append(Paragraph(f"<b>{fornecedor}</b>", texto_style))
                elementos.append(Spacer(1, 5*mm))
                
                # Criar blocos de informações do fornecedor em layout mais organizado
                info_elementos = []
                
                # Primeira linha: Total e Quantidade
                info_elementos.append(Paragraph(f"<b>Total Gasto:</b> {formatar_moeda_br(dados['total'])}", info_style))
                info_elementos.append(Paragraph(f"<b>Quantidade de Lançamentos:</b> {dados['qtd_lancamentos']}", info_style))
                
                # Segunda linha: Média e Percentual
                if dados['qtd_lancamentos'] > 0:
                    media = dados['total'] / dados['qtd_lancamentos']
                    percentual = (dados['total'] / self.total_geral) * 100
                    info_elementos.append(Paragraph(f"<b>Média por Lançamento:</b> {formatar_moeda_br(media)}", info_style))
                    info_elementos.append(Paragraph(f"<b>Percentual do Total:</b> {percentual:.2f}%", info_style))
                
                # Tipos de despesa com formatação mais clara
                tipos_str = ", ".join([str(tipo) for tipo in sorted(dados['tipos_despesa'])])
                info_elementos.append(Paragraph(f"<b>Tipos de Despesa:</b> {tipos_str}", info_style))
                
                # Clientes (se aplicável)
                if len(dados['clientes']) > 0:
                    clientes_str = ", ".join(sorted(dados['clientes']))
                    info_elementos.append(Paragraph(f"<b>Clientes:</b> {clientes_str}", info_style))
                
                # Adicionar espaço após informações
                info_elementos.append(Spacer(1, 10*mm))
                
                # Título da tabela de lançamentos
                info_elementos.append(Paragraph("Lançamentos", subtitulo_style))
                info_elementos.append(Spacer(1, 3*mm))
                
                # Cabeçalho da tabela de lançamentos
                dados_lancamentos = [
                    ['Data', 'Cliente', 'Tipo', 'Referência', 'Valor']
                ]
                
                # Ordenar lançamentos por data
                lancamentos_ordenados = sorted(dados['lancamentos'], key=lambda x: x['data'], reverse=True)
                
                # Determinar número máximo de lançamentos com base no espaço disponível
                max_lancamentos = min(30, len(lancamentos_ordenados))
                
                # Adicionar lançamentos à tabela
                for j, lancamento in enumerate(lancamentos_ordenados[:max_lancamentos]):
                    # Formatar valores e quebrar textos longos
                    referencia = quebrar_texto(lancamento['referencia'], 40)
                    
                    dados_lancamentos.append([
                        lancamento['data'].strftime('%d/%m/%Y'),
                        quebrar_texto(lancamento['cliente'], 40),
                        str(lancamento['tipo_despesa']),
                        referencia,
                        formatar_moeda_br(lancamento['valor'])
                    ])
                
                # Adicionar indicação se há mais lançamentos
                if len(lancamentos_ordenados) > max_lancamentos:
                    info_elementos.append(Paragraph(
                        f"(Mostrando {max_lancamentos} de {len(lancamentos_ordenados)} lançamentos - mais recentes)", 
                        texto_style
                    ))
                
                # Larguras ajustadas para a tabela de lançamentos
                larguras_lancamentos = [20*mm, 95*mm, 10*mm, 100*mm, 30*mm]
                
                # Criar tabela de lançamentos
                tabela_lancamentos = Table(
                    dados_lancamentos, 
                    colWidths=larguras_lancamentos,
                    repeatRows=1
                )
                
                # Estilo da tabela de lançamentos
                estilo_lancamentos = TableStyle([
                    # Cabeçalho
                    ('BACKGROUND', (0, 0), (-1, 0), colors.navy),
                    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
                    ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
                    ('TOPPADDING', (0, 0), (-1, 0), 6),
                    
                    # Conteúdo
                    ('BACKGROUND', (0, 1), (-1, -1), colors.white),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('BOX', (0, 0), (-1, -1), 1, colors.black),
                    
                    # Alinhamentos
                    ('ALIGN', (0, 1), (0, -1), 'CENTER'),  # Data centralizada
                    ('ALIGN', (2, 1), (2, -1), 'CENTER'),  # Tipo centralizado
                    ('ALIGN', (4, 1), (4, -1), 'RIGHT'),   # Valor à direita
                    ('VALIGN', (0, 0), (-1, -1), 'TOP'),   # Alinhar ao topo para textos com quebra
                    
                    # Espaçamento interno
                    ('LEFTPADDING', (0, 0), (-1, -1), 4),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 4),
                    ('TOPPADDING', (0, 1), (-1, -1), 4),
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 4),
                ])
                
                # Zebrar linhas para melhor leitura
                for j in range(1, len(dados_lancamentos)):
                    if j % 2 == 0:
                        estilo_lancamentos.add('BACKGROUND', (0, j), (-1, j), colors.lightgrey)
                
                tabela_lancamentos.setStyle(estilo_lancamentos)
                info_elementos.append(tabela_lancamentos)
                
                # Adicionar todos os elementos de informação
                for elem in info_elementos:
                    elementos.append(elem)
            
            # Adicionar rodapé ao documento
            def adicionar_rodape(canvas, doc):
                canvas.saveState()
                # Desenhar linha do rodapé
                footer_y = 15*mm
                canvas.setStrokeColor(colors.grey)
                canvas.line(15*mm, footer_y, largura-15*mm, footer_y)
                
                # Adicionar texto do rodapé
                canvas.setFont('Helvetica', 8)
                canvas.drawString(15*mm, footer_y-10, f"Relatório gerado em: {data_atual}")
                
                # Adicionar numeração de página
                page_num = canvas.getPageNumber()
                texto_pagina = f"Página {page_num}"
                canvas.drawRightString(largura-15*mm, footer_y-10, texto_pagina)
                
                canvas.restoreState()
            
            # Criar o PDF com rodapé
            doc.build(elementos, onFirstPage=adicionar_rodape, onLaterPages=adicionar_rodape)
            
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{arquivo}")
            
        except ImportError:
            messagebox.showerror("Erro", "Biblioteca ReportLab não encontrada. Instale-a com 'pip install reportlab'.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para PDF: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def voltar_menu(self):
        """Volta ao menu principal"""
        self.root.destroy()
        
        # Mostrar janela principal
        if self.menu_principal:
            self.menu_principal.deiconify()
            self.menu_principal.lift()
            self.menu_principal.focus_force()

def main():
    """Função principal para executar o módulo de forma independente"""
    app = RelatorioFornecedores()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()