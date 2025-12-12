import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tkcalendar import DateEntry
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# Adicionar diretório raiz ao path para importar módulos corretamente
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
        BASE_PATH
    )
    print("Configurações importadas com sucesso")
except ImportError as e:
    print(f"Erro ao importar configurações: {str(e)}")
    # Definir valores padrão em caso de falha
    BASE_PATH = Path(".")
    ARQUIVO_CLIENTES = BASE_PATH / "dados" / "clientes.xlsx"
    PASTA_CLIENTES = BASE_PATH / "dados" / "clientes"

# Importar o utils.py
from src.config.utils import atualizar_combobox_clientes, cliente_esta_ativo, obter_info_cliente


try:
    from src.config.window_config import configurar_janela
    print("window_config importado com sucesso")
except ImportError as e:
    print(f"Erro ao importar window_config: {str(e)}")
    # Implementação simples de configurar_janela como fallback
    def configurar_janela(janela, titulo, largura=900, altura=1000):
        """
        Configura o posicionamento e dimensionamento padrão de uma janela
        
        Args:
            janela: Instância de tk.Tk ou tk.Toplevel
            titulo: Título da janela
            largura: Largura desejada (default 900)
            altura: Altura desejada (default 1000)
        """
        janela.title(titulo)
        
        # Obter dimensões da tela
        screen_width = janela.winfo_screenwidth()
        screen_height = janela.winfo_screenheight()
        
        # Ajustar dimensões para não exceder o tamanho da tela
        largura = min(largura, screen_width)
        altura = min(altura, screen_height)
        
        # Definir posição (sempre no topo esquerdo)
        x = 0
        y = 0
        
        # Configurar geometria
        janela.geometry(f"{largura}x{altura}+{x}+{y}")
        
        # Permitir redimensionamento
        janela.resizable(True, True)
        
        # Configurar peso das linhas/colunas para redimensionamento proporcional
        janela.grid_rowconfigure(0, weight=1)
        janela.grid_columnconfigure(0, weight=1)
        
        # Trazer janela para frente
        janela.lift()
        janela.focus_force()
# Importar funções auxiliares ou definir aqui
def formatar_moeda_br(valor):
    """Formata um valor numérico como moeda brasileira"""
    try:
        valor_float = float(valor)
        return f"R$ {valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except (ValueError, TypeError):
        return f"R$ 0,00"

class RelatorioContratos:
    """Classe para geração de relatórios de posição de contratos"""
    def __init__(self, parent=None):
        """Inicializa a interface do relatório de contratos"""
        self.parent = parent
        
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        configurar_janela(self.root, "Relatório de Contratos por Medição", 1000, 950)
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.data_referencia = datetime.now()
        self.contratos = []
        self.medicoes = []
        
        # Configurar interface
        self.setup_gui()
        
    def setup_gui(self):
        """Configuração da interface gráfica principal"""
        # Frame principal
        self.frame_principal = ttk.Frame(self.root, padding=10)
        self.frame_principal.pack(fill='both', expand=True)
        
        # Frame para seleção
        self.frame_selecao = ttk.LabelFrame(self.frame_principal, text="Seleção de Cliente e Data")
        self.frame_selecao.pack(fill='x', pady=10)
        
        # Container para cliente
        frame_cliente = ttk.Frame(self.frame_selecao)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Selecione o Cliente:", font=('Arial', 11)).pack(side='left', pady=5)
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=40, font=('Arial', 11))
        self.cliente_combobox.pack(side='left', padx=5)
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.selecionar_cliente)
        
        # Container para data
        frame_data = ttk.Frame(self.frame_selecao)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_data, text="Data de Referência:", font=('Arial', 11)).pack(side='left', pady=5)
        self.data_entry = DateEntry(
            frame_data, 
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR',
            font=('Arial', 11)
        )
        self.data_entry.pack(side='left', padx=5)
        self.data_entry.set_date(datetime.now())
        
        # Botão de gerar relatório
        ttk.Button(
            frame_data,
            text="Gerar Relatório",
            command=self.gerar_relatorio,
            style='Big.TButton'
        ).pack(side='left', padx=20)
        
        # Frame para resultados - com notebook para separar visões
        self.frame_resultados = ttk.LabelFrame(self.frame_principal, text="Resultados")
        self.frame_resultados.pack(fill='both', expand=True, pady=10)
        
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
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(
            frame_botoes,
            text="Exportar para Excel",
            command=self.exportar_excel
        ).pack(side='left', padx=5)
        
        ttk.Button(
            frame_botoes,
            text="Voltar ao Menu",
            command=self.voltar_menu
        ).pack(side='right', padx=5)
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 11, 'bold'), padding=(10, 5))
        
        # Carregar lista de clientes
        self.atualizar_lista_clientes()
        
    def setup_aba_resumo(self):
        """Configura a aba de resumo do relatório"""
        # Frame para informações do cliente
        frame_info = ttk.Frame(self.aba_resumo, padding=5)
        frame_info.pack(fill='x', pady=5)
        
        self.lbl_cliente_resumo = ttk.Label(
            frame_info, 
            text="Cliente: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_cliente_resumo.pack(side='left', padx=10)
        
        self.lbl_data_resumo = ttk.Label(
            frame_info, 
            text=f"Data: {datetime.now().strftime('%d/%m/%Y')}", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_data_resumo.pack(side='left', padx=10)
        
        # Frame para totais
        frame_totais = ttk.LabelFrame(self.aba_resumo, text="Totais Consolidados")
        frame_totais.pack(fill='x', pady=5)
        
        # Grid para os campos de totais
        frame_grid_totais = ttk.Frame(frame_totais, padding=10)
        frame_grid_totais.pack(fill='x')
        
        # Labels da primeira linha
        ttk.Label(frame_grid_totais, text="Total de Contratos:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.lbl_qtd_contratos = ttk.Label(frame_grid_totais, text="0", font=('Arial', 10))
        self.lbl_qtd_contratos.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_grid_totais, text="Contratos em Andamento:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=5)
        self.lbl_qtd_em_andamento = ttk.Label(frame_grid_totais, text="0", font=('Arial', 10))
        self.lbl_qtd_em_andamento.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        # Labels da segunda linha
        ttk.Label(frame_grid_totais, text="Valor Total dos Contratos:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.lbl_valor_total = ttk.Label(frame_grid_totais, text="R$ 0,00", font=('Arial', 10))
        self.lbl_valor_total.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_grid_totais, text="Valor já Pago:", font=('Arial', 10, 'bold')).grid(row=1, column=2, sticky='w', padx=5, pady=5)
        self.lbl_valor_pago = ttk.Label(frame_grid_totais, text="R$ 0,00", font=('Arial', 10))
        self.lbl_valor_pago.grid(row=1, column=3, sticky='w', padx=5, pady=5)
        
        # Labels da terceira linha
        ttk.Label(frame_grid_totais, text="Saldo a Pagar:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=5)
        self.lbl_saldo = ttk.Label(frame_grid_totais, text="R$ 0,00", font=('Arial', 10))
        self.lbl_saldo.grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_grid_totais, text="Percentual Executado:", font=('Arial', 10, 'bold')).grid(row=2, column=2, sticky='w', padx=5, pady=5)
        self.lbl_percentual = ttk.Label(frame_grid_totais, text="0%", font=('Arial', 10))
        self.lbl_percentual.grid(row=2, column=3, sticky='w', padx=5, pady=5)
        
        # Tree para tabela resumo
        frame_tabela = ttk.Frame(self.aba_resumo, padding=5)
        frame_tabela.pack(fill='both', expand=True, pady=5)
        
        colunas = ('ID', 'Fornecedor', 'Descrição', 'Valor Global', 'Valor Pago', 'Saldo', '% Executado', 'Status')
        self.tree_resumo = ttk.Treeview(frame_tabela, columns=colunas, show='headings', height=15)
        
        # Configurar colunas
        self.tree_resumo.heading('ID', text='ID')
        self.tree_resumo.heading('Fornecedor', text='Fornecedor')
        self.tree_resumo.heading('Descrição', text='Descrição')
        self.tree_resumo.heading('Valor Global', text='Valor Global')
        self.tree_resumo.heading('Valor Pago', text='Valor Pago')
        self.tree_resumo.heading('Saldo', text='Saldo')
        self.tree_resumo.heading('% Executado', text='% Executado')
        self.tree_resumo.heading('Status', text='Status')
        
        # Ajustar larguras das colunas
        self.tree_resumo.column('ID', width=40, anchor='center')
        self.tree_resumo.column('Fornecedor', width=150)
        self.tree_resumo.column('Descrição', width=200)
        self.tree_resumo.column('Valor Global', width=100, anchor='e')
        self.tree_resumo.column('Valor Pago', width=100, anchor='e')
        self.tree_resumo.column('Saldo', width=100, anchor='e')
        self.tree_resumo.column('% Executado', width=80, anchor='center')
        self.tree_resumo.column('Status', width=80, anchor='center')
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tree_resumo.yview)
        scrollx = ttk.Scrollbar(frame_tabela, orient='horizontal', command=self.tree_resumo.xview)
        self.tree_resumo.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_resumo.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
        # Binding para seleção
        self.tree_resumo.bind('<<TreeviewSelect>>', self.selecionar_contrato_resumo)
        
    def setup_aba_detalhes(self):
        """Configura a aba de detalhes com as medições"""
        # Frame para informações do contrato selecionado
        frame_contrato = ttk.LabelFrame(self.aba_detalhes, text="Contrato Selecionado")
        frame_contrato.pack(fill='x', pady=5)
        
        # Grid para informações
        frame_grid = ttk.Frame(frame_contrato, padding=10)
        frame_grid.pack(fill='x')
        
        # Configurar pesos das colunas para distribuição uniforme
        for col in [1, 3, 5]:
            frame_grid.columnconfigure(col, weight=1)
        
        # Primeira linha: ID | Fornecedor (span 2 colunas)
        ttk.Label(frame_grid, text="ID:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.lbl_id_contrato = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_id_contrato.grid(row=0, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Fornecedor:", font=('Arial', 10, 'bold')).grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.lbl_fornecedor = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_fornecedor.grid(row=0, column=3, columnspan=3, sticky='w', padx=5, pady=2)
        
        # Segunda linha: Descrição (span todas as colunas)
        ttk.Label(frame_grid, text="Descrição:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.lbl_descricao = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_descricao.grid(row=1, column=1, columnspan=5, sticky='w', padx=5, pady=2)
        
        # Terceira linha: Valor Global | Valor Pago | Saldo (3 colunas)
        ttk.Label(frame_grid, text="Valor Global:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.lbl_valor_global = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_valor_global.grid(row=2, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Valor Pago:", font=('Arial', 10, 'bold')).grid(row=2, column=2, sticky='w', padx=5, pady=2)
        self.lbl_valor_pago_contrato = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_valor_pago_contrato.grid(row=2, column=3, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Saldo:", font=('Arial', 10, 'bold')).grid(row=2, column=4, sticky='w', padx=5, pady=2)
        self.lbl_saldo_contrato = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_saldo_contrato.grid(row=2, column=5, sticky='w', padx=5, pady=2)
        
        # Quarta linha: Data Início | Data Final | Status (3 colunas)
        ttk.Label(frame_grid, text="Data Início:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky='w', padx=5, pady=2)
        self.lbl_data_inicio = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_data_inicio.grid(row=3, column=1, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Data Final:", font=('Arial', 10, 'bold')).grid(row=3, column=2, sticky='w', padx=5, pady=2)
        self.lbl_data_final = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_data_final.grid(row=3, column=3, sticky='w', padx=5, pady=2)
        
        ttk.Label(frame_grid, text="Status:", font=('Arial', 10, 'bold')).grid(row=3, column=4, sticky='w', padx=5, pady=2)
        self.lbl_status_contrato = ttk.Label(frame_grid, text="", font=('Arial', 10))
        self.lbl_status_contrato.grid(row=3, column=5, sticky='w', padx=5, pady=2)
        
        # Frame para tabela de medições
        frame_medicoes = ttk.LabelFrame(self.aba_detalhes, text="Medições")
        frame_medicoes.pack(fill='both', expand=True, pady=5)
        
        # Tree para medições
        colunas = ('ID', 'Data Medição', 'Data Pagamento', 'Referência', 'Valor', 'Status')
        self.tree_medicoes = ttk.Treeview(frame_medicoes, columns=colunas, show='headings', height=10)
        
        # Configurar colunas
        self.tree_medicoes.heading('ID', text='ID')
        self.tree_medicoes.heading('Data Medição', text='Data Medição')
        self.tree_medicoes.heading('Data Pagamento', text='Data Pagamento')
        self.tree_medicoes.heading('Referência', text='Referência')
        self.tree_medicoes.heading('Valor', text='Valor')
        self.tree_medicoes.heading('Status', text='Status')
        
        # Ajustar larguras das colunas
        self.tree_medicoes.column('ID', width=30, anchor='center')
        self.tree_medicoes.column('Data Medição', width=80, anchor='center')
        self.tree_medicoes.column('Data Pagamento', width=80, anchor='center')
        self.tree_medicoes.column('Referência', width=300)
        self.tree_medicoes.column('Valor', width=100, anchor='e')
        self.tree_medicoes.column('Status', width=80, anchor='center')
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_medicoes, orient='vertical', command=self.tree_medicoes.yview)
        scrollx = ttk.Scrollbar(frame_medicoes, orient='horizontal', command=self.tree_medicoes.xview)
        self.tree_medicoes.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_medicoes.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
    def setup_aba_grafico(self):
        """Configura a aba de gráficos"""
        # Frame para controles do gráfico
        frame_controles = ttk.Frame(self.aba_grafico, padding=5)
        frame_controles.pack(fill='x', pady=5)
        
        ttk.Label(frame_controles, text="Tipo de Gráfico:").pack(side='left', padx=5)
        self.combo_tipo_grafico = ttk.Combobox(frame_controles, values=[
            "Pizza - Valor por Contrato",
            "Barras - Valor Global vs. Pago",
            "Linha - Evolução de Pagamentos"
        ], state='readonly', width=30)
        self.combo_tipo_grafico.pack(side='left', padx=5)
        self.combo_tipo_grafico.current(0)
        
        ttk.Button(frame_controles, text="Atualizar Gráfico", command=self.atualizar_grafico).pack(side='left', padx=20)
        
        # Frame para o gráfico
        self.frame_grafico = ttk.Frame(self.aba_grafico)
        self.frame_grafico.pack(fill='both', expand=True, pady=5)
        
        # A figura será criada quando houver dados
    
    def atualizar_lista_clientes(self):
        """Atualiza a lista de clientes no combobox usando a função centralizada"""
        try:
            # Usar a função centralizada (apenas clientes ativos)
            self.info_clientes = atualizar_combobox_clientes(self.cliente_combobox, mostrar_inativos=False)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")

    # E modifique o método selecionar_cliente:

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
        if not self.cliente_atual:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
            
        # Obter data de referência
        try:
            self.data_referencia = datetime.strptime(self.data_entry.get(), '%d/%m/%Y')
            self.lbl_data_resumo.config(text=f"Data: {self.data_referencia.strftime('%d/%m/%Y')}")
        except ValueError:
            messagebox.showerror("Erro", "Data inválida!")
            return
            
        # Carregar contratos e medições
        self.carregar_dados()
        
        # Preencher treeview de resumo
        self.preencher_resumo()
        
        # Gerar gráfico inicial
        self.atualizar_grafico()
        
        # Selecionar aba de resumo
        self.notebook.select(0)
    
    def carregar_dados(self):
        """Carrega os dados dos contratos e medições do cliente"""
        try:
            if not os.path.exists(self.arquivo_cliente):
                messagebox.showerror("Erro", f"Arquivo do cliente '{self.cliente_atual}' não encontrado!")
                return False
                
            wb = load_workbook(self.arquivo_cliente)
            
            # Verificar se as abas necessárias existem
            if "Contratos_Medicao" not in wb.sheetnames:
                messagebox.showerror("Erro", "Aba de contratos não encontrada!")
                wb.close()
                return False
                
            if "Medicoes" not in wb.sheetnames:
                messagebox.showerror("Erro", "Aba de medições não encontrada!")
                wb.close()
                return False
                
            # Carregar contratos
            ws_contratos = wb["Contratos_Medicao"]
            self.contratos = []
            
            for row in ws_contratos.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Se tem ID do contrato
                    # Verificar se o contrato já existia na data de referência
                    data_inicio = row[4]
                    if isinstance(data_inicio, datetime) and data_inicio <= self.data_referencia:
                        self.contratos.append({
                            'id': row[0],
                            'cnpj': row[1],
                            'nome': row[2],
                            'descricao': row[3],
                            'data_inicio': row[4],
                            'data_final': row[5],  # NOVO: Data_Final - coluna 6
                            'valor_global': row[6] or 0,
                            'valor_pago': row[7] or 0,
                            'saldo': row[8] or 0,
                            'status': row[9] or 'ATIVO',
                            'observacao': row[10]
                        })
            
            # Carregar medições
            ws_medicoes = wb["Medicoes"]
            self.medicoes = []
            
            for row in ws_medicoes.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Se tem ID do contrato
                    # Verificar se a medição já existia na data de referência
                    data_medicao = row[4]
                    if isinstance(data_medicao, datetime) and data_medicao <= self.data_referencia:
                        self.medicoes.append({
                            'id_contrato': row[0],
                            'id_medicao': row[1],
                            'cnpj': row[2],
                            'nome': row[3],
                            'data_medicao': row[4],
                            'data_pagamento': row[5],
                            'referencia': row[6],
                            'valor': row[7] or 0,
                            'status': row[8] or 'PENDENTE',
                            'data_lancamento': row[9],
                            'observacao': row[10]
                        })
            
            wb.close()
            return True
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            try:
                wb.close()
            except:
                pass
            return False
    
    def preencher_resumo(self):
        """Preenche a tabela de resumo e os totais"""
        # Limpar treeview
        for item in self.tree_resumo.get_children():
            self.tree_resumo.delete(item)
            
        if not self.contratos:
            return
            
        # Variáveis para totais
        total_contratos = len(self.contratos)
        contratos_andamento = 0
        valor_total = 0
        valor_pago = 0
        
        # Preencher tabela
        for contrato in self.contratos:
            # Calcular percentual executado
            valor_global = float(contrato['valor_global']) if contrato['valor_global'] else 0
            valor_pago_contrato = float(contrato['valor_pago']) if contrato['valor_pago'] else 0
            
            if valor_global > 0:
                percentual = (valor_pago_contrato / valor_global) * 100
            else:
                percentual = 0
                
            # Verificar se está em andamento
            if contrato['status'] == 'ATIVO':
                contratos_andamento += 1
                
            # Acumular totais
            valor_total += valor_global
            valor_pago += valor_pago_contrato
            
            # Inserir na treeview
            self.tree_resumo.insert('', 'end', values=(
                contrato['id'],
                contrato['nome'],
                contrato['descricao'],
                formatar_moeda_br(valor_global),
                formatar_moeda_br(valor_pago_contrato),
                formatar_moeda_br(valor_global - valor_pago_contrato),
                f"{percentual:.1f}%",
                contrato['status']
            ))
            
        # Calcular saldo total
        saldo_total = valor_total - valor_pago
        percentual_total = (valor_pago / valor_total) * 100 if valor_total > 0 else 0
        
        # Atualizar labels de totais
        self.lbl_qtd_contratos.config(text=str(total_contratos))
        self.lbl_qtd_em_andamento.config(text=str(contratos_andamento))
        self.lbl_valor_total.config(text=formatar_moeda_br(valor_total))
        self.lbl_valor_pago.config(text=formatar_moeda_br(valor_pago))
        self.lbl_saldo.config(text=formatar_moeda_br(saldo_total))
        self.lbl_percentual.config(text=f"{percentual_total:.1f}%")
    
    def selecionar_contrato_resumo(self, event=None):
        """Quando um contrato é selecionado na aba resumo, atualiza a aba de detalhes"""
        selecionado = self.tree_resumo.selection()
        if not selecionado:
            return
            
        # Obter ID do contrato selecionado
        valores = self.tree_resumo.item(selecionado)['values']
        id_contrato = valores[0]
        
        # Buscar dados do contrato
        contrato = next((c for c in self.contratos if c['id'] == id_contrato), None)
        if not contrato:
            return
            
        # Atualizar informações do contrato na aba de detalhes
        self.lbl_id_contrato.config(text=str(contrato['id']))
        self.lbl_fornecedor.config(text=contrato['nome'])
        self.lbl_descricao.config(text=contrato['descricao'])
        self.lbl_valor_global.config(text=formatar_moeda_br(contrato['valor_global']))
        self.lbl_valor_pago_contrato.config(text=formatar_moeda_br(contrato['valor_pago']))
        self.lbl_saldo_contrato.config(text=formatar_moeda_br(contrato['valor_global'] - contrato['valor_pago']))
        
        # Atualizar datas
        if isinstance(contrato.get('data_inicio'), datetime):
            self.lbl_data_inicio.config(text=contrato['data_inicio'].strftime('%d/%m/%Y'))
        else:
            self.lbl_data_inicio.config(text=str(contrato.get('data_inicio') if contrato.get('data_inicio') else ""))
        
        if isinstance(contrato.get('data_final'), datetime):
            self.lbl_data_final.config(text=contrato['data_final'].strftime('%d/%m/%Y'))
        else:
            data_final_texto = str(contrato.get('data_final') if contrato.get('data_final') else "Não definida")
            self.lbl_data_final.config(text=data_final_texto)
        
        self.lbl_status_contrato.config(text=contrato['status'])
        
        # Limpar e preencher medições do contrato
        self.preencher_medicoes(id_contrato)
        
        # Mudar para a aba de detalhes
        self.notebook.select(1)
    
    def preencher_medicoes(self, id_contrato):
        """Preenche a tabela de medições para o contrato selecionado"""
        # Limpar treeview
        for item in self.tree_medicoes.get_children():
            self.tree_medicoes.delete(item)
            
        # Filtrar medições pelo contrato
        medicoes_contrato = [m for m in self.medicoes if m['id_contrato'] == id_contrato]
        
        if not medicoes_contrato:
            return
            
        # Ordenar por data de medição
        medicoes_contrato.sort(key=lambda x: x['data_medicao'] if isinstance(x['data_medicao'], datetime) else datetime.min)
        
        # Preencher tabela
        for medicao in medicoes_contrato:
            # Formatar datas
            if isinstance(medicao['data_medicao'], datetime):
                data_med = medicao['data_medicao'].strftime('%d/%m/%Y')
            else:
                data_med = str(medicao['data_medicao'] if medicao['data_medicao'] else "")
                
            if isinstance(medicao['data_pagamento'], datetime):
                data_pag = medicao['data_pagamento'].strftime('%d/%m/%Y')
            else:
                data_pag = str(medicao['data_pagamento'] if medicao['data_pagamento'] else "")
            
            # Combinar referência e observação
            referencia_completa = medicao['referencia'] or ""
            if medicao['observacao']:
                if referencia_completa:
                    referencia_completa += " - "
                referencia_completa += medicao['observacao']
            
            # Inserir na treeview
            self.tree_medicoes.insert('', 'end', values=(
                medicao['id_medicao'],
                data_med,
                data_pag,
                referencia_completa,
                formatar_moeda_br(medicao['valor']),
                medicao['status']
            ))
    
    def atualizar_grafico(self):
        """Atualiza o gráfico com base no tipo selecionado"""
        tipo_grafico = self.combo_tipo_grafico.get()
        
        # Limpar frame do gráfico
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
            
        if not self.contratos:
            return
            
        # Criar figura
        fig, ax = plt.subplots(figsize=(8, 6))
        
        if tipo_grafico == "Pizza - Valor por Contrato":
            self.criar_grafico_pizza(fig, ax)
        elif tipo_grafico == "Barras - Valor Global vs. Pago":
            self.criar_grafico_barras(fig, ax)
        elif tipo_grafico == "Linha - Evolução de Pagamentos":
            self.criar_grafico_linha(fig, ax)
            
        # Exibir o gráfico
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    
    def criar_grafico_pizza(self, fig, ax):
        """Cria um gráfico de pizza mostrando a distribuição dos valores dos contratos"""
        # Preparar dados
        labels = []
        valores = []
        
        # Limitar a 5 contratos maiores + "Outros"
        contratos_ordenados = sorted(self.contratos, key=lambda x: float(x['valor_global']), reverse=True)
        
        if len(contratos_ordenados) <= 5:
            for contrato in contratos_ordenados:
                labels.append(f"{contrato['id']}: {contrato['nome'][:20]}")
                valores.append(float(contrato['valor_global']))
        else:
            # 5 maiores
            for i in range(5):
                labels.append(f"{contratos_ordenados[i]['id']}: {contratos_ordenados[i]['nome'][:20]}")
                valores.append(float(contratos_ordenados[i]['valor_global']))
            
            # Resto agrupado como "Outros"
            valor_outros = sum(float(c['valor_global']) for c in contratos_ordenados[5:])
            labels.append("Outros")
            valores.append(valor_outros)
        
        # Criar gráfico
        wedges, texts, autotexts = ax.pie(
            valores, 
            labels=None,
            autopct='%1.1f%%',
            startangle=90,
            shadow=False
        )
        
        # Ajustar legenda
        ax.legend(wedges, labels, loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
        
        ax.set_title('Distribuição do Valor Total por Contrato')
        fig.tight_layout()
    
    def criar_grafico_barras(self, fig, ax):
        """Cria um gráfico de barras comparando valor global vs. valor pago"""
        # Preparar dados
        contratos_ids = []
        valores_globais = []
        valores_pagos = []
        
        # Limitar a 10 contratos
        contratos_exibir = self.contratos[:10] if len(self.contratos) > 10 else self.contratos
        
        for contrato in contratos_exibir:
            contratos_ids.append(str(contrato['id']))
            valores_globais.append(float(contrato['valor_global']))
            valores_pagos.append(float(contrato['valor_pago']))
        
        # Configurar barras
        x = range(len(contratos_ids))
        width = 0.35
        
        ax.bar([p - width/2 for p in x], valores_globais, width, label='Valor Global')
        ax.bar([p + width/2 for p in x], valores_pagos, width, label='Valor Pago')
        
        # Configurar gráfico
        ax.set_title('Valor Global vs. Valor Pago por Contrato')
        ax.set_xticks(x)
        ax.set_xticklabels(contratos_ids)
        ax.legend()
        
        # Adicionar valores nas barras
        for i, v in enumerate(valores_globais):
            ax.text(i - width/2, v + max(valores_globais) * 0.01, f"R${v:.1f}K", ha='center', va='bottom', rotation=90, fontsize=8)
            
        for i, v in enumerate(valores_pagos):
            if v > 0:
                ax.text(i + width/2, v + max(valores_globais) * 0.01, f"R${v:.1f}K", ha='center', va='bottom', rotation=90, fontsize=8)
        
        fig.tight_layout()
    
    def criar_grafico_linha(self, fig, ax):
        """Cria um gráfico de linha mostrando a evolução de pagamentos por data"""
        # Preparar dados - precisa agrupar medições por data
        datas = {}
        
        for medicao in self.medicoes:
            if isinstance(medicao['data_medicao'], datetime):
                data_key = medicao['data_medicao'].strftime('%Y-%m')
                valor = float(medicao['valor'])
                
                if data_key in datas:
                    datas[data_key] += valor
                else:
                    datas[data_key] = valor
        
        # Ordenar por data
        datas_ordenadas = sorted(datas.keys())
        valores_acumulados = []
        acumulado = 0
        
        for data in datas_ordenadas:
            acumulado += datas[data]
            valores_acumulados.append(acumulado)
        
        # Converter datas para formato de exibição
        datas_exibicao = [datetime.strptime(d, '%Y-%m').strftime('%m/%Y') for d in datas_ordenadas]
        
        # Criar gráfico
        ax.plot(datas_exibicao, valores_acumulados, 'o-', linewidth=2)
        
        # Adicionar pontos
        for i, (data, valor) in enumerate(zip(datas_exibicao, valores_acumulados)):
            ax.annotate(f"R${valor:.2f}K", 
                       (data, valor),
                       textcoords="offset points",
                       xytext=(0, 10),
                       ha='center')
        
        # Configurar gráfico
        ax.set_title('Evolução de Pagamentos Acumulados')
        ax.set_xlabel('Data')
        ax.set_ylabel('Valor Acumulado (R$)')
        
        # Rotacionar labels do eixo x para melhor visualização
        plt.xticks(rotation=45, ha='right')
        
        fig.tight_layout()
    
    def exportar_excel(self):
        """Exporta o relatório para um arquivo Excel"""
        if not self.contratos:
            messagebox.showwarning("Aviso", "Não há dados para exportar!")
            return
            
        # Solicitar nome do arquivo ao usuário
        data_str = self.data_referencia.strftime('%d-%m-%Y')
        nome_padrao = f"Relatorio_{self.cliente_atual}_{data_str}.xlsx"
        
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
            
            # Aba de resumo
            ws_resumo = wb.active
            ws_resumo.title = "Resumo Contratos"
            
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
            
            # Título
            ws_resumo['A1'] = f"Relatório de Contratos - {self.cliente_atual}"
            ws_resumo['A1'].font = titulo_font
            ws_resumo.merge_cells('A1:H1')
            
            ws_resumo['A2'] = f"Data de Referência: {self.data_referencia.strftime('%d/%m/%Y')}"
            ws_resumo.merge_cells('A2:H2')
            
            # Informações de totais
            ws_resumo['A4'] = "Total de Contratos:"
            ws_resumo['B4'] = self.lbl_qtd_contratos.cget("text")
            
            ws_resumo['D4'] = "Contratos em Andamento:"
            ws_resumo['E4'] = self.lbl_qtd_em_andamento.cget("text")
            
            ws_resumo['A5'] = "Valor Total dos Contratos:"
            ws_resumo['B5'] = self.lbl_valor_total.cget("text")
            
            ws_resumo['D5'] = "Valor já Pago:"
            ws_resumo['E5'] = self.lbl_valor_pago.cget("text")
            
            ws_resumo['A6'] = "Saldo a Pagar:"
            ws_resumo['B6'] = self.lbl_saldo.cget("text")
            
            ws_resumo['D6'] = "Percentual Executado:"
            ws_resumo['E6'] = self.lbl_percentual.cget("text")
            
            # Cabeçalho da tabela
            cabecalhos = ['ID', 'Fornecedor', 'Descrição', 'Valor Global', 'Valor Pago', 'Saldo', '% Executado', 'Status']
            for col, texto in enumerate(cabecalhos, 1):
                celula = ws_resumo.cell(row=8, column=col, value=texto)
                celula.font = cabecalho_font
                celula.fill = cabecalho_fill
                celula.border = borda
                celula.alignment = Alignment(horizontal='center')
            
            # Dados dos contratos
            linha = 9
            for contrato in self.contratos:
                valor_global = float(contrato['valor_global']) if contrato['valor_global'] else 0
                valor_pago = float(contrato['valor_pago']) if contrato['valor_pago'] else 0
                saldo = valor_global - valor_pago
                percentual = (valor_pago / valor_global) * 100 if valor_global > 0 else 0
                
                ws_resumo.cell(row=linha, column=1, value=contrato['id'])
                ws_resumo.cell(row=linha, column=2, value=contrato['nome'])
                ws_resumo.cell(row=linha, column=3, value=contrato['descricao'])
                ws_resumo.cell(row=linha, column=4, value=valor_global)
                ws_resumo.cell(row=linha, column=5, value=valor_pago)
                ws_resumo.cell(row=linha, column=6, value=saldo)
                ws_resumo.cell(row=linha, column=7, value=f"{percentual:.1f}%")
                ws_resumo.cell(row=linha, column=8, value=contrato['status'])
                
                # Formatar células de valor como moeda
                for col in [4, 5, 6]:
                    ws_resumo.cell(row=linha, column=col).number_format = '#,##0.00'
                
                linha += 1
            
            # Ajustar larguras das colunas
            for col in range(1, len(cabecalhos) + 1):
                ws_resumo.column_dimensions[get_column_letter(col)].width = 15
            ws_resumo.column_dimensions['B'].width = 25
            ws_resumo.column_dimensions['C'].width = 40
            
            # Adicionar aba para cada contrato
            for contrato in self.contratos:
                # Limitar o nome da aba para 31 caracteres (limite do Excel)
                nome_aba = f"Contrato {contrato['id']}"
                ws_contrato = wb.create_sheet(nome_aba)
                
                # Titulo
                ws_contrato['A1'] = f"Detalhamento do Contrato {contrato['id']}"
                ws_contrato['A1'].font = titulo_font
                ws_contrato.merge_cells('A1:F1')
                
                # Informações do contrato
                ws_contrato['A3'] = "ID do Contrato:"
                ws_contrato['B3'] = contrato['id']
                
                ws_contrato['A4'] = "Fornecedor:"
                ws_contrato['B4'] = contrato['nome']
                
                ws_contrato['A5'] = "Descrição:"
                ws_contrato['B5'] = contrato['descricao']
                ws_contrato.merge_cells('B5:F5')
                
                ws_contrato['A6'] = "Data de Início:"
                if isinstance(contrato['data_inicio'], datetime):
                    ws_contrato['B6'] = contrato['data_inicio'].strftime('%d/%m/%Y')
                else:
                    ws_contrato['B6'] = str(contrato['data_inicio'] if contrato['data_inicio'] else "")
                
                
                ws_contrato['D6'] = "Data Final:"
                if isinstance(contrato.get('data_final'), datetime):
                    ws_contrato['E6'] = contrato['data_final'].strftime('%d/%m/%Y')
                else:
                    ws_contrato['E6'] = str(contrato.get('data_final') if contrato.get('data_final') else "Não definida")
                
                ws_contrato['A7'] = "Valor Global:"
                ws_contrato['B7'] = float(contrato['valor_global']) if contrato['valor_global'] else 0
                ws_contrato['B7'].number_format = '#,##0.00'
                
                ws_contrato['D7'] = "Valor Pago:"
                ws_contrato['E7'] = float(contrato['valor_pago']) if contrato['valor_pago'] else 0
                ws_contrato['E7'].number_format = '#,##0.00'
                
                ws_contrato['A8'] = "Saldo:"
                saldo = float(contrato['valor_global'] or 0) - float(contrato['valor_pago'] or 0)
                ws_contrato['B8'] = saldo
                ws_contrato['B8'].number_format = '#,##0.00'
                
                ws_contrato['D8'] = "Status:"
                ws_contrato['E8'] = contrato['status']
                
                # Cabeçalho da tabela de medições
                ws_contrato['A10'] = "Medições do Contrato"
                ws_contrato['A10'].font = titulo_font
                ws_contrato.merge_cells('A10:F10')
                
                cabecalhos_medicoes = ['ID', 'Data Medição', 'Data Pagamento', 'Referência', 'Valor', 'Status']
                for col, texto in enumerate(cabecalhos_medicoes, 1):
                    celula = ws_contrato.cell(row=11, column=col, value=texto)
                    celula.font = cabecalho_font
                    celula.fill = cabecalho_fill
                    celula.border = borda
                    celula.alignment = Alignment(horizontal='center')
                
                # Filtrar medições deste contrato
                medicoes_contrato = [m for m in self.medicoes if m['id_contrato'] == contrato['id']]
                
                # Ordenar por data
                medicoes_contrato.sort(key=lambda x: x['data_medicao'] if isinstance(x['data_medicao'], datetime) else datetime.min)
                
                # Adicionar medições
                linha = 12
                for medicao in medicoes_contrato:
                    ws_contrato.cell(row=linha, column=1, value=medicao['id_medicao'])
                    
                    # Data medição
                    if isinstance(medicao['data_medicao'], datetime):
                        ws_contrato.cell(row=linha, column=2, value=medicao['data_medicao'])
                    else:
                        ws_contrato.cell(row=linha, column=2, value=str(medicao['data_medicao'] if medicao['data_medicao'] else ""))
                    
                    # Data pagamento
                    if isinstance(medicao['data_pagamento'], datetime):
                        ws_contrato.cell(row=linha, column=3, value=medicao['data_pagamento'])
                    else:
                        ws_contrato.cell(row=linha, column=3, value=str(medicao['data_pagamento'] if medicao['data_pagamento'] else ""))
                    
                    ws_contrato.cell(row=linha, column=4, value=medicao['referencia'])
                    ws_contrato.cell(row=linha, column=5, value=float(medicao['valor']) if medicao['valor'] else 0)
                    ws_contrato.cell(row=linha, column=6, value=medicao['status'])
                    
                    # Formatar células
                    for col in [2, 3]:
                        cell = ws_contrato.cell(row=linha, column=col)
                        if isinstance(cell.value, datetime):
                            cell.number_format = 'DD/MM/YYYY'
                    
                    ws_contrato.cell(row=linha, column=5).number_format = '#,##0.00'
                    
                    linha += 1
                
                # Ajustar larguras das colunas
                ws_contrato.column_dimensions['A'].width = 10
                ws_contrato.column_dimensions['B'].width = 15
                ws_contrato.column_dimensions['C'].width = 15
                ws_contrato.column_dimensions['D'].width = 40
                ws_contrato.column_dimensions['E'].width = 15
                ws_contrato.column_dimensions['F'].width = 12
            
            # Salvar o arquivo
            wb.save(arquivo)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {str(e)}")
    
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
    app = RelatorioContratos()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()