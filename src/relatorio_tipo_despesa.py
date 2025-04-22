import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
import pandas as pd
import numpy as np
from pathlib import Path
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.ticker as mticker

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
    def configurar_janela(janela, titulo="Janela", largura=800, altura=600):
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

# Funções auxiliares
def formatar_moeda_br(valor):
    """Formata um valor numérico como moeda brasileira"""
    try:
        valor_float = float(valor)
        return f"R$ {valor_float:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
    except (ValueError, TypeError):
        return f"R$ 0,00"

class RelatorioTipoDespesa:
    """Classe para geração de relatórios por tipo de despesa"""
    
    def __init__(self, parent=None):
        """Inicializa a interface do relatório"""
        self.parent = parent
        
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        configurar_janela(self.root, "Relatório por Tipo de Despesa", 900, 1000)
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.data_referencia = datetime.now()
        self.df_despesas = None
        self.df_tipos_despesa = None
        self.tipos_despesa = []
        self.tipo_despesa_selecionado = None
        self.dados_grafico = {}
        
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
        
        # Usar DateEntry se disponível, caso contrário usar Entry simples
        try:
            from tkcalendar import DateEntry
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
        except ImportError:
            self.data_var = tk.StringVar(value=datetime.now().strftime('%d/%m/%Y'))
            ttk.Entry(
                frame_data,
                textvariable=self.data_var,
                width=12,
                font=('Arial', 11)
            ).pack(side='left', padx=5)
        
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
            text="Exportar para PDF",
            command=self.exportar_pdf
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
        """Configura a aba de resumo do relatório de tipos de despesa"""
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
        
        # Frame para o TreeView com os tipos de despesa
        frame_resumo = ttk.Frame(self.aba_resumo, padding=5)
        frame_resumo.pack(fill='both', expand=True, pady=5)
        
        # Criar TreeView para os tipos de despesa
        colunas = ('tipo_despesa', 'total', 'percentual')
        self.tv_tipos_despesa = ttk.Treeview(frame_resumo, columns=colunas, show='headings', height=15)
        
        # Configurar cabeçalhos
        self.tv_tipos_despesa.heading('tipo_despesa', text='Tipo de Despesa')
        self.tv_tipos_despesa.heading('total', text='Total (R$)')
        self.tv_tipos_despesa.heading('percentual', text='% do Total')
        
        # Configurar colunas
        self.tv_tipos_despesa.column('tipo_despesa', width=300, anchor='w')
        self.tv_tipos_despesa.column('total', width=150, anchor='e')
        self.tv_tipos_despesa.column('percentual', width=100, anchor='e')
        
        # Configurar scrollbar para o TreeView
        scrollbar_y = ttk.Scrollbar(frame_resumo, orient='vertical', command=self.tv_tipos_despesa.yview)
        self.tv_tipos_despesa.configure(yscrollcommand=scrollbar_y.set)
        
        # Adicionar à tela
        self.tv_tipos_despesa.pack(side='left', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        
        # Adicionar evento de seleção
        self.tv_tipos_despesa.bind('<<TreeviewSelect>>', self.selecionar_tipo_despesa)
        
        # Frame para resumo de totais
        frame_totais = ttk.LabelFrame(self.aba_resumo, text="Resumo Financeiro", padding=10)
        frame_totais.pack(fill='x', pady=10, padx=10)
        
        # Adicionar labels para total geral
        ttk.Label(frame_totais, text="Total de Despesas:", font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.lbl_total_geral = ttk.Label(frame_totais, text="R$ 0,00", font=('Arial', 11))
        self.lbl_total_geral.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_totais, text="Número de Tipos de Despesa:", font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.lbl_num_tipos = ttk.Label(frame_totais, text="0", font=('Arial', 11))
        self.lbl_num_tipos.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_totais, text="Média por Tipo de Despesa:", font=('Arial', 11, 'bold')).grid(row=0, column=2, sticky='e', padx=5, pady=5)
        self.lbl_media_tipo = ttk.Label(frame_totais, text="R$ 0,00", font=('Arial', 11))
        self.lbl_media_tipo.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_totais, text="Tipo de Maior Valor:", font=('Arial', 11, 'bold')).grid(row=1, column=2, sticky='e', padx=5, pady=5)
        self.lbl_maior_tipo = ttk.Label(frame_totais, text="Nenhum", font=('Arial', 11))
        self.lbl_maior_tipo.grid(row=1, column=3, sticky='w', padx=5, pady=5)
    
    def setup_aba_detalhes(self):
        """Configura a aba de detalhes do relatório"""
        # Frame para informações do tipo de despesa selecionado
        frame_info_tipo = ttk.Frame(self.aba_detalhes, padding=5)
        frame_info_tipo.pack(fill='x', pady=5)
        
        self.lbl_tipo_detalhe = ttk.Label(
            frame_info_tipo, 
            text="Tipo de Despesa: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_tipo_detalhe.pack(side='left', padx=10)
        
        self.lbl_total_tipo_detalhe = ttk.Label(
            frame_info_tipo, 
            text="Total: R$ 0,00", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_total_tipo_detalhe.pack(side='left', padx=10)
        
        # Frame para a tabela de detalhes
        frame_tabela = ttk.Frame(self.aba_detalhes, padding=5)
        frame_tabela.pack(fill='both', expand=True, pady=5)
        
        # Criar TreeView para os lançamentos do tipo de despesa selecionado
        colunas = ('data', 'descricao', 'valor', 'observacao')
        self.tv_detalhes = ttk.Treeview(frame_tabela, columns=colunas, show='headings', height=15)
        
        # Configurar cabeçalhos
        self.tv_detalhes.heading('data', text='Data')
        self.tv_detalhes.heading('descricao', text='Descrição')
        self.tv_detalhes.heading('valor', text='Valor (R$)')
        self.tv_detalhes.heading('observacao', text='Observação')
        
        # Configurar colunas
        self.tv_detalhes.column('data', width=100, anchor='center')
        self.tv_detalhes.column('descricao', width=300, anchor='w')
        self.tv_detalhes.column('valor', width=120, anchor='e')
        self.tv_detalhes.column('observacao', width=300, anchor='w')
        
        # Configurar scrollbars
        scrollbar_y = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tv_detalhes.yview)
        scrollbar_x = ttk.Scrollbar(frame_tabela, orient='horizontal', command=self.tv_detalhes.xview)
        self.tv_detalhes.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Adicionar à tela
        self.tv_detalhes.pack(side='top', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        
        # Frame para estatísticas do tipo de despesa
        frame_stats = ttk.LabelFrame(self.aba_detalhes, text="Estatísticas", padding=10)
        frame_stats.pack(fill='x', pady=10, padx=10)
        
        # Adicionar labels para estatísticas
        ttk.Label(frame_stats, text="Número de Lançamentos:", font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.lbl_num_lancamentos = ttk.Label(frame_stats, text="0", font=('Arial', 11))
        self.lbl_num_lancamentos.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_stats, text="Média por Lançamento:", font=('Arial', 11, 'bold')).grid(row=0, column=2, sticky='e', padx=5, pady=5)
        self.lbl_media_lancamento = ttk.Label(frame_stats, text="R$ 0,00", font=('Arial', 11))
        self.lbl_media_lancamento.grid(row=0, column=3, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_stats, text="Maior Lançamento:", font=('Arial', 11, 'bold')).grid(row=1, column=0, sticky='e', padx=5, pady=5)
        self.lbl_maior_lancamento = ttk.Label(frame_stats, text="R$ 0,00", font=('Arial', 11))
        self.lbl_maior_lancamento.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        ttk.Label(frame_stats, text="Menor Lançamento:", font=('Arial', 11, 'bold')).grid(row=1, column=2, sticky='e', padx=5, pady=5)
        self.lbl_menor_lancamento = ttk.Label(frame_stats, text="R$ 0,00", font=('Arial', 11))
        self.lbl_menor_lancamento.grid(row=1, column=3, sticky='w', padx=5, pady=5)
    
    def setup_aba_grafico(self):
        """Configura a aba de gráficos"""
        # Frame para controles do gráfico
        frame_controles = ttk.Frame(self.aba_grafico, padding=5)
        frame_controles.pack(fill='x', pady=5)
        
        ttk.Label(frame_controles, text="Tipo de Gráfico:").pack(side='left', padx=5)
        self.combo_tipo_grafico = ttk.Combobox(frame_controles, values=[
            "Gráfico de Pizza",
            "Gráfico de Barras",
            "Gráfico de Linha (Evolução Mensal)"
        ], state='readonly', width=30)
        self.combo_tipo_grafico.pack(side='left', padx=5)
        self.combo_tipo_grafico.current(0)
        
        # Opções adicionais para gráfico
        self.var_mostrar_top = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_controles, 
            text="Mostrar apenas Top 10 tipos",
            variable=self.var_mostrar_top,
            command=self.atualizar_grafico
        ).pack(side='left', padx=20)
        
        ttk.Button(frame_controles, text="Atualizar Gráfico", command=self.atualizar_grafico).pack(side='left', padx=20)
        
        # Frame para o gráfico
        self.frame_grafico = ttk.Frame(self.aba_grafico)
        self.frame_grafico.pack(fill='both', expand=True, pady=5)
    
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
        self.cliente_atual = self.cliente_combobox.get()
        
        if self.cliente_atual:
            # Atualizar label (se já tiver sido criada)
            if hasattr(self, 'lbl_cliente_resumo'):
                self.lbl_cliente_resumo.config(text=f"Cliente: {self.cliente_atual}")
            
            # Definir o caminho do arquivo
            self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
    
    def gerar_relatorio(self):
        """Gera o relatório com base nos dados selecionados"""
        if not self.cliente_atual:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
            
        # Obter data de referência
        try:
            # Verificar se estamos usando DateEntry ou Entry
            if hasattr(self, 'data_entry'):
                self.data_referencia = self.data_entry.get_date()
            else:
                self.data_referencia = datetime.strptime(self.data_var.get(), '%d/%m/%Y')
                
            if hasattr(self, 'lbl_data_resumo'):
                self.lbl_data_resumo.config(text=f"Data: {self.data_referencia.strftime('%d/%m/%Y')}")
        except ValueError:
            messagebox.showerror("Erro", "Data inválida!")
            return
            
        # Carregar dados
        if not self.carregar_dados():
            return
        
        # Preencher resumo
        self.preencher_resumo()
        
        # Limpar detalhes (pois ainda não há tipo selecionado)
        if hasattr(self, 'tv_detalhes'):
            for item in self.tv_detalhes.get_children():
                self.tv_detalhes.delete(item)
        
        # Gerar gráfico inicial
        self.atualizar_grafico()
        
        # Selecionar aba de resumo
        self.notebook.select(0)  # Índice 0 corresponde à primeira aba (resumo)
    
    def carregar_dados(self):
        """Carrega os dados para o relatório a partir da aba 'Dados'"""
        try:
            if not os.path.exists(self.arquivo_cliente):
                messagebox.showerror("Erro", f"Arquivo do cliente '{self.cliente_atual}' não encontrado!")
                return False
            
            # Carregar dados do Excel - da aba 'Dados'
            try:
                # Usar pandas para ler a aba Dados
                self.df_despesas = pd.read_excel(self.arquivo_cliente, sheet_name='Dados')
                print(f"Colunas do DataFrame: {self.df_despesas.columns.tolist()}")
                
                # Verificar se as colunas necessárias existem (em maiúsculas)
                colunas_necessarias = ['TP_DESP', 'VALOR']
                for coluna in colunas_necessarias:
                    if coluna not in self.df_despesas.columns:
                        messagebox.showerror("Erro", f"A coluna '{coluna}' não foi encontrada na aba Dados!")
                        return False
                
                # Filtrar por data se necessário (verificar se há coluna de data)
                if 'DATA_REL' in self.df_despesas.columns:
                    # Converter para datetime (ignorando erros)
                    self.df_despesas['DATA_REL'] = pd.to_datetime(self.df_despesas['DATA_REL'], errors='coerce')
                    
                    # Ordenar os dados por data
                    self.df_despesas = self.df_despesas.sort_values(by='DATA_REL')
                
                # Garantir valores numéricos para a coluna valor
                self.df_despesas['VALOR'] = pd.to_numeric(self.df_despesas['VALOR'], errors='coerce').fillna(0)
                
                # Agrupar por tipo de despesa
                self.df_tipos_despesa = self.df_despesas.groupby('TP_DESP')['VALOR'].sum().reset_index()
                self.df_tipos_despesa = self.df_tipos_despesa.sort_values(by='VALOR', ascending=False)
                
                # Calcular percentual sobre o total
                total_geral = self.df_tipos_despesa['VALOR'].sum()
                if total_geral > 0:
                    self.df_tipos_despesa['percentual'] = (self.df_tipos_despesa['VALOR'] / total_geral) * 100
                else:
                    self.df_tipos_despesa['percentual'] = 0
                
                # Salvar lista de tipos de despesa para uso posterior
                self.tipos_despesa = self.df_tipos_despesa['TP_DESP'].tolist()
                
                # Se houver coluna de data, prepará-la para análise por período
                if 'DATA_REL' in self.df_despesas.columns:
                    # Criar coluna de mês/ano para análise temporal
                    self.df_despesas['periodo'] = self.df_despesas['DATA_REL'].dt.strftime('%m/%Y')
                    
                    # Criar DataFrame com evolução mensal por tipo de despesa
                    self.df_evolucao = self.df_despesas.pivot_table(
                        index='periodo',
                        columns='TP_DESP',
                        values='VALOR',
                        aggfunc='sum'
                    ).fillna(0)
                    
                    # Garantir que os períodos estão em ordem cronológica
                    try:
                        periodos = pd.to_datetime(self.df_evolucao.index, format='%m/%Y')
                        self.df_evolucao['__date'] = periodos
                        self.df_evolucao = self.df_evolucao.sort_values('__date')
                        self.df_evolucao = self.df_evolucao.drop('__date', axis=1)
                    except Exception as e:
                        print(f"Erro ao ordenar períodos: {e}")
                
                # Preparar dados para gráficos - incluindo evolução por período
                self.preparar_dados_grafico()
                
                return True
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados do Excel: {str(e)}")
                import traceback
                traceback.print_exc()
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            return False
    
    def preparar_dados_grafico(self):
        """Prepara os dados para os gráficos"""
        try:
            # Dados para gráfico de pizza e barras
            self.dados_grafico['pizza'] = self.df_tipos_despesa.copy()
            self.dados_grafico['barras'] = self.df_tipos_despesa.copy()
            
            # Para gráfico de linha (evolução mensal por tipo)
            # Verificar se temos coluna de data e df_evolucao
            if hasattr(self, 'df_evolucao') and not self.df_evolucao.empty:
                self.dados_grafico['linha'] = self.df_evolucao.copy()
            else:
                # Se não tivermos data, criar um dataframe vazio
                self.dados_grafico['linha'] = pd.DataFrame()
        
        except Exception as e:
            print(f"Erro ao preparar dados para gráfico: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def preencher_resumo(self):
        """Preenche os dados da aba de resumo"""
        try:
            # Limpar TreeView
            for item in self.tv_tipos_despesa.get_children():
                self.tv_tipos_despesa.delete(item)
            
            # Adicionar dados à TreeView
            for _, row in self.df_tipos_despesa.iterrows():
                self.tv_tipos_despesa.insert(
                    '', 'end', 
                    values=(
                        row['TP_DESP'],#######
                        formatar_moeda_br(row['VALOR']),
                        f"{row['percentual']:.2f}%"
                    )
                )
            
            # Atualizar labels de totais
            total_geral = self.df_tipos_despesa['VALOR'].sum()
            num_tipos = len(self.df_tipos_despesa)
            
            self.lbl_total_geral.config(text=formatar_moeda_br(total_geral))
            self.lbl_num_tipos.config(text=str(num_tipos))
            
            # Calcular média por tipo
            if num_tipos > 0:
                media_tipo = total_geral / num_tipos
                self.lbl_media_tipo.config(text=formatar_moeda_br(media_tipo))
            else:
                self.lbl_media_tipo.config(text="R$ 0,00")
            
            # Identificar tipo de maior valor
            if not self.df_tipos_despesa.empty:
                maior_tipo = self.df_tipos_despesa.iloc[0]['TP_DESP']
                self.lbl_maior_tipo.config(text=maior_tipo)
            else:
                self.lbl_maior_tipo.config(text="Nenhum")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher resumo: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def selecionar_tipo_despesa(self, event=None):
        """Atualiza o tipo de despesa selecionado e os detalhes"""
        try:
            # Obter seleção atual
            selecao = self.tv_tipos_despesa.selection()
            if not selecao:
                return
                
            # Obter tipo de despesa selecionado
            item = self.tv_tipos_despesa.item(selecao[0])
            self.tipo_despesa_selecionado = item['values'][0]  # Primeira coluna é o tipo
            
            # Atualizar label na aba de detalhes
            self.lbl_tipo_detalhe.config(text=f"Tipo de Despesa: {self.tipo_despesa_selecionado}")
            
            # Filtrar dados para o tipo selecionado
            df_tipo = self.df_despesas[self.df_despesas['TP_DESP'] == self.tipo_despesa_selecionado].copy()
            
            # Ordenar por data se disponível
            if 'data' in df_tipo.columns:
                df_tipo = df_tipo.sort_values(by='data', ascending=False)
            
            # Calcular total
            total_tipo = df_tipo['VALOR'].sum()
            self.lbl_total_tipo_detalhe.config(text=f"Total: {formatar_moeda_br(total_tipo)}")
            
            # Preencher tabela de detalhes
            self.preencher_detalhes(df_tipo)
            
            # Mudar para aba de detalhes
            self.notebook.select(1)  # Índice 1 corresponde à aba de detalhes
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar tipo de despesa: {str(e)}")
    
    def preencher_detalhes(self, df_filtrado):
        """Preenche os detalhes para o tipo de despesa selecionado"""
        try:
            # Limpar tabela
            for item in self.tv_detalhes.get_children():
                self.tv_detalhes.delete(item)
            
            # Verificar se o DataFrame está vazio
            if df_filtrado.empty:
                return
            
            # Adicionar dados à tabela
            for _, row in df_filtrado.iterrows():
                # Formatar data se disponível
                if 'data' in row and pd.notna(row['data']):
                    data_str = row['data'].strftime('%d/%m/%Y')
                else:
                    data_str = ''
                
                # Obter descrição e valor
                descricao = row.get('descricao', '') if pd.notna(row.get('descricao', '')) else ''
                valor = formatar_moeda_br(row['VALOR'])
                
                # Obter observação se disponível
                observacao = row.get('observacao', '') if pd.notna(row.get('observacao', '')) else ''
                
                # Inserir na tabela
                self.tv_detalhes.insert(
                    '', 'end', 
                    values=(
                        data_str,
                        descricao,
                        valor,
                        observacao
                    )
                )
            
            # Atualizar estatísticas
            num_lancamentos = len(df_filtrado)
            total_tipo = df_filtrado['VALOR'].sum()
            
            self.lbl_num_lancamentos.config(text=str(num_lancamentos))
            
            # Média por lançamento
            if num_lancamentos > 0:
                media_lancamento = total_tipo / num_lancamentos
                self.lbl_media_lancamento.config(text=formatar_moeda_br(media_lancamento))
            else:
                self.lbl_media_lancamento.config(text="R$ 0,00")
            
            # Maior e menor lançamento
            if not df_filtrado.empty:
                maior_valor = df_filtrado['VALOR'].max()
                menor_valor = df_filtrado['VALOR'].min()
                
                self.lbl_maior_lancamento.config(text=formatar_moeda_br(maior_valor))
                self.lbl_menor_lancamento.config(text=formatar_moeda_br(menor_valor))
            else:
                self.lbl_maior_lancamento.config(text="R$ 0,00")
                self.lbl_menor_lancamento.config(text="R$ 0,00")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher detalhes: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def atualizar_grafico(self, event=None):
        """Atualiza o gráfico com base no tipo selecionado"""
        try:
            tipo_grafico = self.combo_tipo_grafico.get()
            
            # Limpar frame do gráfico
            for widget in self.frame_grafico.winfo_children():
                widget.destroy()
                
            # Verificar se há dados para gerar o gráfico
            if not hasattr(self, 'dados_grafico') or not self.dados_grafico:
                return
                
            # Criar figura
            fig, ax = plt.subplots(figsize=(10, 6))
            
            if tipo_grafico == "Gráfico de Pizza":
                self.criar_grafico_pizza(fig, ax)
            elif tipo_grafico == "Gráfico de Barras":
                self.criar_grafico_barras(fig, ax)
            elif tipo_grafico == "Gráfico de Linha (Evolução Mensal)":
                self.criar_grafico_linha(fig, ax)
                
            # Exibir o gráfico
            canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar gráfico: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def atualizar_grafico(self):
        """Atualiza o gráfico com base no tipo selecionado"""
        tipo_grafico = self.combo_tipo_grafico.get()
        
        # Limpar frame do gráfico
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
            
        # Verificar se há dados para gerar o gráfico
        if not hasattr(self, 'dados_grafico') or not self.dados_grafico:
            return
            
        # Criar figura
        fig, ax = plt.subplots(figsize=(8, 6))
        
        if tipo_grafico == "Gráfico de Pizza":
            self.criar_grafico_pizza(fig, ax)
        elif tipo_grafico == "Gráfico de Barras":
            self.criar_grafico_barras(fig, ax)
        elif tipo_grafico == "Gráfico de Linha":
            self.criar_grafico_linha(fig, ax)
            
        # Exibir o gráfico
        canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
        canvas.draw()
        canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
    
    def criar_grafico_pizza(self, fig, ax):
        """Cria um gráfico de pizza com os tipos de despesa"""
        try:
            # Usar os dados para gráfico de pizza
            df = self.dados_grafico.get('pizza', pd.DataFrame())
            
            if df.empty:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Verificar se deve mostrar apenas o top 10
            if self.var_mostrar_top.get() and len(df) > 10:
                # Usar os 10 maiores tipos e agrupar o resto como "Outros"
                top_df = df.head(10).copy()
                outros_valor = df.iloc[10:]['VALOR'].sum()
                outros_percentual = df.iloc[10:]['percentual'].sum()
                
                # Adicionar linha para "Outros"
                outros_row = pd.DataFrame({
                    'TP_DESP': ['Outros'],
                    'VALOR': [outros_valor],
                    'percentual': [outros_percentual]
                })
                
                df = pd.concat([top_df, outros_row], ignore_index=True)
            
            # Cores para o gráfico
            colors = plt.cm.tab20.colors
            
            # Criar o gráfico de pizza
            wedges, texts, autotexts = ax.pie(
                df['VALOR'], 
                labels=df['TP_DESP'], 
                autopct='%1.1f%%',
                startangle=90,
                colors=colors,
                wedgeprops={'edgecolor': 'w', 'linewidth': 1}
            )
            
            # Melhorar aparência
            for text in texts:
                text.set_fontsize(9)
            
            for autotext in autotexts:
                autotext.set_fontsize(9)
                autotext.set_fontweight('bold')
            
            # Adicionar título
            ax.set_title(f'Distribuição por Tipo de Despesa ({self.cliente_atual})', fontsize=14, pad=20)
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de pizza: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Mostrar erro no gráfico
            ax.text(0.5, 0.5, f"Erro ao gerar gráfico: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12, color='red')
    
    def criar_grafico_barras(self, fig, ax):
        """Cria um gráfico de barras com os tipos de despesa"""
        try:
            # Usar os dados para gráfico de barras
            df = self.dados_grafico.get('barras', pd.DataFrame())
            
            if df.empty:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Verificar se deve mostrar apenas o top 10
            if self.var_mostrar_top.get() and len(df) > 10:
                df = df.head(10)
            
            # Inverter a ordem para o gráfico de barras (do menor para o maior)
            df = df.sort_values(by='VALOR', ascending=True)
            
            # Cores para o gráfico
            colors = plt.cm.tab20.colors[:len(df)]
            
            # Criar o gráfico de barras
            bars = ax.barh(df['TP_DESP'], df['VALOR'], color=colors)
            
            # Adicionar valores nas barras
            for bar in bars:
                width = bar.get_width()
                label_x_pos = width + width * 0.01
                ax.text(label_x_pos, bar.get_y() + bar.get_height()/2, f'R$ {width:,.2f}'.replace(',', '.'),
                       va='center', fontsize=9)
            
            # Ajustar formatação do eixo x (valores)
            def format_real(x, pos):
                return f'R$ {x:,.0f}'.replace(',', '.')
            
            ax.xaxis.set_major_formatter(mticker.FuncFormatter(format_real))
            
            # Adicionar títulos e labels
            ax.set_title(f'Tipos de Despesa por Valor Total ({self.cliente_atual})', fontsize=14)
            ax.set_xlabel('Valor (R$)', fontsize=11)
            ax.set_ylabel('Tipo de Despesa', fontsize=11)
            
            # Adicionar grid
            ax.grid(axis='x', linestyle='--', alpha=0.7)
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de barras: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Mostrar erro no gráfico
            ax.text(0.5, 0.5, f"Erro ao gerar gráfico: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12, color='red')
    
    def criar_grafico_linha(self, fig, ax):
        """Cria um gráfico de linha com a evolução mensal por tipo de despesa"""
        try:
            # Usar os dados para gráfico de linha
            df = self.dados_grafico.get('linha', pd.DataFrame())
            
            if df.empty:
                ax.text(0.5, 0.5, "Não há dados suficientes para exibir evolução mensal", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Verificar se temos o tipo selecionado
            if self.tipo_despesa_selecionado and self.tipo_despesa_selecionado in df.columns:
                # Plotar apenas o tipo selecionado
                df = df[[self.tipo_despesa_selecionado]]
                titulo = f'Evolução Mensal: {self.tipo_despesa_selecionado}'
            else:
                # Selecionar os 5 maiores tipos para plotar
                top_tipos = self.df_tipos_despesa.head(5)['TP_DESP'].tolist()
                tipos_presentes = [t for t in top_tipos if t in df.columns]
                
                if not tipos_presentes:
                    ax.text(0.5, 0.5, "Não há dados suficientes para tipos selecionados", 
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, fontsize=14)
                    return
                
                df = df[tipos_presentes]
                titulo = 'Evolução Mensal dos 5 Principais Tipos de Despesa'
            
            # Plotar cada tipo como uma linha
            for coluna in df.columns:
                ax.plot(df.index, df[coluna], marker='o', linewidth=2, label=coluna)
            
            # Formatação
            ax.set_title(titulo, fontsize=14)
            ax.set_xlabel('Mês/Ano', fontsize=11)
            ax.set_ylabel('Valor (R$)', fontsize=11)
            
            # Formatação do eixo y (valores)
            def format_real(x, pos):
                return f'R$ {x:,.0f}'.replace(',', '.')
            
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(format_real))
            
            # Rotacionar labels do eixo x para melhor visualização
            plt.xticks(rotation=45)
            
            # Adicionar grid e legenda
            ax.grid(linestyle='--', alpha=0.7)
            ax.legend(fontsize=9, loc='best')
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de linha: {str(e)}")
            import traceback
            traceback.print_exc()
            
            # Mostrar erro no gráfico
            ax.text(0.5, 0.5, f"Erro ao gerar gráfico: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12, color='red')
    
    def exportar_excel(self):
        """Exporta o relatório para um arquivo Excel"""
        if not hasattr(self, 'cliente_atual') or not self.cliente_atual:
            messagebox.showwarning("Aviso", "Não há dados para exportar!")
            return
            
        # Solicitar nome do arquivo ao usuário
        data_str = self.data_referencia.strftime('%d-%m-%Y')
        nome_padrao = f"Relatorio_{self.__class__.__name__}_{self.cliente_atual}_{data_str}.xlsx"
        
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
            
            # Remover a aba padrão
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            
            # Criar aba de resumo
            ws_resumo = wb.create_sheet("Resumo")
            
            # Adicionar cabeçalho
            ws_resumo['A1'] = "Relatório por Tipo de Despesa"
            ws_resumo['A1'].font = Font(size=14, bold=True)
            ws_resumo.merge_cells('A1:D1')
            ws_resumo['A1'].alignment = Alignment(horizontal='center')
            
            ws_resumo['A2'] = f"Cliente: {self.cliente_atual}"
            ws_resumo['A2'].font = Font(size=12, bold=True)
            ws_resumo.merge_cells('A2:D2')
            
            ws_resumo['A3'] = f"Data do relatório: {data_str}"
            ws_resumo['A3'].font = Font(size=12)
            ws_resumo.merge_cells('A3:D3')
            
            # Adicionar cabeçalhos da tabela
            headers = ["Tipo de Despesa", "Valor (R$)", "% do Total", "Posição"]
            for col, header in enumerate(headers, start=1):
                cell = ws_resumo.cell(row=5, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.fill = PatternFill(fgColor="DDDDDD", fill_type="solid")
            
            # Adicionar dados
            for i, (_, row) in enumerate(self.df_tipos_despesa.iterrows(), start=6):
                ws_resumo.cell(row=i, column=1, value=row['TP_DESP'])
                
                # Formatar valor como moeda
                ws_resumo.cell(row=i, column=2, value=row['VALOR'])
                ws_resumo.cell(row=i, column=2).number_format = "R$ #,##0.00"
                
                # Percentual
                ws_resumo.cell(row=i, column=3, value=row['percentual'] / 100)  # Dividir por 100 para formato percentual
                ws_resumo.cell(row=i, column=3).number_format = "0.00%"
                
                # Posição
                ws_resumo.cell(row=i, column=4, value=i-5)
            
            # Ajustar largura das colunas
            ws_resumo.column_dimensions['A'].width = 40
            ws_resumo.column_dimensions['B'].width = 15
            ws_resumo.column_dimensions['C'].width = 15
            ws_resumo.column_dimensions['D'].width = 10
            
            # Adicionar totais
            total_row = 6 + len(self.df_tipos_despesa)
            
            ws_resumo.cell(row=total_row, column=1, value="TOTAL")
            ws_resumo.cell(row=total_row, column=1).font = Font(bold=True)
            
            # Total em R$
            total_formula = f"=SUM(B6:B{total_row-1})"
            ws_resumo.cell(row=total_row, column=2, value=total_formula)
            ws_resumo.cell(row=total_row, column=2).font = Font(bold=True)
            ws_resumo.cell(row=total_row, column=2).number_format = "R$ #,##0.00"
            
            # Criar aba de detalhes se tivermos um tipo selecionado
            if hasattr(self, 'tipo_despesa_selecionado') and self.tipo_despesa_selecionado:
                ws_detalhes = wb.create_sheet("Detalhes")
                
                # Adicionar cabeçalho
                ws_detalhes['A1'] = f"Detalhes do Tipo de Despesa: {self.tipo_despesa_selecionado}"
                ws_detalhes['A1'].font = Font(size=14, bold=True)
                ws_detalhes.merge_cells('A1:D1')
                ws_detalhes['A1'].alignment = Alignment(horizontal='center')
                
                # Filtrando dados para o tipo selecionado
                df_filtrado = self.df_despesas[self.df_despesas['TP_DESP'] == self.tipo_despesa_selecionado].copy()
                
                # Ordenar por data se disponível
                if 'data' in df_filtrado.columns:
                    df_filtrado = df_filtrado.sort_values(by='data', ascending=False)
                
                # Adicionar cabeçalhos da tabela
                headers = ["Data", "Descrição", "Valor (R$)", "Observação"]
                for col, header in enumerate(headers, start=1):
                    cell = ws_detalhes.cell(row=3, column=col, value=header)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                    cell.fill = PatternFill(fgColor="DDDDDD", fill_type="solid")
                
                # Adicionar dados
                for i, (_, row) in enumerate(df_filtrado.iterrows(), start=4):
                    # Data formatada
                    if 'data' in row and pd.notna(row['data']):
                        ws_detalhes.cell(row=i, column=1, value=row['data'])
                        ws_detalhes.cell(row=i, column=1).number_format = "dd/mm/yyyy"
                    
                    # Descrição
                    if 'descricao' in row and pd.notna(row['descricao']):
                        ws_detalhes.cell(row=i, column=2, value=row['descricao'])
                    
                    # Valor
                    if 'VALOR' in row:
                        ws_detalhes.cell(row=i, column=3, value=row['VALOR'])
                        ws_detalhes.cell(row=i, column=3).number_format = "R$ #,##0.00"
                    
                    # Observação
                    if 'observacao' in row and pd.notna(row['observacao']):
                        ws_detalhes.cell(row=i, column=4, value=row['observacao'])
                
                # Ajustar largura das colunas
                ws_detalhes.column_dimensions['A'].width = 15
                ws_detalhes.column_dimensions['B'].width = 40
                ws_detalhes.column_dimensions['C'].width = 15
                ws_detalhes.column_dimensions['D'].width = 40
                
                # Adicionar total
                total_row = 4 + len(df_filtrado)
                
                ws_detalhes.cell(row=total_row, column=2, value="TOTAL")
                ws_detalhes.cell(row=total_row, column=2).font = Font(bold=True)
                
                # Total em R$
                total_formula = f"=SUM(C4:C{total_row-1})"
                ws_detalhes.cell(row=total_row, column=3, value=total_formula)
                ws_detalhes.cell(row=total_row, column=3).font = Font(bold=True)
                ws_detalhes.cell(row=total_row, column=3).number_format = "R$ #,##0.00"
            
            # Criar aba para todos os dados
            ws_dados = wb.create_sheet("Todos os Dados")
            
            # Adicionar cabeçalho
            ws_dados['A1'] = "Todos os Lançamentos por Tipo de Despesa"
            ws_dados['A1'].font = Font(size=14, bold=True)
            ws_dados.merge_cells('A1:E1')
            ws_dados['A1'].alignment = Alignment(horizontal='center')
            
            # Adicionar cabeçalhos da tabela
            headers = ["Data", "Tipo de Despesa", "Descrição", "Valor (R$)", "Observação"]
            for col, header in enumerate(headers, start=1):
                cell = ws_dados.cell(row=3, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.fill = PatternFill(fgColor="DDDDDD", fill_type="solid")
            
            # Ordenar todos os dados por data e tipo
            df_todos = self.df_despesas.copy()
            
            if 'data' in df_todos.columns:
                df_todos = df_todos.sort_values(by=['data', 'TP_DESP'], ascending=[False, True])
            else:
                df_todos = df_todos.sort_values(by='TP_DESP')
            
            # Adicionar todos os dados
            for i, (_, row) in enumerate(df_todos.iterrows(), start=4):
                # Data formatada
                if 'data' in row and pd.notna(row['data']):
                    ws_dados.cell(row=i, column=1, value=row['data'])
                    ws_dados.cell(row=i, column=1).number_format = "dd/mm/yyyy"
                
                # Tipo de Despesa
                ws_dados.cell(row=i, column=2, value=row['TP_DESP'])
                
                # Descrição
                if 'descricao' in row and pd.notna(row['descricao']):
                    ws_dados.cell(row=i, column=3, value=row['descricao'])
                
                # Valor
                ws_dados.cell(row=i, column=4, value=row['VALOR'])
                ws_dados.cell(row=i, column=4).number_format = "R$ #,##0.00"
                
                # Observação
                if 'observacao' in row and pd.notna(row['observacao']):
                    ws_dados.cell(row=i, column=5, value=row['observacao'])
            
            # Ajustar largura das colunas
            ws_dados.column_dimensions['A'].width = 15
            ws_dados.column_dimensions['B'].width = 30
            ws_dados.column_dimensions['C'].width = 40
            ws_dados.column_dimensions['D'].width = 15
            ws_dados.column_dimensions['E'].width = 40
            
            # Adicionar total
            total_row = 4 + len(df_todos)
            
            ws_dados.cell(row=total_row, column=3, value="TOTAL")
            ws_dados.cell(row=total_row, column=3).font = Font(bold=True)
            
            # Total em R$
            total_formula = f"=SUM(D4:D{total_row-1})"
            ws_dados.cell(row=total_row, column=4, value=total_formula)
            ws_dados.cell(row=total_row, column=4).font = Font(bold=True)
            ws_dados.cell(row=total_row, column=4).number_format = "R$ #,##0.00"
            
            # Salvar o arquivo
            wb.save(arquivo)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{arquivo}")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao exportar para Excel: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def exportar_pdf(self):
        """Exporta o relatório para um arquivo PDF"""
        if not hasattr(self, 'cliente_atual') or not self.cliente_atual:
            messagebox.showwarning("Aviso", "Não há dados para exportar!")
            return
            
        # Solicitar nome do arquivo ao usuário
        data_str = self.data_referencia.strftime('%d-%m-%Y')
        nome_padrao = f"Relatorio_{self.__class__.__name__}_{self.cliente_atual}_{data_str}.pdf"
        
        arquivo = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Arquivos PDF", "*.pdf")],
            initialfile=nome_padrao
        )
        
        if not arquivo:
            return
            
        try:
            # Verificar se temos a biblioteca reportlab disponível
            try:
                from reportlab.lib.pagesizes import letter, A4
                from reportlab.lib import colors
                from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
                from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
                from reportlab.lib.units import inch
                import matplotlib.pyplot as plt
                import io
            except ImportError:
                messagebox.showerror("Erro", "Biblioteca ReportLab não encontrada. Por favor instale usando 'pip install reportlab'.")
                return
            
            # Criar documento PDF
            doc = SimpleDocTemplate(arquivo, pagesize=A4)
            story = []
            
            # Estilos
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading1_style = styles['Heading1']
            heading2_style = styles['Heading2']
            normal_style = styles['Normal']
            
            # Título
            story.append(Paragraph(f"Relatório por Tipo de Despesa", title_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Informações do cliente
            story.append(Paragraph(f"Cliente: {self.cliente_atual}", heading1_style))
            story.append(Paragraph(f"Data do relatório: {data_str}", normal_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Resumo de valores
            total_geral = self.df_tipos_despesa['VALOR'].sum()
            num_tipos = len(self.df_tipos_despesa)
            
            story.append(Paragraph("Resumo Financeiro", heading2_style))
            story.append(Spacer(1, 0.1*inch))
            
            resumo_data = [
                ["Total de Despesas:", f"R$ {total_geral:,.2f}".replace(',', '.').replace('.', ',')],
                ["Número de Tipos de Despesa:", f"{num_tipos}"],
            ]
            
            if num_tipos > 0:
                media_tipo = total_geral / num_tipos
                maior_tipo = self.df_tipos_despesa.iloc[0]['TP_DESP'] if not self.df_tipos_despesa.empty else "Nenhum"
                
                resumo_data.append(["Média por Tipo de Despesa:", f"R$ {media_tipo:,.2f}".replace(',', '.').replace('.', ',')])
                resumo_data.append(["Tipo de Maior Valor:", maior_tipo])
            
            # Criar tabela de resumo
            resumo_table = Table(resumo_data, colWidths=[3*inch, 3*inch])
            resumo_table.setStyle(TableStyle([
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                ('BACKGROUND', (0, 0), (0, -1), colors.lightgrey),
                ('ALIGN', (0, 0), (0, -1), 'RIGHT'),
                ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                ('FONTNAME', (0, 0), (-1, -1), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
            ]))
            
            story.append(resumo_table)
            story.append(Spacer(1, 0.2*inch))
            
            # Tabela com os tipos de despesa
            story.append(Paragraph("Tipos de Despesa", heading2_style))
            story.append(Spacer(1, 0.1*inch))
            
            # Cabeçalhos
            headers = ["Tipo de Despesa", "Valor (R$)", "% do Total"]
            
            # Dados da tabela
            table_data = [headers]
            
            for _, row in self.df_tipos_despesa.iterrows():
                tipo = row['TP_DESP']
                valor = f"R$ {row['VALOR']:,.2f}".replace(',', '.').replace('.', ',')
                percentual = f"{row['percentual']:.2f}%"
                
                table_data.append([tipo, valor, percentual])
            
            # Adicionar linha de total
            total_valor = f"R$ {total_geral:,.2f}".replace(',', '.').replace('.', ',')
            table_data.append(["TOTAL", total_valor, "100.00%"])
            
            # Criar tabela
            col_widths = [4*inch, 2*inch, 1.5*inch]
            tipos_table = Table(table_data, colWidths=col_widths)
            
            # Estilo da tabela
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (2, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ])
            
            # Destacar a linha de total
            table_style.add('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
            table_style.add('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
            
            tipos_table.setStyle(table_style)
            story.append(tipos_table)
            story.append(Spacer(1, 0.2*inch))
            
            # Adicionar gráfico de pizza
            story.append(Paragraph("Gráfico de Distribuição", heading2_style))
            story.append(Spacer(1, 0.1*inch))
            
            # Gerar gráfico para incluir no PDF
            plt.figure(figsize=(8, 6))
            
            # Verificar se deve mostrar apenas o top 10
            df = self.df_tipos_despesa.copy()
            if len(df) > 10:
                # Usar os 10 maiores tipos e agrupar o resto como "Outros"
                top_df = df.head(10).copy()
                outros_valor = df.iloc[10:]['VALOR'].sum()
                outros_percentual = df.iloc[10:]['percentual'].sum()
                
                # Adicionar linha para "Outros"
                outros_row = pd.DataFrame({
                    'TP_DESP': ['Outros'],
                    'VALOR': [outros_valor],
                    'percentual': [outros_percentual]
                })
                
                df = pd.concat([top_df, outros_row], ignore_index=True)
            
            # Criar o gráfico de pizza
            plt.pie(
                df['VALOR'], 
                labels=df['TP_DESP'], 
                autopct='%1.1f%%',
                startangle=90,
                colors=plt.cm.tab20.colors,
                wedgeprops={'edgecolor': 'w', 'linewidth': 1}
            )
            
            plt.title(f'Distribuição por Tipo de Despesa ({self.cliente_atual})', fontsize=14, pad=20)
            plt.tight_layout()
            
            # Salvar o gráfico em um buffer
            img_buffer = io.BytesIO()
            plt.savefig(img_buffer, format='png', dpi=100)
            img_buffer.seek(0)
            plt.close()
            
            # Adicionar o gráfico ao PDF
            img = Image(img_buffer, width=6*inch, height=4*inch)
            story.append(img)
            story.append(Spacer(1, 0.2*inch))
            
            # Adicionar detalhes do tipo selecionado se houver
            if hasattr(self, 'tipo_despesa_selecionado') and self.tipo_despesa_selecionado:
                story.append(Paragraph(f"Detalhes: {self.tipo_despesa_selecionado}", heading2_style))
                story.append(Spacer(1, 0.1*inch))
                
                # Filtrar dados para o tipo selecionado
                df_filtrado = self.df_despesas[self.df_despesas['TP_DESP'] == self.tipo_despesa_selecionado].copy()
                
                # Verificar se há dados
                if not df_filtrado.empty:
                    # Ordenar por data se disponível
                    if 'data' in df_filtrado.columns:
                        df_filtrado = df_filtrado.sort_values(by='data', ascending=False)
                    
                    # Cabeçalhos e dados para a tabela
                    headers = ["Data", "Descrição", "Valor (R$)"]
                    table_data = [headers]
                    
                    for _, row in df_filtrado.iterrows():
                        # Formatar data se disponível
                        data_str = ''
                        if 'data' in row and pd.notna(row['data']):
                            data_str = row['data'].strftime('%d/%m/%Y')
                        
                        # Obter descrição e valor
                        descricao = row.get('descricao', '') if pd.notna(row.get('descricao', '')) else ''
                        valor = f"R$ {row['VALOR']:,.2f}".replace(',', '.').replace('.', ',')
                        
                        table_data.append([data_str, descricao, valor])
                    
                    # Adicionar total
                    total_tipo = df_filtrado['VALOR'].sum()
                    table_data.append(["", "TOTAL", f"R$ {total_tipo:,.2f}".replace(',', '.').replace('.', ',')])
                    
                    # Criar tabela
                    col_widths = [1.5*inch, 4*inch, 2*inch]
                    detalhes_table = Table(table_data, colWidths=col_widths)
                    
                    # Estilo da tabela
                    table_style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                        ('ALIGN', (1, 0), (1, -1), 'LEFT'),
                        ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, -1), 9),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ])
                    
                    # Destacar a linha de total
                    table_style.add('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
                    table_style.add('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
                    
                    detalhes_table.setStyle(table_style)
                    story.append(detalhes_table)
                else:
                    story.append(Paragraph("Não há lançamentos para este tipo de despesa.", normal_style))
            
            # Construir o PDF
            doc.build(story)
            messagebox.showinfo("Sucesso", f"Relatório exportado com sucesso para:\n{arquivo}")
            
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
    app = RelatorioTipoDespesa()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()
