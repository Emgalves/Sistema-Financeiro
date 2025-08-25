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

class RelatorioCategoria:
    """Classe para geração de relatórios por categoria de despesa"""
    
    def __init__(self, parent=None):
        """Inicializa a interface do relatório"""
        self.parent = parent
        
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        configurar_janela(self.root, "Relatório por Categoria de Despesa", 1200, 1000)
        
        # Definição das categorias de despesa
        self.categorias_despesas = {
            'ADM': "ADMINISTRATIVO",
            'DIV': "DIVERSOS",
            'LOC': "LOCAÇÃO", 
            'MAT': "MATERIAL", 
            'MO': "MÃO-DE-OBRA", 
            'SERV': "SERVIÇOS",
            'TAX': "TAXA ADMINISTRAÇÃO",
            'TP': "TARIFAS/TRIBUTOS PÚBLICOS"
        }
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.data_referencia = datetime.now()
        self.df_despesas = None
        self.df_por_data = None
        self.data_selecionada = None
        self.dados_grafico = {}
        
        # Configurar interface
        self.setup_gui()
    
    def setup_gui(self):
        """Configuração da interface gráfica principal"""
        # Frame principal
        self.frame_principal = ttk.Frame(self.root, padding=10)
        self.frame_principal.pack(fill='both', expand=True)
        
        # Frame para seleção (TOPO - FIXO)
        self.frame_selecao = ttk.LabelFrame(self.frame_principal, text="Seleção de Cliente e Data")
        self.frame_selecao.pack(fill='x', side='top', pady=(0, 10))
        
        # Container para cliente
        frame_cliente = ttk.Frame(self.frame_selecao)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Selecione o Cliente:", font=('Arial', 11)).pack(side='left', pady=5)
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=40, font=('Arial', 11))
        self.cliente_combobox.pack(side='left', padx=5)
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.selecionar_cliente)
        
        # Container para botão de gerar relatório
        frame_data = ttk.Frame(self.frame_selecao)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        # Botão de gerar relatório
        ttk.Button(
            frame_data,
            text="Gerar Relatório",
            command=self.gerar_relatorio,
            style='Big.TButton'
        ).pack(side='left', padx=20)
        
        # BOTÕES NA PARTE INFERIOR (FIXO)
        frame_botoes = ttk.Frame(self.frame_principal)
        frame_botoes.pack(fill='x', side='bottom', pady=(10, 0))
        
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
        
        # Frame para resultados - MEIO (EXPANSÍVEL)
        self.frame_resultados = ttk.LabelFrame(self.frame_principal, text="Resultados")
        self.frame_resultados.pack(fill='both', expand=True, pady=(0, 10))
        
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
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 11, 'bold'), padding=(10, 5))
        
        # Carregar lista de clientes
        self.atualizar_lista_clientes()
    
    def setup_aba_resumo(self):
        """Configura a aba de resumo do relatório por data e categoria de despesa"""
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
        
        # Frame para o TreeView com os dados por data
        frame_resumo = ttk.Frame(self.aba_resumo, padding=5)
        frame_resumo.pack(fill='both', expand=True, pady=5)
        
        # Criar TreeView para os dados por data
        # Colunas: 'data', categorias (ADM, DIV, LOC, MAT, MO, SERV, TAX, TP), 'total'
        colunas = ['data']
        for categoria in self.categorias_despesas.keys():
            colunas.append(f'cat_{categoria}')
        colunas.append('total')
        
        self.tv_resumo = ttk.Treeview(frame_resumo, columns=colunas, show='headings', height=15)
        
        # Configurar cabeçalhos
        self.tv_resumo.heading('data', text='Data')
        for categoria in self.categorias_despesas.keys():
            # Usar a sigla da categoria para o cabeçalho
            self.tv_resumo.heading(f'cat_{categoria}', text=categoria)
            
        self.tv_resumo.heading('total', text='Total (R$)')
        
        # Configurar colunas
        self.tv_resumo.column('data', width=100, anchor='center')
        for categoria in self.categorias_despesas.keys():
            self.tv_resumo.column(f'cat_{categoria}', width=100, anchor='e', stretch=True)
        self.tv_resumo.column('total', width=120, anchor='e')
        
        # Configurar scrollbars
        scrollbar_y = ttk.Scrollbar(frame_resumo, orient='vertical', command=self.tv_resumo.yview)
        scrollbar_x = ttk.Scrollbar(frame_resumo, orient='horizontal', command=self.tv_resumo.xview)
        self.tv_resumo.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Adicionar à tela
        self.tv_resumo.pack(side='top', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        
        # Adicionar evento de seleção
        self.tv_resumo.bind('<<TreeviewSelect>>', self.selecionar_data)
        
        # Frame para resumo de totais
        frame_totais = ttk.LabelFrame(self.aba_resumo, text="Resumo Financeiro", padding=10)
        frame_totais.pack(fill='x', pady=10, padx=10)
        
        # Total de Despesas (primeira linha, primeira coluna)
        ttk.Label(frame_totais, text="Total de Despesas:", font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.lbl_total_geral = ttk.Label(frame_totais, text="R$ 0,00", font=('Arial', 11))
        self.lbl_total_geral.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        # Criar labels dinâmicos para cada categoria
        self.labels_categorias = {}
        row = 1
        col = 0
        
        for categoria, nome_completo in self.categorias_despesas.items():
            # Label do nome da categoria
            ttk.Label(frame_totais, text=f"{categoria}:", font=('Arial', 10, 'bold')).grid(row=row, column=col, sticky='e', padx=5, pady=2)
            
            # Label do valor da categoria
            self.labels_categorias[categoria] = ttk.Label(frame_totais, text="R$ 0,00", font=('Arial', 10))
            self.labels_categorias[categoria].grid(row=row, column=col+1, sticky='w', padx=5, pady=2)
            
            # Avançar para próxima posição
            col += 2
            if col >= 12:  # Máximo de 6 colunas (6 considerando label + valor)
                col = 0
                row += 1
    
    def setup_aba_detalhes(self):
        """Configura a aba de detalhes do relatório para a data selecionada"""
        # Frame para informações da data selecionada
        frame_info_data = ttk.Frame(self.aba_detalhes, padding=5)
        frame_info_data.pack(fill='x', pady=5)
        
        self.lbl_data_detalhe = ttk.Label(
            frame_info_data, 
            text="Data Selecionada: Nenhuma", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_data_detalhe.pack(side='left', padx=10)
        
        self.lbl_total_data_detalhe = ttk.Label(
            frame_info_data, 
            text="Total: R$ 0,00", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_total_data_detalhe.pack(side='left', padx=10)
        
        # Frame para a tabela de detalhes
        frame_tabela = ttk.Frame(self.aba_detalhes, padding=5)
        frame_tabela.pack(fill='both', expand=True, pady=5)
        
        # Criar TreeView para os lançamentos da data selecionada
        colunas = ('data', 'categoria', 'nome', 'referencia', 'dt_vencto', 'valor', 'observacao')
        self.tv_detalhes = ttk.Treeview(frame_tabela, columns=colunas, show='headings', height=15)
        
        # Configurar cabeçalhos
        self.tv_detalhes.heading('data', text='Data')
        self.tv_detalhes.heading('categoria', text='Categoria')
        self.tv_detalhes.heading('nome', text='Nome')
        self.tv_detalhes.heading('referencia', text='Referência')
        self.tv_detalhes.heading('dt_vencto', text='Data Vencto')
        self.tv_detalhes.heading('valor', text='Valor (R$)')
        self.tv_detalhes.heading('observacao', text='Observação')
        
        # Configurar colunas
        self.tv_detalhes.column('data', width=80, anchor='center')
        self.tv_detalhes.column('categoria', width=90, anchor='center')
        self.tv_detalhes.column('nome', width=180, anchor='w')
        self.tv_detalhes.column('referencia', width=220, anchor='w')
        self.tv_detalhes.column('dt_vencto', width=80, anchor='center')
        self.tv_detalhes.column('valor', width=120, anchor='e')
        self.tv_detalhes.column('observacao', width=180, anchor='w')
        
        # Configurar scrollbars
        scrollbar_y = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tv_detalhes.yview)
        scrollbar_x = ttk.Scrollbar(frame_tabela, orient='horizontal', command=self.tv_detalhes.xview)
        self.tv_detalhes.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Adicionar à tela
        self.tv_detalhes.pack(side='top', fill='both', expand=True)
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        
        # Frame para resumo de totais (igual ao da aba resumo)
        frame_totais_detalhes = ttk.LabelFrame(self.aba_detalhes, text="Resumo Financeiro", padding=10)
        frame_totais_detalhes.pack(fill='x', pady=10, padx=10)
        
        # Total de Despesas (primeira linha, primeira coluna)
        ttk.Label(frame_totais_detalhes, text="Total de Despesas:", font=('Arial', 11, 'bold')).grid(row=0, column=0, sticky='e', padx=5, pady=5)
        self.lbl_total_geral_detalhes = ttk.Label(frame_totais_detalhes, text="R$ 0,00", font=('Arial', 11))
        self.lbl_total_geral_detalhes.grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        # Criar labels dinâmicos para cada categoria na aba detalhes
        self.labels_categorias_detalhes = {}
        row = 1
        col = 0
        
        for categoria, nome_completo in self.categorias_despesas.items():
            # Label do nome da categoria
            ttk.Label(frame_totais_detalhes, text=f"{categoria}:", font=('Arial', 10, 'bold')).grid(row=row, column=col, sticky='e', padx=5, pady=2)
            
            # Label do valor da categoria
            self.labels_categorias_detalhes[categoria] = ttk.Label(frame_totais_detalhes, text="R$ 0,00", font=('Arial', 10))
            self.labels_categorias_detalhes[categoria].grid(row=row, column=col+1, sticky='w', padx=5, pady=2)
            
            # Avançar para próxima posição
            col += 2
            if col >= 12:  # Máximo de 6 colunas (considerando label + valor)
                col = 0
                row += 1
    
    def setup_aba_grafico(self):
        """Configura a aba de gráficos"""
        # Frame para controles do gráfico
        frame_controles = ttk.Frame(self.aba_grafico, padding=5)
        frame_controles.pack(fill='x', pady=5)
        
        ttk.Label(frame_controles, text="Tipo de Gráfico:").pack(side='left', padx=5)
        self.combo_tipo_grafico = ttk.Combobox(frame_controles, values=[
            "Gráfico de Pizza - Totais",
            "Gráfico de Barras - Totais", 
            "Gráfico de Linha do Tempo",
            "Gráfico de Pizza - Data Selecionada",
            "Gráfico de Barras - Data Selecionada"
        ], state='readonly', width=35)
        self.combo_tipo_grafico.pack(side='left', padx=5)
        self.combo_tipo_grafico.current(0)
        
        ttk.Button(frame_controles, text="Atualizar Gráfico", command=self.atualizar_grafico).pack(side='left', padx=20)
        
        # Frame para informações da data no gráfico
        frame_info_grafico = ttk.Frame(self.aba_grafico, padding=5)
        frame_info_grafico.pack(fill='x', pady=5)
        
        self.lbl_data_grafico = ttk.Label(
            frame_info_grafico, 
            text="Visualização: Dados Gerais", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_data_grafico.pack(side='left', padx=10)
        
        # Frame para o gráfico
        self.frame_grafico = ttk.Frame(self.aba_grafico)
        self.frame_grafico.pack(fill='both', expand=True, pady=5)
    
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
        if not self.cliente_atual:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
            
        # Definir data de referência como data atual (apenas para o relatório exportado)
        self.data_referencia = datetime.now()
            
        # Carregar dados
        if not self.carregar_dados():
            return
        
        # Preencher resumo
        self.preencher_resumo()
        
        # Resetar resumo financeiro da aba detalhes (ainda não há data selecionada)
        if hasattr(self, 'lbl_total_geral_detalhes'):
            self.lbl_total_geral_detalhes.config(text="R$ 0,00")
            
            # PARA CATEGORIA (usar este bloco no relatorio_categoria.py):
            if hasattr(self, 'labels_categorias_detalhes'):
                for categoria in self.categorias_despesas.keys():
                    if categoria in self.labels_categorias_detalhes:
                        self.labels_categorias_detalhes[categoria].config(text="R$ 0,00")

                        
        # Limpar detalhes (pois ainda não há data selecionada)
        if hasattr(self, 'tv_detalhes'):
            for item in self.tv_detalhes.get_children():
                self.tv_detalhes.delete(item)
        
        # Resetar data selecionada
        self.data_selecionada = None
        
        # Atualizar label do gráfico
        if hasattr(self, 'lbl_data_grafico'):
            self.lbl_data_grafico.config(text="Visualização: Dados Gerais")
        
        # Gerar gráfico inicial de totais
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
                
                # Verificar se as colunas necessárias existem
                colunas_necessarias = ['VALOR', 'DATA_REL']
                for coluna in colunas_necessarias:
                    if coluna not in self.df_despesas.columns:
                        messagebox.showerror("Erro", f"A coluna '{coluna}' não foi encontrada na aba Dados!")
                        return False
                
                # Verificar se existe a coluna K (categoria) - índice 10 (base 0)
                if len(self.df_despesas.columns) < 11:
                    messagebox.showerror("Erro", "A coluna K (categoria) não foi encontrada na aba Dados!")
                    return False
                
                # Nomear a coluna K como 'CATEGORIA' se ainda não tiver nome
                nome_coluna_k = self.df_despesas.columns[10]  # Coluna K (índice 10)
                if nome_coluna_k != 'CATEGORIA':
                    # Renomear a coluna K para CATEGORIA
                    self.df_despesas = self.df_despesas.rename(columns={nome_coluna_k: 'CATEGORIA'})
                
                print(f"Coluna de categoria identificada: {nome_coluna_k} -> CATEGORIA")
                
                # NOVO: Filtrar apenas lançamentos ativos
                if 'STATUS' in self.df_despesas.columns:
                    # Filtrar apenas registros com STATUS = 'ATIVO'
                    df_original_len = len(self.df_despesas)
                    self.df_despesas = self.df_despesas[
                        self.df_despesas['STATUS'].str.upper().str.strip() == 'ATIVO'
                    ].copy()
                    df_filtrado_len = len(self.df_despesas)
                    
                    print(f"Cliente {self.cliente_atual}: {df_original_len} registros totais, {df_filtrado_len} ativos processados")
                    
                    # Se não há registros ativos, mostrar aviso
                    if self.df_despesas.empty:
                        messagebox.showinfo("Aviso", f"Nenhum lançamento ativo encontrado para o cliente {self.cliente_atual}")
                        return False
                else:
                    # Se não existe a coluna STATUS, processar todos (compatibilidade)
                    print(f"Cliente {self.cliente_atual}: Coluna STATUS não encontrada, processando todos os registros")
                
                # Converter DATA_REL para datetime
                self.df_despesas['DATA_REL'] = pd.to_datetime(self.df_despesas['DATA_REL'], errors='coerce')
                
                # Converter DT_VENCTO para datetime (se existir)
                if 'DT_VENCTO' in self.df_despesas.columns:
                    self.df_despesas['DT_VENCTO'] = pd.to_datetime(self.df_despesas['DT_VENCTO'], errors='coerce')
                else:
                    # Se não existir, criar coluna vazia
                    self.df_despesas['DT_VENCTO'] = pd.NaT

                # Garantir valores numéricos para a coluna valor
                self.df_despesas['VALOR'] = pd.to_numeric(self.df_despesas['VALOR'], errors='coerce').fillna(0)
                
                # Processar categorias - converter para string e limpar
                self.df_despesas['CATEGORIA'] = self.df_despesas['CATEGORIA'].astype(str).str.upper().str.strip()
                
                # Substituir valores vazios ou 'nan' por 'DIV' (DIVERSOS)
                self.df_despesas['CATEGORIA'] = self.df_despesas['CATEGORIA'].replace(['', 'NAN', 'NONE'], 'DIV')
                
                # Verificar se existem categorias não mapeadas
                categorias_encontradas = set(self.df_despesas['CATEGORIA'].unique())
                categorias_validas = set(self.categorias_despesas.keys())
                categorias_invalidas = categorias_encontradas - categorias_validas
                
                if categorias_invalidas:
                    print(f"Categorias não mapeadas encontradas: {categorias_invalidas}")
                    # Substituir categorias inválidas por 'DIV'
                    self.df_despesas.loc[self.df_despesas['CATEGORIA'].isin(categorias_invalidas), 'CATEGORIA'] = 'DIV'
                
                # Ordenar dados por data
                self.df_despesas = self.df_despesas.sort_values(by='DATA_REL')
                
                # Agrupar dados por data e categoria
                self.preparar_dados_por_data()
                
                return True
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar dados do Excel: {str(e)}")
                import traceback
                traceback.print_exc()
                return False
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar dados: {str(e)}")
            return False
    
    def preparar_dados_por_data(self):
        """Prepara os dados agrupados por data"""
        try:
            # Criar um DataFrame para armazenar os dados agrupados por data e categoria
            # Agrupar por data e categoria de despesa
            df_pivot = self.df_despesas.pivot_table(
                index='DATA_REL', 
                columns='CATEGORIA', 
                values='VALOR', 
                aggfunc='sum'
            ).fillna(0)
            
            # Resetar o índice para tornar a data uma coluna
            df_pivot = df_pivot.reset_index()
            
            # Criar colunas para cada categoria se não existirem
            for categoria in self.categorias_despesas.keys():
                if categoria not in df_pivot.columns:
                    df_pivot[categoria] = 0.0
            
            # Calcular total por data
            df_pivot['total'] = df_pivot[[cat for cat in self.categorias_despesas.keys() if cat in df_pivot.columns]].sum(axis=1)
            
            # Ordenar por data (ascendente)
            df_pivot = df_pivot.sort_values(by='DATA_REL')
            
            # Armazenar o DataFrame para uso posterior
            self.df_por_data = df_pivot
            
            # Preparar dados para gráficos
            self.preparar_dados_grafico()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preparar dados por data: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def preparar_dados_grafico(self):
        """Prepara os dados para os gráficos"""
        try:
            # Inicializar dicionário de dados para gráficos
            self.dados_grafico = {}
            
            # Se não temos dados ou data selecionada, não há o que fazer
            if not hasattr(self, 'df_por_data') or self.df_por_data.empty:
                return
                
            # Dados para gráfico de pizza e barras por data selecionada
            # Serão preenchidos quando uma data for selecionada
            self.dados_grafico['pizza'] = None
            self.dados_grafico['barras'] = None
            
        except Exception as e:
            print(f"Erro ao preparar dados para gráfico: {str(e)}")
            import traceback
            traceback.print_exc()

    
    def preencher_resumo(self):
        """Preenche os dados da aba de resumo"""
        try:
            # Limpar TreeView
            for item in self.tv_resumo.get_children():
                self.tv_resumo.delete(item)
            
            # Adicionar dados à TreeView
            for _, row in self.df_por_data.iterrows():
                # Formatar data
                data_str = row['DATA_REL'].strftime('%d/%m/%Y')
                
                # Preparar valores para cada categoria
                valores = []
                for categoria in self.categorias_despesas.keys():
                    valor_formatado = formatar_moeda_br(row[categoria]) if categoria in row else "R$ 0,00"
                    valores.append(valor_formatado)
                
                # Adicionar total
                total_formatado = formatar_moeda_br(row['total'])
                
                # Inserir na treeview
                self.tv_resumo.insert(
                    '', 'end', 
                    values=[data_str] + valores + [total_formatado]
                )
            
            # Atualizar labels de totais
            total_geral = self.df_por_data['total'].sum()
            self.lbl_total_geral.config(text=formatar_moeda_br(total_geral))
            
            # Atualizar totais por tipo/categoria DA ABA RESUMO
            for categoria in self.categorias_despesas.keys():
                if categoria in self.df_por_data.columns:
                    total_tipo = self.df_por_data[categoria].sum()
                else:
                    total_tipo = 0
                
                if categoria in self.labels_categorias:
                    self.labels_categorias[categoria].config(text=formatar_moeda_br(total_tipo))
           
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher resumo: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def selecionar_data(self, event=None):
        """Atualiza a data selecionada e preenche as abas de detalhes e gráfico"""
        try:
            # Obter seleção atual
            selecao = self.tv_resumo.selection()
            if not selecao:
                return
                
            # Obter data selecionada
            item = self.tv_resumo.item(selecao[0])
            data_str = item['values'][0]  # Primeira coluna é a data
            
            # Converter string de data para datetime
            try:
                self.data_selecionada = datetime.strptime(data_str, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", f"Formato de data inválido: {data_str}")
                return
            
           
            # Atualizar label na aba de detalhes
            self.lbl_data_detalhe.config(text=f"Data Selecionada: {data_str}")
            
            # Atualizar label na aba de gráfico
            self.lbl_data_grafico.config(text=f"Data Selecionada: {data_str}")
            
            # Encontrar o total da data no DataFrame
            df_data = self.df_por_data[self.df_por_data['DATA_REL'] == self.data_selecionada]
            if not df_data.empty:
                total_data = df_data.iloc[0]['total']
                self.lbl_total_data_detalhe.config(text=f"Total: {formatar_moeda_br(total_data)}")
            
            # Filtrar dados para a data selecionada
            df_filtrado = self.df_despesas[self.df_despesas['DATA_REL'].dt.date == self.data_selecionada.date()].copy()
            
            # Atualizar resumo financeiro da aba detalhes com dados da data selecionada
            self.atualizar_resumo_financeiro_detalhes(df_filtrado)

            # Preencher detalhes
            self.preencher_detalhes(df_filtrado)

             # Atualizar resumo financeiro da aba detalhes com dados da data selecionada
            self.atualizar_resumo_financeiro_detalhes(df_filtrado)

            
            # Preparar dados para gráfico
            self.preparar_grafico_data_selecionada(df_filtrado)
            
            # Atualizar o tipo de gráfico para mostrar a data selecionada se estiver na aba de gráfico
            if "Data Selecionada" not in self.combo_tipo_grafico.get():
                self.combo_tipo_grafico.set("Gráfico de Pizza - Data Selecionada")
            
            # Atualizar gráfico
            self.atualizar_grafico()
            
            # Alternar para a aba de detalhes
            self.notebook.select(1)  # Índice 1 corresponde à aba de detalhes
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao selecionar data: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def atualizar_resumo_financeiro_detalhes(self, df_filtrado):
        """Atualiza o resumo financeiro da aba detalhes com dados da data selecionada"""
        try:
            if df_filtrado.empty:
                # Se não há dados, zerar tudo
                self.lbl_total_geral_detalhes.config(text="R$ 0,00")
                for categoria in self.categorias_despesas.keys():
                    if categoria in self.labels_categorias_detalhes:
                        self.labels_categorias_detalhes[categoria].config(text="R$ 0,00")
                return
            
            # Calcular total da data selecionada
            total_data = df_filtrado['VALOR'].sum()
            if hasattr(self, 'lbl_total_geral_detalhes'):
                self.lbl_total_geral_detalhes.config(text=formatar_moeda_br(total_data))
            
            # Calcular totais por categoria da data selecionada
            totais_por_categoria = df_filtrado.groupby('CATEGORIA')['VALOR'].sum()
            
            # Atualizar labels das categorias
            for categoria in self.categorias_despesas.keys():
                if categoria in self.labels_categorias_detalhes:
                    valor_categoria = totais_por_categoria.get(categoria, 0)
                    self.labels_categorias_detalhes[categoria].config(text=formatar_moeda_br(valor_categoria))
                    
        except Exception as e:
            print(f"Erro ao atualizar resumo financeiro detalhes: {str(e)}")

    def preparar_grafico_data_selecionada(self, df_filtrado):
        """Prepara dados para os gráficos da data selecionada"""
        try:
            if df_filtrado.empty:
                self.dados_grafico['pizza'] = None
                self.dados_grafico['barras'] = None
                return
                
            # Agrupar por categoria
            df_agrupado = df_filtrado.groupby('CATEGORIA')['VALOR'].sum().reset_index()
            
            # Adicionar nome completo da categoria
            df_agrupado['categoria_nome'] = df_agrupado['CATEGORIA'].apply(
                lambda x: f"{x} - {self.categorias_despesas.get(x, 'Não classificado')}" if pd.notna(x) else "Não classificado"
            )
            
            # Ordenar por categoria
            df_agrupado = df_agrupado.sort_values(by='CATEGORIA')
            
            # Armazenar para gráficos
            self.dados_grafico['pizza'] = df_agrupado.copy()
            self.dados_grafico['barras'] = df_agrupado.copy()
            
        except Exception as e:
            print(f"Erro ao preparar gráfico para data selecionada: {str(e)}")
            import traceback
            traceback.print_exc()
    
    def preencher_detalhes(self, df_filtrado):
        """Preenche os detalhes para a data selecionada"""
        try:
            # Limpar tabela
            for item in self.tv_detalhes.get_children():
                self.tv_detalhes.delete(item)
            
            # Verificar se o DataFrame está vazio
            if df_filtrado.empty:
                return
            
            # Ordenar o DataFrame por Categoria (ascendente), Nome (ascendente) e Valor (descendente)
            df_ordenado = df_filtrado.copy()
            
            # Garantir que todos os campos necessários existam
            if 'NOME' not in df_ordenado.columns:
                df_ordenado['NOME'] = ''
                
            # Ordenar primeiro por categoria, depois por nome (asc) e finalmente por valor (desc)
            df_ordenado = df_ordenado.sort_values(
                by=['CATEGORIA', 'NOME', 'VALOR'], 
                ascending=[True, True, False]
            )
            
            # Adicionar dados à tabela
            for _, row in df_ordenado.iterrows():
                # Formatar data
                data_str = row['DATA_REL'].strftime('%d/%m/%Y') if pd.notna(row['DATA_REL']) else ''
                
                # Obter categoria
                categoria = row['CATEGORIA'] if pd.notna(row['CATEGORIA']) else 'DIV'
                categoria_nome = f"{categoria} - {self.categorias_despesas.get(categoria, 'Não classificado')}"
                
                # Obter nome e referência
                nome = row.get('NOME', '') if pd.notna(row.get('NOME', '')) else ''
                
                # Obter referência e NF (juntar referência e NF)
                referencia = row.get('REFERÊNCIA', '') if pd.notna(row.get('REFERÊNCIA', '')) else ''
                nf = row.get('NF', '') if pd.notna(row.get('NF', '')) else ''

                # Concatenar referência e NF se ambos existirem
                if referencia and nf:
                    referencia = f"{referencia} - NF: {nf}"
                elif nf:
                    referencia = f"NF: {nf}"

                # Data de vencimento
                dt_vencto_str = ''
                if 'DT_VENCTO' in row and pd.notna(row['DT_VENCTO']):
                    dt_vencto_str = row['DT_VENCTO'].strftime('%d/%m/%Y')
                
                # Obter valor
                valor = formatar_moeda_br(row['VALOR'])
                
                # Obter observação
                observacao = row.get('OBSERVAÇÃO', '') if pd.notna(row.get('OBSERVAÇÃO', '')) else ''
                
                # Inserir na tabela
                self.tv_detalhes.insert(
                    '', 'end', 
                    values=(
                        data_str,
                        categoria_nome,
                        nome,
                        referencia,
                        dt_vencto_str,
                        valor,
                        observacao
                    )
                )
            
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao preencher detalhes: {str(e)}")
            import traceback
            traceback.print_exc()
 
    def determinar_agrupamento_temporal(self):
        """Determina o melhor agrupamento temporal baseado no período dos dados"""
        if not hasattr(self, 'df_por_data') or self.df_por_data.empty:
            return 'dia'
        
        # Calcular diferença em dias entre primeira e última data
        data_inicio = self.df_por_data['DATA_REL'].min()
        data_fim = self.df_por_data['DATA_REL'].max()
        dias_total = (data_fim - data_inicio).days
        
        # Definir agrupamento baseado no período
        if dias_total <= 31:  # Até 1 mês
            return 'dia'
        elif dias_total <= 93:  # Até 3 meses
            return 'semana'
        elif dias_total <= 365:  # Até 1 ano
            return 'mes'
        elif dias_total <= 1095:  # Até 3 anos
            return 'trimestre'
        else:  # Mais de 3 anos
            return 'ano'

    def preparar_dados_timeline(self):
        """Prepara dados para gráfico de linha do tempo"""
        try:
            if not hasattr(self, 'df_despesas') or self.df_despesas.empty:
                return None
            
            agrupamento = self.determinar_agrupamento_temporal()
            df_timeline = self.df_despesas.copy()
            
            # Criar coluna de agrupamento temporal
            if agrupamento == 'dia':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.date
                formato_periodo = lambda x: x.strftime('%d/%m/%Y')
            elif agrupamento == 'semana':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('W')
                formato_periodo = lambda x: f"Sem {x.week}/{x.year}"
            elif agrupamento == 'mes':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('M')
                formato_periodo = lambda x: f"{x.month:02d}/{x.year}"
            elif agrupamento == 'trimestre':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('Q')
                formato_periodo = lambda x: f"Q{x.quarter}/{x.year}"
            else:  # ano
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('Y')
                formato_periodo = lambda x: str(x.year)
            
            # Agrupar por período e categoria
            df_agrupado = df_timeline.groupby(['periodo', 'CATEGORIA'])['VALOR'].sum().reset_index()
            
            # Criar pivot para ter categorias como colunas
            df_pivot = df_agrupado.pivot(index='periodo', columns='CATEGORIA', values='VALOR').fillna(0)
            
            # Garantir que todas as categorias existam
            for categoria in self.categorias_despesas.keys():
                if categoria not in df_pivot.columns:
                    df_pivot[categoria] = 0
            
            # Resetar índice e formatar período
            df_pivot = df_pivot.reset_index()
            df_pivot['periodo_str'] = df_pivot['periodo'].apply(formato_periodo)
            
            # Calcular total por período
            df_pivot['total'] = df_pivot[[cat for cat in self.categorias_despesas.keys() if cat in df_pivot.columns]].sum(axis=1)
            
            return {
                'dados': df_pivot,
                'agrupamento': agrupamento,
                'categorias': list(self.categorias_despesas.keys())
            }
            
        except Exception as e:
            print(f"Erro ao preparar dados de timeline: {str(e)}")
            import traceback
            traceback.print_exc()
            return None    

    def atualizar_grafico(self, event=None):
        """Atualiza o gráfico com base na seleção"""
        try:
            tipo_grafico = self.combo_tipo_grafico.get()
            
            # Limpar frame do gráfico
            self.limpar_grafico()
            
            # Verificar se há dados
            if not hasattr(self, 'df_por_data') or self.df_por_data.empty:
                return
            
            # Criar figura
            fig, ax = plt.subplots(figsize=(12, 7))
            
            if "Linha do Tempo" in tipo_grafico:
                self.criar_grafico_timeline(fig, ax)
            elif "Data Selecionada" in tipo_grafico:
                # Gráficos para data específica
                if not self.data_selecionada:
                    ax.text(0.5, 0.5, "Selecione uma data na aba Resumo\npara ver os gráficos da data específica", 
                        horizontalalignment='center', verticalalignment='center',
                        transform=ax.transAxes, fontsize=14)
                else:
                    if "Pizza" in tipo_grafico:
                        self.criar_grafico_pizza(fig, ax)
                    elif "Barras" in tipo_grafico:
                        self.criar_grafico_barras(fig, ax)
            else:
                # Gráficos de totais gerais
                if "Pizza" in tipo_grafico:
                    self.criar_grafico_pizza_totais(fig, ax)
                elif "Barras" in tipo_grafico:
                    self.criar_grafico_barras_totais(fig, ax)
            
            # Exibir o gráfico
            canvas = FigureCanvasTkAgg(fig, master=self.frame_grafico)
            canvas.draw()
            canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar gráfico: {str(e)}")
            import traceback
            traceback.print_exc()

    def limpar_grafico(self):
        """Limpa o gráfico atual"""
        for widget in self.frame_grafico.winfo_children():
            widget.destroy()
    
    
    def determinar_agrupamento_temporal(self):
        """Determina o melhor agrupamento temporal baseado no período dos dados"""
        if not hasattr(self, 'df_por_data') or self.df_por_data.empty:
            return 'dia'
        
        # Calcular diferença em dias entre primeira e última data
        data_inicio = self.df_por_data['DATA_REL'].min()
        data_fim = self.df_por_data['DATA_REL'].max()
        dias_total = (data_fim - data_inicio).days
        
        # Definir agrupamento baseado no período
        if dias_total <= 31:  # Até 1 mês
            return 'dia'
        elif dias_total <= 93:  # Até 3 meses
            return 'semana'
        elif dias_total <= 365:  # Até 1 ano
            return 'mes'
        elif dias_total <= 1095:  # Até 3 anos
            return 'trimestre'
        else:  # Mais de 3 anos
            return 'ano'

    def preparar_dados_timeline(self):
        """Prepara dados para gráfico de linha do tempo"""
        try:
            if not hasattr(self, 'df_despesas') or self.df_despesas.empty:
                return None
            
            agrupamento = self.determinar_agrupamento_temporal()
            df_timeline = self.df_despesas.copy()
            
            # Criar coluna de agrupamento temporal
            if agrupamento == 'dia':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.date
                formato_periodo = lambda x: x.strftime('%d/%m/%Y')
            elif agrupamento == 'semana':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('W')
                formato_periodo = lambda x: f"Sem {x.week}/{x.year}"
            elif agrupamento == 'mes':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('M')
                formato_periodo = lambda x: f"{x.month:02d}/{x.year}"
            elif agrupamento == 'trimestre':
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('Q')
                formato_periodo = lambda x: f"Q{x.quarter}/{x.year}"
            else:  # ano
                df_timeline['periodo'] = df_timeline['DATA_REL'].dt.to_period('Y')
                formato_periodo = lambda x: str(x.year)
            
            # Agrupar por período e categoria
            df_agrupado = df_timeline.groupby(['periodo', 'CATEGORIA'])['VALOR'].sum().reset_index()
            
            # Criar pivot para ter categorias como colunas
            df_pivot = df_agrupado.pivot(index='periodo', columns='CATEGORIA', values='VALOR').fillna(0)
            
            # Garantir que todas as categorias existam
            for categoria in self.categorias_despesas.keys():
                if categoria not in df_pivot.columns:
                    df_pivot[categoria] = 0
            
            # Resetar índice e formatar período
            df_pivot = df_pivot.reset_index()
            df_pivot['periodo_str'] = df_pivot['periodo'].apply(formato_periodo)
            
            # Calcular total por período
            df_pivot['total'] = df_pivot[[cat for cat in self.categorias_despesas.keys() if cat in df_pivot.columns]].sum(axis=1)
            
            return {
                'dados': df_pivot,
                'agrupamento': agrupamento,
                'categorias': list(self.categorias_despesas.keys())
            }
            
        except Exception as e:
            print(f"Erro ao preparar dados de timeline: {str(e)}")
            import traceback
            traceback.print_exc()
            return None
    
    def criar_grafico_pizza(self, fig, ax):
        """Cria um gráfico de pizza com as categorias da data selecionada"""
        try:
            # Usar os dados para gráfico de pizza
            df = self.dados_grafico.get('pizza')
            
            if df is None or df.empty:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Cores para o gráfico
            colors = plt.cm.tab10.colors
            
            # Criar o gráfico de pizza
            wedges, texts, autotexts = ax.pie(
                df['VALOR'], 
                labels=df['categoria_nome'], 
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
            data_str = self.data_selecionada.strftime('%d/%m/%Y')
            ax.set_title(f'Distribuição por Categoria de Despesa - {data_str}', fontsize=14, pad=20)
            
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
        """Cria um gráfico de barras com as categorias da data selecionada"""
        try:
            # Usar os dados para gráfico de barras
            df = self.dados_grafico.get('barras')
            
            if df is None or df.empty:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Ordenar por valor para melhor visualização
            df = df.sort_values(by='VALOR', ascending=True)
            
            # Cores para o gráfico
            colors = plt.cm.tab10.colors[:len(df)]
            
            # Criar o gráfico de barras
            bars = ax.barh(df['categoria_nome'], df['VALOR'], color=colors)
            
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
            data_str = self.data_selecionada.strftime('%d/%m/%Y')
            ax.set_title(f'Valores por Categoria de Despesa - {data_str}', fontsize=14)
            ax.set_xlabel('Valor (R$)', fontsize=11)
            ax.set_ylabel('Categoria de Despesa', fontsize=11)
            
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
            
    def criar_grafico_pizza_totais(self, fig, ax):
        """Cria um gráfico de pizza com os totais por categoria"""
        try:
            # Calcular totais por categoria
            totais = {}
            for categoria in self.categorias_despesas.keys():
                if categoria in self.df_por_data.columns:
                    totais[categoria] = self.df_por_data[categoria].sum()
                else:
                    totais[categoria] = 0
            
            # Filtrar categorias com valor > 0
            totais_filtrados = {k: v for k, v in totais.items() if v > 0}
            
            if not totais_filtrados:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Preparar dados para o gráfico
            categorias = list(totais_filtrados.keys())
            valores = list(totais_filtrados.values())
            labels = [f"{cat} - {self.categorias_despesas[cat]}" for cat in categorias]
            
            # Cores para o gráfico
            colors = plt.cm.tab10.colors
            
            # Criar o gráfico de pizza
            wedges, texts, autotexts = ax.pie(
                valores, 
                labels=labels, 
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
            total_geral = sum(valores)
            ax.set_title(f'Distribuição Total por Categoria - {formatar_moeda_br(total_geral)}', 
                        fontsize=14, pad=20)
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de pizza totais: {str(e)}")
            ax.text(0.5, 0.5, f"Erro ao gerar gráfico: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12, color='red')

    def criar_grafico_barras_totais(self, fig, ax):
        """Cria um gráfico de barras com os totais por categoria"""
        try:
            # Calcular totais por categoria
            totais = {}
            for categoria in self.categorias_despesas.keys():
                if categoria in self.df_por_data.columns:
                    totais[categoria] = self.df_por_data[categoria].sum()
                else:
                    totais[categoria] = 0
            
            # Filtrar categorias com valor > 0 e ordenar por valor
            totais_filtrados = {k: v for k, v in totais.items() if v > 0}
            totais_ordenados = dict(sorted(totais_filtrados.items(), key=lambda x: x[1], reverse=True))
            
            if not totais_ordenados:
                ax.text(0.5, 0.5, "Não há dados para exibir", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            # Preparar dados para o gráfico
            categorias = list(totais_ordenados.keys())
            valores = list(totais_ordenados.values())
            labels = [f"{cat} - {self.categorias_despesas[cat]}" for cat in categorias]
            
            # Cores para o gráfico
            colors = plt.cm.tab10.colors[:len(categorias)]
            
            # Criar o gráfico de barras
            bars = ax.barh(labels, valores, color=colors)
            
            # Adicionar valores nas barras
            for bar in bars:
                width = bar.get_width()
                label_x_pos = width + width * 0.01
                ax.text(label_x_pos, bar.get_y() + bar.get_height()/2, 
                    formatar_moeda_br(width), va='center', fontsize=9)
            
            # Ajustar formatação do eixo x (valores)
            def format_real(x, pos):
                return f'R$ {x:,.0f}'.replace(',', '.')
            
            ax.xaxis.set_major_formatter(mticker.FuncFormatter(format_real))
            
            # Adicionar títulos e labels
            total_geral = sum(valores)
            ax.set_title(f'Totais por Categoria - {formatar_moeda_br(total_geral)}', fontsize=14)
            ax.set_xlabel('Valor (R$)', fontsize=11)
            ax.set_ylabel('Categoria de Despesa', fontsize=11)
            
            # Adicionar grid
            ax.grid(axis='x', linestyle='--', alpha=0.7)
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de barras totais: {str(e)}")
            ax.text(0.5, 0.5, f"Erro ao gerar gráfico: {str(e)}", 
                horizontalalignment='center', verticalalignment='center',
                transform=ax.transAxes, fontsize=12, color='red')

    def criar_grafico_timeline(self, fig, ax):
        """Cria um gráfico de linha do tempo"""
        try:
            dados_timeline = self.preparar_dados_timeline()
            
            if not dados_timeline or dados_timeline['dados'].empty:
                ax.text(0.5, 0.5, "Não há dados suficientes para linha do tempo", 
                    horizontalalignment='center', verticalalignment='center',
                    transform=ax.transAxes, fontsize=14)
                return
            
            df = dados_timeline['dados']
            agrupamento = dados_timeline['agrupamento']
            categorias = dados_timeline['categorias']
            
            # Cores para cada categoria
            colors = plt.cm.tab10.colors
            
            # Plotar linha para cada categoria
            for i, categoria in enumerate(categorias):
                if categoria in df.columns and df[categoria].sum() > 0:
                    ax.plot(df['periodo_str'], df[categoria], 
                        marker='o', linewidth=2, markersize=4,
                        color=colors[i % len(colors)],
                        label=f"{categoria} - {self.categorias_despesas[categoria]}")
            
            # Plotar linha do total
            ax.plot(df['periodo_str'], df['total'], 
                marker='s', linewidth=3, markersize=5,
                color='black', linestyle='--',
                label='Total Geral')
            
            # Configurar eixos
            ax.set_xlabel(f'Período ({agrupamento.title()})', fontsize=11)
            ax.set_ylabel('Valor (R$)', fontsize=11)
            
            # Formatar eixo Y
            def format_real(x, pos):
                return f'R$ {x:,.0f}'.replace(',', '.')
            ax.yaxis.set_major_formatter(mticker.FuncFormatter(format_real))
            
            # Rotacionar labels do eixo X se necessário
            if len(df) > 10:
                ax.tick_params(axis='x', rotation=45)
            
            # Adicionar grid
            ax.grid(True, linestyle='--', alpha=0.7)
            
            # Adicionar legenda
            ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left')
            
            # Título
            data_inicio = self.df_por_data['DATA_REL'].min().strftime('%d/%m/%Y')
            data_fim = self.df_por_data['DATA_REL'].max().strftime('%d/%m/%Y')
            ax.set_title(f'Evolução das Despesas por Categoria\n{data_inicio} a {data_fim}', 
                        fontsize=14, pad=20)
            
            # Ajustar layout
            fig.tight_layout()
            
        except Exception as e:
            print(f"Erro ao criar gráfico de timeline: {str(e)}")
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
        nome_padrao = f"Relatorio_Categoria_{self.cliente_atual}_{data_str}.xlsx"
        
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
            ws_resumo['A1'] = "Relatório por Categoria de Despesa"
            ws_resumo['A1'].font = Font(size=14, bold=True)
            ws_resumo.merge_cells('A1:I1')
            ws_resumo['A1'].alignment = Alignment(horizontal='center')
            
            ws_resumo['A2'] = f"Cliente: {self.cliente_atual}"
            ws_resumo['A2'].font = Font(size=12, bold=True)
            ws_resumo.merge_cells('A2:I2')
            
            ws_resumo['A3'] = f"Data do relatório: {data_str}"
            ws_resumo['A3'].font = Font(size=12)
            ws_resumo.merge_cells('A3:I3')
            
            # Adicionar cabeçalhos da tabela
            headers = ["Data"] + list(self.categorias_despesas.keys()) + ["Total (R$)"]
            for col, header in enumerate(headers, start=1):
                cell = ws_resumo.cell(row=5, column=col, value=header)
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
                cell.fill = PatternFill(fgColor="DDDDDD", fill_type="solid")
            
            # Adicionar dados
            for i, (_, row) in enumerate(self.df_por_data.iterrows(), start=6):
                # Data formatada
                ws_resumo.cell(row=i, column=1, value=row['DATA_REL'])
                ws_resumo.cell(row=i, column=1).number_format = "dd/mm/yyyy"
                
                # Valores por categoria
                for j, categoria in enumerate(self.categorias_despesas.keys(), start=2):
                    ws_resumo.cell(row=i, column=j, value=row[categoria] if categoria in row else 0)
                    ws_resumo.cell(row=i, column=j).number_format = "R$ #,##0.00"
                
                # Total
                ws_resumo.cell(row=i, column=len(headers), value=row['total'])
                ws_resumo.cell(row=i, column=len(headers)).number_format = "R$ #,##0.00"
            
            # Ajustar largura das colunas
            ws_resumo.column_dimensions['A'].width = 15
            for col in range(2, len(headers) + 1):
                ws_resumo.column_dimensions[get_column_letter(col)].width = 15
            
            # Adicionar totais
            total_row = 6 + len(self.df_por_data)
            
            ws_resumo.cell(row=total_row, column=1, value="TOTAL")
            ws_resumo.cell(row=total_row, column=1).font = Font(bold=True)
            
            # Totais por categoria
            for j, categoria in enumerate(self.categorias_despesas.keys(), start=2):
                formula = f"=SUM({get_column_letter(j)}6:{get_column_letter(j)}{total_row-1})"
                ws_resumo.cell(row=total_row, column=j, value=formula)
                ws_resumo.cell(row=total_row, column=j).font = Font(bold=True)
                ws_resumo.cell(row=total_row, column=j).number_format = "R$ #,##0.00"
            
            # Total geral
            total_formula = f"=SUM({get_column_letter(len(headers))}6:{get_column_letter(len(headers))}{total_row-1})"
            ws_resumo.cell(row=total_row, column=len(headers), value=total_formula)
            ws_resumo.cell(row=total_row, column=len(headers)).font = Font(bold=True)
            ws_resumo.cell(row=total_row, column=len(headers)).number_format = "R$ #,##0.00"
            
            # Criar aba de detalhes se tivermos uma data selecionada
            if hasattr(self, 'data_selecionada') and self.data_selecionada:
                ws_detalhes = wb.create_sheet("Detalhes")
                
                # Adicionar cabeçalho
                data_str_detalhe = self.data_selecionada.strftime('%d/%m/%Y')
                ws_detalhes['A1'] = f"Detalhes da Data: {data_str_detalhe}"
                ws_detalhes['A1'].font = Font(size=14, bold=True)
                ws_detalhes.merge_cells('A1:G1')
                ws_detalhes['A1'].alignment = Alignment(horizontal='center')
                
                # Filtrando dados para a data selecionada
                df_filtrado = self.df_despesas[self.df_despesas['DATA_REL'].dt.date == self.data_selecionada.date()].copy()
                
                # Adicionar cabeçalhos da tabela
                headers = ["Data", "Categoria", "Nome", "Referência", "Valor (R$)", "Observação"]
                for col, header in enumerate(headers, start=1):
                    cell = ws_detalhes.cell(row=3, column=col, value=header)
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')
                    cell.fill = PatternFill(fgColor="DDDDDD", fill_type="solid")
                
                # Adicionar dados
                for i, (_, row) in enumerate(df_filtrado.iterrows(), start=4):
                    # Data formatada
                    if pd.notna(row['DATA_REL']):
                        ws_detalhes.cell(row=i, column=1, value=row['DATA_REL'])
                        ws_detalhes.cell(row=i, column=1).number_format = "dd/mm/yyyy"
                    
                    # Categoria
                    categoria = row['CATEGORIA'] if pd.notna(row['CATEGORIA']) else 'DIV'
                    categoria_nome = f"{categoria} - {self.categorias_despesas.get(categoria, 'Não classificado')}"
                    ws_detalhes.cell(row=i, column=2, value=categoria_nome)
                    
                    # Nome
                    if 'NOME' in row and pd.notna(row['NOME']):
                        ws_detalhes.cell(row=i, column=3, value=row['NOME'])
                    
                    # Referência
                    if 'REFERÊNCIA' in row and pd.notna(row['REFERÊNCIA']):
                        ws_detalhes.cell(row=i, column=4, value=row['REFERÊNCIA'])
                    
                    # Valor
                    ws_detalhes.cell(row=i, column=5, value=row['VALOR'])
                    ws_detalhes.cell(row=i, column=5).number_format = "R$ #,##0.00"
                    
                    # Observação
                    if 'OBSERVAÇÃO' in row and pd.notna(row['OBSERVAÇÃO']):
                        ws_detalhes.cell(row=i, column=6, value=row['OBSERVAÇÃO'])
                
                # Ajustar largura das colunas
                ws_detalhes.column_dimensions['A'].width = 15
                ws_detalhes.column_dimensions['B'].width = 25
                ws_detalhes.column_dimensions['C'].width = 30
                ws_detalhes.column_dimensions['D'].width = 30
                ws_detalhes.column_dimensions['E'].width = 15
                ws_detalhes.column_dimensions['F'].width = 40
                
                # Adicionar total
                total_row = 4 + len(df_filtrado)
                
                ws_detalhes.cell(row=total_row, column=4, value="TOTAL")
                ws_detalhes.cell(row=total_row, column=4).font = Font(bold=True)
                
                # Total em R$
                total_formula = f"=SUM(E4:E{total_row-1})"
                ws_detalhes.cell(row=total_row, column=5, value=total_formula)
                ws_detalhes.cell(row=total_row, column=5).font = Font(bold=True)
                ws_detalhes.cell(row=total_row, column=5).number_format = "R$ #,##0.00"
            
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
        nome_padrao = f"Relatorio_Categoria_{self.cliente_atual}_{data_str}.pdf"
        
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
            doc = SimpleDocTemplate(arquivo, pagesize=A4, leftMargin=0.5*inch, rightMargin=0.5*inch)
            story = []
            
            # Estilos
            styles = getSampleStyleSheet()
            title_style = styles['Title']
            heading1_style = styles['Heading1']
            heading2_style = styles['Heading2']
            normal_style = styles['Normal']
            
            # Título
            story.append(Paragraph(f"Relatório por Categoria de Despesa", title_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Informações do cliente
            story.append(Paragraph(f"Cliente: {self.cliente_atual}", heading1_style))
            story.append(Paragraph(f"Data do relatório: {data_str}", normal_style))
            story.append(Spacer(1, 0.2*inch))
            
            # Resumo por data
            story.append(Paragraph("Resumo por Data", heading2_style))
            story.append(Spacer(1, 0.1*inch))
            
            # Cabeçalhos
            headers = ["Data"] + list(self.categorias_despesas.keys()) + ["Total (R$)"]
            
            # Dados da tabela de resumo
            table_data = [headers]
            
            for _, row in self.df_por_data.iterrows():
                # Formatar data
                data_str = row['DATA_REL'].strftime('%d/%m/%Y')
                
                # Preparar valores para cada categoria
                valores = [data_str]
                for categoria in self.categorias_despesas.keys():
                    valor_formatado = f"R$ {row[categoria]:,.2f}".replace(',', '.').replace('.', ',') if categoria in row else "R$ 0,00"
                    valores.append(valor_formatado)
                
                # Adicionar total
                total_formatado = f"R$ {row['total']:,.2f}".replace(',', '.').replace('.', ',')
                valores.append(total_formatado)
                
                table_data.append(valores)
            
            # Adicionar linha de total
            if not self.df_por_data.empty:
                total_row = ["TOTAL"]
                for categoria in self.categorias_despesas.keys():
                    total_cat = self.df_por_data[categoria].sum() if categoria in self.df_por_data.columns else 0
                    total_formatado = f"R$ {total_cat:,.2f}".replace(',', '.').replace('.', ',')
                    total_row.append(total_formatado)
                
                # Total geral
                total_geral = self.df_por_data['total'].sum()
                total_geral_formatado = f"R$ {total_geral:,.2f}".replace(',', '.').replace('.', ',')
                total_row.append(total_geral_formatado)
                
                table_data.append(total_row)
            
            # Criar tabela de resumo
            # Calcular larguras de colunas
            num_categorias = len(self.categorias_despesas)
            col_width_data = 1.0*inch  # Data
            col_width_cat = (6.0*inch) / num_categorias  # Dividir espaço restante entre categorias
            col_width_total = 1.0*inch  # Total
            
            col_widths = [col_width_data] + [col_width_cat] * num_categorias + [col_width_total]
            
            resumo_table = Table(table_data, colWidths=col_widths)
            
            # Estilo da tabela
            table_style = TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                ('ALIGN', (0, 0), (0, -1), 'CENTER'),
                ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, -1), 8),  # Fonte menor para caber
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ])
            
            # Destacar a linha de total
            if not self.df_por_data.empty:
                table_style.add('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
                table_style.add('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
            
            resumo_table.setStyle(table_style)
            story.append(resumo_table)
            story.append(Spacer(1, 0.2*inch))
            
            # Se tiver uma data selecionada, adicionar detalhes
            if hasattr(self, 'data_selecionada') and self.data_selecionada:
                # Detalhes da data selecionada
                data_str_detalhe = self.data_selecionada.strftime('%d/%m/%Y')
                story.append(Paragraph(f"Detalhes - Data: {data_str_detalhe}", heading2_style))
                story.append(Spacer(1, 0.1*inch))
                
                # Filtrar dados para a data selecionada
                df_filtrado = self.df_despesas[self.df_despesas['DATA_REL'].dt.date == self.data_selecionada.date()].copy()
                
                if not df_filtrado.empty:
                    # Cabeçalhos
                    headers = ["Categoria", "Nome", "Referência", "Valor (R$)"]
                    
                    # Dados da tabela de detalhes
                    table_data = [headers]
                    
                    for _, row in df_filtrado.iterrows():
                        # Obter categoria
                        categoria = row['CATEGORIA'] if pd.notna(row['CATEGORIA']) else 'DIV'
                        categoria_nome = f"{categoria} - {self.categorias_despesas.get(categoria, 'Não classificado')}"
                        
                        # Obter nome e referência
                        nome = row.get('NOME', '') if pd.notna(row.get('NOME', '')) else ''
                        referencia = row.get('REFERÊNCIA', '') if pd.notna(row.get('REFERÊNCIA', '')) else ''
                        
                        # Formatar valor
                        valor = f"R$ {row['VALOR']:,.2f}".replace(',', '.').replace('.', ',')
                        
                        # Adicionar linha
                        table_data.append([categoria_nome, nome, referencia, valor])
                    
                    # Adicionar linha de total
                    total_data = df_filtrado['VALOR'].sum()
                    total_formatado = f"R$ {total_data:,.2f}".replace(',', '.').replace('.', ',')
                    table_data.append(["TOTAL", "", "", total_formatado])
                    
                    # Criar tabela de detalhes
                    col_widths = [1.8*inch, 1.8*inch, 2.0*inch, 1.0*inch]
                    detalhes_table = Table(table_data, colWidths=col_widths)
                    
                    # Estilo da tabela
                    table_style = TableStyle([
                        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
                        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
                        ('ALIGN', (0, 0), (2, -1), 'LEFT'),
                        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
                        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                        ('FONTSIZE', (0, 0), (-1, -1), 8),
                        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ])
                    
                    # Destacar a linha de total
                    table_style.add('BACKGROUND', (0, -1), (-1, -1), colors.lightgrey)
                    table_style.add('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold')
                    
                    detalhes_table.setStyle(table_style)
                    story.append(detalhes_table)
                    story.append(Spacer(1, 0.2*inch))
                    
                    # Adicionar gráfico
                    story.append(Paragraph("Gráfico de Distribuição por Categoria de Despesa", heading2_style))
                    story.append(Spacer(1, 0.1*inch))
                    
                    # Gerar gráfico para incluir no PDF
                    plt.figure(figsize=(7, 5))
                    
                    # Preparar dados para o gráfico
                    df_grafico = df_filtrado.groupby('CATEGORIA')['VALOR'].sum().reset_index()
                    
                    # Adicionar nome da categoria
                    df_grafico['categoria_nome'] = df_grafico['CATEGORIA'].apply(
                        lambda x: f"{x} - {self.categorias_despesas.get(x, 'Não classificado')}"
                    )
                    
                    # Criar gráfico de pizza
                    if not df_grafico.empty:
                        plt.pie(
                            df_grafico['VALOR'], 
                            labels=df_grafico['categoria_nome'], 
                            autopct='%1.1f%%',
                            startangle=90,
                            colors=plt.cm.tab10.colors,
                            wedgeprops={'edgecolor': 'w', 'linewidth': 1}
                        )
                        
                        plt.title(f'Distribuição por Categoria de Despesa - {data_str_detalhe}', fontsize=12, pad=20)
                        plt.tight_layout()
                        
                        # Salvar o gráfico em um buffer
                        img_buffer = io.BytesIO()
                        plt.savefig(img_buffer, format='png', dpi=100)
                        img_buffer.seek(0)
                        plt.close()
                        
                        # Adicionar o gráfico ao PDF
                        img = Image(img_buffer, width=6*inch, height=4*inch)
                        story.append(img)
                    else:
                        story.append(Paragraph("Não há dados suficientes para gerar o gráfico.", normal_style))
                else:
                    story.append(Paragraph("Não há lançamentos para esta data.", normal_style))
            
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
    app = RelatorioCategoria()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()