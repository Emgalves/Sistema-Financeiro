import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib
from datetime import datetime, date
from pathlib import Path

# Adicionar diretório raiz ao path ANTES de qualquer importação
def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# SISTEMA DE LOGGING ROBUSTO usando o sistema existente
def setup_logging_safe():
    """
    Configura logging usando o sistema existente com fallback seguro
    """
    try:
        # Tentar usar o sistema de logging existente
        from src.config.logger_config import system_logger, log_action
        
        # Configurar usuário padrão se não estiver definido
        system_logger.set_user('sistema_relatorios')
        
        # Obter logger
        logger = system_logger.get_logger()
        logger.info("Sistema de relatórios inicializando usando logger configurado")
        
        return logger, log_action
        
    except ImportError as e:
        print(f"Aviso: Não foi possível importar sistema de logging configurado: {str(e)}")
        
        # Fallback: criar sistema de logging simples
        import logging
        
        # Configurar logger básico
        logger = logging.getLogger("sistema_relatorios")
        logger.setLevel(logging.INFO)
        
        # Evitar handlers duplicados
        if not logger.handlers:
            # Handler para console
            console_handler = logging.StreamHandler()
            console_handler.setFormatter(
                logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
            )
            logger.addHandler(console_handler)
            
            # Tentar handler para arquivo
            try:
                # Determinar diretório base
                if getattr(sys, 'frozen', False):
                    base_dir = os.path.dirname(sys.executable)
                else:
                    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
                
                logs_dir = os.path.join(base_dir, 'logs')
                os.makedirs(logs_dir, exist_ok=True)
                
                log_file = os.path.join(logs_dir, f"sistema_relatorios_{datetime.now().strftime('%Y%m%d')}.log")
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(
                    logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
                )
                logger.addHandler(file_handler)
                logger.info(f"Log de fallback criado: {log_file}")
                
            except Exception as file_error:
                logger.warning(f"Não foi possível criar log em arquivo: {str(file_error)}")
        
        # Criar decorator simples para compatibilidade
        def log_action_fallback(description):
            def decorator(func):
                def wrapper(*args, **kwargs):
                    logger.info(f"Executando: {description}")
                    try:
                        result = func(*args, **kwargs)
                        logger.info(f"Concluído: {description}")
                        return result
                    except Exception as e:
                        logger.error(f"Erro em {description}: {str(e)}")
                        raise
                return wrapper
            return decorator
        
        logger.info("Sistema de logging fallback configurado")
        return logger, log_action_fallback
    
    except Exception as e:
        print(f"Erro crítico ao configurar logging: {str(e)}")
        
        # Último recurso: logging mínimo para console
        import logging
        logger = logging.getLogger("sistema_relatorios")
        logger.setLevel(logging.INFO)
        
        if not logger.handlers:
            handler = logging.StreamHandler()
            handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
            logger.addHandler(handler)
        
        def no_op_decorator(description):
            def decorator(func):
                return func
            return decorator
        
        logger.warning("Usando sistema de logging mínimo")
        return logger, no_op_decorator

# Configurar logging
logger, log_action = setup_logging_safe()

# Importar configurações (com fallback)
try:
    from src.config.window_config import configurar_janela
    logger.info("Configurações de janela importadas com sucesso")
except ImportError:
    logger.warning("Usando configuração de janela fallback")
    # Implementação básica caso o módulo não seja encontrado
    def configurar_janela(janela, titulo, largura=700, altura=1000):
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

# Log de inicialização
logger.info("=== Sistema de Relatórios Inicializando ===")

class SistemaRelatorios:
    """Interface centralizada para todos os relatórios do sistema"""
    
    def __init__(self, parent=None):
        """Inicializa a interface do sistema de relatórios"""
        self.parent = parent
        
        # Configurar janela principal
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        # Configurar a janela
        configurar_janela(self.root, "Sistema Integrado de Relatórios", 800, 1000)
        
        # Acompanhar quais módulos foram carregados
        self.modulos_carregados = {}
        
        # Inicializar os atributos para os comboboxes
        self.cliente_combobox = None
        self.cliente_contratos = None
        
        # Configurar interface
        self.setup_ui()
    
    def setup_ui(self):
        """Configura a interface gráfica do sistema"""
        # Frame principal dividido em esquerda e direita
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill='both', expand=True)
        
        # Frame esquerdo para lista de relatórios
        self.left_frame = ttk.LabelFrame(self.main_frame, text="Tipos de Relatórios")
        self.left_frame.pack(side='left', fill='y', padx=10, pady=10)
        
        # Frame direito para opções do relatório selecionado
        self.right_frame = ttk.LabelFrame(self.main_frame, text="Configurações do Relatório")
        self.right_frame.pack(side='right', fill='both', expand=True, padx=10, pady=10)
        
        # Lista de relatórios disponíveis
        self.setup_relatorios_list()
        
        # Frame inferior para botões de ação
        self.bottom_frame = ttk.Frame(self.root, padding=10)
        self.bottom_frame.pack(side='bottom', fill='x')
        
        # Botão para voltar ao menu principal
        ttk.Button(
            self.bottom_frame, 
            text="Voltar ao Menu Principal", 
            command=self.voltar_menu
        ).pack(side='right', padx=5)

        # Carregar lista de clientes
        self.atualizar_lista_clientes()
        
        # Configurar período inicial
        # self.alterar_periodo()
        
        # Forçar atualização da interface para garantir que todos os widgets estejam prontos
        self.root.update_idletasks()
    
    def setup_relatorios_list(self):
        """Configura a lista de relatórios disponíveis"""
        # Definir os relatórios disponíveis
        self.relatorios = [
            {
                "id": "despesas",
                "nome": "Relatório de Despesas",
                "descricao": "Relatório financeiro de despesas por cliente",
                "modulo": "relatorio_despesas_aprimorado",
                "classe": "RelatorioHandler",
                "disponivel": True
            },
            {
                "id": "contratos",
                "nome": "Relatório de Contratos e Medições",
                "descricao": "Relatório de contratos por medição e status",
                "modulo": "relatorio_contratos_medicoes",
                "classe": "RelatorioContratos",
                "disponivel": True
            },
            {
                "id": "administracao",
                "nome": "Relatório de Contratos de Administração",
                "descricao": "Relatório de contratos de administração de obra",
                "modulo": "relatorio_administracao",
                "classe": "RelatorioAdministracao",
                "disponivel": False
            },
            {
                "id": "categoria",
                "nome": "Relatório por Categoria",
                "descricao": "Análise de despesas agrupadas por categoria",
                "modulo": "relatorio_categoria",
                "classe": "RelatorioCategoria",
                "disponivel": True
            },
            {
                "id": "tipo_despesa",
                "nome": "Relatório por Tipo de Despesa",
                "descricao": "Análise detalhada por tipo de despesa",
                "modulo": "relatorio_tipo_despesa",
                "classe": "RelatorioTipoDespesa",
                "disponivel": True
            },
            {
                "id": "fornecedores",
                "nome": "Relatório de Principais Fornecedores",
                "descricao": "Resumo de fornecedores por cliente e global",
                "modulo": "relatorio_fornecedores",
                "classe": "RelatorioFornecedores",
                "disponivel": True
            },
            {
                "id": "lancamentos_pendentes",
                "nome": "Relatório de Lançamentos Pendentes",
                "descricao": "Relatório de lançamentos pendentes de múltiplos clientes",
                "modulo": "relatorio_despesas_aprimorado",
                "classe": "RelatorioLancamentosPendentes",
                "disponivel": True
            }
        ]
        
        # Criar o Treeview para a lista de relatórios
        columns = ('nome', 'status')
        self.tree_relatorios = ttk.Treeview(self.left_frame, columns=columns, show='headings', height=15)
        
        # Configurar cabeçalhos
        self.tree_relatorios.heading('nome', text='Relatório')
        self.tree_relatorios.heading('status', text='Status')
        
        # Configurar colunas
        self.tree_relatorios.column('nome', width=200)
        self.tree_relatorios.column('status', width=100, anchor='center')
        
        # Preencher a treeview
        for relatorio in self.relatorios:
            status = "Disponível" if relatorio["disponivel"] else "Em Desenvolvimento"
            self.tree_relatorios.insert('', 'end', iid=relatorio["id"], values=(relatorio["nome"], status))
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(self.left_frame, orient="vertical", command=self.tree_relatorios.yview)
        self.tree_relatorios.configure(yscrollcommand=scrollbar.set)
        
        # Colocar widgets na tela
        self.tree_relatorios.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Bind para seleção
        self.tree_relatorios.bind('<<TreeviewSelect>>', self.mostrar_opcoes_relatorio)
    
    def mostrar_opcoes_relatorio(self, event=None):
        """Mostra as opções do relatório selecionado"""
        # Limpar frame direito
        for widget in self.right_frame.winfo_children():
            widget.destroy()
        
        # Obter relatório selecionado
        selecao = self.tree_relatorios.selection()
        if not selecao:
            return
            
        rel_id = selecao[0]
        relatorio = next((r for r in self.relatorios if r["id"] == rel_id), None)
        
        if not relatorio:
            return
        
        # Mostrar informações do relatório
        ttk.Label(
            self.right_frame, 
            text=relatorio["nome"], 
            font=('Arial', 14, 'bold')
        ).pack(pady=(10,5), anchor='w')
        
        ttk.Label(
            self.right_frame, 
            text=relatorio["descricao"],
            wraplength=400
        ).pack(pady=(0,20), anchor='w')
        
        # Se o relatório não estiver disponível
        if not relatorio["disponivel"]:
            ttk.Label(
                self.right_frame,
                text="Este relatório está em desenvolvimento e ainda não está disponível.",
                foreground='red'
            ).pack(pady=20)
            return
        
        # Frame para as opções do relatório
        opcoes_frame = ttk.LabelFrame(self.right_frame, text="Opções do Relatório")
        opcoes_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Botões de ação específicos para cada tipo de relatório
        if relatorio["id"] == "despesas":
            self.setup_opcoes_despesas(opcoes_frame)
        elif relatorio["id"] == "contratos":
            self.setup_opcoes_contratos(opcoes_frame)
        elif relatorio["id"] == "categoria":
            self.setup_opcoes_categoria(opcoes_frame)
        elif relatorio["id"] == "tipo_despesa":
            self.setup_opcoes_tipo_despesa(opcoes_frame)
        elif relatorio["id"] == "fornecedores":
            self.setup_opcoes_fornecedores(opcoes_frame)  # Adicionar esta condição
        elif relatorio["id"] == "lancamentos_pendentes":  # ADICIONAR este elif
            self.setup_opcoes_lancamentos_pendentes(opcoes_frame)
        else:
            ttk.Label(
                opcoes_frame,
                text="Opções específicas para este relatório serão implementadas em breve."
            ).pack(pady=20)
        
        # Botão para gerar relatório
        btn_frame = ttk.Frame(self.right_frame)
        btn_frame.pack(fill='x', pady=20)
        
        ttk.Button(
            btn_frame,
            text="Gerar Relatório",
            command=lambda: self.gerar_relatorio(relatorio),
            style='Accentuated.TButton'
        ).pack(side='right', padx=5)
    
    def preencher_combobox_clientes(self, combobox):
        """Preenche um combobox com a lista de clientes ativos"""
        try:
            if hasattr(self, 'lista_clientes') and self.lista_clientes:
                clientes = self.lista_clientes
            else:
                clientes = self.carregar_clientes()
                self.lista_clientes = clientes  # Cache da lista
            
            combobox['values'] = clientes
            combobox.current(0)  # Selecionar "Todos os Clientes"
            
        except Exception as e:
            logger.error(f"Erro ao preencher combobox de clientes: {str(e)}")
            combobox['values'] = ['Todos os Clientes']
            combobox.current(0)

    def setup_opcoes_despesas(self, parent_frame):
        """Configura as opções específicas para relatório de despesas"""
        # Frame para data
        frame_data = ttk.Frame(parent_frame)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_data, text="Data do Relatório:").pack(side='left', padx=5)
        
        # Importar DateEntry apenas quando necessário
        try:
            from tkcalendar import DateEntry
            self.data_entry = DateEntry(
                frame_data,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR'
            )
            self.data_entry.pack(side='left', padx=5)
        except ImportError:
            # Fallback se tkcalendar não estiver instalado
            ttk.Label(frame_data, text="Módulo tkcalendar não encontrado. Data atual será usada.").pack(side='left')
        
        # Checkbox para incluir lançamentos futuros
        self.incluir_futuros = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            parent_frame,
            text="Incluir lançamentos futuros",
            variable=self.incluir_futuros
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para tipo de geração (individual ou lote)
        frame_tipo = ttk.LabelFrame(parent_frame, text="Tipo de Geração")
        frame_tipo.pack(fill='x', padx=10, pady=10)
        
        self.tipo_geracao = tk.StringVar(value="individual")
        
        # Radio button para relatório individual
        ttk.Radiobutton(
            frame_tipo,
            text="Relatório Individual",
            variable=self.tipo_geracao,
            value="individual",
            command=self.alternar_tipo_geracao
        ).pack(anchor='w', padx=15, pady=5)
        
        # Radio button para relatório em lote
        ttk.Radiobutton(
            frame_tipo,
            text="Relatório em Lote",
            variable=self.tipo_geracao,
            value="lote",
            command=self.alternar_tipo_geracao
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para seleção individual
        self.frame_individual = ttk.Frame(parent_frame)
        self.frame_individual.pack(fill='x', padx=10, pady=10)
        
        # Frame para seleção de cliente
        frame_cliente = ttk.Frame(self.frame_individual)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=40)
        self.cliente_combobox.pack(side='left', padx=5)
        
        # Preencher com clientes reais
        self.preencher_combobox_clientes(self.cliente_combobox)
        
        # Botão para selecionar arquivo individual
        ttk.Button(
            self.frame_individual,
            text="Selecionar Arquivo de Cliente",
            command=self.selecionar_arquivo_cliente
        ).pack(anchor='w', padx=15, pady=10)
        
        # Frame para seleção em lote (inicialmente oculto)
        self.frame_lote = ttk.Frame(parent_frame)
        
        # Botão para selecionar arquivos em lote
        ttk.Button(
            self.frame_lote,
            text="Selecionar Arquivos para Lote",
            command=self.selecionar_arquivos_lote
        ).pack(anchor='w', padx=15, pady=10)
        
        # Label para mostrar quantidade de arquivos selecionados
        self.lbl_arquivos_lote = ttk.Label(self.frame_lote, text="")
        self.lbl_arquivos_lote.pack(anchor='w', padx=15, pady=5)
        
        # Inicializar lista de arquivos em lote
        self.arquivos_lote = []
        
        # Frame para formato de saída
        frame_formato = ttk.LabelFrame(parent_frame, text="Formato de Saída")
        frame_formato.pack(fill='x', padx=10, pady=10)
        
        self.formato_saida = tk.StringVar(value="pdf")
        ttk.Radiobutton(
            frame_formato,
            text="PDF",
            variable=self.formato_saida,
            value="pdf"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_formato,
            text="Excel",
            variable=self.formato_saida,
            value="excel"
        ).pack(side='left', padx=20, pady=5)
        
        # Inicializar mostrando apenas a opção individual
        self.frame_lote.pack_forget()

    def alternar_tipo_geracao(self):
        """Alterna entre opções de geração individual e em lote"""
        if self.tipo_geracao.get() == "individual":
            self.frame_lote.pack_forget()
            self.frame_individual.pack(fill='x', padx=10, pady=10)
        else:
            self.frame_individual.pack_forget()
            self.frame_lote.pack(fill='x', padx=10, pady=10)

    def selecionar_arquivos_lote(self):
        """Abre diálogo para selecionar múltiplos arquivos para geração em lote"""
        arquivos = filedialog.askopenfilenames(
            title="Selecione os arquivos Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivos:
            self.arquivos_lote = arquivos
            self.lbl_arquivos_lote.config(text=f"{len(arquivos)} arquivos selecionados")

    
    def setup_opcoes_contratos(self, parent_frame):
        """Configura as opções específicas para relatório de contratos e medições"""
        # Frame para data
        frame_data = ttk.Frame(parent_frame)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_data, text="Data de Referência:").pack(side='left', padx=5)
        
        # Importar DateEntry apenas quando necessário
        try:
            from tkcalendar import DateEntry
            self.data_referencia = DateEntry(
                frame_data,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR'
            )
            self.data_referencia.pack(side='left', padx=5)
        except ImportError:
            # Fallback se tkcalendar não estiver instalado
            ttk.Label(frame_data, text="Módulo tkcalendar não encontrado. Data atual será usada.").pack(side='left')
        
        # Frame para seleção de cliente
        frame_cliente = ttk.Frame(parent_frame)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_contratos = ttk.Combobox(frame_cliente, width=40)
        self.cliente_contratos.pack(side='left', padx=5)
        
        # Preencher com clientes reais
        self.preencher_combobox_clientes(self.cliente_contratos)
        
        # Opções de visualização
        frame_visualizacao = ttk.LabelFrame(parent_frame, text="Opções de Visualização")
        frame_visualizacao.pack(fill='x', padx=10, pady=10)
        
        # Checkboxes para diferentes visualizações
        self.mostrar_resumo = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Resumo",
            variable=self.mostrar_resumo
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_detalhes = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Detalhes",
            variable=self.mostrar_detalhes
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_grafico = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Incluir Gráficos",
            variable=self.mostrar_grafico
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para formato de saída
        frame_formato = ttk.LabelFrame(parent_frame, text="Formato de Saída")
        frame_formato.pack(fill='x', padx=10, pady=10)
        
        self.formato_contratos = tk.StringVar(value="excel")
        ttk.Radiobutton(
            frame_formato,
            text="Excel",
            variable=self.formato_contratos,
            value="excel"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_formato,
            text="PDF",
            variable=self.formato_contratos,
            value="pdf"
        ).pack(side='left', padx=20, pady=5)

    def setup_opcoes_categoria(self, parent_frame):
        """Configura as opções específicas para relatório por tipo de despesa"""
        # Frame para seleção de cliente
        frame_cliente = ttk.Frame(parent_frame)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_categoria = ttk.Combobox(frame_cliente, width=40)
        self.cliente_categoria.pack(side='left', padx=5)
        
        # Preencher com clientes reais
        self.preencher_combobox_clientes(self.cliente_categoria)
        
        # Descrição do relatório
        ttk.Label(
            parent_frame,
            text="Este relatório mostra os dados agrupados por data,\n" +
                "com colunas para cada tipo de categoria e seus totais.",
            justify='center',
            font=('Arial', 10),
            foreground='gray'
        ).pack(pady=10)
        
        # Opções de visualização (opcional)
        frame_visualizacao = ttk.LabelFrame(parent_frame, text="Opções de Visualização")
        frame_visualizacao.pack(fill='x', padx=10, pady=10)
        
        # Checkboxes para diferentes visualizações
        self.mostrar_resumo_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Resumo",
            variable=self.mostrar_resumo_td
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_detalhes_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Detalhes",
            variable=self.mostrar_detalhes_td
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_grafico_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Incluir Gráficos",
            variable=self.mostrar_grafico_td
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para formato de saída
        frame_formato = ttk.LabelFrame(parent_frame, text="Formato de Saída")
        frame_formato.pack(fill='x', padx=10, pady=10)
        
        self.formato_categoria = tk.StringVar(value="excel")
        ttk.Radiobutton(
            frame_formato,
            text="Excel",
            variable=self.formato_categoria,
            value="excel"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_formato,
            text="PDF",
            variable=self.formato_categoria,
            value="pdf"
        ).pack(side='left', padx=20, pady=5)

    def setup_opcoes_tipo_despesa(self, parent_frame):
        """Configura as opções específicas para relatório por tipo de despesa"""
        # Frame para seleção de cliente
        frame_cliente = ttk.Frame(parent_frame)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_tipo_despesa = ttk.Combobox(frame_cliente, width=40)
        self.cliente_tipo_despesa.pack(side='left', padx=5)
        
        # Preencher com clientes reais
        self.preencher_combobox_clientes(self.cliente_tipo_despesa)
        
        # Descrição do relatório
        ttk.Label(
            parent_frame,
            text="Este relatório mostra os dados agrupados por data, \n" +
                "com colunas para cada tipo de despesa e seus totais.",
            justify='center',
            font=('Arial', 10),
            foreground='gray'
        ).pack(pady=10)
        
        # Opções de visualização (opcional)
        frame_visualizacao = ttk.LabelFrame(parent_frame, text="Opções de Visualização")
        frame_visualizacao.pack(fill='x', padx=10, pady=10)
        
        # Checkboxes para diferentes visualizações
        self.mostrar_resumo_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Resumo",
            variable=self.mostrar_resumo_td
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_detalhes_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Detalhes",
            variable=self.mostrar_detalhes_td
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_grafico_td = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Incluir Gráficos",
            variable=self.mostrar_grafico_td
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para formato de saída
        frame_formato = ttk.LabelFrame(parent_frame, text="Formato de Saída")
        frame_formato.pack(fill='x', padx=10, pady=10)
        
        self.formato_tipo_despesa = tk.StringVar(value="excel")
        ttk.Radiobutton(
            frame_formato,
            text="Excel",
            variable=self.formato_tipo_despesa,
            value="excel"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_formato,
            text="PDF",
            variable=self.formato_tipo_despesa,
            value="pdf"
        ).pack(side='left', padx=20, pady=5)

    def setup_opcoes_fornecedores(self, parent_frame):
        """Configura as opções específicas para relatório de fornecedores"""
        # Frame para data
        frame_data = ttk.Frame(parent_frame)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_data, text="Data de Referência:").pack(side='left', padx=5)
        
        # Importar DateEntry apenas quando necessário
        try:
            from tkcalendar import DateEntry
            self.data_referencia = DateEntry(
                frame_data,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR'
            )
            self.data_referencia.pack(side='left', padx=5)
        except ImportError:
            # Fallback se tkcalendar não estiver instalado
            ttk.Label(frame_data, text="Módulo tkcalendar não encontrado. Data atual será usada.").pack(side='left')
       # Frame para seleção de cliente
        frame_cliente = ttk.Frame(parent_frame)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_contratos = ttk.Combobox(frame_cliente, width=40)
        self.cliente_contratos.pack(side='left', padx=5)
        
        # Preencher com clientes reais
        self.preencher_combobox_clientes(self.cliente_contratos)
        
        # Opções de visualização
        frame_visualizacao = ttk.LabelFrame(parent_frame, text="Opções de Visualização")
        frame_visualizacao.pack(fill='x', padx=10, pady=10)
        
        # Checkboxes para diferentes visualizações
        self.mostrar_resumo = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Resumo",
            variable=self.mostrar_resumo
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_detalhes = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Mostrar Detalhes",
            variable=self.mostrar_detalhes
        ).pack(anchor='w', padx=15, pady=5)
        
        self.mostrar_grafico = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_visualizacao,
            text="Incluir Gráficos",
            variable=self.mostrar_grafico
        ).pack(anchor='w', padx=15, pady=5)
        
        # Frame para formato de saída
        frame_formato = ttk.LabelFrame(parent_frame, text="Formato de Saída")
        frame_formato.pack(fill='x', padx=10, pady=10)
        
        self.formato_contratos = tk.StringVar(value="excel")
        ttk.Radiobutton(
            frame_formato,
            text="Excel",
            variable=self.formato_contratos,
            value="excel"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_formato,
            text="PDF",
            variable=self.formato_contratos,
            value="pdf"
        ).pack(side='left', padx=20, pady=5)

    def setup_opcoes_lancamentos_pendentes(self, parent_frame):
        """
        Configura as opções específicas para relatório de lançamentos pendentes
        """
        # Frame para data de referência
        frame_data = ttk.Frame(parent_frame)
        frame_data.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_data, text="Data de Referência:").pack(side='left', padx=5)
        
        try:
            from tkcalendar import DateEntry
            self.data_referencia_pendentes = DateEntry(
                frame_data,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR'
            )
            self.data_referencia_pendentes.pack(side='left', padx=5)
        except ImportError:
            ttk.Label(frame_data, text="Módulo tkcalendar não encontrado.").pack(side='left')
        
        # Frame para seleção de pasta
        frame_pasta = ttk.Frame(parent_frame)
        frame_pasta.pack(fill='x', padx=10, pady=10)
        
        # Botão para selecionar pasta
        ttk.Button(
            frame_pasta,
            text="Selecionar Pasta com Arquivos",
            command=self.selecionar_pasta_lancamentos
        ).pack(side='left', padx=5)
        
        # Label para mostrar pasta selecionada
        self.pasta_selecionada_label = ttk.Label(
            frame_pasta, 
            text="Nenhuma pasta selecionada",
            wraplength=400
        )
        self.pasta_selecionada_label.pack(side='left', padx=5)
        
        # Descrição do processo
        ttk.Label(
            parent_frame,
            text="Este relatório processará todos os arquivos Excel \n"
                "na pasta selecionada e gerará um relatório consolidado\n"
                "em HTML com os lançamentos pendentes.",
            justify='center',
            font=('Arial', 10),
            foreground='gray'
        ).pack(pady=20)

    # 2. Corrigir o método selecionar_pasta_lancamentos:
    def selecionar_pasta_lancamentos(self):
        """
        Seleciona pasta com arquivos para relatório de lançamentos pendentes
        """
        pasta = filedialog.askdirectory(
            title="Selecione a pasta com os arquivos dos clientes"
        )
        if pasta:
            self.pasta_lancamentos = pasta
            # Verificar se o label existe antes de tentar atualizar
            if hasattr(self, 'pasta_selecionada_label'):
                # Mostrar apenas o nome da pasta, não o caminho completo para melhor visualização
                nome_pasta = os.path.basename(pasta) or pasta
                self.pasta_selecionada_label.config(text=f"Pasta: {nome_pasta}")
            else:
                print(f"Pasta selecionada: {pasta}")  # Fallback caso o label não exista
                messagebox.showinfo("Pasta Selecionada", f"Pasta selecionada: {pasta}")

    
    def selecionar_arquivo_cliente(self):
        """Abre diálogo para selecionar arquivo de cliente individual"""
        arquivo = filedialog.askopenfilename(
            title="Selecione o arquivo do cliente",
            filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
        )
        if arquivo:
            # Extrair nome do cliente do arquivo
            nome_arquivo = os.path.basename(arquivo)
            self.cliente_combobox.set(f"Arquivo: {nome_arquivo}")
            self.arquivo_cliente_selecionado = arquivo
    
    def carregar_modulo(self, nome_modulo):
        """Carrega ou recarrega um módulo e retorna a classe especificada"""
        try:
            print(f"Tentando carregar módulo: {nome_modulo}")
            # Se o módulo já foi carregado, recarregá-lo
            if nome_modulo in sys.modules:
                print(f"Recarregando módulo existente: {nome_modulo}")
                modulo = importlib.reload(sys.modules[nome_modulo])
            else:
                # Tentar importar do caminho atual
                try:
                    print(f"Tentando importar direto: {nome_modulo}")
                    modulo = importlib.import_module(nome_modulo)
                except ImportError as e1:
                    print(f"Erro importando direto: {str(e1)}")
                    # Tentar importar de src
                    try:
                        print(f"Tentando importar de src: src.{nome_modulo}")
                        modulo = importlib.import_module(f"src.{nome_modulo}")
                    except ImportError as e2:
                        print(f"Erro importando de src: {str(e2)}")
                        raise ImportError(f"Não foi possível importar {nome_modulo}: {str(e1)}, {str(e2)}")
            
            # Armazenar módulo carregado
            self.modulos_carregados[nome_modulo] = modulo
            print(f"Módulo carregado com sucesso: {nome_modulo}")
            return modulo
            
        except Exception as e:
            print(f"Erro ao carregar módulo {nome_modulo}: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror(
                "Erro ao carregar módulo", 
                f"Não foi possível carregar o módulo {nome_modulo}.\nErro: {str(e)}"
            )
            return None
    
    def gerar_relatorio(self, relatorio):
        """Gera o relatório selecionado"""
        try:
            # Verificar se o relatório está disponível
            if not relatorio["disponivel"]:
                messagebox.showinfo(
                    "Em desenvolvimento",
                    "Este relatório ainda está em desenvolvimento e não está disponível."
                )
                return
            
            # Para o relatório de lançamentos pendentes, usamos uma abordagem específica
            if relatorio["id"] == "lancamentos_pendentes":
                modulo = self.carregar_modulo(relatorio["modulo"])
                if not modulo:
                    return
                    
                # Obter a classe do relatório
                try:
                    classe_relatorio = getattr(modulo, relatorio["classe"])
                    # Iniciar o relatório de lançamentos pendentes
                    self.iniciar_relatorio_lancamentos_pendentes(classe_relatorio)
                    return
                except AttributeError:
                    messagebox.showerror(
                        "Erro",
                        f"Classe {relatorio['classe']} não encontrada no módulo {relatorio['modulo']}"
                    )
                    return
            
            # Para o relatório de fornecedores, usar uma abordagem mais direta
            if relatorio["id"] == "fornecedores":
                print("Iniciando relatório de fornecedores")
                self.root.withdraw()
                
                try:
                    # Importação direta
                    from relatorio_fornecedores import RelatorioFornecedores
                    app = RelatorioFornecedores(parent=self.root)
                    app.menu_principal = self.root
                    
                    # IMPORTANTE: Configurar o cliente selecionado ANTES de iniciar o mainloop
                    if hasattr(self, 'cliente_contratos') and self.cliente_contratos:
                        cliente_selecionado = self.cliente_contratos.get()
                        print(f"Cliente selecionado na interface: {cliente_selecionado}")
                        
                        if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                            # Atualizar a lista de clientes primeiro
                            app.atualizar_lista_clientes()
                            
                            # Aguardar um momento para garantir que a lista foi carregada
                            app.root.update()
                            
                            # Configurar o cliente no relatório de fornecedores
                            if cliente_selecionado in app.cliente_combobox['values']:
                                app.cliente_combobox.set(cliente_selecionado)
                                app.selecionar_cliente()
                                print(f"Cliente configurado: {cliente_selecionado}")
                            else:
                                print(f"Cliente {cliente_selecionado} não encontrado na lista")
                                
                    app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
                    app.root.lift()
                    app.root.focus_force()
                    app.root.mainloop()
                    return
                except ImportError as e:
                    try:
                        from src.relatorio_fornecedores import RelatorioFornecedores
                        app = RelatorioFornecedores(parent=self.root)
                        app.menu_principal = self.root
                        
                        # Repetir a configuração do cliente para o segundo caso de import
                        if hasattr(self, 'cliente_contratos') and self.cliente_contratos:
                            cliente_selecionado = self.cliente_contratos.get()
                            
                            if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                                app.atualizar_lista_clientes()
                                app.root.update()
                                
                                if cliente_selecionado in app.cliente_combobox['values']:
                                    app.cliente_combobox.set(cliente_selecionado)
                                    app.selecionar_cliente()
                        
                        app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
                        app.root.lift()
                        app.root.focus_force()
                        app.root.mainloop()
                        return
                    except ImportError as e:
                        messagebox.showerror(
                            "Erro", 
                            f"Não foi possível importar o módulo de relatório de fornecedores.\nErro: {str(e)}"
                        )
                        self.root.deiconify()
                        return
            
            # Código existente para outros tipos de relatório
            modulo = self.carregar_modulo(relatorio["modulo"])
            if not modulo:
                return
            
            # Obter a classe do relatório
            try:
                classe_relatorio = getattr(modulo, relatorio["classe"])
            except AttributeError:
                messagebox.showerror(
                    "Erro",
                    f"Classe {relatorio['classe']} não encontrada no módulo {relatorio['modulo']}"
                )
                return
            
            # Iniciar interface conforme o tipo de relatório
            if relatorio["id"] == "despesas":
                self.iniciar_relatorio_despesas(classe_relatorio)
            elif relatorio["id"] == "contratos":
                self.iniciar_relatorio_contratos(classe_relatorio)
            elif relatorio["id"] == "categoria":
                self.iniciar_relatorio_categoria(classe_relatorio)
            elif relatorio["id"] == "tipo_despesa":
                self.iniciar_relatorio_tipo_despesa(classe_relatorio)
            else:
                messagebox.showinfo(
                    "Em desenvolvimento",
                    "As opções específicas para este relatório ainda estão sendo implementadas."
                )
                
        except Exception as e:
            messagebox.showerror(
                "Erro", 
                f"Ocorreu um erro ao gerar o relatório.\nErro: {str(e)}"
            )
            self.root.deiconify()
    
    def iniciar_relatorio_despesas(self, classe_relatorio):
        """Inicia a geração do relatório de despesas diretamente sem abrir nova janela"""
        try:
            # Coletar dados da interface
            data_selecionada = self.data_entry.get_date() if hasattr(self, 'data_entry') else datetime.now()
            incluir_futuros = self.incluir_futuros.get() if hasattr(self, 'incluir_futuros') else True
            
            # Verificar se é geração individual ou em lote
            if self.tipo_geracao.get() == "individual":
                # Processar relatório individual
                if hasattr(self, 'arquivo_cliente_selecionado'):
                    arquivo = self.arquivo_cliente_selecionado
                else:
                    # Selecionar arquivo
                    arquivo = filedialog.askopenfilename(
                        title="Selecione o arquivo Excel",
                        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
                    )
                    if not arquivo:
                        return
                
                # Instanciar o handler
                handler = classe_relatorio()
                
                # Função de callback para exibir status
                def status_callback(msg):
                    messagebox.showinfo("Status", msg)
                
                # Gerar relatório
                handler.gerar_relatorio_direto(
                    arquivo_path=arquivo,
                    data_relatorio=data_selecionada,
                    incluir_futuros=incluir_futuros,
                    output_callback=status_callback
                )
            else:
                # Processar relatório em lote
                if not self.arquivos_lote:
                    messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para processamento em lote.")
                    return
                
                self.processar_relatorios_lote(classe_relatorio, data_selecionada, incluir_futuros)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")

    def processar_relatorios_lote(self, classe_relatorio, data_selecionada, incluir_futuros):
        """Processa geração de relatórios em lote com barra de progresso"""
        try:
            # Criar janela de progresso
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Gerando Relatórios em Lote")
            progress_window.geometry("700x550")
            progress_window.transient(self.root)
            
            # Frame principal
            main_frame = ttk.Frame(progress_window, padding=20)
            main_frame.pack(fill='both', expand=True)
            
            # Label para mostrar progresso
            progress_label = ttk.Label(main_frame, text="Iniciando processamento...", font=('Arial', 12))
            progress_label.pack(pady=10)
            
            # Barra de progresso
            progress_bar = ttk.Progressbar(main_frame, length=600, mode='determinate')
            progress_bar.pack(pady=20)
            
            # Lista de resultados
            result_frame = ttk.LabelFrame(main_frame, text="Relatórios Processados")
            result_frame.pack(fill='both', expand=True, pady=10)
            
            result_list = tk.Listbox(result_frame, font=('Courier', 10), height=15)
            scrollbar = ttk.Scrollbar(result_frame, orient='vertical', command=result_list.yview)
            result_list.configure(yscrollcommand=scrollbar.set)
            result_list.pack(side='left', fill='both', expand=True, padx=5, pady=5)
            scrollbar.pack(side='right', fill='y')
            
            # Configurar barra de progresso
            total_arquivos = len(self.arquivos_lote)
            progress_bar['maximum'] = total_arquivos
            
            # Instanciar o handler
            handler = classe_relatorio()
            
            # Processar cada arquivo
            for i, arquivo in enumerate(self.arquivos_lote, 1):
                try:
                    nome_arquivo = os.path.basename(arquivo)
                    progress_label.config(text=f"Processando {i}/{total_arquivos}: {nome_arquivo}")
                    progress_bar['value'] = i - 0.5
                    progress_window.update()
                    
                    # Gerar relatório
                    resultado = handler.gerar_relatorio_direto(
                        arquivo_path=arquivo,
                        data_relatorio=data_selecionada,
                        incluir_futuros=incluir_futuros
                    )
                    
                    # Atualizar lista de resultados
                    status = "✓" if resultado else "✗"
                    result_list.insert(tk.END, f"{status} {nome_arquivo}")
                    result_list.itemconfig(tk.END, fg="green" if resultado else "red")
                    result_list.see(tk.END)
                    
                    # Atualizar barra de progresso
                    progress_bar['value'] = i
                    progress_window.update()
                    
                except Exception as e:
                    # Registrar erro na lista
                    result_list.insert(tk.END, f"✗ {nome_arquivo} - Erro: {str(e)}")
                    result_list.itemconfig(tk.END, fg="red")
                    result_list.see(tk.END)
                    continue
            
            # Finalização
            progress_label.config(text="Processamento concluído!")
            
            # Botão para fechar
            ttk.Button(
                main_frame,
                text="Fechar",
                command=progress_window.destroy
            ).pack(pady=20)
            
            # Tornar a janela modal
            progress_window.grab_set()
            progress_window.focus_set()
            progress_window.wait_window()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar relatórios em lote: {str(e)}")
            if 'progress_window' in locals():
                progress_window.destroy()
    
    def iniciar_relatorio_contratos(self, classe_relatorio):
        """Inicia a geração do relatório de contratos e medições"""
        # Esconder a janela atual
        self.root.withdraw()
        
        # Inicializar o relatório passando a janela atual como parent
        app_relatorio = classe_relatorio(self.root)
        
        # Verificar se app_relatorio tem os atributos esperados
        if not hasattr(app_relatorio, 'root'):
            messagebox.showerror(
                "Erro", 
                "Erro ao inicializar relatório. A classe do relatório não retornou o objeto esperado."
            )
            self.root.deiconify()
            return
        
        # Configurar menu principal para retornar
        app_relatorio.menu_principal = self.root
        
        # Se houver cliente selecionado e o método adequado existir, selecioná-lo
        if hasattr(app_relatorio, 'cliente_combobox') and self.cliente_contratos.get() != 'Todos os Clientes':
            try:
                app_relatorio.cliente_combobox.set(self.cliente_contratos.get())
                # Se existir um método específico para selecionar o cliente, chamá-lo
                if hasattr(app_relatorio, 'selecionar_cliente'):
                    app_relatorio.selecionar_cliente()
            except Exception as e:
                logger.warning(f"Não foi possível selecionar o cliente: {str(e)}")
        
        # Se houver data selecionada, configurá-la
        if hasattr(self, 'data_referencia') and hasattr(app_relatorio, 'data_entry'):
            try:
                app_relatorio.data_entry.set_date(self.data_referencia.get_date())
            except Exception as e:
                logger.warning(f"Não foi possível configurar a data: {str(e)}")
        
        # Configurar comportamento ao fechar
        app_relatorio.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app_relatorio.root))
        
        # Exibir janela
        app_relatorio.root.lift()
        app_relatorio.root.focus_force()
        app_relatorio.root.mainloop()

    def iniciar_relatorio_categoria(self, classe_relatorio):
        """Inicia a geração do relatório por tipo de despesa"""
        # Esconder a janela atual
        self.root.withdraw()
        
        # Inicializar o relatório passando a janela atual como parent
        app_relatorio = classe_relatorio(self.root)
        
        # Verificar se app_relatorio tem os atributos esperados
        if not hasattr(app_relatorio, 'root'):
            messagebox.showerror(
                "Erro", 
                "Erro ao inicializar relatório. A classe do relatório não retornou o objeto esperado."
            )
            self.root.deiconify()
            return
        
        # Configurar menu principal para retornar
        app_relatorio.menu_principal = self.root
        
        # Se houver cliente selecionado, configurá-lo
        if hasattr(app_relatorio, 'cliente_combobox') and hasattr(self, 'cliente_categoria'):
            cliente_selecionado = self.cliente_categoria.get()
            if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                try:
                    # Atualizar lista de clientes primeiro
                    app_relatorio.atualizar_lista_clientes()
                    
                    # Configurar o cliente no combobox
                    app_relatorio.cliente_combobox.set(cliente_selecionado)
                    
                    # Chamar o método para selecionar cliente
                    if hasattr(app_relatorio, 'selecionar_cliente'):
                        app_relatorio.selecionar_cliente()
                except Exception as e:
                    logger.warning(f"Não foi possível selecionar o cliente: {str(e)}")
        
        # Configurar comportamento ao fechar
        app_relatorio.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app_relatorio.root))
        
        # Exibir janela
        app_relatorio.root.lift()
        app_relatorio.root.focus_force()
        app_relatorio.root.mainloop()

    def iniciar_relatorio_tipo_despesa(self, classe_relatorio):
        """Inicia a geração do relatório por tipo de despesa"""
        # Esconder a janela atual
        self.root.withdraw()
        
        # Inicializar o relatório passando a janela atual como parent
        app_relatorio = classe_relatorio(self.root)
        
        # Verificar se app_relatorio tem os atributos esperados
        if not hasattr(app_relatorio, 'root'):
            messagebox.showerror(
                "Erro", 
                "Erro ao inicializar relatório. A classe do relatório não retornou o objeto esperado."
            )
            self.root.deiconify()
            return
        
        # Configurar menu principal para retornar
        app_relatorio.menu_principal = self.root
        
        # Se houver cliente selecionado, configurá-lo
        if hasattr(app_relatorio, 'cliente_combobox') and hasattr(self, 'cliente_tipo_despesa'):
            cliente_selecionado = self.cliente_tipo_despesa.get()
            if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                try:
                    # Atualizar lista de clientes primeiro
                    app_relatorio.atualizar_lista_clientes()
                    
                    # Configurar o cliente no combobox
                    app_relatorio.cliente_combobox.set(cliente_selecionado)
                    
                    # Chamar o método para selecionar cliente
                    if hasattr(app_relatorio, 'selecionar_cliente'):
                        app_relatorio.selecionar_cliente()
                except Exception as e:
                    logger.warning(f"Não foi possível selecionar o cliente: {str(e)}")
        
        # Configurar comportamento ao fechar
        app_relatorio.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app_relatorio.root))
        
        # Exibir janela
        app_relatorio.root.lift()
        app_relatorio.root.focus_force()
        app_relatorio.root.mainloop()

    def iniciar_relatorio_fornecedores(self, classe_relatorio):
        """Inicia a geração do relatório de fornecedores"""
        try:
            print("Iniciando método iniciar_relatorio_fornecedores")
            # Esconder a janela atual
            self.root.withdraw()
            
            # Criar uma nova janela para o relatório
            print("Criando instância do relatório de fornecedores")
            app_relatorio = classe_relatorio(self.root)
            
            # Configurar menu principal para retornar
            app_relatorio.menu_principal = self.root
            
            # IMPORTANTE: Verificar se há um cliente selecionado e configurá-lo
            if hasattr(self, 'cliente_contratos') and self.cliente_contratos:
                cliente_selecionado = self.cliente_contratos.get()
                if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                    try:
                        # Configurar o cliente no relatório de fornecedores
                        print(f"Configurando cliente: {cliente_selecionado}")
                        app_relatorio.cliente_combobox.set(cliente_selecionado)
                        
                        # Chamar o método selecionar_cliente diretamente
                        app_relatorio.cliente_atual = cliente_selecionado
                        app_relatorio.arquivo_cliente = PASTA_CLIENTES / f"{cliente_selecionado}.xlsx"
                        app_relatorio.lbl_cliente_resumo.config(text=f"Cliente: {cliente_selecionado}")
                        
                        # Desmarcar checkbox de todos os clientes
                        app_relatorio.var_todos_clientes.set(False)
                        app_relatorio.todos_clientes = False
                        
                    except Exception as e:
                        print(f"Erro ao configurar cliente: {str(e)}")
            
            # Configurar comportamento ao fechar
            app_relatorio.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app_relatorio.root))
            
            # Exibir janela
            app_relatorio.root.lift()
            app_relatorio.root.focus_force()
            print("Iniciando mainloop do relatório de fornecedores")
            app_relatorio.root.mainloop()
        except Exception as e:
            import traceback
            print(f"Erro em iniciar_relatorio_fornecedores: {str(e)}")
            traceback.print_exc()
            messagebox.showerror(
                "Erro", 
                f"Ocorreu um erro ao iniciar o relatório de fornecedores.\nErro: {str(e)}"
            )
            self.root.deiconify()

    def iniciar_relatorio_lancamentos_pendentes(self, classe_relatorio):
        """
        Inicia a geração do relatório de lançamentos pendentes
        """
        try:
            # Verificar se pasta foi selecionada
            if not hasattr(self, 'pasta_lancamentos'):
                messagebox.showerror("Erro", "Por favor, selecione uma pasta primeiro.")
                return
            
            # Data de referência
            data_ref = self.data_referencia_pendentes.get_date() if hasattr(self, 'data_referencia_pendentes') else datetime.now()
        
            # Garantir que data_ref é datetime e não apenas date
            if isinstance(data_ref, date) and not isinstance(data_ref, datetime):
                data_ref = datetime.combine(data_ref, datetime.min.time())
            
            # Instanciar relatório
            relatorio = classe_relatorio()
            
            # Gerar relatório
            arquivo_saida = os.path.join(self.pasta_lancamentos, "relatorio_lancamentos_pendentes.html")
            
            # Usar o método gerar_relatorio_pendentes que já existe na classe
            if relatorio.gerar_relatorio_pendentes(self.pasta_lancamentos, arquivo_saida, data_ref):
                messagebox.showinfo(
                    "Sucesso",
                    f"Relatório gerado com sucesso!\nSalvo em: {arquivo_saida}"
                )
            else:
                messagebox.showwarning(
                    "Aviso",
                    "Nenhum lançamento pendente encontrado."
                )
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
        
    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        janela.destroy()
        self.root.deiconify()
        self.root.lift()
    
    def voltar_menu(self):
        """Volta ao menu principal"""
        print("Finalizando interface de relatórios...")
        
        # Destruir a janela
        self.root.destroy()
        
        # Mostrar janela principal
        if self.menu_principal:
            self.menu_principal.deiconify()
            self.menu_principal.lift()
            self.menu_principal.focus_force()
    
    def carregar_clientes(self):
        """Carrega a lista de clientes ativos do arquivo de clientes"""
        try:
            # Importar bibliotecas necessárias
            import pandas as pd
            from openpyxl import load_workbook
            
            # Caminho para o arquivo de clientes
            try:
                from src.config.config import ARQUIVO_CLIENTES
                logger.info(f"Carregando clientes de: {ARQUIVO_CLIENTES}")
            except ImportError:
                # Caminho padrão se não conseguir importar das configurações
                ARQUIVO_CLIENTES = "dados/clientes.xlsx"
                logger.warning(f"Usando caminho padrão para clientes: {ARQUIVO_CLIENTES}")
            
            # Verificar se o arquivo existe
            if not os.path.exists(ARQUIVO_CLIENTES):
                logger.warning(f"Arquivo de clientes não encontrado: {ARQUIVO_CLIENTES}")
                return ['Todos os Clientes']
            
            # Carregar o arquivo usando pandas
            try:
                # Ler o arquivo Excel
                df = pd.read_excel(ARQUIVO_CLIENTES, sheet_name='Clientes')
                
                # Debug: mostrar as colunas disponíveis
                logger.info(f"Colunas disponíveis: {df.columns.tolist()}")
                
                # Verificar se a coluna E existe (coluna 4 em índice baseado em 0)
                # Ou verificar pelo nome da coluna se existir
                if len(df.columns) >= 5:  # Verifica se tem pelo menos 5 colunas (A-E)
                    # Filtrar clientes ativos (coluna E vazia)
                    coluna_status = df.columns[4]  # Coluna E (índice 4)
                    logger.info(f"Coluna de status: {coluna_status}")
                    
                    # Considera como vazio: None, NaN, '', etc.
                    df_ativos = df[df[coluna_status].isna() | (df[coluna_status] == '')]
                    
                    # Verificar se a primeira coluna contém os nomes dos clientes
                    coluna_nome = df.columns[0]  # Coluna A
                    logger.info(f"Coluna de nome: {coluna_nome}")
                    
                    # Extrair nomes dos clientes ativos (assumindo que estão na primeira coluna)
                    clientes_ativos = df_ativos[coluna_nome].dropna().tolist()
                    
                    logger.info(f"Total de clientes ativos encontrados: {len(clientes_ativos)}")
                    
                    # Ordenar alfabeticamente
                    clientes_ativos.sort()
                    
                    # Adicionar "Todos os Clientes" no início
                    clientes = ['Todos os Clientes'] + clientes_ativos
                    
                    return clientes
                else:
                    logger.warning("Arquivo não tem colunas suficientes (precisa de pelo menos 5 colunas - A até E)")
                    return ['Todos os Clientes']
                
            except Exception as e:
                logger.error(f"Erro ao ler arquivo Excel com pandas: {str(e)}")
                # Tentar com openpyxl como fallback
                try:
                    workbook = load_workbook(ARQUIVO_CLIENTES)
                    sheet = workbook['Clientes']
                    
                    clientes = ['Todos os Clientes']
                    for row in sheet.iter_rows(min_row=2, values_only=True):
                        # Verifica se a coluna E (índice 4) está vazia
                        if row[0] and (len(row) < 5 or not row[4]):
                            clientes.append(row[0])
                    
                    workbook.close()
                    clientes.sort()  # Ordenar alfabeticamente (mantendo "Todos os Clientes" primeiro)
                    return clientes
                    
                except Exception as inner_e:
                    logger.error(f"Erro ao ler arquivo Excel com openpyxl: {str(inner_e)}")
                    return ['Todos os Clientes']
                
        except Exception as e:
            logger.error(f"Erro ao carregar clientes: {str(e)}", exc_info=True)
            return ['Todos os Clientes']

    # Também vamos adicionar um método para atualizar o combobox quando necessário
    def atualizar_lista_clientes(self):
        """Atualiza a lista de clientes na combobox"""
        try:
            clientes = self.carregar_clientes()
            
            # Atualizar todos os comboboxes que mostram clientes
            if hasattr(self, 'cliente_combobox') and self.cliente_combobox is not None:
                self.cliente_combobox['values'] = clientes
                self.cliente_combobox.current(0)  # Selecionar "Todos os Clientes"
            
            if hasattr(self, 'cliente_contratos') and self.cliente_contratos is not None:
                self.cliente_contratos['values'] = clientes
                self.cliente_contratos.current(0)
                
            logger.info(f"Lista de clientes atualizada com {len(clientes)} clientes")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar lista de clientes: {str(e)}")

    # E adicionar um botão para recarregar a lista de clientes (opcional)
    def adicionar_botao_atualizar_clientes(self, parent_frame):
        """Adiciona botão para atualizar a lista de clientes"""
        ttk.Button(
            parent_frame,
            text="Atualizar Lista de Clientes",
            command=self.atualizar_lista_clientes
        ).pack(side='right', padx=5, pady=5)
    
    def selecionar_cliente_nome(self, nome_cliente):
        """Método stub para selecionar cliente por nome"""
        pass
    
    def selecionar_arquivo_direto(self, caminho_arquivo):
        """Método stub para selecionar arquivo diretamente"""
        pass
    
    def run(self):
        """Inicia a execução do sistema"""
        try:
            # Pré-carregar lista de clientes
            self.lista_clientes = self.carregar_clientes()
            logger.info(f"Lista de clientes carregada com {len(self.lista_clientes)} itens")
            
        except Exception as e:
            logger.error(f"Erro ao carregar lista de clientes: {str(e)}", exc_info=True)
            self.lista_clientes = ['Todos os Clientes']
        
        # Configurar estilos
        style = ttk.Style()
        style.configure('Accentuated.TButton', font=('Arial', 11, 'bold'))
        
        # Iniciar mainloop
        self.root.mainloop()

# Função para executar o sistema como módulo independente
def main():
    app = SistemaRelatorios()
    app.run()

if __name__ == "__main__":
    main()