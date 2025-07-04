import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib
import logging
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
from pathlib import Path

# from correcoes_emergenciais import aplicar_todas_correcoes 
# aplicar_todas_correcoes()

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

        # Inicializar variáveis para controle do relatório de despesas
        self.arquivo_cliente_selecionado = None
        self.arquivos_lote = []
        self.pasta_lancamentos = None
        
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

        # Configurar validações
        self.backup_metodo_original()
        
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
        """Versão corrigida que usa estrutura original"""
        
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
        
        # Configurar opções específicas (CÓDIGO ORIGINAL)
        if relatorio["id"] == "despesas":
            self.setup_opcoes_despesas(opcoes_frame)
        elif relatorio["id"] == "contratos":
            self.setup_opcoes_contratos(opcoes_frame)
        elif relatorio["id"] == "categoria":
            self.setup_opcoes_categoria(opcoes_frame)
        elif relatorio["id"] == "tipo_despesa":
            self.setup_opcoes_tipo_despesa(opcoes_frame)
        elif relatorio["id"] == "fornecedores":
            self.setup_opcoes_fornecedores(opcoes_frame)
        elif relatorio["id"] == "lancamentos_pendentes":
            self.setup_opcoes_lancamentos_pendentes(opcoes_frame)
        else:
            ttk.Label(
                opcoes_frame,
                text="Opções específicas para este relatório serão implementadas em breve."
            ).pack(pady=20)
        
        # === ÚNICA MUDANÇA: BOTÃO PERSONALIZADO APENAS PARA DESPESAS ===
        btn_frame = ttk.Frame(self.right_frame)
        btn_frame.pack(fill='x', pady=20)
        
        if relatorio["id"] == "despesas":
            # NOVO: Botão otimizado para despesas
            # Label explicativo
            ttk.Label(
                btn_frame,
                text="💡 O sistema processará os dados e abrirá diretamente o preview ou gerará o PDF conforme configurado.",
                font=('Arial', 9),
                foreground='blue',
                wraplength=400
            ).pack(pady=(0, 10))
            
            # Botão de validação (opcional)
            ttk.Button(
                btn_frame,
                text="✅ Validar Configurações",
                command=lambda: self.validar_e_mostrar_resumo(),
                style='TButton'
            ).pack(side='left', padx=5)
            
            # Botão principal otimizado
            ttk.Button(
                btn_frame,
                text="🚀 Processar e Gerar Relatório",
                command=lambda: self.gerar_relatorio(relatorio),
                style='Accentuated.TButton'
            ).pack(side='right', padx=5)
            
        else:
            # ORIGINAL: Botão padrão para outros relatórios
            ttk.Button(
                btn_frame,
                text="Gerar Relatório",
                command=lambda: self.gerar_relatorio(relatorio),
                style='Accentuated.TButton'
            ).pack(side='right', padx=5)

    def criar_botao_despesas_otimizado(self, btn_frame, relatorio):
        """Cria botão otimizado específico para relatório de despesas"""
        
        # Label explicativo
        info_label = ttk.Label(
            btn_frame,
            text="💡 O sistema processará os dados e abrirá diretamente o preview ou gerará o PDF conforme configurado.",
            font=('Arial', 9),
            foreground='blue',
            wraplength=400
        )
        info_label.pack(pady=(0, 10))
        
        # Botão principal otimizado
        botao_principal = ttk.Button(
            btn_frame,
            text="🚀 Processar e Gerar Relatório",
            command=lambda: self.gerar_relatorio(relatorio),
            style='Accentuated.TButton'
        )
        botao_principal.pack(side='right', padx=5)
        
        # OPCIONAL: Botão de validação prévia
        botao_validar = ttk.Button(
            btn_frame,
            text="✅ Validar Configurações",
            command=lambda: self.validar_e_mostrar_resumo(),
            style='TButton'
        )
        botao_validar.pack(side='left', padx=5)

    def criar_botao_padrao(self, btn_frame, relatorio):
        """Cria botão padrão para outros tipos de relatório"""
        
        # Botão padrão (comportamento original)
        ttk.Button(
            btn_frame,
            text="Gerar Relatório",
            command=lambda: self.gerar_relatorio(relatorio),
            style='Accentuated.TButton'
        ).pack(side='right', padx=5)

    def validar_e_mostrar_resumo(self):
        """Valida configurações e mostra resumo antes da geração"""
        try:
            # Validar configurações
            if not self.validar_configuracoes_despesas():
                return
            
            # Coletar configurações
            configuracoes = self.coletar_configuracoes_completas()
            
            # Gerar resumo
            resumo = self.gerar_resumo_configuracoes(configuracoes)
            
            # Mostrar resumo em janela separada
            self.mostrar_janela_resumo(resumo, configuracoes)
            
        except Exception as e:
            logger.error(f"Erro na validação prévia: {str(e)}")
            messagebox.showerror("Erro", f"Erro na validação: {str(e)}")

    def mostrar_janela_resumo(self, resumo, configuracoes):
        """Mostra janela com resumo das configurações"""
        
        # Criar janela
        resumo_window = tk.Toplevel(self.root)
        resumo_window.title("Resumo das Configurações")
        resumo_window.geometry("500x400")
        resumo_window.transient(self.root)
        resumo_window.grab_set()
        
        # Frame principal
        main_frame = ttk.Frame(resumo_window, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            main_frame,
            text="📋 Resumo das Configurações",
            font=('Arial', 14, 'bold')
        ).pack(pady=(0, 20))
        
        # Área de texto com scroll
        text_frame = ttk.Frame(main_frame)
        text_frame.pack(fill='both', expand=True)
        
        text_widget = tk.Text(
            text_frame,
            wrap='word',
            font=('Courier', 10),
            state='disabled'
        )
        
        scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
        text_widget.configure(yscrollcommand=scrollbar.set)
        
        text_widget.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Inserir resumo
        text_widget.config(state='normal')
        text_widget.insert('1.0', resumo)
        text_widget.config(state='disabled')
        
        # Frame para botões
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill='x', pady=(20, 0))
        
        # Botões
        ttk.Button(
            btn_frame,
            text="❌ Cancelar",
            command=resumo_window.destroy
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame,
            text="✏️ Editar Configurações",
            command=resumo_window.destroy
        ).pack(side='left', padx=5)
        
        ttk.Button(
            btn_frame,
            text="🚀 Continuar com Geração",
            command=lambda: self.continuar_geracao_apos_resumo(resumo_window, configuracoes)
        ).pack(side='right', padx=5)

    def continuar_geracao_apos_resumo(self, resumo_window, configuracoes):
        """Continua com a geração após confirmação do resumo"""
        try:
            resumo_window.destroy()
            
            # Criar o relatório mock para compatibilidade
            relatorio_mock = {
                "id": "despesas",
                "nome": "Relatório de Despesas",
                "disponivel": True
            }
            
            # Proceder com geração otimizada
            self.gerar_relatorio(relatorio_mock)
            
        except Exception as e:
            logger.error(f"Erro ao continuar geração: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")
    
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

    def calcular_data_rel_automatica(self):
        """Calcula automaticamente a data do relatório baseado na regra dos dias 5 e 20"""
        try:
            hoje = datetime.now()
            
            if 6 <= hoje.day <= 20:
                # Entre dia 6 e 20: relatório do dia 20 do mês atual
                data_rel = hoje.replace(day=20)
            else:
                if hoje.day > 20:
                    # Após dia 20: relatório do dia 5 do próximo mês
                    data_rel = (hoje + relativedelta(months=1)).replace(day=5)
                else:
                    # Antes do dia 6: relatório do dia 5 do mês atual
                    data_rel = hoje.replace(day=5)
            
            logger.info(f"Data calculada automaticamente: {data_rel.strftime('%d/%m/%Y')}")
            return data_rel
            
        except Exception as e:
            logger.error(f"Erro ao calcular data automática: {str(e)}")
            # Fallback: retorna data atual
            return datetime.now()

    def explicar_regra_data(self):
        """Retorna explicação da regra de cálculo de data"""
        hoje = datetime.now()
        data_calculada = self.calcular_data_rel_automatica()
        
        if 6 <= hoje.day <= 20:
            explicacao = f"📅 Hoje é dia {hoje.day}: período para relatório do dia 20"
        elif hoje.day > 20:
            explicacao = f"📅 Hoje é dia {hoje.day}: período para relatório do dia 5 do próximo mês"
        else:
            explicacao = f"📅 Hoje é dia {hoje.day}: período para relatório do dia 5"
        
        return f"{explicacao}\n🎯 Data calculada: {data_calculada.strftime('%d/%m/%Y')}"

    def validar_data_relatorio(self, data_selecionada):
        """Valida se a data selecionada está correta conforme a regra"""
        try:
            if isinstance(data_selecionada, str):
                data_selecionada = datetime.strptime(data_selecionada, '%d/%m/%Y')
            
            # Verificar se é dia 5 ou 20
            if data_selecionada.day not in [5, 20]:
                return False, f"❌ Data deve ser dia 5 ou 20 do mês.\nData selecionada: {data_selecionada.strftime('%d/%m/%Y')}"
            
            # Verificar se está no período correto
            data_automatica = self.calcular_data_rel_automatica()
            
            if data_selecionada.date() == data_automatica.date():
                return True, f"✅ Data correta para o período atual"
            else:
                return True, f"⚠️ Data válida, mas não é a sugerida para hoje.\nSugerida: {data_automatica.strftime('%d/%m/%Y')}"
            
        except Exception as e:
            return False, f"❌ Erro ao validar data: {str(e)}"

    def setup_opcoes_despesas(self, parent_frame):
        """Versão otimizada com seleção de cliente via combobox"""
        
        # Frame para data com cálculo automático
        frame_data = ttk.LabelFrame(parent_frame, text="Data do Relatório")
        frame_data.pack(fill='x', padx=10, pady=10)
        
        # Calcular data automática
        data_automatica = self.calcular_data_rel_automatica()
        
        # Área de informações sobre a regra
        info_frame = ttk.Frame(frame_data)
        info_frame.pack(fill='x', padx=10, pady=5)
        
        # Label explicativa
        explicacao = self.explicar_regra_data()
        ttk.Label(info_frame, text=explicacao, font=('Arial', 9), foreground='blue').pack(anchor='w')
        
        # Frame para seleção de data
        selecao_frame = ttk.Frame(frame_data)
        selecao_frame.pack(fill='x', padx=10, pady=5)
        
        # Opção de usar data automática (padrão)
        self.usar_data_automatica = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(
            selecao_frame,
            text="Usar data calculada automaticamente",
            variable=self.usar_data_automatica,
            command=self.alternar_modo_data
        ).pack(anchor='w', pady=2)
        
        # Frame para data manual (inicialmente oculto)
        self.frame_data_manual = ttk.Frame(frame_data)
        
        ttk.Label(self.frame_data_manual, text="Data manual:").pack(side='left', padx=5)
        
        # DateEntry para seleção manual
        try:
            from tkcalendar import DateEntry
            self.data_entry = DateEntry(
                self.frame_data_manual,
                width=12,
                background='darkblue',
                foreground='white',
                borderwidth=2,
                date_pattern='dd/mm/yyyy',
                locale='pt_BR'
            )
            self.data_entry.pack(side='left', padx=5)
            
            # Botão para validar data manual
            ttk.Button(
                self.frame_data_manual,
                text="Validar Data",
                command=self.validar_data_manual
            ).pack(side='left', padx=5)
            
        except ImportError:
            ttk.Label(
                self.frame_data_manual, 
                text="Módulo tkcalendar não encontrado"
            ).pack(side='left')
        
        # Configurar data inicial
        self.data_automatica_calculada = data_automatica
        if hasattr(self, 'data_entry'):
            self.data_entry.set_date(data_automatica)

        # === NOVA SEÇÃO: SELEÇÃO DE CLIENTE ===
        frame_cliente = ttk.LabelFrame(parent_frame, text="Seleção de Cliente")
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        # Frame interno para organizar melhor
        cliente_inner_frame = ttk.Frame(frame_cliente)
        cliente_inner_frame.pack(fill='x', padx=10, pady=10)
        
        # Label e Combobox de cliente
        ttk.Label(cliente_inner_frame, text="Cliente:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0, 5))
        
        self.cliente_combobox = ttk.Combobox(
            cliente_inner_frame, 
            width=50,
            state='readonly',  # Apenas seleção, não digitação
            font=('Arial', 10)
        )
        self.cliente_combobox.pack(fill='x', pady=(0, 10))
        
        # Preencher combobox com clientes
        self.preencher_combobox_clientes(self.cliente_combobox)
        
        # Bind para evento de seleção
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.on_cliente_selecionado)
        
        # Label para mostrar status da seleção
        self.status_cliente_label = ttk.Label(
            cliente_inner_frame, 
            text="Selecione um cliente para continuar",
            font=('Arial', 9),
            foreground='gray'
        )
        self.status_cliente_label.pack(anchor='w', pady=(0, 10))
        
        # Frame para botões adicionais de cliente
        botoes_cliente_frame = ttk.Frame(cliente_inner_frame)
        botoes_cliente_frame.pack(fill='x')
        
        # Botão para atualizar lista de clientes
        ttk.Button(
            botoes_cliente_frame,
            text="🔄 Atualizar Lista",
            command=self.atualizar_lista_clientes_despesas,
            width=15
        ).pack(side='left', padx=(0, 10))
        
        # Botão para seleção manual de arquivo (fallback)
        ttk.Button(
            botoes_cliente_frame,
            text="📁 Selecionar Arquivo Manual",
            command=self.selecionar_arquivo_manual_despesas,
            width=25
        ).pack(side='left')
        
        # === OPÇÕES DE PROCESSAMENTO ===
        frame_opcoes = ttk.LabelFrame(parent_frame, text="Opções de Processamento")
        frame_opcoes.pack(fill='x', padx=10, pady=10)
        
        # Checkbox para incluir lançamentos futuros
        self.incluir_futuros = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            frame_opcoes,
            text="Incluir lançamentos futuros",
            variable=self.incluir_futuros
        ).pack(anchor='w', padx=15, pady=5)
        
        # Checkbox para incluir lançamentos excluídos
        self.incluir_excluidos = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame_opcoes,
            text="Incluir lançamentos excluídos no relatório",
            variable=self.incluir_excluidos
        ).pack(anchor='w', padx=15, pady=5)
        
        # === TIPO DE GERAÇÃO ===
        frame_tipo = ttk.LabelFrame(parent_frame, text="Tipo de Geração")
        frame_tipo.pack(fill='x', padx=10, pady=10)
        
        self.tipo_geracao = tk.StringVar(value="individual")
        
        ttk.Radiobutton(
            frame_tipo,
            text="Relatório Individual",
            variable=self.tipo_geracao,
            value="individual",
            command=self.alternar_tipo_geracao
        ).pack(anchor='w', padx=15, pady=5)
        
        ttk.Radiobutton(
            frame_tipo,
            text="Relatório em Lote",
            variable=self.tipo_geracao,
            value="lote",
            command=self.alternar_tipo_geracao
        ).pack(anchor='w', padx=15, pady=5)
        
        # === FRAMES PARA TIPOS ESPECÍFICOS ===
        
        # Frame para seleção individual (já preenchido com cliente selecionado)
        self.frame_individual = ttk.Frame(parent_frame)
        self.frame_individual.pack(fill='x', padx=10, pady=10)
        
        # Label de status para individual
        self.status_individual_label = ttk.Label(
            self.frame_individual,
            text="Cliente será selecionado através da combobox acima",
            font=('Arial', 9),
            foreground='blue'
        )
        self.status_individual_label.pack(anchor='w', padx=15, pady=5)
        
        # Frame para seleção em lote
        self.frame_lote = ttk.Frame(parent_frame)
        
        ttk.Button(
            self.frame_lote,
            text="Selecionar Arquivos para Lote",
            command=self.selecionar_arquivos_lote
        ).pack(anchor='w', padx=15, pady=10)
        
        self.lbl_arquivos_lote = ttk.Label(self.frame_lote, text="")
        self.lbl_arquivos_lote.pack(anchor='w', padx=15, pady=5)
        
        self.arquivos_lote = []
        
        # === MODO DE VISUALIZAÇÃO ===
        frame_visualizacao = ttk.LabelFrame(parent_frame, text="Modo de Visualização")
        frame_visualizacao.pack(fill='x', padx=10, pady=10)
        
        self.modo_visualizacao = tk.StringVar(value="preview")
        ttk.Radiobutton(
            frame_visualizacao,
            text="Gerar com Preview",
            variable=self.modo_visualizacao,
            value="preview"
        ).pack(side='left', padx=20, pady=5)
        
        ttk.Radiobutton(
            frame_visualizacao,
            text="Gerar Direto",
            variable=self.modo_visualizacao,
            value="direto"
        ).pack(side='left', padx=20, pady=5)
        
        # === FORMATO DE SAÍDA ===
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
        
        # Configurar variáveis de controle
        self.arquivo_cliente_selecionado = None
        self.cliente_atual = None

    def on_cliente_selecionado(self, event=None):
        """Trata a seleção de um cliente na combobox"""
        try:
            cliente_selecionado = self.cliente_combobox.get()
            logger.info(f"Cliente selecionado: {cliente_selecionado}")
            
            if not cliente_selecionado or cliente_selecionado == 'Todos os Clientes':
                self.limpar_selecao_cliente()
                return
            
            # Buscar arquivo do cliente
            caminho_arquivo = self.buscar_arquivo_cliente(cliente_selecionado)
            
            if caminho_arquivo and os.path.exists(caminho_arquivo):
                self.arquivo_cliente_selecionado = caminho_arquivo
                self.cliente_atual = cliente_selecionado
                
                # Atualizar status
                self.status_cliente_label.config(
                    text=f"✅ Cliente: {cliente_selecionado} | Arquivo: {os.path.basename(caminho_arquivo)}",
                    foreground='green'
                )
                
                # Atualizar status individual
                if hasattr(self, 'status_individual_label'):
                    self.status_individual_label.config(
                        text=f"✅ Arquivo selecionado: {os.path.basename(caminho_arquivo)}",
                        foreground='green'
                    )
                
                logger.info(f"Arquivo encontrado: {caminho_arquivo}")
                
            else:
                self.status_cliente_label.config(
                    text=f"❌ Arquivo não encontrado para {cliente_selecionado}",
                    foreground='red'
                )
                
                # Oferecer seleção manual
                resposta = messagebox.askyesno(
                    "Arquivo não encontrado",
                    f"Não foi encontrado arquivo para o cliente '{cliente_selecionado}'.\n\n"
                    f"Deseja selecionar manualmente o arquivo deste cliente?"
                )
                
                if resposta:
                    self.selecionar_arquivo_manual_despesas()
                else:
                    self.limpar_selecao_cliente()
                    
        except Exception as e:
            logger.error(f"Erro ao selecionar cliente: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao selecionar cliente: {str(e)}")

    def buscar_arquivo_cliente(self, nome_cliente):
        """Busca o arquivo Excel do cliente especificado"""
        try:
            # Importar configurações de pasta
            try:
                from src.config.config import PASTA_CLIENTES
            except ImportError:
                try:
                    from config.config import PASTA_CLIENTES
                except ImportError:
                    # Pasta padrão
                    PASTA_CLIENTES = "clientes"
            
            # Verificar se PASTA_CLIENTES existe
            if not os.path.exists(PASTA_CLIENTES):
                logger.warning(f"Pasta de clientes não encontrada: {PASTA_CLIENTES}")
                # Tentar pasta relativa
                pasta_alternativa = os.path.join(os.path.dirname(__file__), "..", "clientes")
                if os.path.exists(pasta_alternativa):
                    PASTA_CLIENTES = pasta_alternativa
                else:
                    return None
            
            # Possíveis nomes de arquivo
            possíveis_nomes = [
                f"{nome_cliente}.xlsx",
                f"{nome_cliente}.xls",
                f"{nome_cliente.upper()}.xlsx",
                f"{nome_cliente.lower()}.xlsx",
                f"{nome_cliente.replace(' ', '_')}.xlsx",
                f"{nome_cliente.replace(' ', '')}.xlsx"
            ]
            
            # Buscar arquivo
            for nome_arquivo in possíveis_nomes:
                caminho_completo = os.path.join(PASTA_CLIENTES, nome_arquivo)
                if os.path.exists(caminho_completo):
                    logger.info(f"Arquivo encontrado: {caminho_completo}")
                    return caminho_completo
            
            # Se não encontrou, listar arquivos na pasta para debug
            try:
                arquivos_existentes = os.listdir(PASTA_CLIENTES)
                logger.debug(f"Arquivos na pasta {PASTA_CLIENTES}: {arquivos_existentes}")
            except:
                pass
            
            logger.warning(f"Arquivo não encontrado para cliente: {nome_cliente}")
            return None
            
        except Exception as e:
            logger.error(f"Erro ao buscar arquivo do cliente: {str(e)}")
            return None

    def selecionar_arquivo_manual_despesas(self):
        """Permite seleção manual de arquivo (fallback)"""
        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo Excel do cliente",
                filetypes=[("Arquivos Excel", "*.xlsx *.xls")],
                initialdir=self.obter_pasta_clientes()
            )
            
            if arquivo:
                # Verificar se o arquivo é válido
                if not os.path.exists(arquivo):
                    messagebox.showerror("Erro", "Arquivo não encontrado.")
                    return
                    
                try:
                    # Tentar abrir o arquivo para verificar se é válido
                    from openpyxl import load_workbook
                    wb = load_workbook(arquivo, data_only=True)
                    
                    # Tentar obter nome do cliente do arquivo
                    try:
                        ws_resumo = wb['RESUMO']
                        nome_cliente_arquivo = ws_resumo['A3'].value
                        if nome_cliente_arquivo:
                            self.cliente_atual = nome_cliente_arquivo
                            # Atualizar combobox para mostrar o cliente correto
                            self.cliente_combobox.set(nome_cliente_arquivo)
                    except:
                        # Se não conseguir obter nome, usar nome do arquivo
                        self.cliente_atual = os.path.splitext(os.path.basename(arquivo))[0]
                    
                    wb.close()
                    
                    # Configurar arquivo selecionado
                    self.arquivo_cliente_selecionado = arquivo
                    
                    # Atualizar status
                    self.status_cliente_label.config(
                        text=f"✅ Arquivo selecionado manualmente: {os.path.basename(arquivo)}",
                        foreground='blue'
                    )
                    
                    if hasattr(self, 'status_individual_label'):
                        self.status_individual_label.config(
                            text=f"✅ Arquivo: {os.path.basename(arquivo)}",
                            foreground='blue'
                        )
                    
                    logger.info(f"Arquivo selecionado manualmente: {arquivo}")
                    
                except Exception as e:
                    messagebox.showerror(
                        "Erro", 
                        f"Arquivo inválido ou corrompido.\nErro: {str(e)}"
                    )
                    
        except Exception as e:
            logger.error(f"Erro na seleção manual: {str(e)}")
            messagebox.showerror("Erro", f"Erro na seleção manual: {str(e)}")

    def limpar_selecao_cliente(self):
        """Limpa a seleção de cliente atual"""
        self.arquivo_cliente_selecionado = None
        self.cliente_atual = None
        self.cliente_combobox.set('Todos os Clientes')
        
        self.status_cliente_label.config(
            text="Selecione um cliente para continuar",
            foreground='gray'
        )
        
        if hasattr(self, 'status_individual_label'):
            self.status_individual_label.config(
                text="Cliente será selecionado através da combobox acima",
                foreground='blue'
            )

    def atualizar_lista_clientes_despesas(self):
        """Atualiza a lista de clientes especificamente para despesas"""
        try:
            # Salvar seleção atual
            cliente_atual = self.cliente_combobox.get()
            
            # Recarregar lista
            self.atualizar_lista_clientes()
            
            # Tentar restaurar seleção
            if cliente_atual and cliente_atual in self.cliente_combobox['values']:
                self.cliente_combobox.set(cliente_atual)
            else:
                self.cliente_combobox.set('Todos os Clientes')
            
            messagebox.showinfo("Sucesso", "Lista de clientes atualizada!")
            logger.info("Lista de clientes atualizada na interface de despesas")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar lista: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao atualizar lista: {str(e)}")

    def obter_pasta_clientes(self):
        """Obtém o caminho da pasta de clientes"""
        try:
            from src.config.config import PASTA_CLIENTES
            return PASTA_CLIENTES
        except ImportError:
            try:
                from config.config import PASTA_CLIENTES
                return PASTA_CLIENTES
            except ImportError:
                return "clientes"

    def alternar_modo_data(self):
        """Alterna entre data automática e manual"""
        try:
            if self.usar_data_automatica.get():
                # Usar data automática - ocultar seleção manual
                self.frame_data_manual.pack_forget()
                
                # Recalcular data automática
                data_auto = self.calcular_data_rel_automatica()
                self.data_automatica_calculada = data_auto
                
                if hasattr(self, 'data_entry'):
                    self.data_entry.set_date(data_auto)
                    
                logger.info(f"Modo automático ativado: {data_auto.strftime('%d/%m/%Y')}")
                
            else:
                # Usar data manual - mostrar seleção
                self.frame_data_manual.pack(fill='x', padx=10, pady=5)
                logger.info("Modo manual ativado")
                
        except Exception as e:
            logger.error(f"Erro ao alternar modo de data: {str(e)}")

    def validar_data_manual(self):
        """Valida a data inserida manualmente"""
        try:
            if hasattr(self, 'data_entry'):
                data_selecionada = self.data_entry.get_date()
                valida, mensagem = self.validar_data_relatorio(data_selecionada)
                
                if valida:
                    messagebox.showinfo("Validação de Data", mensagem)
                else:
                    messagebox.showerror("Data Inválida", mensagem)
                    
        except Exception as e:
            logger.error(f"Erro ao validar data manual: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao validar data: {str(e)}")

    def obter_data_relatorio_final(self):
        """Obtém a data final a ser usada no relatório"""
        try:
            if self.usar_data_automatica.get():
                return self.data_automatica_calculada
            else:
                if hasattr(self, 'data_entry'):
                    return self.data_entry.get_date()
                else:
                    return self.data_automatica_calculada
                    
        except Exception as e:
            logger.error(f"Erro ao obter data final: {str(e)}")
            return self.calcular_data_rel_automatica()

    def alternar_tipo_geracao(self):
        """Alterna entre opções de geração individual e em lote"""
        try:
            if self.tipo_geracao.get() == "individual":
                self.frame_lote.pack_forget()
                self.frame_individual.pack(fill='x', padx=10, pady=10)
            else:
                self.frame_individual.pack_forget()
                self.frame_lote.pack(fill='x', padx=10, pady=10)
                
            # NOVO: Atualizar botão de geração
            self.atualizar_botao_geracao()
            
        except Exception as e:
            logger.error(f"Erro ao alternar tipo de geração: {str(e)}")

    def selecionar_arquivos_lote(self):
        """Abre diálogo para selecionar múltiplos arquivos para geração em lote"""
        try:
            arquivos = filedialog.askopenfilenames(
                title="Selecione os arquivos Excel",
                filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
            )
            if arquivos:
                # Validar todos os arquivos
                arquivos_validos = []
                arquivos_invalidos = []
                
                for arquivo in arquivos:
                    if os.path.exists(arquivo):
                        try:
                            # Verificar se é um arquivo Excel válido
                            from openpyxl import load_workbook
                            wb = load_workbook(arquivo, data_only=True)
                            wb.close()
                            arquivos_validos.append(arquivo)
                        except:
                            arquivos_invalidos.append(os.path.basename(arquivo))
                    else:
                        arquivos_invalidos.append(os.path.basename(arquivo))
                
                if arquivos_invalidos:
                    messagebox.showwarning(
                        "Arquivos Inválidos",
                        f"Os seguintes arquivos não puderam ser carregados:\n" +
                        "\n".join(arquivos_invalidos)
                    )
                
                if arquivos_validos:
                    self.arquivos_lote = arquivos_validos
                    self.lbl_arquivos_lote.config(
                        text=f"{len(arquivos_validos)} arquivos válidos selecionados"
                    )
                    logger.info(f"Selecionados {len(arquivos_validos)} arquivos para lote")
                else:
                    messagebox.showerror("Erro", "Nenhum arquivo válido foi selecionado.")
                    
        except Exception as e:
            logger.error(f"Erro ao selecionar arquivos em lote: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao selecionar arquivos: {str(e)}")
    
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
        try:
            arquivo = filedialog.askopenfilename(
                title="Selecione o arquivo do cliente",
                filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
            )
            if arquivo:
                # Verificar se o arquivo existe e é acessível
                if not os.path.exists(arquivo):
                    messagebox.showerror("Erro", "Arquivo não encontrado.")
                    return
                    
                try:
                    # Tentar abrir o arquivo para verificar se é válido
                    from openpyxl import load_workbook
                    wb = load_workbook(arquivo, data_only=True)
                    wb.close()
                    
                    # Extrair nome do cliente do arquivo
                    nome_arquivo = os.path.basename(arquivo)
                    self.cliente_combobox.set(f"Arquivo: {nome_arquivo}")
                    self.arquivo_cliente_selecionado = arquivo
                    
                    logger.info(f"Arquivo selecionado: {arquivo}")
                    
                except Exception as e:
                    messagebox.showerror(
                        "Erro", 
                        f"Arquivo inválido ou corrompido.\nErro: {str(e)}"
                    )
                    
        except Exception as e:
            logger.error(f"Erro ao selecionar arquivo: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao selecionar arquivo: {str(e)}")
    
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
        """Versão corrigida que mantém o código original para outros relatórios"""
        try:
            # Verificar se o relatório está disponível
            if not relatorio["disponivel"]:
                messagebox.showinfo("Em desenvolvimento", "Este relatório ainda está em desenvolvimento.")
                return
            
            # === APENAS DESPESAS TEM TRATAMENTO ESPECIAL ===
            if relatorio["id"] == "despesas":
                # NOVO: Fluxo otimizado apenas para despesas
                logger.info("Iniciando relatório de despesas - fluxo otimizado")
                
                # Validar configurações
                if not self.validar_configuracoes_despesas():
                    return
                
                # Coletar configurações
                configuracoes = self.coletar_configuracoes_completas()
                
                # Confirmar geração
                if not self.confirmar_geracao_relatorio():
                    return
                
                # Verificar modo selecionado
                usar_preview = hasattr(self, 'modo_visualizacao') and self.modo_visualizacao.get() == "preview"
                
                if usar_preview:
                    # Ir direto para preview
                    self.gerar_direto_com_preview(configuracoes)
                else:
                    # Geração direta
                    self.gerar_direto_sem_interface(configuracoes)
                
                return
            
            # === TODOS OS OUTROS RELATÓRIOS: CÓDIGO ORIGINAL INALTERADO ===
            
            # Para o relatório de fornecedores, usar uma abordagem mais direta
            if relatorio["id"] == "fornecedores":
                logger.info("Iniciando relatório de fornecedores")
                self.root.withdraw()
                from relatorios_interface import definir_menu_principal
                definir_menu_principal(self.root)   
                
                try:
                    # Importação direta
                    from relatorio_fornecedores import RelatorioFornecedores
                    app = RelatorioFornecedores(parent=self.root)
                    app.menu_principal = self.root
                    
                    # IMPORTANTE: Configurar o cliente selecionado ANTES de iniciar o mainloop
                    if hasattr(self, 'cliente_contratos') and self.cliente_contratos:
                        cliente_selecionado = self.cliente_contratos.get()
                        logger.info(f"Cliente selecionado na interface: {cliente_selecionado}")
                        
                        if cliente_selecionado and cliente_selecionado != 'Todos os Clientes':
                            # Atualizar a lista de clientes primeiro
                            app.atualizar_lista_clientes()
                            
                            # Aguardar um momento para garantir que a lista foi carregada
                            app.root.update()
                            
                            # Configurar o cliente no relatório de fornecedores
                            if cliente_selecionado in app.cliente_combobox['values']:
                                app.cliente_combobox.set(cliente_selecionado)
                                app.selecionar_cliente()
                                logger.info(f"Cliente configurado: {cliente_selecionado}")
                            else:
                                logger.info(f"Cliente {cliente_selecionado} não encontrado na lista")
                                
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
            
            # Código para outros tipos de relatório (ORIGINAL MANTIDO)
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
            
            # Iniciar interface conforme o tipo de relatório (ORIGINAL MANTIDO)
            if relatorio["id"] == "contratos":
                self.iniciar_relatorio_contratos(classe_relatorio)
            elif relatorio["id"] == "categoria":
                self.iniciar_relatorio_categoria(classe_relatorio)
            elif relatorio["id"] == "tipo_despesa":
                self.iniciar_relatorio_tipo_despesa(classe_relatorio)
            elif relatorio["id"] == "lancamentos_pendentes":
                self.iniciar_relatorio_lancamentos_pendentes(classe_relatorio)
            else:
                messagebox.showinfo(
                    "Em desenvolvimento",
                    "As opções específicas para este relatório ainda estão sendo implementadas."
                )
                    
        except Exception as e:
            logger.error(f"Erro ao gerar relatório: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
            self.root.deiconify()

    def processar_relatorio_despesas_otimizado(self):
        """Processamento otimizado específico para relatório de despesas"""
        try:
            logger.info("Iniciando relatório de despesas - fluxo otimizado")
            
            # Validar configurações
            if not self.validar_configuracoes_despesas():
                return
            
            # Coletar configurações
            configuracoes = self.coletar_configuracoes_completas()
            
            # Verificar modo selecionado
            usar_preview = hasattr(self, 'modo_visualizacao') and self.modo_visualizacao.get() == "preview"
            
            if usar_preview:
                # Ir direto para preview
                self.gerar_direto_com_preview(configuracoes)
            else:
                # Geração direta
                self.gerar_direto_sem_interface(configuracoes)
                
        except Exception as e:
            logger.error(f"Erro no processamento otimizado: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def gerar_direto_com_preview(self, configuracoes):
        """Gera dados e vai direto para o visualizador de preview"""
        try:
            logger.info("=== GERAÇÃO DIRETA COM PREVIEW ===")
            
            # Criar janela de progresso
            progress_window = self.criar_janela_progresso()
            
            def processar_em_thread():
                """Processa dados em thread separada"""
                try:
                    # Atualizar progresso
                    self.atualizar_progresso(progress_window, "Carregando dados do Excel...", 20)
                    
                    # Importar e criar handler
                    from src.relatorio_despesas_aprimorado import RelatorioHandler
                    handler = RelatorioHandler()
                    
                    # Carregar dados
                    df = handler.carregar_dados_excel(
                        configuracoes['arquivo'], 
                        configuracoes['incluir_excluidos']
                    )
                    
                    self.atualizar_progresso(progress_window, "Processando dados...", 40)
                    
                    # Processar dados
                    df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2 = handler.processar_dados(
                        df, configuracoes['data'], configuracoes['incluir_excluidos']
                    )
                    
                    self.atualizar_progresso(progress_window, "Processando lançamentos futuros...", 60)
                    
                    # Processar lançamentos futuros
                    df_futuro = None
                    if configuracoes['incluir_futuros']:
                        if hasattr(handler, 'processar_lancamentos_futuros'):
                            df_futuro = handler.processar_lancamentos_futuros(
                                df, configuracoes['data'], configuracoes['incluir_excluidos']
                            )
                    
                    self.atualizar_progresso(progress_window, "Obtendo dados do cliente...", 80)
                    
                    # Obter dados do cliente
                    from openpyxl import load_workbook
                    workbook = load_workbook(configuracoes['arquivo'], data_only=True)
                    ws_resumo = workbook['RESUMO']
                    nome_cliente = ws_resumo['A3'].value
                    
                    # Calcular valores
                    numero_relatorio = handler.obter_numero_relatorio(ws_resumo, configuracoes['data'])
                    valor_acumulado = handler.calcular_acumulado_dados(
                        df, configuracoes['data'], configuracoes['incluir_excluidos']
                    )
                    
                    # Montar dados completos
                    dados_completos = {
                        'df_filtrado': df_filtrado,
                        'df_diaria': df_diaria,
                        'df_tp_desp_1': df_tp_desp_1,
                        'df_tp_desp_2': df_tp_desp_2,
                        'df_futuro': df_futuro,
                        'df_original': df,
                        'incluir_futuros': configuracoes['incluir_futuros'],
                        'incluir_excluidos': configuracoes['incluir_excluidos'],
                        'data_relatorio': configuracoes['data'],
                        'nome_cliente': nome_cliente,
                        'endereco_cliente': ws_resumo['A4'].value,
                        'numero_relatorio': numero_relatorio,
                        'acumulado': valor_acumulado
                    }
                    
                    self.atualizar_progresso(progress_window, "Finalizando processamento...", 100)
                    
                    # Fechar janela de progresso
                    progress_window.destroy()
                    
                    # === IR DIRETO PARA O PREVIEW ===
                    self.abrir_visualizador_direto(dados_completos, configuracoes['arquivo'])
                    
                except Exception as e:
                    progress_window.destroy()
                    logger.error(f"Erro no processamento: {str(e)}", exc_info=True)
                    messagebox.showerror("Erro", f"Erro ao processar dados: {str(e)}")
            
            # Executar processamento
            import threading
            thread = threading.Thread(target=processar_em_thread)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            logger.error(f"Erro na geração direta com preview: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def criar_janela_progresso(self):
        """Cria janela de progresso elegante"""
        progress_window = tk.Toplevel(self.root)
        progress_window.title("Processando Relatório")
        progress_window.geometry("400x200")
        progress_window.transient(self.root)
        progress_window.grab_set()
        
        # Centralizar janela
        progress_window.update_idletasks()
        x = (progress_window.winfo_screenwidth() // 2) - (400 // 2)
        y = (progress_window.winfo_screenheight() // 2) - (200 // 2)
        progress_window.geometry(f"400x200+{x}+{y}")
        
        # Frame principal
        main_frame = ttk.Frame(progress_window, padding=20)
        main_frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            main_frame, 
            text="Gerando Relatório de Despesas", 
            font=('Arial', 12, 'bold')
        ).pack(pady=(0, 20))
        
        # Label de status
        progress_window.status_label = ttk.Label(
            main_frame, 
            text="Iniciando processamento...",
            font=('Arial', 10)
        )
        progress_window.status_label.pack(pady=(0, 10))
        
        # Barra de progresso
        progress_window.progress_bar = ttk.Progressbar(
            main_frame, 
            length=300, 
            mode='determinate'
        )
        progress_window.progress_bar.pack(pady=(0, 20))
        
        # Label de porcentagem
        progress_window.percent_label = ttk.Label(
            main_frame, 
            text="0%",
            font=('Arial', 9)
        )
        progress_window.percent_label.pack()
        
        return progress_window

    def atualizar_progresso(self, progress_window, mensagem, porcentagem):
        """Atualiza o progresso da janela"""
        try:
            if progress_window.winfo_exists():
                progress_window.status_label.config(text=mensagem)
                progress_window.progress_bar['value'] = porcentagem
                progress_window.percent_label.config(text=f"{porcentagem}%")
                progress_window.update()
                
                # Pequena pausa para visualização
                import time
                time.sleep(0.1)
                
        except Exception as e:
            logger.debug(f"Erro ao atualizar progresso: {str(e)}")

    def abrir_visualizador_direto(self, dados_completos, arquivo_path):
        """Abre o visualizador diretamente com os dados processados"""
        try:
            logger.info("Abrindo visualizador direto")
            
            # Importar classes necessárias
            from src.relatorio_despesas_aprimorado import VisualizadorRelatorio
            
            # Criar janela para o visualizador
            visualizador_window = tk.Toplevel(self.root)
            visualizador_window.title("Preview do Relatório")
            visualizador_window.geometry("900x700")
            
            # Configurar referências para navegação
            visualizador_window.menu_principal = self.root
            
            # Criar visualizador
            visualizador = VisualizadorRelatorio(visualizador_window)
            visualizador.arquivo_path = arquivo_path
            
            # === CONFIGURAR FECHAMENTO CORRETO ===
            def ao_fechar():
                """Comportamento ao fechar visualizador"""
                try:
                    visualizador_window.destroy()
                    # Voltar para interface de relatórios (manter aberta)
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                except Exception as e:
                    logger.error(f"Erro ao fechar visualizador: {str(e)}")
            
            visualizador_window.protocol("WM_DELETE_WINDOW", ao_fechar)
            
            # Ocultar interface atual temporariamente
            self.root.withdraw()
            
            # Mostrar preview direto
            preview_window = visualizador.mostrar_preview(dados_completos)
            
            # Configurar retorno correto
            def preview_fechado():
                """Quando preview é fechado, volta para interface de relatórios"""
                try:
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                except:
                    pass
            
            # Interceptar fechamento do preview
            original_destroy = preview_window.destroy
            def destroy_with_callback():
                preview_fechado()
                original_destroy()
            preview_window.destroy = destroy_with_callback
            
            logger.info("Visualizador direto aberto com sucesso")
            
        except Exception as e:
            logger.error(f"Erro ao abrir visualizador direto: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir visualizador: {str(e)}")
            # Em caso de erro, voltar à interface
            self.root.deiconify()

    def gerar_direto_sem_interface(self, configuracoes):
        """Gera relatório direto sem preview"""
        try:
            logger.info("=== GERAÇÃO DIRETA SEM PREVIEW ===")
            
            # Criar janela de progresso
            progress_window = self.criar_janela_progresso()
            
            def processar_e_gerar():
                """Processa e gera o PDF diretamente"""
                try:
                    # [Mesmo código de processamento da função anterior]
                    # ... processamento dos dados ...
                    
                    self.atualizar_progresso(progress_window, "Gerando arquivo PDF...", 90)
                    
                    # Gerar PDF direto
                    data_formatada = configuracoes['data'].strftime('%d-%m-%Y')
                    nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
                    
                    if configuracoes['incluir_excluidos']:
                        nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
                        
                    caminho_output = os.path.join(os.path.dirname(configuracoes['arquivo']), nome_arquivo)
                    
                    # Gerar PDF
                    handler.gerar_relatorio_pdf(dados_completos, caminho_output, configuracoes['arquivo'])
                    
                    self.atualizar_progresso(progress_window, "Relatório gerado com sucesso!", 100)
                    
                    # Fechar progresso
                    progress_window.destroy()
                    
                    # Mostrar resultado
                    self.mostrar_resultado_geracao(nome_cliente, nome_arquivo, caminho_output)
                    
                except Exception as e:
                    progress_window.destroy()
                    logger.error(f"Erro na geração direta: {str(e)}")
                    messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")
            
            # Executar em thread
            import threading
            thread = threading.Thread(target=processar_e_gerar)
            thread.daemon = True
            thread.start()
            
        except Exception as e:
            logger.error(f"Erro na geração direta: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def mostrar_resultado_geracao(self, nome_cliente, nome_arquivo, caminho_output):
        """Mostra resultado da geração com opções"""
        try:
            resposta = messagebox.askyesnocancel(
                "Relatório Gerado!",
                f"✅ Relatório gerado com sucesso!\n\n"
                f"Cliente: {nome_cliente}\n"
                f"Arquivo: {nome_arquivo}\n\n"
                f"🔄 Opções:\n"
                f"• Sim: Abrir PDF\n"
                f"• Não: Continuar sem abrir\n"
                f"• Cancelar: Gerar outro relatório",
                icon='question'
            )
            
            if resposta is True:  # Abrir PDF
                self.abrir_arquivo(caminho_output)
            elif resposta is False:  # Não abrir
                pass  # Continua na interface
            # resposta is None = Cancelar = continua na interface
            
        except Exception as e:
            logger.error(f"Erro ao mostrar resultado: {str(e)}")

    def abrir_arquivo(self, caminho):
        """Abre arquivo com programa padrão do sistema"""
        try:
            import platform
            import subprocess
            
            if platform.system() == 'Darwin':       # macOS
                subprocess.run(['open', caminho])
            elif platform.system() == 'Windows':    # Windows
                os.startfile(caminho)
            else:                                   # Linux
                subprocess.run(['xdg-open', caminho])
                
        except Exception as e:
            logger.error(f"Erro ao abrir arquivo: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir arquivo: {str(e)}")
    
    def iniciar_relatorio_despesas(self, classe_relatorio):
        """Versão final: Executa conforme configuração selecionada"""
        try:
            logger.info("Iniciando relatório de despesas - versão final")
            
            # Verificar modo selecionado na interface (sem perguntar novamente)
            usar_preview = False
            if hasattr(self, 'modo_visualizacao'):
                usar_preview = (self.modo_visualizacao.get() == "preview")
            
            if usar_preview:
                # Abrir interface completa COM dados já configurados
                self.abrir_interface_com_dados_transferidos(classe_relatorio)
            else:
                # Gerar direto com dados já configurados
                self.gerar_direto_simples_v2(classe_relatorio)
                
        except Exception as e:
            logger.error(f"Erro: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def abrir_interface_com_dados_transferidos(self, classe_relatorio):
        """Versão que sempre funciona - usa interface integrada"""
        try:
            # Coletar configurações
            configuracoes = self.coletar_configuracoes_completas()
            
            if not configuracoes['arquivo']:
                messagebox.showerror("Erro", "Selecione um arquivo primeiro.")
                return
            
            # Mostrar resumo
            resumo = self.gerar_resumo_configuracoes(configuracoes)
            
            resposta = messagebox.askyesno(
                "Abrir Interface Completa",
                f"Configurações que serão transferidas:\n\n{resumo}\n\n" +
                "Continuar?"
            )
            
            if not resposta:
                return
            
            # Usar interface integrada (mais confiável)
            self.usar_interface_integrada(classe_relatorio, configuracoes)
            
        except Exception as e:
            logger.error(f"Erro: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def usar_interface_integrada(self, classe_relatorio, configuracoes):
        """Usa interface integrada - VERSÃO COM REFERÊNCIAS CORRETAS"""
        try:
            self.root.withdraw()
            
            # Importar e criar interface diretamente
            from relatorio_despesas_aprimorado import RelatorioUI
            
            # Criar nova janela
            nova_root = tk.Tk()
            app = RelatorioUI(nova_root)
            
            # ===== CONFIGURAR REFERÊNCIAS AO MENU PRINCIPAL (CRÍTICO) =====
            app.menu_principal = self.root  # Referência na instância
            nova_root.menu_principal = self.root  # Referência na janela
            
            # Configurar também no handler se existir
            if hasattr(app, 'handler'):
                app.handler.menu_principal = self.root
            
            print(f"✅ Referências configuradas:")
            print(f"   app.menu_principal = {self.root}")
            print(f"   nova_root.menu_principal = {self.root}")
            
            # Aplicar configurações
            try:
                app.data_selecionada.set(configuracoes['data'].strftime('%d/%m/%Y'))
                app.incluir_futuros.set(configuracoes['incluir_futuros'])
                app.incluir_excluidos.set(configuracoes['incluir_excluidos'])
                
                if configuracoes['arquivo']:
                    app.arquivo_path = configuracoes['arquivo']
                    app.arquivo_selecionado.set(os.path.basename(configuracoes['arquivo']))
                
                if configuracoes['arquivos_lote']:
                    app.arquivos_lote = configuracoes['arquivos_lote']
                    
                logger.info("Configurações aplicadas na interface integrada")
                
            except Exception as e:
                logger.warning(f"Erro ao aplicar configurações: {str(e)}")
            
            # Configurar fechamento CORRETO
            def ao_fechar():
                print("🔄 Fechando interface e retornando ao menu...")
                nova_root.destroy()
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
                print("✅ Retornado ao menu principal")
            
            nova_root.protocol("WM_DELETE_WINDOW", ao_fechar)
            
            # Mostrar interface
            nova_root.lift()
            nova_root.focus_force()
            
            messagebox.showinfo(
                "Interface Carregada",
                "Interface carregada com suas configurações!\n\n" +
                "✅ Configurações aplicadas\n" +
                "✅ Arquivo selecionado\n" +
                "✅ Menu principal vinculado\n" +
                "✅ Pronto para usar"
            )
            
        except Exception as e:
            logger.error(f"Erro na interface integrada: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")
            self.root.deiconify()

    def coletar_configuracoes_completas(self):
        """Versão modificada que usa data calculada automaticamente"""
        config = {
            'data': datetime.now(),
            'incluir_futuros': True,
            'incluir_excluidos': False,
            'arquivo': None,
            'tipo_geracao': 'individual',
            'arquivos_lote': [],
            'formato_saida': 'pdf',
            'cliente_selecionado': None,
            'data_automatica': True  # Nova flag
        }
        
        try:
            # Data - usar o método que considera automático/manual
            config['data'] = self.obter_data_relatorio_final()
            config['data_automatica'] = self.usar_data_automatica.get() if hasattr(self, 'usar_data_automatica') else True
        except Exception as e:
            logger.debug(f"Erro ao coletar data: {str(e)}")
            config['data'] = self.calcular_data_rel_automatica()
        
        try:
            # Flags
            if hasattr(self, 'incluir_futuros'):
                config['incluir_futuros'] = self.incluir_futuros.get()
        except Exception as e:
            logger.debug(f"Erro ao coletar incluir_futuros: {str(e)}")
        
        try:
            if hasattr(self, 'incluir_excluidos'):
                config['incluir_excluidos'] = self.incluir_excluidos.get()
        except Exception as e:
            logger.debug(f"Erro ao coletar incluir_excluidos: {str(e)}")
        
        try:
            # Arquivo individual
            if hasattr(self, 'arquivo_cliente_selecionado'):
                config['arquivo'] = self.arquivo_cliente_selecionado
        except Exception as e:
            logger.debug(f"Erro ao coletar arquivo: {str(e)}")
        
        try:
            # Tipo de geração
            if hasattr(self, 'tipo_geracao'):
                config['tipo_geracao'] = self.tipo_geracao.get()
        except Exception as e:
            logger.debug(f"Erro ao coletar tipo_geracao: {str(e)}")
        
        try:
            # Arquivos em lote
            if hasattr(self, 'arquivos_lote') and self.arquivos_lote:
                config['arquivos_lote'] = self.arquivos_lote
        except Exception as e:
            logger.debug(f"Erro ao coletar arquivos_lote: {str(e)}")
        
        try:
            # Formato de saída
            if hasattr(self, 'formato_saida'):
                config['formato_saida'] = self.formato_saida.get()
        except Exception as e:
            logger.debug(f"Erro ao coletar formato_saida: {str(e)}")
        
        try:
            # Cliente selecionado
            if hasattr(self, 'cliente_combobox'):
                config['cliente_selecionado'] = self.cliente_combobox.get()
        except Exception as e:
            logger.debug(f"Erro ao coletar cliente: {str(e)}")
        
        return config

    def gerar_resumo_configuracoes(self, config):
        """Gera resumo legível das configurações"""
        try:
            resumo_lines = []
            
            # Data
            resumo_lines.append(f"📅 Data: {config['data'].strftime('%d/%m/%Y')}")
            
            # Arquivo/Cliente
            if config['arquivo']:
                nome_arquivo = os.path.basename(config['arquivo'])
                resumo_lines.append(f"📁 Arquivo: {nome_arquivo}")
            
            if config['cliente_selecionado'] and "Arquivo:" not in config['cliente_selecionado']:
                resumo_lines.append(f"👤 Cliente: {config['cliente_selecionado']}")
            
            # Tipo de geração
            tipo_texto = "Individual" if config['tipo_geracao'] == 'individual' else "Lote"
            resumo_lines.append(f"🔄 Tipo: {tipo_texto}")
            
            if config['arquivos_lote']:
                resumo_lines.append(f"📂 Arquivos em lote: {len(config['arquivos_lote'])} arquivos")
            
            # Opções
            opcoes = []
            if config['incluir_futuros']:
                opcoes.append("Lançamentos futuros")
            if config['incluir_excluidos']:
                opcoes.append("Lançamentos excluídos")
            
            if opcoes:
                resumo_lines.append(f"⚙️ Incluir: {', '.join(opcoes)}")
            
            # Formato
            resumo_lines.append(f"📄 Formato: {config['formato_saida'].upper()}")
            
            return "\n".join(resumo_lines)
            
        except Exception as e:
            logger.error(f"Erro ao gerar resumo: {str(e)}")
            return "Erro ao gerar resumo das configurações"

    def executar_interface_com_configuracoes(self, classe_relatorio, configuracoes):
        """Executa a interface externa com configurações pré-definidas"""
        try:
            import subprocess
            import sys
            import tempfile
            import json
            
            # Criar arquivo temporário com configurações
            config_data = {
                'data': configuracoes['data'].strftime('%d/%m/%Y'),
                'incluir_futuros': configuracoes['incluir_futuros'],
                'incluir_excluidos': configuracoes['incluir_excluidos'],
                'arquivo': configuracoes['arquivo'],
                'tipo_geracao': configuracoes['tipo_geracao'],
                'arquivos_lote': configuracoes['arquivos_lote'],
                'formato_saida': configuracoes['formato_saida']
            }
            
            # Criar arquivo temporário
            with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
                json.dump(config_data, f, ensure_ascii=False, indent=2)
                config_file_path = f.name
            
            # Tentar executar como processo separado
            script_path = os.path.join(os.path.dirname(__file__), 'relatorio_despesas_aprimorado.py')
            
            if os.path.exists(script_path):
                # Executar passando o arquivo de configuração como parâmetro
                processo = subprocess.Popen([
                    sys.executable, 
                    script_path, 
                    '--config', 
                    config_file_path
                ])
                
                messagebox.showinfo(
                    "Interface Aberta",
                    "A interface completa foi aberta com suas configurações!\n\n" +
                    "✅ Todas as configurações foram transferidas automaticamente.\n" +
                    "✅ O arquivo já está selecionado.\n" +
                    "✅ As opções estão pré-configuradas.\n\n" +
                    "Agora você pode:\n" +
                    "• Revisar as configurações\n" +
                    "• Gerar com preview ou direto\n" +
                    "• Ajustar se necessário"
                )
                
                # Agendar limpeza do arquivo temporário
                def limpar_temp_file():
                    try:
                        os.unlink(config_file_path)
                    except:
                        pass
                
                self.root.after(30000, limpar_temp_file)  # Limpar após 30 segundos
                
            else:
                # Fallback: abrir interface direta
                messagebox.showwarning(
                    "Aviso",
                    "Interface externa não encontrada.\n" +
                    "Usando interface simplificada..."
                )
                self.fallback_interface_com_configuracoes(classe_relatorio, configuracoes)
            
            # Restaurar janela principal
            self.root.deiconify()
            
        except Exception as e:
            logger.error(f"Erro ao executar interface: {str(e)}")
            # Tentar fallback
            try:
                self.fallback_interface_com_configuracoes(classe_relatorio, configuracoes)
            except:
                messagebox.showerror("Erro", f"Erro ao abrir interface: {str(e)}")
            finally:
                self.root.deiconify()

    def fallback_interface_com_configuracoes(self, classe_relatorio, configuracoes):
        """Fallback: Abre interface direta com configurações aplicadas"""
        try:
            from relatorio_despesas_aprimorado import RelatorioUI
            
            # Criar nova janela independente
            nova_root = tk.Tk()
            app = RelatorioUI(nova_root)
            
            # Aplicar TODAS as configurações coletadas
            try:
                # Data
                app.data_selecionada.set(configuracoes['data'].strftime('%d/%m/%Y'))
                
                # Arquivo
                if configuracoes['arquivo']:
                    app.arquivo_path = configuracoes['arquivo']
                    app.arquivo_selecionado.set(os.path.basename(configuracoes['arquivo']))
                
                # Flags
                app.incluir_futuros.set(configuracoes['incluir_futuros'])
                app.incluir_excluidos.set(configuracoes['incluir_excluidos'])
                
                # Arquivos em lote (se aplicável)
                if configuracoes['arquivos_lote']:
                    app.arquivos_lote = configuracoes['arquivos_lote']
                
                logger.info("Configurações aplicadas na interface fallback")
                
            except Exception as e:
                logger.warning(f"Erro ao aplicar algumas configurações: {str(e)}")
            
            # Configurar fechamento
            def ao_fechar():
                nova_root.destroy()
                self.root.deiconify()
            
            nova_root.protocol("WM_DELETE_WINDOW", ao_fechar)
            
            # Mostrar janela
            nova_root.lift()
            nova_root.focus_force()
            
            # Informar o usuário
            messagebox.showinfo(
                "Interface Carregada",
                "Interface carregada com suas configurações!\n\n" +
                "✅ Configurações transferidas\n" +
                "✅ Arquivo selecionado\n" +
                "✅ Opções aplicadas"
            )
            
        except Exception as e:
            logger.error(f"Erro no fallback com configurações: {str(e)}")
            # Último recurso: geração direta
            messagebox.showwarning(
                "Problema na Interface",
                "Não foi possível abrir a interface completa.\n" +
                "Será executada a geração direta."
            )
            self.gerar_direto_simples_v2(classe_relatorio)

    def abrir_interface_completa_v2(self, classe_relatorio):
        """Abre interface completa de forma mais segura"""
        try:
            # Confirmar fechamento da janela atual
            resposta = messagebox.askyesno(
                "Confirmar",
                "A interface atual será fechada.\n" +
                "A interface completa será aberta em nova janela.\n\n" +
                "Continuar?"
            )
            
            if not resposta:
                return
            
            # Coletar configurações antes de fechar
            configuracoes = self.coletar_configuracoes_seguro()
            
            # Fechar janela atual
            self.root.withdraw()
            
            # Criar nova aplicação independente
            def criar_nova_aplicacao():
                try:
                    import subprocess
                    import sys
                    
                    # Opção 1: Executar como processo separado
                    script_path = os.path.join(os.path.dirname(__file__), 'relatorio_despesas_aprimorado.py')
                    
                    if os.path.exists(script_path):
                        # Executar como processo independente
                        subprocess.Popen([sys.executable, script_path])
                        messagebox.showinfo(
                            "Interface Aberta",
                            "A interface completa foi aberta em janela separada.\n" +
                            "Configure as opções e gere o relatório."
                        )
                    else:
                        # Fallback: importar diretamente (mais arriscado)
                        self.fallback_interface_direta(classe_relatorio, configuracoes)
                    
                except Exception as e:
                    logger.error(f"Erro ao criar nova aplicação: {str(e)}")
                    messagebox.showerror("Erro", f"Erro ao abrir interface: {str(e)}")
                finally:
                    # Sempre mostrar janela principal novamente
                    self.root.deiconify()
            
            # Executar após delay
            self.root.after(100, criar_nova_aplicacao)
            
        except Exception as e:
            logger.error(f"Erro ao abrir interface completa: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")
            self.root.deiconify()

    def coletar_configuracoes_seguro(self):
        """Coleta configurações de forma segura"""
        config = {
            'data': datetime.now(),
            'incluir_futuros': True,
            'incluir_excluidos': False,
            'arquivo': None
        }
        
        try:
            if hasattr(self, 'data_entry'):
                config['data'] = self.data_entry.get_date()
        except:
            pass
        
        try:
            if hasattr(self, 'incluir_futuros'):
                config['incluir_futuros'] = self.incluir_futuros.get()
        except:
            pass
        
        try:
            if hasattr(self, 'incluir_excluidos'):
                config['incluir_excluidos'] = self.incluir_excluidos.get()
        except:
            pass
        
        try:
            if hasattr(self, 'arquivo_cliente_selecionado'):
                config['arquivo'] = self.arquivo_cliente_selecionado
        except:
            pass
        
        return config

    def fallback_interface_direta(self, classe_relatorio, configuracoes):
        """Fallback para abrir interface diretamente"""
        try:
            # CUIDADO: Esta é a versão arriscada - só usar se subprocess falhar
            from relatorio_despesas_aprimorado import RelatorioUI
            
            # Criar janela independente
            nova_root = tk.Tk()
            app = RelatorioUI(nova_root)
            
            # Aplicar configurações
            try:
                if configuracoes['data']:
                    app.data_selecionada.set(configuracoes['data'].strftime('%d/%m/%Y'))
                if configuracoes['arquivo']:
                    app.arquivo_path = configuracoes['arquivo']
                    app.arquivo_selecionado.set(os.path.basename(configuracoes['arquivo']))
                app.incluir_futuros.set(configuracoes['incluir_futuros'])
                app.incluir_excluidos.set(configuracoes['incluir_excluidos'])
            except Exception as e:
                logger.warning(f"Erro ao aplicar configurações: {str(e)}")
            
            # Configurar fechamento
            def ao_fechar():
                nova_root.destroy()
                self.root.deiconify()
            
            nova_root.protocol("WM_DELETE_WINDOW", ao_fechar)
            
            # IMPORTANTE: Não chamar mainloop aqui!
            # A janela vai funcionar no mesmo loop da aplicação principal
            nova_root.lift()
            nova_root.focus_force()
            
        except Exception as e:
            logger.error(f"Erro no fallback: {str(e)}")
            raise

    def gerar_direto_simples_v2(self, classe_relatorio):
        """Versão melhorada da geração direta"""
        try:
            # Validações
            if not hasattr(self, 'arquivo_cliente_selecionado') or not self.arquivo_cliente_selecionado:
                messagebox.showerror("Erro", "Selecione um arquivo primeiro.")
                return
            
            if not os.path.exists(self.arquivo_cliente_selecionado):
                messagebox.showerror("Erro", "Arquivo não encontrado.")
                return
            
            # Coletar configurações
            config = self.coletar_configuracoes_seguro()
            
            # Confirmar
            resumo = f"""Configurações do relatório:
            
    Arquivo: {os.path.basename(config['arquivo'])}
    Data: {config['data'].strftime('%d/%m/%Y')}
    Incluir futuros: {'Sim' if config['incluir_futuros'] else 'Não'}
    Incluir excluídos: {'Sim' if config['incluir_excluidos'] else 'Não'}

    Gerar relatório?"""
            
            if not messagebox.askyesno("Confirmar Geração", resumo):
                return
            
            # Processar
            self.processar_relatorio_direto(classe_relatorio, config)
            
        except Exception as e:
            logger.error(f"Erro na geração direta: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def processar_relatorio_direto(self, classe_relatorio, config):
        """Processa o relatório diretamente"""
        try:
            # Janela de progresso
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Gerando Relatório")
            progress_window.geometry("400x300")
            progress_window.transient(self.root)
            progress_window.grab_set()
            
            # Interface de progresso
            ttk.Label(progress_window, text="Processando relatório...").pack(pady=20)
            
            progress_bar = ttk.Progressbar(progress_window, mode='indeterminate')
            progress_bar.pack(pady=10)
            progress_bar.start()
            
            status_label = ttk.Label(progress_window, text="Carregando...")
            status_label.pack(pady=10)
            
            # Frame para resultado
            result_frame = ttk.Frame(progress_window)
            result_frame.pack(fill='x', padx=20, pady=10)
            
            def processar_async():
                try:
                    status_label.config(text="Carregando dados do Excel...")
                    progress_window.update()
                    
                    # Usar o handler
                    handler = classe_relatorio()
                    
                    status_label.config(text="Processando dados...")
                    progress_window.update()
                    
                    # Processar usando método completo
                    df = handler.carregar_dados_excel(config['arquivo'], config['incluir_excluidos'])
                    df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2 = handler.processar_dados(
                        df, config['data'], config['incluir_excluidos']
                    )
                    
                    df_futuro = None
                    if config['incluir_futuros']:
                        df_futuro = handler.processar_lancamentos_futuros(df, config['data'], config['incluir_excluidos'])
                    
                    status_label.config(text="Obtendo dados do cliente...")
                    progress_window.update()
                    
                    # Dados do cliente
                    from openpyxl import load_workbook
                    workbook = load_workbook(config['arquivo'], data_only=True)
                    ws_resumo = workbook['RESUMO']
                    nome_cliente = ws_resumo['A3'].value
                    
                    numero_relatorio = handler.obter_numero_relatorio(ws_resumo, config['data'])
                    valor_acumulado = handler.calcular_acumulado_dados(df, config['data'], config['incluir_excluidos'])
                    
                    dados_completos = {
                        'df_filtrado': df_filtrado,
                        'df_diaria': df_diaria,
                        'df_tp_desp_1': df_tp_desp_1,
                        'df_tp_desp_2': df_tp_desp_2,
                        'df_futuro': df_futuro,
                        'df_original': df,
                        'incluir_futuros': config['incluir_futuros'],
                        'incluir_excluidos': config['incluir_excluidos'],
                        'data_relatorio': config['data'],
                        'nome_cliente': nome_cliente,
                        'endereco_cliente': ws_resumo['A4'].value,
                        'numero_relatorio': numero_relatorio,
                        'acumulado': valor_acumulado
                    }
                    
                    status_label.config(text="Gerando arquivo PDF...")
                    progress_window.update()
                    
                    # Gerar PDF
                    data_formatada = config['data'].strftime('%d-%m-%Y')
                    nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
                    
                    if config['incluir_excluidos']:
                        nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
                    
                    caminho_output = os.path.join(os.path.dirname(config['arquivo']), nome_arquivo)
                    
                    handler.gerar_relatorio_pdf(dados_completos, caminho_output, config['arquivo'])
                    
                    # Finalizar
                    progress_bar.stop()
                    status_label.config(text="Relatório gerado com sucesso!")
                    
                    ttk.Label(result_frame, text=f"Cliente: {nome_cliente}").pack(anchor='w')
                    ttk.Label(result_frame, text=f"Arquivo: {nome_arquivo}").pack(anchor='w')
                    
                    btn_frame = ttk.Frame(result_frame)
                    btn_frame.pack(fill='x', pady=10)
                    
                    ttk.Button(
                        btn_frame, 
                        text="Abrir Relatório",
                        command=lambda: os.startfile(caminho_output)
                    ).pack(side='left', padx=5)
                    
                    ttk.Button(
                        btn_frame,
                        text="Fechar",
                        command=progress_window.destroy
                    ).pack(side='right', padx=5)
                    
                except Exception as e:
                    progress_bar.stop()
                    status_label.config(text="Erro no processamento!")
                    ttk.Label(result_frame, text=f"Erro: {str(e)}", foreground='red').pack()
                    ttk.Button(result_frame, text="Fechar", command=progress_window.destroy).pack(pady=10)
                    logger.error(f"Erro no processamento: {str(e)}", exc_info=True)
            
            # Executar processamento após delay
            progress_window.after(500, processar_async)
            
        except Exception as e:
            try:
                progress_window.destroy()
            except:
                pass
            raise

    def gerar_relatorio_direto_despesas(self, classe_relatorio):
        """Gera o relatório de despesas diretamente sem interface adicional"""
        try:
            # Coletar dados da interface
            data_selecionada = self.data_entry.get_date() if hasattr(self, 'data_entry') else datetime.now()
            incluir_futuros = self.incluir_futuros.get() if hasattr(self, 'incluir_futuros') else True
            incluir_excluidos = self.incluir_excluidos.get() if hasattr(self, 'incluir_excluidos') else False
            
            # Verificar se é geração individual ou em lote
            if hasattr(self, 'tipo_geracao') and self.tipo_geracao.get() == "lote":
                # Processar relatório em lote
                if not hasattr(self, 'arquivos_lote') or not self.arquivos_lote:
                    messagebox.showwarning("Aviso", "Nenhum arquivo selecionado para processamento em lote.")
                    return
                
                self.processar_relatorios_lote_direto(classe_relatorio, data_selecionada, incluir_futuros, incluir_excluidos)
            else:
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
                
                # Gerar relatório individual
                self.gerar_relatorio_individual_direto(classe_relatorio, arquivo, data_selecionada, incluir_futuros, incluir_excluidos)
            
        except Exception as e:
            logger.error(f"Erro ao gerar relatório direto: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")

    def gerar_relatorio_individual_direto(self, classe_relatorio, arquivo, data_selecionada, incluir_futuros, incluir_excluidos):
        """Gera um relatório individual diretamente"""
        try:
            from openpyxl import load_workbook
            
            # Instanciar o handler
            handler = classe_relatorio()
            
            # Carregar e processar dados
            df = handler.carregar_dados_excel(arquivo, incluir_excluidos)
            df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2 = handler.processar_dados(
                df, data_selecionada, incluir_excluidos
            )
            
            # Processar lançamentos futuros
            df_futuro = None
            if incluir_futuros:
                df_futuro = handler.processar_lancamentos_futuros(df, data_selecionada, incluir_excluidos)
            
            # Processar workbook
            workbook = load_workbook(arquivo, data_only=True)
            ws_resumo = workbook['RESUMO']
            nome_cliente = ws_resumo['A3'].value
            
            # Obter número do relatório e valor acumulado
            numero_relatorio = handler.obter_numero_relatorio(ws_resumo, data_selecionada)
            valor_acumulado = handler.calcular_acumulado_dados(df, data_selecionada, incluir_excluidos)
            
            dados_completos = {
                'df_filtrado': df_filtrado,
                'df_diaria': df_diaria,
                'df_tp_desp_1': df_tp_desp_1,
                'df_tp_desp_2': df_tp_desp_2,
                'df_futuro': df_futuro,
                'df_original': df,
                'incluir_futuros': incluir_futuros,
                'incluir_excluidos': incluir_excluidos,
                'data_relatorio': data_selecionada,
                'nome_cliente': nome_cliente,
                'endereco_cliente': ws_resumo['A4'].value,
                'numero_relatorio': numero_relatorio,
                'acumulado': valor_acumulado
            }
            
            # Gerar nome do arquivo
            data_formatada = data_selecionada.strftime('%d-%m-%Y')
            nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
            
            # Adicionar sufixo se incluir excluídos
            if incluir_excluidos:
                nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
                
            caminho_output = os.path.join(os.path.dirname(arquivo), nome_arquivo)
            
            # Gerar o PDF
            handler.gerar_relatorio_pdf(dados_completos, caminho_output, arquivo)
            
            # Mostrar mensagem de sucesso
            messagebox.showinfo(
                "Sucesso",
                f"Relatório gerado com sucesso!\n"
                f"Cliente: {nome_cliente}\n"
                f"Arquivo: {nome_arquivo}"
            )
            
            # Abrir o arquivo se desejado
            resposta = messagebox.askyesno(
                "Abrir Arquivo",
                "Deseja abrir o relatório gerado?"
            )
            
            if resposta:
                self.abrir_arquivo(caminho_output)
            
        except Exception as e:
            logger.error(f"Erro ao gerar relatório individual: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")

    def processar_relatorios_lote_direto(self, classe_relatorio, data_selecionada, incluir_futuros, incluir_excluidos):
        """Processa geração de relatórios em lote com melhor tratamento de parâmetros"""
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
            
            sucessos = 0
            erros = 0
            
            # Processar cada arquivo
            for i, arquivo in enumerate(self.arquivos_lote, 1):
                try:
                    nome_arquivo = os.path.basename(arquivo)
                    progress_label.config(text=f"Processando {i}/{total_arquivos}: {nome_arquivo}")
                    progress_bar['value'] = i - 0.5
                    progress_window.update()
                    
                    # Gerar relatório usando o mesmo método do individual
                    self.gerar_relatorio_individual_direto(
                        classe_relatorio, arquivo, data_selecionada, incluir_futuros, incluir_excluidos
                    )
                    
                    # Atualizar lista de resultados
                    result_list.insert(tk.END, f"✓ {nome_arquivo} - Concluído")
                    result_list.itemconfig(tk.END, fg="green")
                    result_list.see(tk.END)
                    sucessos += 1
                    
                    # Atualizar barra de progresso
                    progress_bar['value'] = i
                    progress_window.update()
                    
                except Exception as e:
                    # Registrar erro na lista
                    result_list.insert(tk.END, f"✗ {nome_arquivo} - Erro: {str(e)}")
                    result_list.itemconfig(tk.END, fg="red")
                    result_list.see(tk.END)
                    erros += 1
                    continue
            
            # Finalização
            progress_label.config(text=f"Processamento concluído! Sucessos: {sucessos}, Erros: {erros}")
            
            # Botão para fechar
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(pady=20)
            
            ttk.Button(
                btn_frame,
                text="Fechar",
                command=progress_window.destroy
            ).pack(side='right', padx=5)
            
            ttk.Button(
                btn_frame,
                text="Abrir Pasta",
                command=lambda: self.abrir_pasta_arquivos()
            ).pack(side='left', padx=5)
            
            # Tornar a janela modal
            progress_window.grab_set()
            progress_window.focus_set()
            progress_window.wait_window()
            
        except Exception as e:
            logger.error(f"Erro ao processar relatórios em lote: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao processar relatórios em lote: {str(e)}")
            if 'progress_window' in locals():
                progress_window.destroy()

    def abrir_arquivo(self, caminho):
        """Abre arquivo com o programa padrão do sistema"""
        try:
            import platform
            import subprocess
            
            if platform.system() == 'Darwin':       # macOS
                subprocess.run(['open', caminho])
            elif platform.system() == 'Windows':    # Windows
                os.startfile(caminho)
            else:                                   # Linux
                subprocess.run(['xdg-open', caminho])
        except Exception as e:
            logger.error(f"Erro ao abrir arquivo: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir arquivo: {str(e)}")

    def abrir_pasta_arquivos(self):
        """Abre a pasta onde estão os arquivos processados"""
        try:
            if hasattr(self, 'arquivos_lote') and self.arquivos_lote:
                pasta = os.path.dirname(self.arquivos_lote[0])
                self.abrir_arquivo(pasta)
        except Exception as e:
            logger.error(f"Erro ao abrir pasta: {str(e)}")

    def validar_configuracoes_despesas(self):
        """Versão atualizada da validação que considera a nova seleção"""
        try:
            # Verificar data
            if hasattr(self, 'data_entry'):
                try:
                    data = self.data_entry.get_date()
                    if not data:
                        messagebox.showerror("Erro", "Por favor, selecione uma data válida.")
                        return False
                except Exception:
                    messagebox.showerror("Erro", "Data selecionada é inválida.")
                    return False
            
            # Verificar tipo de geração
            if hasattr(self, 'tipo_geracao'):
                tipo = self.tipo_geracao.get()
                
                if tipo == "individual":
                    # Verificar se há cliente/arquivo selecionado
                    if not hasattr(self, 'arquivo_cliente_selecionado') or not self.arquivo_cliente_selecionado:
                        messagebox.showerror(
                            "Erro", 
                            "Por favor, selecione um cliente na combobox ou use a seleção manual de arquivo."
                        )
                        return False
                        
                    # Verificar se o arquivo existe
                    if not os.path.exists(self.arquivo_cliente_selecionado):
                        messagebox.showerror(
                            "Erro", 
                            "O arquivo do cliente selecionado não existe ou não pode ser acessado.\n"
                            "Tente atualizar a lista de clientes ou selecionar manualmente."
                        )
                        return False
                        
                elif tipo == "lote":
                    # Verificar se há arquivos selecionados para lote
                    if not hasattr(self, 'arquivos_lote') or not self.arquivos_lote:
                        messagebox.showerror(
                            "Erro", 
                            "Por favor, selecione arquivos para processamento em lote."
                        )
                        return False
                        
                    # Verificar se todos os arquivos existem
                    arquivos_inexistentes = []
                    for arquivo in self.arquivos_lote:
                        if not os.path.exists(arquivo):
                            arquivos_inexistentes.append(os.path.basename(arquivo))
                    
                    if arquivos_inexistentes:
                        messagebox.showerror(
                            "Erro", 
                            f"Os seguintes arquivos não existem:\n" + 
                            "\n".join(arquivos_inexistentes)
                        )
                        return False
            
            return True
            
        except Exception as e:
            logger.error(f"Erro ao validar configurações: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao validar configurações: {str(e)}")
            return False

    def mostrar_resumo_configuracoes(self):
        """Mostra um resumo das configurações antes de gerar o relatório"""
        try:
            resumo = []
            
            # Data
            if hasattr(self, 'data_entry'):
                try:
                    data = self.data_entry.get_date().strftime('%d/%m/%Y')
                    resumo.append(f"Data do relatório: {data}")
                except:
                    resumo.append("Data do relatório: Data atual")
            
            # Lançamentos futuros
            if hasattr(self, 'incluir_futuros'):
                status = "Sim" if self.incluir_futuros.get() else "Não"
                resumo.append(f"Incluir lançamentos futuros: {status}")
            
            # Lançamentos excluídos
            if hasattr(self, 'incluir_excluidos'):
                status = "Sim" if self.incluir_excluidos.get() else "Não"
                resumo.append(f"Incluir lançamentos excluídos: {status}")
            
            # Tipo de geração
            if hasattr(self, 'tipo_geracao'):
                tipo = self.tipo_geracao.get()
                resumo.append(f"Tipo de geração: {tipo.title()}")
                
                if tipo == "individual" and hasattr(self, 'arquivo_cliente_selecionado'):
                    nome_arquivo = os.path.basename(self.arquivo_cliente_selecionado)
                    resumo.append(f"Arquivo: {nome_arquivo}")
                elif tipo == "lote" and hasattr(self, 'arquivos_lote'):
                    resumo.append(f"Arquivos em lote: {len(self.arquivos_lote)} arquivos")
            
            # Modo de visualização
            if hasattr(self, 'modo_visualizacao'):
                modo = self.modo_visualizacao.get()
                modo_texto = "Com Preview" if modo == "preview" else "Direto"
                resumo.append(f"Modo de visualização: {modo_texto}")
            
            # Formato de saída
            if hasattr(self, 'formato_saida'):
                formato = self.formato_saida.get().upper()
                resumo.append(f"Formato de saída: {formato}")
            
            return "\n".join(resumo)
            
        except Exception as e:
            logger.error(f"Erro ao gerar resumo: {str(e)}")
            return "Erro ao gerar resumo das configurações"

    def confirmar_geracao_relatorio(self):
        """Confirma a geração do relatório mostrando um resumo"""
        try:
            resumo = self.mostrar_resumo_configuracoes()
            
            resposta = messagebox.askyesno(
                "Confirmar Geração",
                f"Confirma a geração do relatório com as seguintes configurações?\n\n{resumo}",
                icon='question'
            )
            
            return resposta
            
        except Exception as e:
            logger.error(f"Erro ao confirmar geração: {str(e)}")
            return True  # Em caso de erro, prosseguir

    def atualizar_botao_geracao(self):
        """Atualiza o texto e estado do botão de geração conforme as configurações"""
        try:
            # Este método pode ser chamado quando há mudanças nas configurações
            # para atualizar dinamicamente a interface
            
            if hasattr(self, 'tipo_geracao'):
                tipo = self.tipo_geracao.get()
                
                if tipo == "individual":
                    texto_botao = "Gerar Relatório Individual"
                else:
                    texto_botao = "Gerar Relatórios em Lote"
                    
                # Se houver um botão específico, atualizar seu texto
                # (Este código pode ser adaptado conforme a estrutura real da interface)
                
            logger.debug(f"Botão atualizado para: {texto_botao}")
            
        except Exception as e:
            logger.debug(f"Erro ao atualizar botão: {str(e)}")

    def limpar_selecoes(self):
        """Limpa as seleções de arquivos"""
        try:
            if hasattr(self, 'arquivo_cliente_selecionado'):
                delattr(self, 'arquivo_cliente_selecionado')
                
            if hasattr(self, 'arquivos_lote'):
                self.arquivos_lote = []
                
            if hasattr(self, 'lbl_arquivos_lote'):
                self.lbl_arquivos_lote.config(text="")
                
            if hasattr(self, 'cliente_combobox'):
                self.cliente_combobox.set("Todos os Clientes")
                
            logger.info("Seleções de arquivos limpas")
            
        except Exception as e:
            logger.error(f"Erro ao limpar seleções: {str(e)}")

    def resetar_configuracoes_despesas(self):
        """Reseta todas as configurações para valores padrão"""
        try:
            # Data atual
            if hasattr(self, 'data_entry'):
                from datetime import datetime
                self.data_entry.set_date(datetime.now())
            
            # Lançamentos futuros: True
            if hasattr(self, 'incluir_futuros'):
                self.incluir_futuros.set(True)
                
            # Lançamentos excluídos: False
            if hasattr(self, 'incluir_excluidos'):
                self.incluir_excluidos.set(False)
                
            # Tipo individual
            if hasattr(self, 'tipo_geracao'):
                self.tipo_geracao.set("individual")
                self.alternar_tipo_geracao()
                
            # Modo preview
            if hasattr(self, 'modo_visualizacao'):
                self.modo_visualizacao.set("preview")
                
            # Formato PDF
            if hasattr(self, 'formato_saida'):
                self.formato_saida.set("pdf")
                
            # Limpar seleções
            self.limpar_selecoes()
            
            logger.info("Configurações resetadas para padrão")
            
        except Exception as e:
            logger.error(f"Erro ao resetar configurações: {str(e)}")

    def adicionar_botoes_auxiliares(self, parent_frame):
        """Adiciona botões auxiliares para gerenciar configurações"""
        try:
            # Frame para botões auxiliares
            frame_botoes = ttk.LabelFrame(parent_frame, text="Ações Auxiliares")
            frame_botoes.pack(fill='x', padx=10, pady=10)
            
            # Botão para limpar seleções
            ttk.Button(
                frame_botoes,
                text="Limpar Seleções",
                command=self.limpar_selecoes
            ).pack(side='left', padx=5, pady=5)
            
            # Botão para resetar configurações
            ttk.Button(
                frame_botoes,
                text="Resetar Configurações",
                command=self.resetar_configuracoes_despesas
            ).pack(side='left', padx=5, pady=5)
            
            # Botão para mostrar resumo
            ttk.Button(
                frame_botoes,
                text="Ver Resumo",
                command=lambda: messagebox.showinfo(
                    "Resumo das Configurações", 
                    self.mostrar_resumo_configuracoes()
                )
            ).pack(side='left', padx=5, pady=5)
            
            logger.debug("Botões auxiliares adicionados")
            
        except Exception as e:
            logger.error(f"Erro ao adicionar botões auxiliares: {str(e)}")

    def mostrar_opcoes_relatorio_com_validacao(self, event=None):
        """Versão melhorada do mostrar_opcoes_relatorio com validação"""
        try:
            # Chamar o método original
            self.mostrar_opcoes_relatorio_original(event)
            
            # Adicionar validações específicas após mostrar as opções
            selecao = self.tree_relatorios.selection()
            if selecao:
                rel_id = selecao[0]
                
                if rel_id == "despesas":
                    # Adicionar botões auxiliares para relatório de despesas
                    if hasattr(self, 'right_frame'):
                        # Verificar se já foi adicionado
                        botoes_existem = any(
                            isinstance(widget, ttk.LabelFrame) and 
                            "Ações Auxiliares" in str(widget.cget('text', ''))
                            for widget in self.right_frame.winfo_children()
                        )
                        
                        if not botoes_existem:
                            self.adicionar_botoes_auxiliares(self.right_frame)
                            
        except Exception as e:
            logger.error(f"Erro na validação de opções: {str(e)}")

    def backup_metodo_original(self):
        """Cria backup do método original se necessário"""
        if not hasattr(self, 'mostrar_opcoes_relatorio_original'):
            self.mostrar_opcoes_relatorio_original = self.mostrar_opcoes_relatorio
            self.mostrar_opcoes_relatorio = self.mostrar_opcoes_relatorio_com_validacao

    def atualizar_interface_despesas(self):
        """Atualiza a interface baseada nas configurações atuais"""
        try:
            # Atualizar visibilidade dos frames
            if hasattr(self, 'tipo_geracao'):
                self.alternar_tipo_geracao()
                
            # Atualizar textos informativos
            if hasattr(self, 'lbl_arquivos_lote') and hasattr(self, 'arquivos_lote'):
                if self.arquivos_lote:
                    self.lbl_arquivos_lote.config(
                        text=f"{len(self.arquivos_lote)} arquivos selecionados"
                    )
                else:
                    self.lbl_arquivos_lote.config(text="")
                    
            logger.debug("Interface de despesas atualizada")
            
        except Exception as e:
            logger.error(f"Erro ao atualizar interface: {str(e)}")

    def gerar_relatorio_com_validacao(self, relatorio):
        """Wrapper para gerar_relatorio com validação prévia"""
        try:
            # Validações específicas para relatório de despesas
            if relatorio["id"] == "despesas":
                if not self.validar_configuracoes_despesas():
                    return
                    
                # Confirmar geração se configurado
                if not self.confirmar_geracao_relatorio():
                    return
            
            # Chamar método original
            self.gerar_relatorio(relatorio)
            
        except Exception as e:
            logger.error(f"Erro na validação para geração: {str(e)}")
            messagebox.showerror("Erro", f"Erro na validação: {str(e)}")

    def aplicar_melhorias_despesas(self):
        """Aplica todas as melhorias ao relatório de despesas"""
        try:
            # Substituir método de geração por versão com validação
            self.gerar_relatorio_original = self.gerar_relatorio
            self.gerar_relatorio = self.gerar_relatorio_com_validacao
            
            logger.info("Melhorias aplicadas ao relatório de despesas")
            
        except Exception as e:
            logger.error(f"Erro ao aplicar melhorias: {str(e)}")

    # MARCADO PARA EXCLUIR
    # def processar_relatorios_lote(self, classe_relatorio, data_selecionada, incluir_futuros):
    #     """Processa geração de relatórios em lote com barra de progresso"""
    #     try:
    #         # Criar janela de progresso
    #         progress_window = tk.Toplevel(self.root)
    #         progress_window.title("Gerando Relatórios em Lote")
    #         progress_window.geometry("700x550")
    #         progress_window.transient(self.root)
            
    #         # Frame principal
    #         main_frame = ttk.Frame(progress_window, padding=20)
    #         main_frame.pack(fill='both', expand=True)
            
    #         # Label para mostrar progresso
    #         progress_label = ttk.Label(main_frame, text="Iniciando processamento...", font=('Arial', 12))
    #         progress_label.pack(pady=10)
            
    #         # Barra de progresso
    #         progress_bar = ttk.Progressbar(main_frame, length=600, mode='determinate')
    #         progress_bar.pack(pady=20)
            
    #         # Lista de resultados
    #         result_frame = ttk.LabelFrame(main_frame, text="Relatórios Processados")
    #         result_frame.pack(fill='both', expand=True, pady=10)
            
    #         result_list = tk.Listbox(result_frame, font=('Courier', 10), height=15)
    #         scrollbar = ttk.Scrollbar(result_frame, orient='vertical', command=result_list.yview)
    #         result_list.configure(yscrollcommand=scrollbar.set)
    #         result_list.pack(side='left', fill='both', expand=True, padx=5, pady=5)
    #         scrollbar.pack(side='right', fill='y')
            
    #         # Configurar barra de progresso
    #         total_arquivos = len(self.arquivos_lote)
    #         progress_bar['maximum'] = total_arquivos
            
    #         # Instanciar o handler
    #         handler = classe_relatorio()
            
    #         # Processar cada arquivo
    #         for i, arquivo in enumerate(self.arquivos_lote, 1):
    #             try:
    #                 nome_arquivo = os.path.basename(arquivo)
    #                 progress_label.config(text=f"Processando {i}/{total_arquivos}: {nome_arquivo}")
    #                 progress_bar['value'] = i - 0.5
    #                 progress_window.update()
                    
    #                 # Gerar relatório
    #                 resultado = handler.gerar_relatorio_direto(
    #                     arquivo_path=arquivo,
    #                     data_relatorio=data_selecionada,
    #                     incluir_futuros=incluir_futuros
    #                 )
                    
    #                 # Atualizar lista de resultados
    #                 status = "✓" if resultado else "✗"
    #                 result_list.insert(tk.END, f"{status} {nome_arquivo}")
    #                 result_list.itemconfig(tk.END, fg="green" if resultado else "red")
    #                 result_list.see(tk.END)
                    
    #                 # Atualizar barra de progresso
    #                 progress_bar['value'] = i
    #                 progress_window.update()
                    
    #             except Exception as e:
    #                 # Registrar erro na lista
    #                 result_list.insert(tk.END, f"✗ {nome_arquivo} - Erro: {str(e)}")
    #                 result_list.itemconfig(tk.END, fg="red")
    #                 result_list.see(tk.END)
    #                 continue
            
    #         # Finalização
    #         progress_label.config(text="Processamento concluído!")
            
    #         # Botão para fechar
    #         ttk.Button(
    #             main_frame,
    #             text="Fechar",
    #             command=progress_window.destroy
    #         ).pack(pady=20)
            
    #         # Tornar a janela modal
    #         progress_window.grab_set()
    #         progress_window.focus_set()
    #         progress_window.wait_window()
            
    #     except Exception as e:
    #         messagebox.showerror("Erro", f"Erro ao processar relatórios em lote: {str(e)}")
    #         if 'progress_window' in locals():
    #             progress_window.destroy()
    
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
        """Volta ao menu principal de forma segura"""
        try:
            logger.info("Solicitado retorno ao menu principal")
            
            # Verificar se existe menu principal para retornar
            if self.menu_principal and hasattr(self.menu_principal, 'winfo_exists'):
                try:
                    # Verificar se a janela do menu principal ainda existe
                    if self.menu_principal.winfo_exists():
                        logger.info("Retornando ao menu principal existente")
                        
                        # Destruir janela atual
                        self.root.destroy()
                        
                        # Restaurar e focar no menu principal
                        self.menu_principal.deiconify()
                        self.menu_principal.lift()
                        self.menu_principal.focus_force()
                        
                        logger.info("Retorno ao menu principal concluído")
                        return
                    else:
                        logger.warning("Menu principal não existe mais")
                except Exception as e:
                    logger.error(f"Erro ao verificar menu principal: {str(e)}")
            
            # Se não há menu principal válido, fechar aplicação completamente
            logger.info("Não há menu principal válido, fechando aplicação")
            
            # Tentar fechar de forma segura
            try:
                self.root.quit()
                self.root.destroy()
            except:
                pass
            
            # Forçar saída se necessário
            import sys
            import os
            os._exit(0)
            
        except Exception as e:
            logger.error(f"Erro crítico no voltar_menu: {str(e)}")
            # Último recurso: forçar saída
            try:
                import os
                os._exit(0)
            except:
                pass
    
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