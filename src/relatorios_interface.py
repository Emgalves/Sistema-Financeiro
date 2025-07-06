import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib
import logging
import pandas as pd
import psutil  # Para verificação de memória
import gc
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

        # # === APLICAR MELHORIAS ===
        # try:
        #     # Aplicar melhorias principais
        #     self.aplicar_melhorias_sistema()
            
        #     # Aplicar correções específicas  
        #     self.aplicar_correcoes_especificas()
            
        #     logger.info("✅ Todas as melhorias aplicadas com sucesso")
            
        # except Exception as e:
        #     logger.error(f"Erro ao aplicar melhorias: {str(e)}")
    
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
            # ttk.Label(
            #     btn_frame,
            #     text="💡 O sistema processará os dados e abrirá diretamente o preview ou gerará o PDF conforme configurado.",
            #     font=('Arial', 9),
            #     foreground='blue',
            #     wraplength=400
            # ).pack(pady=(0, 10))
            
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
        """Versão corrigida que retorna data sem hora"""
        try:
            if self.usar_data_automatica.get():
                data = self.data_automatica_calculada
            else:
                if hasattr(self, 'data_entry'):
                    data = self.data_entry.get_date()
                else:
                    data = self.data_automatica_calculada
            
            # CORREÇÃO: Garantir que retorna apenas a data sem hora
            from datetime import datetime, date
            
            if isinstance(data, datetime):
                # Se é datetime, pegar apenas a parte da data
                data = data.date()
            
            # Converter para datetime no início do dia para processamento
            if isinstance(data, date):
                data = datetime.combine(data, datetime.min.time())
            
            logger.info(f"Data final obtida: {data} (tipo: {type(data)})")
            return data
            
        except Exception as e:
            logger.error(f"Erro ao obter data final: {str(e)}")
            from datetime import datetime
            return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

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
        """VERSÃO LIMPA E SIMPLIFICADA - Remove sobreposições"""
        try:
            logger.info(f"🔍 INICIANDO gerar_relatorio para: {relatorio['id']}")
            
            # Verificar disponibilidade
            if not relatorio["disponivel"]:
                messagebox.showinfo("Em desenvolvimento", "Este relatório ainda está em desenvolvimento.")
                return
            
            # === TRATAMENTO ESPECÍFICO PARA DESPESAS ===
            if relatorio["id"] == "despesas":
                logger.info("🎯 PROCESSANDO: Relatório de despesas")
                
                # 1. Validar configurações básicas
                if not self.validar_configuracoes_despesas():
                    logger.warning("❌ Validação de configurações falhou")
                    return
                
                # 2. Coletar todas as configurações
                configuracoes = self.coletar_configuracoes_completas()
                logger.info(f"✅ Configurações coletadas: arquivo={bool(configuracoes.get('arquivo'))}")
                
                # 3. Verificar se arquivo foi selecionado
                if not configuracoes.get('arquivo'):
                    messagebox.showerror("Erro", "Nenhum arquivo foi selecionado. Selecione um cliente ou use a seleção manual.")
                    return
                
                # 4. Confirmar geração
                if not self.confirmar_geracao_relatorio():
                    logger.info("❌ Geração cancelada pelo usuário")
                    return
                
                # 5. Verificar modo de visualização
                usar_preview = hasattr(self, 'modo_visualizacao') and self.modo_visualizacao.get() == "preview"
                
                # 6. Executar conforme modo selecionado
                if usar_preview:
                    logger.info("🚀 Executando com PREVIEW")
                    self.executar_relatorio_com_preview(configuracoes)
                else:
                    logger.info("🚀 Executando DIRETO")
                    self.executar_relatorio_direto(configuracoes)
                    
                return  # ⚠️ IMPORTANTE: Para aqui para despesas
            
            # === OUTROS RELATÓRIOS ===
            logger.info(f"📋 Processando outros relatórios: {relatorio['id']}")
            
            if relatorio["id"] == "lancamentos_pendentes":
                self.processar_lancamentos_pendentes()
            elif relatorio["id"] == "fornecedores":
                self.processar_fornecedores()
            else:
                self.processar_outros_relatorios(relatorio)
                
        except Exception as e:
            logger.error(f"💥 ERRO em gerar_relatorio: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao gerar relatório: {str(e)}")


    def executar_relatorio_com_preview(self, configuracoes):
        """VERSÃO SEM THREADING - Executa direto no thread principal"""
        try:
            logger.info("🎯 EXECUTANDO RELATÓRIO SEM THREADING")
            
            # CORREÇÃO: Fazer tudo direto sem threads
            
            # 1. Mostrar progresso simples
            progress_label = tk.Label(self.root, text="Processando dados...", 
                                    font=('Arial', 12), bg='lightblue', 
                                    relief='raised', padx=20, pady=10)
            progress_label.place(relx=0.5, rely=0.5, anchor='center')
            self.root.update()
            
            try:
                # 2. Processar dados DIRETO (sem thread)
                logger.info("🔧 Processando dados direto...")
                dados_processados = self.processar_dados_completo_otimizado(configuracoes)
                logger.info("✅ Dados processados com sucesso")
                
                # 3. Remover label de progresso
                progress_label.destroy()
                
                # 4. Abrir preview DIRETO
                logger.info("🔧 Abrindo preview direto...")
                self.abrir_preview_estavel(dados_processados, configuracoes['arquivo'])
                
            except Exception as e:
                # Limpar progresso em caso de erro
                try:
                    progress_label.destroy()
                except:
                    pass
                raise e
                
        except Exception as e:
            logger.error(f"💥 ERRO no executar_relatorio_com_preview: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def obter_handler_despesas_limpo(self):
        """Obtém handler de despesas de forma limpa"""
        try:
            # Tentar importação do módulo
            try:
                from src.relatorio_despesas_aprimorado import RelatorioHandler
            except ImportError:
                from relatorio_despesas_aprimorado import RelatorioHandler
            
            return RelatorioHandler()
            
        except Exception as e:
            logger.error(f"💥 ERRO ao obter handler: {str(e)}")
            raise Exception(f"Não foi possível importar RelatorioHandler: {str(e)}")


    def processar_dados_completo_otimizado(self, configuracoes):
        """VERSÃO CORRIGIDA - Processa dados uma única vez com validação"""
        try:
            arquivo_path = configuracoes['arquivo']
            data_relatorio = configuracoes['data']
            incluir_excluidos = configuracoes['incluir_excluidos']
            incluir_futuros = configuracoes['incluir_futuros']
            
            logger.info(f"📁 Processando arquivo: {os.path.basename(arquivo_path)}")
            
            # 1. Obter handler
            handler = self.obter_handler_despesas_limpo()
            
            # 2. CORREÇÃO: Carregar dados com validação robusta
            try:
                df_original = handler.carregar_dados_excel(arquivo_path, incluir_excluidos)
                logger.info(f"✅ Dados carregados: {len(df_original)} registros")
                
                # Verificar se df_original tem dados válidos
                if df_original.empty:
                    raise Exception("Arquivo não contém dados válidos")
                    
                # Verificar colunas essenciais
                colunas_essenciais = ['DATA_REL', 'TP_DESP', 'REFERÊNCIA', 'VALOR', 'NOME']
                colunas_faltantes = [col for col in colunas_essenciais if col not in df_original.columns]
                if colunas_faltantes:
                    raise Exception(f"Colunas essenciais ausentes: {colunas_faltantes}")
                    
            except Exception as e:
                logger.error(f"Erro ao carregar dados: {str(e)}")
                raise Exception(f"Erro ao carregar arquivo: {str(e)}")
            
            # 3. CORREÇÃO: Processar dados com validação
            try:
                df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2 = handler.processar_dados(
                    df_original, data_relatorio, incluir_excluidos
                )
                
                # Log detalhado dos dados processados
                logger.info(f"📊 Dados processados:")
                logger.info(f"   - df_filtrado: {len(df_filtrado)} registros")
                logger.info(f"   - df_diaria: {len(df_diaria)} registros") 
                logger.info(f"   - df_tp_desp_1: {len(df_tp_desp_1)} registros")
                logger.info(f"   - df_tp_desp_2: {len(df_tp_desp_2)} registros")
                
            except Exception as e:
                logger.error(f"Erro ao processar dados: {str(e)}")
                raise Exception(f"Erro no processamento: {str(e)}")
            
            # 4. CORREÇÃO: Processar lançamentos futuros com verificação
            df_futuro = None
            if incluir_futuros:
                try:
                    if hasattr(handler, 'processar_lancamentos_futuros'):
                        df_futuro = handler.processar_lancamentos_futuros(df_original, data_relatorio, incluir_excluidos)
                        logger.info(f"   - df_futuro: {len(df_futuro) if df_futuro is not None else 0} registros")
                    else:
                        logger.warning("Método processar_lancamentos_futuros não encontrado")
                except Exception as e:
                    logger.warning(f"Erro ao processar lançamentos futuros: {str(e)}")
                    df_futuro = pd.DataFrame()
            
            # 5. CORREÇÃO: Obter dados do cliente com validação
            try:
                from openpyxl import load_workbook
                workbook = load_workbook(arquivo_path, data_only=True)
                
                if 'RESUMO' not in workbook.sheetnames:
                    raise Exception("Planilha 'RESUMO' não encontrada no arquivo")
                    
                ws_resumo = workbook['RESUMO']
                
                # Verificar se as células essenciais existem
                nome_cliente = ws_resumo['A3'].value
                endereco_cliente = ws_resumo['A4'].value
                
                if not nome_cliente:
                    raise Exception("Nome do cliente não encontrado na célula A3")
                
                numero_relatorio = handler.obter_numero_relatorio(ws_resumo, data_relatorio)
                valor_acumulado = handler.calcular_acumulado_dados(df_original, data_relatorio, incluir_excluidos)
                
                workbook.close()
                
                logger.info(f"📋 Cliente: {nome_cliente}")
                logger.info(f"📋 Relatório nº: {numero_relatorio}")
                logger.info(f"📋 Acumulado: R$ {valor_acumulado:,.2f}")
                
            except Exception as e:
                logger.error(f"Erro ao obter dados do cliente: {str(e)}")
                raise Exception(f"Erro nos dados do cliente: {str(e)}")
            
            # 6. CORREÇÃO: Montar dados completos com validação
            dados_completos = {
                # DataFrames processados
                'df_filtrado': df_filtrado,
                'df_diaria': df_diaria,
                'df_tp_desp_1': df_tp_desp_1,
                'df_tp_desp_2': df_tp_desp_2,
                'df_futuro': df_futuro,
                'df_original': df_original,
                
                # Configurações
                'incluir_futuros': incluir_futuros,
                'incluir_excluidos': incluir_excluidos,
                'data_relatorio': data_relatorio,
                
                # Informações do cliente
                'nome_cliente': nome_cliente,
                'endereco_cliente': endereco_cliente,
                'numero_relatorio': numero_relatorio,
                'acumulado': valor_acumulado,
                
                # Metadados para debug
                'arquivo_path': arquivo_path,
                'timestamp_processamento': datetime.now()
            }
            
            logger.info(f"✅ Dados processados com sucesso para: {nome_cliente}")
            return dados_completos
            
        except Exception as e:
            logger.error(f"💥 ERRO no processamento: {str(e)}", exc_info=True)
            raise

    def gerar_pdf_temporario_preview(self, dados_completos, arquivo_path):
        """Gera PDF temporário para visualização antes de salvar definitivo"""
        try:
            import tempfile
            logger.info("🔍 GERANDO PDF TEMPORÁRIO PARA PREVIEW")
            
            # 1. Obter handler
            handler = self.obter_handler_despesas_limpo()
            
            # 2. Criar arquivo temporário
            temp_dir = tempfile.gettempdir()
            nome_temp = f"PREVIEW_REL_{dados_completos['nome_cliente']}_{datetime.now().strftime('%H%M%S')}.pdf"
            caminho_temp = os.path.join(temp_dir, nome_temp)
            
            # 3. Gerar PDF temporário
            logger.info(f"📄 Criando PDF temporário: {nome_temp}")
            handler.gerar_relatorio_pdf(dados_completos, caminho_temp, arquivo_path)
            
            # 4. Verificar se foi criado
            if os.path.exists(caminho_temp):
                tamanho = os.path.getsize(caminho_temp)
                logger.info(f"✅ PDF temporário criado: {tamanho} bytes")
                return caminho_temp, nome_temp
            else:
                raise Exception("PDF temporário não foi criado")
            
        except Exception as e:
            logger.error(f"💥 ERRO ao gerar PDF temporário: {str(e)}")
            raise Exception(f"Erro no PDF temporário: {str(e)}")


    def gerar_pdf_definitivo(self, dados_completos, arquivo_path):
        """Gera PDF definitivo na pasta do cliente"""
        try:
            logger.info("💾 GERANDO PDF DEFINITIVO")
            
            # 1. Obter handler
            handler = self.obter_handler_despesas_limpo()
            
            # 2. Preparar nome definitivo
            data_formatada = dados_completos['data_relatorio'].strftime('%d-%m-%Y')
            nome_cliente = dados_completos['nome_cliente']
            nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
            
            if dados_completos.get('incluir_excluidos', False):
                nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
            
            # 3. Pasta definitiva (mesma do arquivo original)
            pasta_definitiva = os.path.dirname(arquivo_path) if arquivo_path else os.path.expanduser("~/Desktop")
            caminho_definitivo = os.path.join(pasta_definitiva, nome_arquivo)
            
            # 4. Gerar PDF definitivo
            logger.info(f"📄 Salvando PDF definitivo: {nome_arquivo}")
            handler.gerar_relatorio_pdf(dados_completos, caminho_definitivo, arquivo_path)
            
            # 5. Verificar se foi criado
            if os.path.exists(caminho_definitivo):
                tamanho = os.path.getsize(caminho_definitivo)
                logger.info(f"✅ PDF definitivo salvo: {tamanho} bytes")
                return caminho_definitivo, nome_arquivo
            else:
                raise Exception("PDF definitivo não foi criado")
            
        except Exception as e:
            logger.error(f"💥 ERRO ao gerar PDF definitivo: {str(e)}")
            raise Exception(f"Erro no PDF definitivo: {str(e)}")


    def limpar_arquivo_temporario(self, caminho_temp):
        """Remove arquivo temporário após uso"""
        try:
            if os.path.exists(caminho_temp):
                os.remove(caminho_temp)
                logger.info(f"🗑️ Arquivo temporário removido: {os.path.basename(caminho_temp)}")
        except Exception as e:
            logger.warning(f"Aviso: Não foi possível remover arquivo temporário: {str(e)}")


    def abrir_preview_estavel(self, dados_completos, arquivo_path):
        """Versão que GARANTE que a janela seja visível"""
        try:
            logger.info("🎯 ABRINDO PREVIEW ESTÁVEL - GARANTINDO VISIBILIDADE")
            
            # Armazenar referências na instância
            if not hasattr(self, '_preview_refs'):
                self._preview_refs = {}
            
            # Importar tkinter
            import tkinter as tk
            from tkinter import ttk, messagebox
            
            # CORREÇÃO 1: NÃO ocultar interface principal ainda
            # self.root.withdraw()  # REMOVIDO TEMPORARIAMENTE
            
            # Criar janela de preview
            preview_window = tk.Toplevel(self.root)
            self._preview_refs['window'] = preview_window
            
            preview_window.title("Preview do Relatório de Despesas")
            preview_window.geometry("1000x800")  # Maior para garantir visibilidade
            
            # CORREÇÃO 2: Configurar janela para aparecer na frente
            preview_window.transient(self.root)
            preview_window.lift()
            preview_window.focus_force()
            
            # CORREÇÃO 3: Centralizar janela na tela
            preview_window.update_idletasks()
            width = preview_window.winfo_width()
            height = preview_window.winfo_height()
            x = (preview_window.winfo_screenwidth() // 2) - (width // 2)
            y = (preview_window.winfo_screenheight() // 2) - (height // 2)
            preview_window.geometry(f"1000x800+{x}+{y}")
            
            # CORREÇÃO 4: Forçar janela para frente
            preview_window.attributes('-topmost', True)  # Sempre no topo
            preview_window.update()
            preview_window.attributes('-topmost', False)  # Depois permite outras janelas
            
            logger.info("✅ Janela de preview criada e posicionada")
            
            # Criar widgets
            main_frame = ttk.Frame(preview_window, padding=15)
            main_frame.pack(fill='both', expand=True)
            
            # Título destacado
            title_label = ttk.Label(main_frame, 
                                text="🔍 PREVIEW DO RELATÓRIO DE DESPESAS", 
                                font=('Arial', 16, 'bold'),
                                foreground='blue')
            title_label.pack(pady=(0, 15))
            
            # Informações do cliente em destaque
            info_frame = ttk.LabelFrame(main_frame, text="Informações do Relatório", padding=10)
            info_frame.pack(fill='x', pady=(0, 15))
            
            # Grid para organizar informações
            info_frame.grid_columnconfigure(1, weight=1)
            
            ttk.Label(info_frame, text="Cliente:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=(0, 10))
            ttk.Label(info_frame, text=f"{dados_completos.get('nome_cliente', 'N/A')}", 
                    font=('Arial', 10)).grid(row=0, column=1, sticky='w')
            
            ttk.Label(info_frame, text="Data:", font=('Arial', 10, 'bold')).grid(row=1, column=0, sticky='w', padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text=f"{dados_completos.get('data_relatorio', 'N/A')}", 
                    font=('Arial', 10)).grid(row=1, column=1, sticky='w', pady=(5, 0))
            
            ttk.Label(info_frame, text="Relatório nº:", font=('Arial', 10, 'bold')).grid(row=2, column=0, sticky='w', padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text=f"{dados_completos.get('numero_relatorio', 'N/A')}", 
                    font=('Arial', 10)).grid(row=2, column=1, sticky='w', pady=(5, 0))
            
            ttk.Label(info_frame, text="Status:", font=('Arial', 10, 'bold')).grid(row=3, column=0, sticky='w', padx=(0, 10), pady=(5, 0))
            ttk.Label(info_frame, text="✅ Preview Funcionando e Visível!", 
                    font=('Arial', 10), foreground='green').grid(row=3, column=1, sticky='w', pady=(5, 0))
            
            # Área de texto com título
            text_frame = ttk.LabelFrame(main_frame, text="Conteúdo do Relatório", padding=10)
            text_frame.pack(fill='both', expand=True, pady=(0, 15))
            
            # Widget de texto
            text_widget = tk.Text(text_frame, wrap='word', font=('Courier', 9), 
                                bg='white', fg='black', relief='sunken', bd=2)
            scrollbar = ttk.Scrollbar(text_frame, orient='vertical', command=text_widget.yview)
            text_widget.configure(yscrollcommand=scrollbar.set)
            
            text_widget.pack(side='left', fill='both', expand=True)
            scrollbar.pack(side='right', fill='y')
            
            # Gerar e inserir preview textual
            try:
                logger.info("🔧 Gerando conteúdo do preview...")
                preview_text = self.gerar_preview_textual_simples(dados_completos)
                text_widget.insert('1.0', preview_text)
                text_widget.config(state='disabled')
                logger.info("✅ Conteúdo inserido com sucesso")
            except Exception as e:
                logger.error(f"Erro ao gerar preview textual: {str(e)}")
                text_widget.insert('1.0', f"Erro ao gerar preview: {str(e)}")
                text_widget.config(state='disabled')
            
            # Frame para botões com destaque
            button_frame = ttk.LabelFrame(main_frame, text="Ações Disponíveis", padding=10)
            button_frame.pack(fill='x')
            
            # CORREÇÃO 5: Função para ocultar interface principal DEPOIS
            def ocultar_interface_principal():
                try:
                    logger.info("🔧 Ocultando interface principal...")
                    self.root.withdraw()
                    logger.info("✅ Interface principal ocultada")
                except Exception as e:
                    logger.error(f"Erro ao ocultar interface: {str(e)}")
            
            # Função para voltar melhorada
            def voltar_relatorios():
                try:
                    logger.info("🔄 Voltando para interface de relatórios...")
                    
                    # Limpar referências
                    if hasattr(self, '_preview_refs'):
                        self._preview_refs.clear()
                    
                    # Restaurar interface principal ANTES de destruir preview
                    if self.root and hasattr(self.root, 'deiconify'):
                        self.root.deiconify()
                        self.root.lift()
                        self.root.focus_force()
                        logger.info("✅ Interface principal restaurada")
                    
                    # Destruir janela de preview
                    preview_window.destroy()
                    logger.info("✅ Preview fechado")
                    
                except Exception as e:
                    logger.error(f"Erro ao voltar: {str(e)}")
                    # Garantir que interface principal seja restaurada
                    try:
                        if hasattr(self, 'root') and self.root:
                            self.root.deiconify()
                            self.root.lift()
                            self.root.focus_force()
                    except:
                        pass
            
            # Função PDF simplificada para teste
            def gerar_pdf():
                """Versão com janela de decisão sempre visível"""
                try:
                    # Mostrar progresso inicial
                    progress_label = tk.Label(btn_frame, text="⏳ Gerando PDF temporário...", 
                                            font=('Arial', 10), foreground='blue')
                    progress_label.pack(pady=5)
                    preview_window.update()
                    
                    # 1. Gerar PDF temporário
                    caminho_temp, nome_temp = self.gerar_pdf_temporario_preview(dados_completos, arquivo_path)
                    
                    # Atualizar progresso
                    progress_label.config(text="🔍 Abrindo PDF para visualização...")
                    preview_window.update()
                    
                    # 2. Abrir PDF temporário
                    self.abrir_arquivo(caminho_temp)
                    
                    # Pequena pausa para o PDF abrir
                    import time
                    time.sleep(1.5)
                    
                    # Atualizar progresso
                    progress_label.config(text="📄 PDF aberto! Aguardando sua decisão...")
                    preview_window.update()
                    
                    # 3. Mostrar janela de decisão SEMPRE VISÍVEL
                    decisao = self.mostrar_decisao_pdf_visivel(nome_temp, preview_window)
                    
                    # 4. Processar decisão
                    if decisao == 'salvar':
                        progress_label.config(text="💾 Salvando PDF definitivo...")
                        preview_window.update()
                        
                        # Gerar PDF definitivo
                        caminho_definitivo, nome_definitivo = self.gerar_pdf_definitivo(dados_completos, arquivo_path)
                        
                        # Limpar temporário
                        self.limpar_arquivo_temporario(caminho_temp)
                        
                        progress_label.destroy()
                        
                        # Criar janela de sucesso também visível
                        sucesso_window = tk.Toplevel(preview_window)
                        sucesso_window.title("Sucesso!")
                        sucesso_window.geometry("400x200")
                        sucesso_window.attributes('-topmost', True)
                        sucesso_window.transient(preview_window)
                        
                        # Posicionar no centro da tela
                        sucesso_window.update_idletasks()
                        x = (sucesso_window.winfo_screenwidth() // 2) - 200
                        y = (sucesso_window.winfo_screenheight() // 2) - 100
                        sucesso_window.geometry(f"400x200+{x}+{y}")
                        
                        frame = ttk.Frame(sucesso_window, padding=20)
                        frame.pack(fill='both', expand=True)
                        
                        ttk.Label(frame, text="✅", font=('Arial', 24)).pack(pady=(0, 10))
                        ttk.Label(frame, text="PDF Salvo com Sucesso!", 
                                font=('Arial', 12, 'bold')).pack(pady=(0, 10))
                        ttk.Label(frame, text=f"Arquivo: {nome_definitivo}", 
                                wraplength=350).pack(pady=(0, 10))
                        ttk.Label(frame, text=f"Local: {os.path.dirname(caminho_definitivo)}", 
                                wraplength=350).pack(pady=(0, 15))
                        
                        ttk.Button(frame, text="OK", command=sucesso_window.destroy).pack()
                        
                        # Auto-fechar após 3 segundos
                        sucesso_window.after(3000, sucesso_window.destroy)
                        
                        logger.info(f"✅ PDF definitivo salvo: {caminho_definitivo}")
                        
                    elif decisao == 'temporario':
                        progress_label.destroy()
                        
                        # Criar janela informativa também visível
                        info_window = tk.Toplevel(preview_window)
                        info_window.title("PDF Temporário")
                        info_window.geometry("400x180")
                        info_window.attributes('-topmost', True)
                        info_window.transient(preview_window)
                        
                        # Posicionar no centro
                        info_window.update_idletasks()
                        x = (info_window.winfo_screenwidth() // 2) - 200
                        y = (info_window.winfo_screenheight() // 2) - 90
                        info_window.geometry(f"400x180+{x}+{y}")
                        
                        frame = ttk.Frame(info_window, padding=20)
                        frame.pack(fill='both', expand=True)
                        
                        ttk.Label(frame, text="📄", font=('Arial', 24)).pack(pady=(0, 10))
                        ttk.Label(frame, text="PDF Temporário Mantido", 
                                font=('Arial', 12, 'bold')).pack(pady=(0, 10))
                        ttk.Label(frame, text="O arquivo será removido ao fechar o sistema", 
                                wraplength=350).pack(pady=(0, 15))
                        
                        ttk.Button(frame, text="OK", command=info_window.destroy).pack()
                        
                        # Agendar remoção quando fechar o preview
                        def remover_temp_ao_fechar():
                            self.limpar_arquivo_temporario(caminho_temp)
                            voltar_action()
                        
                        preview_window.protocol("WM_DELETE_WINDOW", remover_temp_ao_fechar)
                        
                        logger.info(f"✅ PDF temporário mantido: {caminho_temp}")
                        
                    else:  # cancelar
                        progress_label.destroy()
                        # Limpar arquivo temporário
                        self.limpar_arquivo_temporario(caminho_temp)
                        logger.info("❌ Geração de PDF cancelada pelo usuário")
                    
                except Exception as e:
                    # Limpar progresso em caso de erro
                    try:
                        progress_label.destroy()
                    except:
                        pass
                    
                    logger.error(f"Erro ao gerar PDF: {str(e)}")
                    messagebox.showerror("Erro ao Gerar PDF", f"Erro: {str(e)}")

            
            # Criar botões
            btn_frame = ttk.Frame(button_frame)
            btn_frame.pack(fill='x')
            
            btn_ocultar = ttk.Button(btn_frame, text="🔻 Ocultar Interface Principal", 
                                    command=ocultar_interface_principal)
            btn_ocultar.pack(side='left', padx=(0, 10))
            
            btn_pdf = ttk.Button(btn_frame, text="🚀 Gerar PDF", command=gerar_pdf)
            btn_pdf.pack(side='left', padx=(0, 10))
            
            btn_voltar = ttk.Button(btn_frame, text="⬅️ Voltar", command=voltar_relatorios)
            btn_voltar.pack(side='right')
            
            # Configurar fechamento
            preview_window.protocol("WM_DELETE_WINDOW", voltar_relatorios)
            
            # CORREÇÃO 6: Garantir que janela apareça
            preview_window.deiconify()  # Garantir que está visível
            preview_window.lift()       # Trazer para frente
            preview_window.focus_force() # Forçar foco
            preview_window.update()     # Atualizar imediatamente
            
            logger.info("✅ Preview criado e DEVE ESTAR VISÍVEL na tela!")
            
            # Verificação de visibilidade
            def verificar_visibilidade():
                try:
                    if preview_window.winfo_exists():
                        if preview_window.winfo_viewable():
                            logger.info("✅ CONFIRMADO: Preview está visível na tela!")
                        else:
                            logger.warning("⚠️ Preview existe mas não está visível!")
                            # Tentar forçar visibilidade
                            preview_window.deiconify()
                            preview_window.lift()
                            preview_window.focus_force()
                    else:
                        logger.error("❌ Preview não existe!")
                except Exception as e:
                    logger.error(f"Erro na verificação de visibilidade: {str(e)}")
            
            # Verificar visibilidade após 1 segundo
            preview_window.after(1000, verificar_visibilidade)
            
            # Armazenar referências
            self._preview_refs.update({
                'window': preview_window,
                'main_frame': main_frame,
                'text_widget': text_widget,
                'scrollbar': scrollbar
            })
            
        except Exception as e:
            logger.error(f"💥 ERRO CRÍTICO no preview: {str(e)}", exc_info=True)
            try:
                messagebox.showerror("Erro Crítico", f"Erro ao criar preview: {str(e)}")
            except:
                pass

    def executar_relatorio_direto(self, configuracoes):
        """Executa relatório direto sem preview"""
        try:
            logger.info("🎯 EXECUTANDO RELATÓRIO DIRETO")
            
            # Processar dados
            handler = self.obter_handler_despesas_limpo()
            dados_processados = self.processar_dados_completo(handler, configuracoes)
            
            # Gerar nome do arquivo
            data_formatada = configuracoes['data'].strftime('%d-%m-%Y')
            nome_arquivo = f"REL - {dados_processados['nome_cliente']} - {data_formatada}.pdf"
            
            if configuracoes['incluir_excluidos']:
                nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
            
            caminho_output = os.path.join(os.path.dirname(configuracoes['arquivo']), nome_arquivo)
            
            # Gerar PDF
            handler.gerar_relatorio_pdf(dados_processados, caminho_output, configuracoes['arquivo'])
            
            # Mostrar resultado
            resposta = messagebox.askyesno(
                "Relatório Gerado!",
                f"Relatório gerado com sucesso!\n\n"
                f"Cliente: {dados_processados['nome_cliente']}\n"
                f"Arquivo: {nome_arquivo}\n\n"
                f"Deseja abrir o PDF?"
            )
            
            if resposta:
                self.abrir_arquivo(caminho_output)
            
            logger.info("✅ Relatório direto concluído")
            
        except Exception as e:
            logger.error(f"💥 ERRO no relatório direto: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def criar_janela_progresso_simples(self):
        """Cria janela de progresso simples"""
        try:
            window = tk.Toplevel(self.root)
            window.title("Processando...")
            window.geometry("400x120")
            window.transient(self.root)
            window.grab_set()
            window.resizable(False, False)
            
            # Centralizar
            window.update_idletasks()
            x = (window.winfo_screenwidth() // 2) - 200
            y = (window.winfo_screenheight() // 2) - 60
            window.geometry(f"400x120+{x}+{y}")
            
            # Widgets
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill='both', expand=True)
            
            ttk.Label(frame, text="Processando Relatório...", font=('Arial', 12)).pack(pady=10)
            
            window.status_label = ttk.Label(frame, text="Iniciando...")
            window.status_label.pack(pady=5)
            
            window.progress_bar = ttk.Progressbar(frame, length=300, mode='determinate')
            window.progress_bar.pack(pady=10)
            
            return window
            
        except Exception as e:
            logger.error(f"💥 ERRO ao criar progresso: {str(e)}")
            return None

    def formatar_numero(self, valor):
        """Formata número de forma segura"""
        try:
            import pandas as pd
            
            if valor is None or pd.isna(valor):
                return "0,00"
            
            # Converter para float se for string
            if isinstance(valor, str):
                valor = valor.replace('R$', '').replace(' ', '').replace(',', '.')
                valor = float(valor)
            
            # Formatar no padrão brasileiro
            return f"{float(valor):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        except:
            return "0,00"            

    def gerar_preview_textual_simples(self, dados):
        """Versão melhorada SIMPLES - apenas substitui o método existente"""
        try:
            preview_lines = []
            
            # Cabeçalho básico
            preview_lines.append("=" * 80)
            preview_lines.append("PREVIEW DO RELATÓRIO DE DESPESAS")
            preview_lines.append("=" * 80)
            preview_lines.append("")
            
            # Informações básicas
            preview_lines.append(f"CLIENTE: {dados.get('nome_cliente', 'N/A')}")
            preview_lines.append(f"ENDEREÇO: {dados.get('endereco_cliente', 'N/A')}")
            
            # Formatar data
            data_relatorio = dados.get('data_relatorio')
            if hasattr(data_relatorio, 'strftime'):
                data_formatada = data_relatorio.strftime('%d/%m/%Y')
            else:
                data_formatada = str(data_relatorio) if data_relatorio else 'N/A'
            
            preview_lines.append(f"RELATÓRIO Nº: {dados.get('numero_relatorio', 'N/A')}")
            preview_lines.append(f"DATA: {data_formatada}")
            preview_lines.append("")
            
            # NOVA SEÇÃO: Resumo dos dados processados
            preview_lines.append("-" * 60)
            preview_lines.append("RESUMO DOS DADOS PROCESSADOS")
            preview_lines.append("-" * 60)
            
            # Verificar DataFrames
            total_registros = 0
            dataframes_info = {
                'df_filtrado': 'Despesas principais (tipos 2-7)',
                'df_tp_desp_1': 'Colaboradores (salário/transporte/café)', 
                'df_tp_desp_2': 'Colaboradores (13º/férias/rescisão)',
                'df_diaria': 'Diárias',
                'df_futuro': 'Lançamentos futuros'
            }
            
            for df_name, descricao in dataframes_info.items():
                df = dados.get(df_name)
                
                if df is None:
                    status = "❌ Não processado"
                    count = 0
                elif not hasattr(df, 'empty'):
                    status = "❓ Formato inválido"
                    count = 0
                elif df.empty:
                    status = "⚪ Vazio"
                    count = 0
                else:
                    status = "✅ OK"
                    count = len(df)
                    total_registros += count
                
                preview_lines.append(f"{descricao}: {count} registros {status}")
            
            preview_lines.append("")
            preview_lines.append(f"📊 TOTAL GERAL: {total_registros} registros processados")
            preview_lines.append("")
            
            # NOVA SEÇÃO: Totais financeiros
            preview_lines.append("-" * 60)
            preview_lines.append("TOTAIS FINANCEIROS")
            preview_lines.append("-" * 60)
            
            # Calcular totais
            total_quinzena = 0
            
            for df_name in ['df_filtrado', 'df_tp_desp_1', 'df_tp_desp_2', 'df_diaria']:
                df = dados.get(df_name)
                if df is not None and hasattr(df, 'empty') and not df.empty and 'VALOR' in df.columns:
                    try:
                        import pandas as pd
                        valores = pd.to_numeric(df['VALOR'], errors='coerce').fillna(0)
                        total_quinzena += valores.sum()
                    except:
                        pass
            
            acumulado = dados.get('acumulado', 0)
            try:
                if isinstance(acumulado, str):
                    acumulado = float(acumulado.replace(',', '.'))
            except:
                acumulado = 0
                
            total_obra = total_quinzena + acumulado
            
            preview_lines.append(f"💵 TOTAL DA QUINZENA: R$ {self.formatar_numero(total_quinzena)}")
            preview_lines.append(f"📈 TOTAL ACUMULADO: R$ {self.formatar_numero(acumulado)}")
            preview_lines.append(f"🏗️ TOTAL DA OBRA: R$ {self.formatar_numero(total_obra)}")
            preview_lines.append("")
            
            # NOVA SEÇÃO: Amostra dos dados
            preview_lines.append("-" * 60)
            preview_lines.append("AMOSTRA DOS DADOS (primeiros registros)")
            preview_lines.append("-" * 60)
            
            # Mostrar amostra apenas dos principais
            for df_name, descricao in [('df_filtrado', 'Despesas principais'), ('df_tp_desp_1', 'Colaboradores')]:
                df = dados.get(df_name)
                
                if df is not None and hasattr(df, 'empty') and not df.empty:
                    preview_lines.append(f"\n🔸 {descricao.upper()}:")
                    preview_lines.append("Nome".ljust(25) + "Referência".ljust(30) + "Valor".rjust(12))
                    preview_lines.append("-" * 67)
                    
                    # Primeiros 3 registros
                    count = 0
                    for _, row in df.head(3).iterrows():
                        nome = str(row.get('NOME', ''))[:24]
                        referencia = str(row.get('REFERÊNCIA', ''))[:29]
                        
                        try:
                            import pandas as pd
                            valor = pd.to_numeric(row.get('VALOR', 0), errors='coerce')
                            if pd.isna(valor):
                                valor = 0
                            valor_fmt = f"R$ {self.formatar_numero(valor)}"
                        except:
                            valor_fmt = "R$ 0,00"
                        
                        preview_lines.append(f"{nome.ljust(25)} {referencia.ljust(30)} {valor_fmt.rjust(12)}")
                        count += 1
                    
                    if len(df) > 3:
                        preview_lines.append(f"... e mais {len(df) - 3} registros")
            
            # Configurações
            preview_lines.append("")
            preview_lines.append("-" * 60)
            preview_lines.append("CONFIGURAÇÕES DO RELATÓRIO")
            preview_lines.append("-" * 60)
            
            if dados.get('incluir_futuros'):
                preview_lines.append("✅ Lançamentos futuros incluídos")
            else:
                preview_lines.append("❌ Lançamentos futuros excluídos")
                
            if dados.get('incluir_excluidos'):
                preview_lines.append("✅ Lançamentos excluídos incluídos")
            else:
                preview_lines.append("❌ Lançamentos excluídos filtrados")
            
            # Rodapé
            preview_lines.append("")
            preview_lines.append("=" * 80)
            preview_lines.append("Use o botão 'Gerar PDF' para criar o relatório completo")
            preview_lines.append("=" * 80)
            
            return "\n".join(preview_lines)
            
        except Exception as e:
            logger.error(f"Erro ao gerar preview: {str(e)}")
            return f"""ERRO AO GERAR PREVIEW

    Detalhes do erro: {str(e)}

    Dados disponíveis: {list(dados.keys()) if isinstance(dados, dict) else 'Formato inválido'}

    Verifique os logs do sistema para mais detalhes."""

    def atualizar_progresso_simples(self, window, texto, valor):
        """Atualiza progresso de forma simples"""
        try:
            if window and hasattr(window, 'winfo_exists') and window.winfo_exists():
                window.status_label.config(text=texto)
                window.progress_bar['value'] = valor
                window.update()
        except Exception as e:
            logger.error(f"💥 ERRO ao atualizar progresso: {str(e)}")

    def limpar_threads_ativas(self):
        """Limpa threads que podem estar causando problemas"""
        try:
            import threading
            
            threads_ativas = threading.enumerate()
            thread_principal = threading.main_thread()
            
            logger.info(f"🧹 Verificando threads ativas: {len(threads_ativas)}")
            
            for thread in threads_ativas:
                if thread != thread_principal and thread.is_alive():
                    logger.warning(f"⚠️ Thread ativa detectada: {thread.name}")
                    # Não force kill threads, apenas log para debug
            
        except Exception as e:
            logger.error(f"Erro ao verificar threads: {str(e)}")
        
    def gerar_pdf_do_preview_simples(self, dados_completos, arquivo_path):
        """Gera PDF usando os dados já processados no preview - VERSÃO SIMPLES"""
        try:
            logger.info("🎯 GERANDO PDF A PARTIR DO PREVIEW")
            
            # 1. Obter handler existente
            handler = self.obter_handler_despesas_limpo()
            
            # 2. Preparar nome do arquivo
            data_formatada = dados_completos['data_relatorio'].strftime('%d-%m-%Y')
            nome_cliente = dados_completos['nome_cliente']
            nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
            
            # Adicionar sufixo se incluir excluídos
            if dados_completos.get('incluir_excluidos', False):
                nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
            
            # 3. Determinar pasta de saída
            if arquivo_path and os.path.exists(arquivo_path):
                pasta_saida = os.path.dirname(arquivo_path)
            else:
                pasta_saida = os.path.expanduser("~/Desktop")  # Fallback para Desktop
            
            caminho_output = os.path.join(pasta_saida, nome_arquivo)
            
            # 4. Gerar PDF usando o handler existente
            logger.info(f"📄 Gerando PDF: {nome_arquivo}")
            handler.gerar_relatorio_pdf(dados_completos, caminho_output, arquivo_path)
            
            # 5. Verificar se foi criado
            if os.path.exists(caminho_output):
                tamanho_arquivo = os.path.getsize(caminho_output)
                logger.info(f"✅ PDF criado com sucesso: {tamanho_arquivo} bytes")
                return caminho_output, nome_arquivo
            else:
                raise Exception("PDF não foi criado")
            
        except Exception as e:
            logger.error(f"💥 ERRO ao gerar PDF: {str(e)}")
            raise Exception(f"Erro na geração do PDF: {str(e)}")

    def obter_handler_despesas(self):
        """Obtém handler de despesas de forma limpa"""
        try:
            # Tentar importação hierárquica
            try:
                from src.relatorio_despesas_aprimorado import RelatorioHandler
                logger.info("✅ Handler importado de src.relatorio_despesas_aprimorado")
            except ImportError:
                from relatorio_despesas_aprimorado import RelatorioHandler
                logger.info("✅ Handler importado de relatorio_despesas_aprimorado")
            
            return RelatorioHandler()
            
        except Exception as e:
            logger.error(f"💥 ERRO ao obter handler: {str(e)}")
            raise Exception(f"Não foi possível importar RelatorioHandler: {str(e)}")

    def limpar_data(self, data_input):
        """Limpa e normaliza data de forma definitiva"""
        try:
            from datetime import datetime, date
            
            logger.info(f"🔧 Limpando data: {data_input} (tipo: {type(data_input)})")
            
            # Converter conforme tipo
            if isinstance(data_input, str):
                data_limpa = datetime.strptime(data_input, '%d/%m/%Y')
            elif isinstance(data_input, datetime):
                data_limpa = data_input
            elif isinstance(data_input, date):
                data_limpa = datetime.combine(data_input, datetime.min.time())
            else:
                logger.warning(f"⚠️ Tipo de data não reconhecido: {type(data_input)}")
                data_limpa = datetime.now()
            
            # Normalizar para início do dia
            data_final = data_limpa.replace(hour=0, minute=0, second=0, microsecond=0)
            
            logger.info(f"✅ Data limpa: {data_final}")
            return data_final
            
        except Exception as e:
            logger.error(f"💥 ERRO ao limpar data: {str(e)}")
            return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    def obter_info_cliente(self, arquivo, data_limpa, handler, df, incluir_excluidos):
        """Obtém informações do cliente de forma limpa"""
        try:
            from openpyxl import load_workbook
            
            workbook = load_workbook(arquivo, data_only=True)
            ws_resumo = workbook['RESUMO']
            
            nome_cliente = ws_resumo['A3'].value
            endereco_cliente = ws_resumo['A4'].value
            numero_relatorio = handler.obter_numero_relatorio(ws_resumo, data_limpa)
            valor_acumulado = handler.calcular_acumulado_dados(df, data_limpa, incluir_excluidos)
            
            workbook.close()
            
            logger.info(f"📋 Cliente: {nome_cliente}")
            logger.info(f"📋 Relatório nº: {numero_relatorio}")
            logger.info(f"📋 Acumulado: R$ {valor_acumulado:,.2f}")
            
            return {
                'nome_cliente': nome_cliente,
                'endereco_cliente': endereco_cliente,
                'numero_relatorio': numero_relatorio,
                'acumulado': valor_acumulado
            }
            
        except Exception as e:
            logger.error(f"💥 ERRO ao obter info cliente: {str(e)}")
            raise

    def montar_dados_completos(self, dados_processados, info_cliente, configuracoes, data_limpa, df_original):
        """Monta dados completos de forma limpa"""
        try:
            dados_completos = {
                # Dados processados
                **dados_processados,
                
                # DataFrame original
                'df_original': df_original,
                
                # Configurações
                'incluir_futuros': configuracoes['incluir_futuros'],
                'incluir_excluidos': configuracoes['incluir_excluidos'],
                'data_relatorio': data_limpa,
                
                # Informações do cliente
                **info_cliente
            }
            
            logger.info("✅ Dados completos montados")
            return dados_completos
            
        except Exception as e:
            logger.error(f"💥 ERRO ao montar dados: {str(e)}")
            raise

    # def abrir_preview_limpo(self, dados_completos, arquivo_path):
    #     """Abre preview de forma limpa - VERSÃO CORRIGIDA"""
    #     try:
    #         logger.info("🎯 ABRINDO PREVIEW LIMPO")
            
    #         # Importar visualizador
    #         try:
    #             from src.relatorio_despesas_aprimorado import VisualizadorRelatorio
    #             logger.info("✅ VisualizadorRelatorio importado de src")
    #         except ImportError:
    #             try:
    #                 from relatorio_despesas_aprimorado import VisualizadorRelatorio
    #                 logger.info("✅ VisualizadorRelatorio importado direto")
    #             except ImportError as e:
    #                 logger.error(f"❌ Erro ao importar VisualizadorRelatorio: {str(e)}")
    #                 raise Exception("Não foi possível importar VisualizadorRelatorio")
            
    #         # CORREÇÃO: Não ocultar a interface ainda - aguardar preview abrir
    #         logger.info("🔧 Preparando para criar visualizador")
            
    #         # CORREÇÃO: Criar visualizador com tratamento de erro
    #         try:
    #             visualizador = VisualizadorRelatorio(self.root)
    #             visualizador.arquivo_path = arquivo_path
    #             logger.info("✅ Visualizador criado com sucesso")
    #         except Exception as e:
    #             logger.error(f"❌ Erro ao criar visualizador: {str(e)}")
    #             messagebox.showerror("Erro", f"Erro ao criar visualizador: {str(e)}")
    #             return
            
    #         # CORREÇÃO: Verificar dados antes de abrir preview
    #         logger.info("🔍 Verificando dados antes do preview...")
    #         logger.info(f"   - df_filtrado: {len(dados_completos.get('df_filtrado', []))}")
    #         logger.info(f"   - df_diaria: {len(dados_completos.get('df_diaria', []))}")
    #         logger.info(f"   - df_tp_desp_1: {len(dados_completos.get('df_tp_desp_1', []))}")
    #         logger.info(f"   - Cliente: {dados_completos.get('nome_cliente', 'N/A')}")
    #         logger.info(f"   - Data: {dados_completos.get('data_relatorio', 'N/A')}")
            
    #         # CORREÇÃO: Abrir preview com tratamento de erro robusto
    #         try:
    #             logger.info("🚀 Chamando mostrar_preview...")
    #             preview_window = visualizador.mostrar_preview(dados_completos)
    #             logger.info("✅ mostrar_preview executado")
                
    #             # Verificar se preview_window foi criado
    #             if preview_window is None:
    #                 logger.error("❌ mostrar_preview retornou None")
    #                 messagebox.showerror("Erro", "Erro ao criar janela de preview")
    #                 return
                
    #             logger.info(f"✅ Preview window criado: {preview_window}")
                
    #         except Exception as e:
    #             logger.error(f"❌ ERRO no mostrar_preview: {str(e)}", exc_info=True)
    #             messagebox.showerror("Erro", f"Erro ao mostrar preview: {str(e)}")
    #             return
            
    #         # CORREÇÃO: Só ocultar interface APÓS preview estar aberto
    #         try:
    #             logger.info("🔧 Ocultando interface principal")
    #             self.root.withdraw()
    #             logger.info("✅ Interface principal ocultada")
    #         except Exception as e:
    #             logger.error(f"❌ Erro ao ocultar interface: {str(e)}")
            
    #         # CORREÇÃO: Configurar retorno com mais robustez
    #         def voltar_relatorios():
    #             """Volta para interface de relatórios de forma robusta"""
    #             try:
    #                 logger.info("🔄 Retornando para interface de relatórios")
                    
    #                 # Destruir preview se ainda existir
    #                 try:
    #                     if preview_window and hasattr(preview_window, 'winfo_exists'):
    #                         if preview_window.winfo_exists():
    #                             preview_window.destroy()
    #                             logger.info("✅ Preview window destruído")
    #                 except Exception as e:
    #                     logger.warning(f"⚠️ Erro ao destruir preview: {str(e)}")
                    
    #                 # Restaurar interface principal
    #                 try:
    #                     self.root.deiconify()
    #                     self.root.lift()
    #                     self.root.focus_force()
    #                     logger.info("✅ Interface principal restaurada")
    #                 except Exception as e:
    #                     logger.error(f"❌ Erro ao restaurar interface: {str(e)}")
                    
    #             except Exception as e:
    #                 logger.error(f"❌ Erro geral no voltar_relatorios: {str(e)}")
            
    #         # CORREÇÃO: Aplicar configuração de fechamento de forma mais robusta
    #         try:
    #             if hasattr(preview_window, 'protocol'):
    #                 preview_window.protocol("WM_DELETE_WINDOW", voltar_relatorios)
    #                 logger.info("✅ Protocolo de fechamento configurado")
    #             else:
    #                 logger.warning("⚠️ Preview window não tem método protocol")
                    
    #         except Exception as e:
    #             logger.error(f"❌ Erro ao configurar protocolo: {str(e)}")
            
    #         # CORREÇÃO: Focar na janela de preview
    #         try:
    #             if hasattr(preview_window, 'lift'):
    #                 preview_window.lift()
    #                 preview_window.focus_force()
    #                 logger.info("✅ Preview focado")
    #         except Exception as e:
    #             logger.warning(f"⚠️ Erro ao focar preview: {str(e)}")
            
    #         logger.info("✅ Preview limpo aberto com sucesso")
            
    #     except Exception as e:
    #         logger.error(f"💥 ERRO GERAL no abrir_preview_limpo: {str(e)}", exc_info=True)
    #         messagebox.showerror("Erro", f"Erro crítico no preview: {str(e)}")
            
    #         # Em caso de erro, restaurar interface
    #         try:
    #             self.root.deiconify()
    #             self.root.lift()
    #             self.root.focus_force()
    #         except:
    #             pass
        
    def criar_progress_window(self):
        """Cria janela de progresso simples"""
        try:
            window = tk.Toplevel(self.root)
            window.title("Processando...")
            window.geometry("400x150")
            window.transient(self.root)
            window.grab_set()
            
            # Centralizar
            window.update_idletasks()
            x = (window.winfo_screenwidth() // 2) - 200
            y = (window.winfo_screenheight() // 2) - 75
            window.geometry(f"400x150+{x}+{y}")
            
            # Widgets
            frame = ttk.Frame(window, padding=20)
            frame.pack(fill='both', expand=True)
            
            ttk.Label(frame, text="Processando Relatório...", font=('Arial', 12, 'bold')).pack(pady=10)
            
            window.status_label = ttk.Label(frame, text="Iniciando...")
            window.status_label.pack(pady=5)
            
            window.progress_bar = ttk.Progressbar(frame, length=300, mode='determinate')
            window.progress_bar.pack(pady=10)
            
            return window
            
        except Exception as e:
            logger.error(f"💥 ERRO ao criar progress: {str(e)}")
            return None

    def update_progress(self, window, texto, valor):
        """Atualiza progresso de forma segura"""
        try:
            if window and hasattr(window, 'winfo_exists') and window.winfo_exists():
                window.status_label.config(text=texto)
                window.progress_bar['value'] = valor
                window.update()
                logger.info(f"📈 {valor}% - {texto}")
        except Exception as e:
            logger.error(f"💥 ERRO ao atualizar progress: {str(e)}")

    def processar_lancamentos_pendentes(self):
        """Processa lançamentos pendentes - mantém original"""
        try:
            if not hasattr(self, 'pasta_lancamentos'):
                messagebox.showerror("Erro", "Selecione uma pasta primeiro.")
                return
            
            data_ref = self.data_referencia_pendentes.get_date() if hasattr(self, 'data_referencia_pendentes') else datetime.now()
            
            from src.relatorio_despesas_aprimorado import RelatorioLancamentosPendentes
            relatorio = RelatorioLancamentosPendentes()
            arquivo_saida = os.path.join(self.pasta_lancamentos, "relatorio_lancamentos_pendentes.html")
            
            if relatorio.gerar_relatorio_pendentes(self.pasta_lancamentos, arquivo_saida, data_ref):
                messagebox.showinfo("Sucesso", f"Relatório gerado: {arquivo_saida}")
            else:
                messagebox.showwarning("Aviso", "Nenhum lançamento pendente encontrado.")
                
        except Exception as e:
            logger.error(f"💥 ERRO lançamentos pendentes: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def processar_fornecedores(self):
        """Processa fornecedores - mantém original simplificado"""
        try:
            self.root.withdraw()
            
            from src.relatorio_fornecedores import RelatorioFornecedores
            app = RelatorioFornecedores(parent=self.root)
            app.menu_principal = self.root
            
            app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()
            
        except Exception as e:
            logger.error(f"💥 ERRO fornecedores: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")
            self.root.deiconify()

    def processar_outros_relatorios(self, relatorio):
        """Processa outros relatórios - mantém original"""
        try:
            modulo = self.carregar_modulo(relatorio["modulo"])
            if not modulo:
                return
            
            classe_relatorio = getattr(modulo, relatorio["classe"])
            
            if relatorio["id"] == "contratos":
                self.iniciar_relatorio_contratos(classe_relatorio)
            elif relatorio["id"] == "categoria":
                self.iniciar_relatorio_categoria(classe_relatorio)
            elif relatorio["id"] == "tipo_despesa":
                self.iniciar_relatorio_tipo_despesa(classe_relatorio)
            else:
                messagebox.showinfo("Em desenvolvimento", "Em desenvolvimento.")
                
        except Exception as e:
            logger.error(f"💥 ERRO outros relatórios: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")

    def mostrar_decisao_pdf_visivel(self, nome_temp, preview_window):
        """Cria janela de decisão sempre visível"""
        try:
            # Criar janela de decisão personalizada
            decisao_window = tk.Toplevel()
            decisao_window.title("Decisão sobre PDF")
            decisao_window.geometry("450x400")
            
            # CONFIGURAÇÕES PARA FICAR SEMPRE VISÍVEL
            decisao_window.attributes('-topmost', True)  # Sempre no topo
            decisao_window.lift()
            decisao_window.focus_force()
            decisao_window.grab_set()  # Modal
            
            # Posicionar no canto superior direito (longe do PDF)
            screen_width = decisao_window.winfo_screenwidth()
            x = screen_width - 470  # 450 + margem
            y = 50  # Topo da tela
            decisao_window.geometry(f"450x400+{x}+{y}")
            
            # Frame principal
            main_frame = ttk.Frame(decisao_window, padding=25)
            main_frame.pack(fill='both', expand=True)
            
            # Ícone e título
            ttk.Label(main_frame, text="📄", font=('Arial', 32)).pack(pady=(0, 10))
            ttk.Label(main_frame, text="PDF Temporário Gerado!", 
                    font=('Arial', 14, 'bold')).pack(pady=(0, 15))
            
            # Informações
            info_frame = ttk.LabelFrame(main_frame, text="📋 Informações", padding=10)
            info_frame.pack(fill='x', pady=(0, 20))
            
            ttk.Label(info_frame, text="✅ PDF temporário criado e aberto", 
                    font=('Arial', 10)).pack(anchor='w')
            ttk.Label(info_frame, text=f"📄 Arquivo: {nome_temp}", 
                    font=('Arial', 9), wraplength=350).pack(anchor='w', pady=(5, 0))
            ttk.Label(info_frame, text="🔍 Analise o PDF e escolha uma opção:", 
                    font=('Arial', 10, 'bold')).pack(anchor='w', pady=(10, 0))
            
            # Variável para capturar a decisão
            decisao = {'resultado': None}
            
            # Frame para botões
            button_frame = ttk.Frame(main_frame)
            button_frame.pack(fill='x')
            
            def salvar_definitivo():
                decisao['resultado'] = 'salvar'
                decisao_window.destroy()
            
            def manter_temporario():
                decisao['resultado'] = 'temporario'
                decisao_window.destroy()
            
            def cancelar():
                decisao['resultado'] = 'cancelar'
                decisao_window.destroy()
            
            # Botões organizados verticalmente para melhor visibilidade
            ttk.Button(button_frame, text="💾 Salvar PDF na Pasta do Cliente", 
                    command=salvar_definitivo).pack(fill='x', pady=(0, 10))
            
            ttk.Button(button_frame, text="📄 Manter Apenas Temporário", 
                    command=manter_temporario).pack(fill='x', pady=(0, 10))
            
            ttk.Button(button_frame, text="❌ Cancelar", 
                    command=cancelar).pack(fill='x')
            
            # Configurar fechamento (equivale a cancelar)
            decisao_window.protocol("WM_DELETE_WINDOW", cancelar)
            
            # Aguardar decisão do usuário
            decisao_window.wait_window()
            
            return decisao['resultado']
            
        except Exception as e:
            logger.error(f"Erro na janela de decisão: {str(e)}")
            return 'cancelar'

    def normalizar_data_relatorio(self, data_input):
        """Normaliza data para comparação correta"""
        try:
            from datetime import datetime, date
            
            # Se é string, converter
            if isinstance(data_input, str):
                data_input = datetime.strptime(data_input, '%d/%m/%Y')
            
            # Se é datetime, pegar apenas a data
            if isinstance(data_input, datetime):
                data_input = data_input.date()
            
            # Converter para datetime no início do dia
            if isinstance(data_input, date):
                data_normalizada = datetime.combine(data_input, datetime.min.time())
            else:
                data_normalizada = data_input
            
            logger.info(f"Data normalizada de {data_input} para {data_normalizada}")
            return data_normalizada
            
        except Exception as e:
            logger.error(f"Erro ao normalizar data: {str(e)}")
            return datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    def criar_janela_progresso_melhorada(self):
        """Cria janela de progresso mais robusta"""
        try:
            progress_window = tk.Toplevel(self.root)
            progress_window.title("Processando Relatório de Despesas")
            progress_window.geometry("450x250")
            progress_window.transient(self.root)
            progress_window.grab_set()
            progress_window.resizable(False, False)
            
            # Centralizar janela
            progress_window.update_idletasks()
            x = (progress_window.winfo_screenwidth() // 2) - (450 // 2)
            y = (progress_window.winfo_screenheight() // 2) - (250 // 2)
            progress_window.geometry(f"450x250+{x}+{y}")
            
            # Frame principal
            main_frame = ttk.Frame(progress_window, padding=30)
            main_frame.pack(fill='both', expand=True)
            
            # Ícone e título
            title_frame = ttk.Frame(main_frame)
            title_frame.pack(fill='x', pady=(0, 20))
            
            ttk.Label(
                title_frame, 
                text="🚀", 
                font=('Arial', 24)
            ).pack()
            
            ttk.Label(
                title_frame, 
                text="Gerando Relatório de Despesas", 
                font=('Arial', 14, 'bold')
            ).pack(pady=(10, 0))
            
            # Status
            progress_window.status_label = ttk.Label(
                main_frame, 
                text="Iniciando processamento...",
                font=('Arial', 10),
                justify='center'
            )
            progress_window.status_label.pack(pady=(0, 15))
            
            # Barra de progresso
            progress_window.progress_bar = ttk.Progressbar(
                main_frame, 
                length=350, 
                mode='determinate',
                style='TProgressbar'
            )
            progress_window.progress_bar.pack(pady=(0, 10))
            
            # Porcentagem
            progress_window.percent_label = ttk.Label(
                main_frame, 
                text="0%",
                font=('Arial', 9, 'bold')
            )
            progress_window.percent_label.pack()
            
            # Configurar fechamento (impedir fechamento manual)
            progress_window.protocol("WM_DELETE_WINDOW", lambda: None)
            
            return progress_window
            
        except Exception as e:
            logger.error(f"Erro ao criar janela de progresso: {str(e)}")
            return None

    def atualizar_progresso_seguro(self, progress_window, mensagem, porcentagem):
        """Atualiza progresso de forma segura"""
        try:
            if progress_window and hasattr(progress_window, 'winfo_exists'):
                if progress_window.winfo_exists():
                    progress_window.status_label.config(text=mensagem)
                    progress_window.progress_bar['value'] = porcentagem
                    progress_window.percent_label.config(text=f"{porcentagem}%")
                    progress_window.update()
                    
                    # Log do progresso
                    logger.info(f"Progresso: {porcentagem}% - {mensagem}")
                    
        except Exception as e:
            logger.error(f"Erro ao atualizar progresso: {str(e)}")

    # def abrir_preview_final(self, dados_completos, arquivo_path):
    #     """Abre o preview final de forma limpa"""
    #     try:
    #         logger.info("=== ABRINDO PREVIEW FINAL ===")
            
    #         # Importar visualizador
    #         try:
    #             from src.relatorio_despesas_aprimorado import VisualizadorRelatorio
    #         except ImportError:
    #             from relatorio_despesas_aprimorado import VisualizadorRelatorio
            
    #         # Criar visualizador DIRETO no root (sem janela intermediária)
    #         visualizador = VisualizadorRelatorio(self.root)
    #         visualizador.arquivo_path = arquivo_path
            
    #         # Ocultar interface atual
    #         self.root.withdraw()
            
    #         # Mostrar preview
    #         preview_window = visualizador.mostrar_preview(dados_completos)
            
    #         # Configurar retorno correto
    #         def voltar_interface():
    #             """Volta para interface de relatórios"""
    #             try:
    #                 self.root.deiconify()
    #                 self.root.lift()
    #                 self.root.focus_force()
    #                 logger.info("Retornado para interface de relatórios")
    #             except Exception as e:
    #                 logger.error(f"Erro ao voltar: {str(e)}")
            
    #         # Configurar fechamento
    #         def fechar_preview():
    #             try:
    #                 preview_window.destroy()
    #             except:
    #                 pass
    #             voltar_interface()
            
    #         # Aplicar configuração de fechamento
    #         preview_window.protocol("WM_DELETE_WINDOW", fechar_preview)
            
    #         # Interceptar método destroy
    #         original_destroy = preview_window.destroy
    #         preview_window.destroy = fechar_preview
            
    #         logger.info("Preview final aberto com sucesso")
            
    #     except Exception as e:
    #         logger.error(f"Erro ao abrir preview final: {str(e)}")
    #         messagebox.showerror("Erro", f"Erro ao abrir preview: {str(e)}")
    #         self.root.deiconify()

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

    
    def abrir_visualizador_unico(self, dados_completos, arquivo_path):
        """Abre apenas um visualizador - corrige problema das duas janelas"""
        try:
            logger.info("Abrindo visualizador único")
            
            # CORREÇÃO: Importar de forma mais robusta
            try:
                from src.relatorio_despesas_aprimorado import VisualizadorRelatorio
            except ImportError:
                from relatorio_despesas_aprimorado import VisualizadorRelatorio
            
            # CORREÇÃO: Criar visualizador diretamente sem janela adicional
            visualizador = VisualizadorRelatorio(self.root)
            visualizador.arquivo_path = arquivo_path
            
            # === CONFIGURAR FECHAMENTO CORRETO ===
            def ao_fechar_preview():
                """Comportamento ao fechar preview"""
                try:
                    # Voltar para interface de relatórios
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
                    logger.info("Voltou para interface de relatórios após fechar preview")
                except Exception as e:
                    logger.error(f"Erro ao voltar para interface: {str(e)}")
            
            # Ocultar interface atual temporariamente
            self.root.withdraw()
            
            # CORREÇÃO: Mostrar preview DIRETO sem criar janela intermediária
            preview_window = visualizador.mostrar_preview(dados_completos)
            
            # CORREÇÃO: Interceptar todos os métodos de fechamento
            def fechar_e_voltar():
                try:
                    preview_window.destroy()
                    ao_fechar_preview()
                except Exception as e:
                    logger.error(f"Erro ao fechar preview: {str(e)}")
                    ao_fechar_preview()
            
            # Configurar fechamento para todos os casos
            preview_window.protocol("WM_DELETE_WINDOW", fechar_e_voltar)
            
            # Interceptar método destroy original
            original_destroy = preview_window.destroy
            preview_window.destroy = fechar_e_voltar
            
            logger.info("Visualizador único aberto com sucesso")
            
        except Exception as e:
            logger.error(f"Erro ao abrir visualizador único: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir visualizador: {str(e)}")
            # Em caso de erro, voltar à interface
            self.root.deiconify()

    
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
        """Versão corrigida que GARANTE que arquivo esteja nas configurações"""
        config = {
            'data': datetime.now(),
            'incluir_futuros': True,
            'incluir_excluidos': False,
            'arquivo': None,  # IMPORTANTE: Inicializar como None
            'tipo_geracao': 'individual',
            'arquivos_lote': [],
            'formato_saida': 'pdf',
            'cliente_selecionado': None,
            'data_automatica': True
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
            # CORREÇÃO: Arquivo individual - verificar múltiplas fontes
            if hasattr(self, 'arquivo_cliente_selecionado') and self.arquivo_cliente_selecionado:
                config['arquivo'] = self.arquivo_cliente_selecionado
                logger.info(f"✅ Arquivo encontrado em arquivo_cliente_selecionado: {config['arquivo']}")
            elif hasattr(self, 'arquivo_path') and self.arquivo_path:
                config['arquivo'] = self.arquivo_path
                logger.info(f"✅ Arquivo encontrado em arquivo_path: {config['arquivo']}")
            else:
                logger.warning("❌ Arquivo não encontrado em nenhuma variável")
                
        except Exception as e:
            logger.error(f"Erro ao coletar arquivo: {str(e)}")
        
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
        
        # CORREÇÃO: Verificação final crítica
        if not config['arquivo']:
            logger.error("❌ ERRO CRÍTICO: Nenhum arquivo foi encontrado nas configurações!")
            logger.error("Variáveis verificadas:")
            logger.error(f"  - hasattr(self, 'arquivo_cliente_selecionado'): {hasattr(self, 'arquivo_cliente_selecionado')}")
            if hasattr(self, 'arquivo_cliente_selecionado'):
                logger.error(f"  - self.arquivo_cliente_selecionado: {self.arquivo_cliente_selecionado}")
            logger.error(f"  - hasattr(self, 'arquivo_path'): {hasattr(self, 'arquivo_path')}")
            if hasattr(self, 'arquivo_path'):
                logger.error(f"  - self.arquivo_path: {self.arquivo_path}")
        else:
            logger.info(f"✅ Arquivo final nas configurações: {config['arquivo']}")
        
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

    def processar_despesas_otimizado(self):
        """Método específico para processar despesas - VERSÃO COM DEBUG"""
        try:
            logger.info("=== PROCESSAMENTO OTIMIZADO DE DESPESAS ===")
            
            # DEBUG: Verificar variáveis de arquivo ANTES de qualquer coisa
            logger.info("🔍 VERIFICANDO VARIÁVEIS DE ARQUIVO:")
            logger.info(f"  - hasattr(arquivo_cliente_selecionado): {hasattr(self, 'arquivo_cliente_selecionado')}")
            if hasattr(self, 'arquivo_cliente_selecionado'):
                logger.info(f"  - arquivo_cliente_selecionado: {self.arquivo_cliente_selecionado}")
            
            logger.info(f"  - hasattr(arquivo_path): {hasattr(self, 'arquivo_path')}")
            if hasattr(self, 'arquivo_path'):
                logger.info(f"  - arquivo_path: {self.arquivo_path}")
            
            # 1. Validar configurações
            if not self.validar_configuracoes_despesas():
                logger.warning("Validação de configurações falhou")
                return
            
            # 2. Coletar configurações
            configuracoes = self.coletar_configuracoes_completas()
            logger.info(f"Configurações coletadas: {list(configuracoes.keys())}")
            
            # CORREÇÃO: Verificação crítica do arquivo nas configurações
            if not configuracoes.get('arquivo'):
                logger.error("❌ ERRO: Arquivo não encontrado nas configurações!")
                messagebox.showerror(
                    "Erro", 
                    "Arquivo não encontrado. Verifique se um cliente foi selecionado ou se o arquivo foi escolhido manualmente."
                )
                return
            
            logger.info(f"✅ Arquivo confirmado: {configuracoes['arquivo']}")
            
            # 3. Confirmar geração
            if not self.confirmar_geracao_relatorio():
                logger.info("Geração cancelada pelo usuário")
                return
            
            # 4. Verificar modo selecionado
            usar_preview = hasattr(self, 'modo_visualizacao') and self.modo_visualizacao.get() == "preview"
            logger.info(f"Modo selecionado: {'Preview' if usar_preview else 'Direto'}")
            
            # 5. Processar conforme modo
            if usar_preview:
                # FLUXO COM PREVIEW
                self.executar_fluxo_preview_despesas(configuracoes)
            else:
                # FLUXO DIRETO
                self.gerar_direto_otimizado(configuracoes)
                
        except Exception as e:
            logger.error(f"Erro no processamento otimizado: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Erro no processamento: {str(e)}")

    def teste_visualizador_isolado(self):
        """Teste isolado do visualizador - MÉTODO TEMPORÁRIO PARA DEBUG"""
        try:
            logger.info("🧪 === TESTE ISOLADO DO VISUALIZADOR ===")
            
            # 1. Testar import
            try:
                from src.relatorio_despesas_aprimorado import VisualizadorRelatorio
                logger.info("✅ TESTE 1: Import bem-sucedido")
            except Exception as e:
                logger.error(f"❌ TESTE 1: Falha no import: {str(e)}")
                return False
            
            # 2. Testar criação
            try:
                visualizador = VisualizadorRelatorio(self.root)
                logger.info("✅ TESTE 2: Criação bem-sucedida")
            except Exception as e:
                logger.error(f"❌ TESTE 2: Falha na criação: {str(e)}")
                return False
            
            # 3. Criar dados mínimos de teste
            dados_teste = {
                'df_filtrado': pd.DataFrame(),
                'df_diaria': pd.DataFrame(),
                'df_tp_desp_1': pd.DataFrame(),
                'df_tp_desp_2': pd.DataFrame(),
                'df_futuro': None,
                'df_original': pd.DataFrame(),
                'incluir_futuros': True,
                'incluir_excluidos': False,
                'data_relatorio': datetime(2025, 7, 5),
                'nome_cliente': 'TESTE CLIENTE',
                'endereco_cliente': 'ENDEREÇO TESTE',
                'numero_relatorio': 1,
                'acumulado': 1000.0
            }
            logger.info("✅ TESTE 3: Dados de teste criados")
            
            # 4. Testar mostrar_preview
            try:
                logger.info("🚀 TESTE 4: Chamando mostrar_preview...")
                preview_window = visualizador.mostrar_preview(dados_teste)
                logger.info(f"✅ TESTE 4: mostrar_preview retornou: {preview_window}")
                
                if preview_window is None:
                    logger.error("❌ TESTE 4: mostrar_preview retornou None")
                    return False
                    
                # 5. Testar se a janela existe
                try:
                    if hasattr(preview_window, 'winfo_exists'):
                        existe = preview_window.winfo_exists()
                        logger.info(f"✅ TESTE 5: Janela existe: {existe}")
                    else:
                        logger.warning("⚠️ TESTE 5: Janela não tem winfo_exists")
                    
                except Exception as e:
                    logger.error(f"❌ TESTE 5: Erro ao verificar janela: {str(e)}")
                
                return True
                
            except Exception as e:
                logger.error(f"❌ TESTE 4: Falha no mostrar_preview: {str(e)}", exc_info=True)
                return False
                
        except Exception as e:
            logger.error(f"💥 ERRO GERAL NO TESTE: {str(e)}", exc_info=True)
            return False

    def abrir_preview_minimalista(self, dados_completos, arquivo_path):
        """Versão minimalista para testar o básico"""
        try:
            logger.info("🧪 === PREVIEW MINIMALISTA ===")
            
            # Criar janela simples de preview
            preview_window = tk.Toplevel(self.root)
            preview_window.title("Preview Teste")
            preview_window.geometry("600x400")
            
            # Adicionar conteúdo básico
            frame = ttk.Frame(preview_window, padding=20)
            frame.pack(fill='both', expand=True)
            
            ttk.Label(frame, text="PREVIEW DE TESTE", font=('Arial', 16, 'bold')).pack(pady=10)
            ttk.Label(frame, text=f"Cliente: {dados_completos.get('nome_cliente', 'N/A')}").pack(pady=5)
            ttk.Label(frame, text=f"Data: {dados_completos.get('data_relatorio', 'N/A')}").pack(pady=5)
            ttk.Label(frame, text=f"Relatório nº: {dados_completos.get('numero_relatorio', 'N/A')}").pack(pady=5)
            ttk.Label(frame, text=f"Acumulado: R$ {dados_completos.get('acumulado', 0):,.2f}").pack(pady=5)
            
            # Mostrar dados processados
            ttk.Label(frame, text="DADOS PROCESSADOS:", font=('Arial', 12, 'bold')).pack(pady=(20,5))
            ttk.Label(frame, text=f"df_filtrado: {len(dados_completos.get('df_filtrado', []))} registros").pack()
            ttk.Label(frame, text=f"df_diaria: {len(dados_completos.get('df_diaria', []))} registros").pack()
            ttk.Label(frame, text=f"df_tp_desp_1: {len(dados_completos.get('df_tp_desp_1', []))} registros").pack()
            
            # Botão para fechar
            def fechar():
                preview_window.destroy()
                self.root.deiconify()
                self.root.lift()
                self.root.focus_force()
            
            ttk.Button(frame, text="Fechar e Voltar", command=fechar).pack(pady=20)
            
            # Configurar fechamento
            preview_window.protocol("WM_DELETE_WINDOW", fechar)
            
            # Ocultar interface principal
            self.root.withdraw()
            
            # Focar preview
            preview_window.lift()
            preview_window.focus_force()
            
            logger.info("✅ Preview minimalista aberto")
            
        except Exception as e:
            logger.error(f"💥 ERRO no preview minimalista: {str(e)}")
            messagebox.showerror("Erro", f"Erro: {str(e)}")
            self.root.deiconify()



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