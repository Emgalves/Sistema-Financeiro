import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
import importlib
from datetime import datetime
from pathlib import Path
import logging

# Configuração básica de logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(f"logs/sistema_relatorios_{datetime.now().strftime('%Y%m%d')}.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)

logger = logging.getLogger("sistema_relatorios")

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
    from config.window_config import configurar_janela
except ImportError:
    # Implementação básica caso o módulo não seja encontrado
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
        configurar_janela(self.root, "Sistema Integrado de Relatórios", 900, 980)
        
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
    
    def setup_relatorios_list(self):
        """Configura a lista de relatórios disponíveis"""
        # Definir os relatórios disponíveis
        self.relatorios = [
            {
                "id": "despesas",
                "nome": "Relatório de Despesas",
                "descricao": "Relatório financeiro de despesas por cliente",
                "modulo": "relatorio_despesas_aprimorado",
                "classe": "RelatorioUI",
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
                "disponivel": False
            },
            {
                "id": "tipo_despesa",
                "nome": "Relatório por Tipo de Despesa",
                "descricao": "Análise detalhada por tipo de despesa",
                "modulo": "relatorio_tipo_despesa",
                "classe": "RelatorioTipoDespesa",
                "disponivel": False
            },
            {
                "id": "fornecedores",
                "nome": "Relatório de Principais Fornecedores",
                "descricao": "Resumo de fornecedores por cliente e global",
                "modulo": "relatorio_fornecedores",
                "classe": "RelatorioFornecedores",
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
        elif relatorio["id"] == "fornecedores":
            self.setup_opcoes_fornecedores(opcoes_frame)  # Adicionar esta condição
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
        
        # Frame para seleção de cliente
        frame_cliente = ttk.Frame(parent_frame)
        frame_cliente.pack(fill='x', padx=10, pady=10)
        
        ttk.Label(frame_cliente, text="Cliente:").pack(side='left', padx=5)
        
        # Combobox para seleção de cliente
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=40)
        self.cliente_combobox.pack(side='left', padx=5)
        
        # Preencher com alguns clientes exemplo (você deve carregá-los do arquivo real)
        self.cliente_combobox['values'] = ['Todos os Clientes', 'Cliente 1', 'Cliente 2', 'Cliente 3']
        self.cliente_combobox.current(0)
        
        # Botão para selecionar arquivo individual
        ttk.Button(
            parent_frame,
            text="Selecionar Arquivo de Cliente",
            command=self.selecionar_arquivo_cliente
        ).pack(anchor='w', padx=15, pady=10)
        
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
        
        # Preencher com alguns clientes exemplo (você deve carregá-los do arquivo real)
        self.cliente_contratos['values'] = ['Todos os Clientes', 'Cliente 1', 'Cliente 2', 'Cliente 3']
        self.cliente_contratos.current(0)
        
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
        
        # Preencher com alguns clientes exemplo (você deve carregá-los do arquivo real)
        self.cliente_contratos['values'] = ['Todos os Clientes', 'Cliente 1', 'Cliente 2', 'Cliente 3']
        self.cliente_contratos.current(0)
        
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
            
            # Para o relatório de fornecedores, usar uma abordagem mais direta
            if relatorio["id"] == "fornecedores":
                print("Iniciando relatório de fornecedores")
                self.root.withdraw()
                
                try:
                    # Importação direta
                    from relatorio_fornecedores import RelatorioFornecedores
                    app = RelatorioFornecedores(parent=self.root)
                    app.menu_principal = self.root
                    app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
                    app.root.lift()
                    app.root.focus_force()
                    app.root.mainloop()
                    return
                except ImportError:
                    # Tentar da pasta src
                    try:
                        from src.relatorio_fornecedores import RelatorioFornecedores
                        app = RelatorioFornecedores(parent=self.root)
                        app.menu_principal = self.root
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
            
            # Continuar com o fluxo normal para outros relatórios
            modulo = self.carregar_modulo(relatorio["modulo"])
            if not modulo:
                return
                
            # Obter a classe do relatório
            classe_relatorio = getattr(modulo, relatorio["classe"])
            
            # Iniciar interface conforme o tipo de relatório
            if relatorio["id"] == "despesas":
                self.iniciar_relatorio_despesas(classe_relatorio)
            elif relatorio["id"] == "contratos":
                self.iniciar_relatorio_contratos(classe_relatorio)
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
        """Inicia a geração do relatório de despesas"""
        # Esconder a janela atual
        self.root.withdraw()
        
        # Criar uma nova janela para o relatório
        relatorio_window = tk.Toplevel(self.root)
        
        # Inicializar o relatório
        app_relatorio = classe_relatorio(relatorio_window)
        
        # Configurar menu principal para retornar
        app_relatorio.menu_principal = self.root
        
        # Se houver data selecionada, passá-la
        if hasattr(self, 'data_entry'):
            app_relatorio.data_selecionada.set(self.data_entry.get())
        
        # Configurar inclusão de lançamentos futuros
        if hasattr(self, 'incluir_futuros'):
            app_relatorio.incluir_futuros.set(self.incluir_futuros.get())
        
        # Atualizar cliente se específico foi selecionado
        if self.cliente_combobox.get() != 'Todos os Clientes' and not self.cliente_combobox.get().startswith('Arquivo:'):
            app_relatorio.selecionar_cliente_nome(self.cliente_combobox.get())
        elif hasattr(self, 'arquivo_cliente_selecionado'):
            app_relatorio.selecionar_arquivo_direto(self.arquivo_cliente_selecionado)
        
        # Configurar comportamento ao fechar
        relatorio_window.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(relatorio_window))
        
        # Exibir janela
        relatorio_window.lift()
        relatorio_window.focus_force()
        relatorio_window.mainloop()
    
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
            self.root.deiconify()  # Mostrar a janela principal novamente
        
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
        """Carrega a lista de clientes do arquivo de clientes"""
        try:
            # Importar openpyxl apenas quando necessário
            from openpyxl import load_workbook
            
            # Caminho para o arquivo de clientes (ajuste conforme necessário)
            try:
                from config.config import ARQUIVO_CLIENTES
            except ImportError:
                # Caminho padrão se não conseguir importar das configurações
                ARQUIVO_CLIENTES = "dados/clientes.xlsx"
            
            # Verificar se o arquivo existe
            if not os.path.exists(ARQUIVO_CLIENTES):
                logger.warning(f"Arquivo de clientes não encontrado: {ARQUIVO_CLIENTES}")
                return ['Todos os Clientes']
            
            # Carregar workbook
            workbook = load_workbook(ARQUIVO_CLIENTES)
            sheet = workbook['Clientes']  # Assumindo que existe uma aba chamada 'Clientes'
            
            # Extrair nomes dos clientes (pulando o cabeçalho)
            clientes = ['Todos os Clientes']
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Nome do cliente está na primeira coluna
                    clientes.append(row[0])
            
            workbook.close()
            return clientes
            
        except Exception as e:
            logger.error(f"Erro ao carregar clientes: {str(e)}", exc_info=True)
            return ['Todos os Clientes']
    
    def selecionar_cliente_nome(self, nome_cliente):
        """Método stub para selecionar cliente por nome"""
        pass
    
    def selecionar_arquivo_direto(self, caminho_arquivo):
        """Método stub para selecionar arquivo diretamente"""
        pass
    
    def run(self):
        """Inicia a execução do sistema"""
        # Carregar lista de clientes
        try:
            clientes = self.carregar_clientes()
            
            # Verificar se os comboboxes foram criados
            if hasattr(self, 'cliente_combobox') and self.cliente_combobox is not None:
                self.cliente_combobox['values'] = clientes
                self.cliente_combobox.current(0)
                
            if hasattr(self, 'cliente_contratos') and self.cliente_contratos is not None:
                self.cliente_contratos['values'] = clientes
                self.cliente_contratos.current(0)
                
        except Exception as e:
            logger.error(f"Erro ao carregar lista de clientes: {str(e)}", exc_info=True)
        
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