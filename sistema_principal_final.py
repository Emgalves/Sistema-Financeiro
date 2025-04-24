"""
Sistema de Gestão Financeira - Versão Final Compilada
"""
import os
import sys
import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import importlib
import logging
from datetime import datetime
from pathlib import Path

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("sistema")

# Função para verificar se um arquivo ou diretório existe
def verificar_caminho(caminho):
    return os.path.exists(caminho)

# Adicionar os diretórios ao path do Python
def configurar_caminhos():
    # Determinar o diretório base
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    logger.info(f"Diretório base: {base_dir}")
    
    # Definir caminhos
    src_dir = os.path.join(base_dir, 'src')
    config_dir = os.path.join(src_dir, 'config')
    
    # Adicionar caminhos ao sys.path
    caminhos = [
        base_dir,
        src_dir,
        config_dir,
    ]
    
    for caminho in caminhos:
        if caminho not in sys.path and verificar_caminho(caminho):
            sys.path.insert(0, caminho)
            logger.info(f"Adicionado ao path: {caminho}")
    
    # Criar o arquivo __init__.py se não existir
    for diretorio in [src_dir, config_dir]:
        init_file = os.path.join(diretorio, '__init__.py')
        if verificar_caminho(diretorio) and not verificar_caminho(init_file):
            try:
                with open(init_file, 'w') as f:
                    f.write("# Arquivo criado automaticamente\n")
                logger.info(f"Criado arquivo {init_file}")
            except Exception as e:
                logger.warning(f"Não foi possível criar {init_file}: {str(e)}")
    
    return base_dir

# Configurar caminhos logo no início
base_dir = configurar_caminhos()

# Classes básicas para substituir módulos que podem estar ausentes
class SystemLogger:
    def __init__(self):
        self.logger = logger
        self.log_format = "%(asctime)s - %(levelname)s - %(message)s"
    
    def get_logger(self):
        return self.logger
    
    def set_user(self, username):
        logger.info(f"Usuário definido: {username}")

# Criar instância global
system_logger = SystemLogger()

# Decorator simplificado para log de ação
def log_action(action_name):
    def decorator(func):
        def wrapper(*args, **kwargs):
            logger.info(f"Ação: {action_name}")
            return func(*args, **kwargs)
        return wrapper
    return decorator

# Função para localizar recursos
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

# Stub para controle de versão
class VersionControl:
    @staticmethod
    def get_version_string():
        return "1.2.0"
    
    @staticmethod
    def save_version_history():
        # Versão simplificada que não precisa salvar em arquivo
        return []
    
    @staticmethod
    def show_version_dialog(parent):
        dialog = tk.Toplevel(parent)
        dialog.title(f"Sobre o Sistema - Versão 1.2.0")
        dialog.geometry("500x450")
        dialog.transient(parent)
        dialog.grab_set()
        
        # Centralizar a janela
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')
        
        # Frame principal
        main_frame = ttk.Frame(dialog, padding=10)
        main_frame.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            main_frame, 
            text=f"Sistema de Gestão Financeira", 
            font=('Helvetica', 16, 'bold')
        ).pack(pady=(0, 5))
        
        # Versão
        ttk.Label(
            main_frame, 
            text=f"Versão 1.2.0", 
            font=('Helvetica', 12)
        ).pack(pady=(0, 5))
        
        # Data de lançamento
        ttk.Label(
            main_frame, 
            text=f"Lançado em: 24/04/2025", 
            font=('Helvetica', 10)
        ).pack(pady=(0, 10))
        
        # Lista de mudanças
        changes_frame = ttk.LabelFrame(main_frame, text="Mudanças nesta versão", padding=10)
        changes_frame.pack(fill='both', expand=True, pady=10)
        
        # Lista simplificada de mudanças
        changes = [
            "Incluído botão para importar dados da folha a partir de planilha",
            "Criação do módulo para rateio de despesas comuns a todos os clientes",
            "Criação do módulo de controle de medições de empreiteiros",
            "Ajuste na aba Clientes, permitindo a inclusão do témino da obra ou do contrato",
            "Criação de uma interface com todos os relatórios",
            "Criação de relatório de Fornecedores",
            "Criação de relatório de Medições",
            "Criação de relatório de Tipos de Despesas",
            "Separação de relatório de Lançamentos Futuros",
            "Inclusão das etapas de obra nos contratos de administração de obras"
        ]
        
        # Usar Text em vez de ScrolledText para evitar dependências
        changes_text = tk.Text(changes_frame, wrap=tk.WORD, height=10)
        changes_text.pack(fill='both', expand=True)
        changes_text.insert(tk.END, "\n".join(f"• {change}" for change in changes))
        changes_text.config(state='disabled')  # Torna o texto somente leitura
        
        # Copyright
        ttk.Label(
            main_frame, 
            text="© 2025 Todos os direitos reservados.", 
            font=('Helvetica', 8)
        ).pack(pady=(10, 0))
        
        # Botão fechar
        ttk.Button(
            main_frame, 
            text="Fechar", 
            command=dialog.destroy
        ).pack(pady=10)

# Usar nossa implementação personalizada
version_control = VersionControl()

# Configuração de janela simplificada
def configurar_janela(janela, titulo, largura=900, altura=1000):
    """
    Configura o posicionamento e dimensionamento padrão de uma janela
    
    Args:
        janela: Instância de tk.Tk ou tk.Toplevel
        titulo: Título da janela
        largura: Largura desejada (default 900)
        altura: Altura desejada (default 900)
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
    
# Funções utilitárias que substituem o módulo config.utils
def obter_base_path():
    """Versão simplificada para obter o base_path"""
    # Dependendo do ambiente, usar um caminho padrão ou o do Google Drive
    drive_path = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios")
    if drive_path.exists():
        return str(drive_path / "Financeiro" / "Planilhas_Base")
    else:
        return os.path.join(base_dir, "dados")

def obter_pasta_clientes():
    """Versão simplificada para obter a pasta de clientes"""
    drive_path = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios")
    if drive_path.exists():
        return str(drive_path / "Financeiro" / "Clientes")
    else:
        pasta = os.path.join(base_dir, "dados", "clientes")
        os.makedirs(pasta, exist_ok=True)
        return pasta

# Configurações que substituem o módulo config.config
ARQUIVO_CLIENTES = os.path.join(obter_base_path(), "clientes.xlsx")
ARQUIVO_MODELO = os.path.join(obter_base_path(), "MODELO.xlsx")
PASTA_CLIENTES = obter_pasta_clientes()
BASE_PATH = obter_base_path()

# Importar módulos de forma segura
def importar_modulo_seguro(nome_modulo, pacote=None):
    """Tenta importar um módulo de várias formas e retorna None se falhar"""
    tentativas = [
        lambda: importlib.import_module(nome_modulo),
        lambda: importlib.import_module(f"src.{nome_modulo}"),
        lambda: importlib.import_module(nome_modulo, package=pacote),
    ]
    
    for tentativa in tentativas:
        try:
            return tentativa()
        except ImportError:
            continue
        except Exception as e:
            logger.error(f"Erro ao importar {nome_modulo}: {str(e)}")
            break
    
    logger.error(f"Não foi possível importar o módulo {nome_modulo}")
    return None

# Montar configurações globais para os módulos
configuracoes_globais = {
    "ARQUIVO_CLIENTES": ARQUIVO_CLIENTES,
    "ARQUIVO_MODELO": ARQUIVO_MODELO,
    "PASTA_CLIENTES": PASTA_CLIENTES,
    "BASE_PATH": BASE_PATH,
    "system_logger": system_logger,
    "log_action": log_action,
    "logger": logger,
    "configurar_janela": configurar_janela,
}

# Colocar configurações no sys.modules para simular imports
class ModuloSimulado:
    def __init__(self, atributos):
        for nome, valor in atributos.items():
            setattr(self, nome, valor)

# Criar módulos simulados
config_module = ModuloSimulado({
    "ARQUIVO_CLIENTES": ARQUIVO_CLIENTES,
    "ARQUIVO_MODELO": ARQUIVO_MODELO,
    "PASTA_CLIENTES": PASTA_CLIENTES,
    "BASE_PATH": BASE_PATH,
})

utils_module = ModuloSimulado({
    "obter_base_path": obter_base_path,
    "obter_pasta_clientes": obter_pasta_clientes,
})

logger_config_module = ModuloSimulado({
    "system_logger": system_logger,
    "log_action": log_action,
})

# Adicionar módulos simulados ao sys.modules
sys.modules["config"] = ModuloSimulado({})
sys.modules["config.config"] = config_module
sys.modules["config.utils"] = utils_module
sys.modules["config.logger_config"] = logger_config_module

# Classe principal do sistema
class SistemaPrincipal:
    def __init__(self):
        self.usuario_atual = None
        self.root = tk.Tk()
        
        # Configurar a janela principal
        titulo_com_versao = f"Sistema de Gestão Financeira v{version_control.get_version_string()}"
        configurar_janela(self.root, titulo_com_versao)
        
        # Salvar histórico de versões
        version_control.save_version_history()
        
        # Inicializar gerenciador de taxas
        self.controlador_taxas = self.criar_controlador_taxas()
        
        # Configurar estilos e conteúdo
        self.setup_style()
        self.create_main_content()
    
    def criar_controlador_taxas(self):
        """Cria uma instância do controlador de taxas"""
        try:
            modulo = importar_modulo_seguro("controle_pagamentos_taxas")
            if modulo and hasattr(modulo, 'ControlePagamentos'):
                return modulo.ControlePagamentos(self.root)
        except Exception as e:
            logger.error(f"Erro ao criar controlador de taxas: {str(e)}")
        
        # Controlador stub se falhar
        class ControladorTaxasStub:
            def __init__(self, parent=None):
                self.parent = parent
            
            def abrir_janela_controle(self):
                messagebox.showinfo("Em Desenvolvimento", 
                                  "Esta funcionalidade está sendo atualizada.\n\n"
                                  "Por favor, tente novamente mais tarde.")
        
        return ControladorTaxasStub(self.root)
    
    def setup_style(self):
        """Configura o estilo visual do aplicativo"""
        style = ttk.Style()
        style.configure('Menu.TFrame', background='#f0f0f0')
        style.configure('Card.TFrame', background='white')
        style.configure('CardTitle.TLabel', 
                      font=('Helvetica', 14, 'bold'),
                      background='white')
        style.configure('CardDesc.TLabel',
                      font=('Helvetica', 10),
                      background='white',
                      wraplength=300)
        style.configure('Action.TButton',
                      font=('Helvetica', 12),
                      padding=10)
    
    def create_main_content(self):
        """Cria o conteúdo principal da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        # Logo
        try:
            self.logo_path = resource_path("logo.png")
            self.logo = PhotoImage(file=self.logo_path)
            logo_label = ttk.Label(main_frame, image=self.logo)
            logo_label.pack(pady=10)
        except Exception as e:
            logger.error(f"Erro ao carregar logo: {str(e)}")
        
        # Título
        title_label = ttk.Label(
            main_frame,
            text="Sistema de Gestão Financeira",
            font=('Helvetica', 24, 'bold')
        )
        title_label.pack(pady=(0, 30))
        
        # Grid para cards
        grid = ttk.Frame(main_frame)
        grid.pack(expand=True, pady=20)
        
        # Verificar quais módulos estão disponíveis
        modulos_disponiveis = {
            "Sistema_Entrada_Dados": self.verificar_modulo("Sistema_Entrada_Dados", "SistemaEntradaDados"),
            "controle_pagamentos_taxas": True,  # Já criamos o stub
            "despesas_rateadas": self.verificar_modulo("despesas_rateadas", "InterfaceDespesasRateadas"),
            "relatorios_interface": self.verificar_modulo("relatorios_interface", "SistemaRelatorios"),
            "gestao_medicoes": self.verificar_modulo("gestao_medicoes", "GestaoMedicoes"),
            "configuracoes_sistema": self.verificar_modulo("configuracoes_sistema", "GerenciadorConfiguracoes")
        }
        
        # Status dos módulos
        status_texto = {True: "", False: " [Em Manutenção]"}
        
        # Cards do sistema
        self.create_card(grid, f"Entrada de Dados{status_texto[modulos_disponiveis['Sistema_Entrada_Dados']]}", 
                       "Cadastro e gestão de dados", 
                       self.abrir_entrada_dados if modulos_disponiveis['Sistema_Entrada_Dados'] else self.modulo_em_manutencao, 
                       0, 0)
        
        self.create_card(grid, f"Taxas de Administração{status_texto[modulos_disponiveis['controle_pagamentos_taxas']]}", 
                       "Gestão completa de taxas administrativas",
                       self.abrir_gestao_taxas, 0, 1)
        
        self.create_card(grid, f"Despesas Rateadas{status_texto[modulos_disponiveis['despesas_rateadas']]}", 
                       "Gerenciamento de despesas compartilhadas entre clientes", 
                       self.abrir_despesas_rateadas if modulos_disponiveis['despesas_rateadas'] else self.modulo_em_manutencao, 
                       1, 0)
        
        self.create_card(grid, f"Geração de Relatórios{status_texto[modulos_disponiveis['relatorios_interface']]}", 
                       "Visualização de relatórios",
                       self.abrir_relatorios if modulos_disponiveis['relatorios_interface'] else self.modulo_em_manutencao, 
                       1, 1)
        
        self.create_card(grid, f"Gestão de Medições{status_texto[modulos_disponiveis['gestao_medicoes']]}", 
                       "Gerenciar contratos com empreiteros e por entregas",
                       self.abrir_gestao_medicoes if modulos_disponiveis['gestao_medicoes'] else self.modulo_em_manutencao, 
                       2, 0)
        
        self.create_card(grid, f"Configurações do Sistema{status_texto[modulos_disponiveis['configuracoes_sistema']]}", 
                       "Gerenciar parâmetros básicos",
                       self.abrir_configuracoes if modulos_disponiveis['configuracoes_sistema'] else self.modulo_em_manutencao, 
                       2, 1)
        
        # Frame para botões inferiores
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(pady=20)
        
        # Versão e botão Sobre à esquerda do botão Sair
        version_frame = ttk.Frame(bottom_frame)
        version_frame.pack(side='left', padx=20)
        
        # Label com a versão
        version_label = ttk.Label(
            version_frame,
            text=f"Versão {version_control.get_version_string()}",
            font=('Helvetica', 9),
            foreground='#555555'
        )
        version_label.pack(pady=5)
        
        # Botão Sobre
        about_button = ttk.Button(
            bottom_frame,
            text="Sobre",
            command=lambda: version_control.show_version_dialog(self.root)
        )
        about_button.pack(side='left', padx=10)
        
        # Botão Sair
        exit_btn = ttk.Button(bottom_frame, text="Sair", 
                            command=self.sair_sistema)
        exit_btn.pack(side='right', padx=5)
    
    def create_card(self, parent, title, description, command, row, col):
        """Cria um card na interface"""
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
        
        title_label = ttk.Label(
            card,
            text=title,
            style='CardTitle.TLabel'
        )
        title_label.pack(pady=(20, 10), padx=20)
        
        desc_label = ttk.Label(
            card,
            text=description,
            style='CardDesc.TLabel'
        )
        desc_label.pack(pady=(0, 20), padx=20)
        
        button = ttk.Button(
            card,
            text="Acessar",
            style='Action.TButton',
            command=command
        )
        button.pack(pady=(0, 20))
    
    def verificar_modulo(self, nome_modulo, classe_modulo):
        """Verifica se um módulo pode ser carregado"""
        try:
            modulo = importar_modulo_seguro(nome_modulo)
            return modulo is not None and hasattr(modulo, classe_modulo)
        except:
            return False
    
    def modulo_em_manutencao(self):
        """Exibe mensagem de módulo em manutenção"""
        messagebox.showinfo("Em Manutenção", 
                          "Este módulo está temporariamente indisponível.\n\n"
                          "Por favor, tente novamente mais tarde.")
    
    def abrir_modulo(self, nome_modulo, classe_modulo, *args, **kwargs):
        """Abre um módulo se disponível, caso contrário mostra mensagem"""
        try:
            modulo = importar_modulo_seguro(nome_modulo)
            if modulo and hasattr(modulo, classe_modulo):
                # Ocultar janela principal
                self.root.withdraw()
                
                # Injetar configurações globais no módulo
                for nome, valor in configuracoes_globais.items():
                    if not hasattr(modulo, nome):
                        setattr(modulo, nome, valor)
                
                # Inicializar o módulo
                classe = getattr(modulo, classe_modulo)
                instancia = classe(*args, **kwargs)
                
                # Configurações adicionais
                if hasattr(instancia, 'menu_principal'):
                    instancia.menu_principal = self.root
                
                # Se o módulo tem uma janela root própria
                if hasattr(instancia, 'root'):
                    # Configurar comportamento ao fechar
                    instancia.root.protocol("WM_DELETE_WINDOW", 
                        lambda: self.finalizar_sistema(instancia.root))
                    
                    # Exibir janela
                    instancia.root.lift()
                    instancia.root.focus_force()
                    
                    # Se o módulo tem um método run, chamá-lo
                    if hasattr(instancia, 'run') and callable(instancia.run):
                        instancia.run()
                    else:
                        instancia.root.mainloop()
                else:
                    # Se não tiver root próprio, reexibir a janela principal
                    self.root.deiconify()
                
                return True
            else:
                raise ImportError(f"Módulo {nome_modulo} não encontrado ou classe {classe_modulo} não disponível")
        except Exception as e:
            logger.error(f"Erro ao abrir módulo {nome_modulo}: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir módulo: {str(e)}")
            self.root.deiconify()
            return False
    
    def abrir_entrada_dados(self):
        """Abre o sistema de entrada de dados"""
        self.abrir_modulo("Sistema_Entrada_Dados", "SistemaEntradaDados", parent=self.root)
    
    def abrir_gestao_taxas(self):
        """Abre o menu de gestão de taxas"""
        try:
            self.controlador_taxas.abrir_janela_controle()
        except Exception as e:
            messagebox.showerror("Erro",
                               f"Erro ao abrir gestão de taxas: {str(e)}")
    
    @log_action("Gerar relatório")
    def abrir_relatorios(self):
        """Abre o sistema integrado de relatórios"""
        # Primeiro tenta o sistema integrado
        if self.abrir_modulo("relatorios_interface", "SistemaRelatorios", parent=self.root):
            return
            
        # Fallback para o relatório legado
        if self.abrir_modulo("relatorio_despesas_aprimorado", "RelatorioUI", self.root):
            return
            
        # Se ambos falharem
        messagebox.showerror("Erro", "Não foi possível carregar o módulo de relatórios.")
    
    def abrir_despesas_rateadas(self):
        """Abre o sistema de despesas rateadas"""
        self.abrir_modulo("despesas_rateadas", "InterfaceDespesasRateadas", self.root)
    
    def abrir_gestao_medicoes(self):
        """Abre o sistema de gestão de medições"""
        self.abrir_modulo("gestao_medicoes", "GestaoMedicoes", parent=self.root)
    
    def abrir_configuracoes(self):
        """Abre as configurações do sistema"""
        self.abrir_modulo("configuracoes_sistema", "GerenciadorConfiguracoes", parent=self.root)
    
    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        logger.info("Finalizando janela secundária...")
        try:
            # Primeiro destruir a janela
            janela.destroy()
        except Exception as e:
            logger.error(f"Erro ao destruir janela: {str(e)}")
        
        # Mostrar a janela principal novamente
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
    
    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            logger.info("Saída confirmada, finalizando sistema...")
            self.root.destroy()
    
    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()

def force_exit():
    """Força a saída do programa"""
    import os
    os._exit(0)

def main():
    try:
        app = SistemaPrincipal()
        app.run()
    except Exception as e:
        error_msg = f"Erro ao iniciar o sistema: {str(e)}"
        logger.error(error_msg)
        import traceback
        logger.error(traceback.format_exc())
        messagebox.showerror("Erro", error_msg)
    finally:
        # Garantir que o programa será encerrado
        try:
            import threading
            threading.Timer(1.0, force_exit).start()
        except:
            pass

# Executar o aplicativo
if __name__ == "__main__":
    main()