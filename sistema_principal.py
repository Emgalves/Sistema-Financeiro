"""
Script principal simplificado com foco em caminhos corretos para PyInstaller
"""
import os
import sys
import logging
from pathlib import Path

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("sistema")

# Ajustar caminhos baseado no modo de execução
if getattr(sys, 'frozen', False):
    # PyInstaller
    base_dir = Path(sys._MEIPASS)
    logger.info(f"Executando a partir do PyInstaller em: {base_dir}")
    
    # Verificar arquivos empacotados
    logger.info("Arquivos no pacote:")
    for root, dirs, files in os.walk(base_dir):
        if 'src/config' in root or 'src\\config' in root:
            for file in files:
                logger.info(f"  - {os.path.join(root, file)}")
else:
    # Execução normal
    base_dir = Path(__file__).resolve().parent
    logger.info(f"Executando em modo normal a partir de: {base_dir}")

# Garantir que caminhos importantes estão no sys.path
src_dir = base_dir / 'src'
config_dir = src_dir / 'config'

for path in [str(base_dir), str(src_dir), str(config_dir)]:
    if path not in sys.path:
        sys.path.insert(0, path)
        logger.info(f"Adicionado ao path: {path}")

# Verificar módulos primeiro
try:
    logger.info("Verificando módulos de configuração...")
    
    # Tentativa 1: Importação absoluta com src
    try:
        import src.config
        logger.info("Módulo src.config importado com sucesso")
        
        # Tentar importar submódulos
        import src.config.utils
        import src.config.logger_config
        import src.config.window_config
        import src.config.config
        logger.info("Todos os submódulos de src.config importados com sucesso")
    except ImportError as e:
        logger.error(f"Erro ao importar via src.config: {e}")
        
        # Tentativa 2: Importação direta
        try:
            import config
            logger.info("Módulo config importado com sucesso")
            
            import config.utils
            import config.logger_config
            import config.window_config
            import config.config
            logger.info("Todos os submódulos de config importados com sucesso")
        except ImportError as e:
            logger.error(f"Erro ao importar via config direta: {e}")
            
            # Se chegou aqui, há um problema sério
            logger.error("FALHA CRÍTICA: Não foi possível importar os módulos de configuração")
            if getattr(sys, 'frozen', False):
                # Mostrar mensagem e pausar no modo compilado
                print("\nERRO CRÍTICO: Módulos de configuração não encontrados!")
                print("Este problema geralmente ocorre quando o PyInstaller não empacota corretamente os arquivos.")
                input("Pressione ENTER para sair...")
                sys.exit(1)
except Exception as e:
    logger.error(f"Erro ao verificar módulos: {str(e)}")
    if getattr(sys, 'frozen', False):
        input("Erro crítico. Pressione ENTER para sair...")
        sys.exit(1)

import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import importlib
import sys
import os
import logging
from io import StringIO
from datetime import datetime
from dotenv import load_dotenv
load_dotenv()

def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

from src.config.window_config import configurar_janela
from src.controle_pagamentos_taxas import ControlePagamentos as ControladorTaxas

# Onde você importa o logger
from src.config.logger_config import system_logger, log_action
print("Logger importado de src.config com sucesso")

from src.config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )

def force_exit():
    """Força a saída do programa"""
    print("Forçando encerramento do programa...")
    import os
    os._exit(0)

try:
    from src.controle_pagamentos_taxas import ControlePagamentos as ControladorTaxas
except ImportError:
    try:
        from controle_pagamentos_taxas import ControlePagamentos as ControladorTaxas
    except ImportError as e:
        print(f"Erro ao importar ControlePagamentos: {str(e)}")
        # Criar stub básico se o módulo não existir
        class ControladorTaxasStub:
            def __init__(self, parent=None):
                self.parent = parent
                
            def abrir_janela_controle(self):
                import tkinter.messagebox as messagebox
                messagebox.showerror("Erro", "Módulo de Controle de Pagamentos não encontrado")
                
        ControladorTaxas = ControladorTaxasStub

# Importar o módulo de controle de versões
try:
    import version_control
except ImportError:
    try:
        from src import version_control
    except ImportError:
        # Criar stub básico se o módulo não existir
        class VersionControlStub:
            @staticmethod
            def get_version_string():
                return "1.0.0"
            @staticmethod
            def show_version_dialog(parent):
                messagebox.showinfo("Versão", "Sistema de Gestão Financeira v1.0.0")
            @staticmethod
            def save_version_history():
                return []
        
        version_control = VersionControlStub()


def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)


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
        self.controlador_taxas = ControladorTaxas(self.root)
        
        # Configurar estilos e conteúdo
        self.setup_style()
        self.create_main_content()
        
    def login(self, username):
        self.usuario_atual = username
        system_logger.set_user(username)
        logger.info(f"Login realizado") # type: ignore


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
        self.logo_path = resource_path(os.path.join("recursos", "imagens", "logo.png"))
        self.logo = PhotoImage(file=self.logo_path)
        logo_label = ttk.Label(main_frame, image=self.logo)
        logo_label.pack(pady=10)

        # Título (sem a versão ao lado)
        title_label = ttk.Label(
            main_frame,
            text="Sistema de Gestão Financeira",
            font=('Helvetica', 24, 'bold'),
            background='#f0f0f0'
        )
        title_label.pack(pady=(0, 30))

        # Grid para cards
        grid = ttk.Frame(main_frame)
        grid.pack(expand=True, pady=20)

        # Cards do sistema
        self.create_card(grid, "Entrada de Dados", 
                        "Cadastro e gestão de dados", 
                        self.abrir_entrada_dados, 0, 0)
        
        self.create_card(grid, "Taxas de Administração",
                        "Gestão completa de taxas administrativas",
                        self.abrir_gestao_taxas, 0, 1)
        
        self.create_card(grid, "Despesas Rateadas", 
                "Gerenciamento de despesas compartilhadas entre clientes", 
                self.abrir_despesas_rateadas, 1, 0)
        
        self.create_card(grid, "Geração de Relatórios",
                        "Visualização de relatórios",
                        self.abrir_relatorios, 1, 1)
                        
        self.create_card(grid, "Gestão de Medições",
                        "Gerenciar contratos com empreiteros e por entregas",
                        self.abrir_gestao_medicoes, 2, 0)
                        
        self.create_card(grid, "Configurações do Sistema",
                        "Gerenciar parâmetros básicos",
                        self.abrir_configuracoes, 2, 1)
        
        
        # Frame para botões inferiores (Sobre, Versão e Sair)
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
        
        # Botão Sair em destaque (lado direito)
        adicionar_btn = ttk.Button(bottom_frame, text="Sair", 
                                command=self.sair_sistema,
                                style='Medium.TButton')
        adicionar_btn.pack(side='right', padx=5)
        
        # Configurar um estilo especial para o botão Adicionar (opcional)
        style = ttk.Style()
        style.configure('Destaque.TButton', 
                    background='#0056b3',  # Esta propriedade pode não ter efeito em todos os temas
                    font=('Arial', 11, 'bold'))
        adicionar_btn.configure(style='Destaque.TButton')

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

    def abrir_entrada_dados(self):
        """Abre o sistema de entrada de dados"""
        try:
            logger = system_logger.get_logger()
            logger.debug("Iniciando abertura do sistema de entrada de dados")
            
            try:
                # Primeira tentativa: importar diretamente
                logger.debug("Tentando importar diretamente...")
                from Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                # Segunda tentativa: importar de src
                logger.debug("Tentando importar de src...")
                from src.Sistema_Entrada_Dados import SistemaEntradaDados
            
            self.root.withdraw()
            
            app = SistemaEntradaDados(parent=self.root)
            
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()

        except Exception as e:
            logger = system_logger.get_logger()
            logger.error(f"Erro ao abrir sistema de entrada de dados: {str(e)}", exc_info=True)
            messagebox.showerror("Erro",
                "Erro ao abrir sistema de entrada de dados. Por favor, contate o suporte.")
            self.root.deiconify()

    def abrir_gestao_taxas(self):
        """Abre o menu de gestão de taxas"""
        try:
            # Agora chama diretamente o método abrir_janela_controle do ControladorTaxas
            self.controlador_taxas.abrir_janela_controle()
        except Exception as e:
            messagebox.showerror("Erro",
                f"Erro ao abrir gestão de taxas: {str(e)}")

    @log_action("Gerar relatório")
    def abrir_relatorios(self):
        """Abre o sistema integrado de relatórios"""
        try:
            # Importar o novo módulo de sistema de relatórios
            modulo = self.reload_module('relatorios_interface')
            if not modulo:
                # Fallback para o sistema de relatórios antigo
                modulo = self.reload_module('relatorio_despesas_aprimorado')
                if not modulo:
                    messagebox.showerror("Erro", "Não foi possível carregar o módulo de relatórios.")
                    return
                
                # Se o sistema integrado falhou, mas o módulo de relatório de despesas funcionou
                self.root.withdraw()
                relatorio_window = tk.Toplevel(self.root)
                
                app = modulo.RelatorioUI(relatorio_window)
                app.menu_principal = self.root
                
                relatorio_window.protocol("WM_DELETE_WINDOW", 
                    lambda: self.finalizar_sistema(relatorio_window))
                
                relatorio_window.lift()
                relatorio_window.focus_force()
                relatorio_window.mainloop()
                return
                
            # Se o sistema integrado foi carregado com sucesso
            self.root.withdraw()
            
            # Iniciar o sistema de relatórios integrado
            app = modulo.SistemaRelatorios(parent=self.root)
            
            # Definir comportamento ao fechar
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            
            # Exibir janela
            app.root.lift()
            app.root.focus_force()
            app.run()
            
        except Exception as e:
            messagebox.showerror("Erro",
                f"Erro ao abrir sistema de relatórios: {str(e)}")
            self.root.deiconify()

    def abrir_despesas_rateadas(self):
        """Abre o sistema de despesas rateadas"""
        try:
            modulo = self.reload_module('despesas_rateadas')
            if not modulo:
                return

            self.root.withdraw()
            rateio_window = tk.Toplevel(self.root)
            
            app = modulo.InterfaceDespesasRateadas(rateio_window)
            app.menu_principal = self.root
            
            rateio_window.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(rateio_window))
            
            rateio_window.lift()
            rateio_window.focus_force()
            rateio_window.mainloop()
            
        except Exception as e:
            messagebox.showerror("Erro",
                f"Erro ao abrir sistema de despesas rateadas: {str(e)}")
            self.root.deiconify()

    def abrir_gestao_medicoes(self):
        """Abre o sistema de gestão de medições"""
        try:
            # Recarregar o módulo para garantir que as alterações sejam aplicadas
            modulo = self.reload_module('gestao_medicoes')
            if not modulo:
                return

            # Inicializar a classe GestaoMedicoes
            self.root.withdraw()  # Ocultar a janela principal
            app = modulo.GestaoMedicoes(parent=self.root)  # Passar self.root como parent
            
            # Configurar o comportamento ao fechar a janela
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            
            # Exibir a janela
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()
            
        except Exception as e:
            # Exibir mensagem de erro e reexibir a janela principal
            messagebox.showerror("Erro", f"Erro ao abrir sistema de gestão de medições: {str(e)}")
            self.root.deiconify()


    def reload_module(self, module_name):
        """
        Recarrega um módulo e retorna a versão atualizada
        Args:
            module_name (str): Nome do módulo a ser recarregado
        Returns:
            module: Módulo recarregado
        """
        try:
            # Remover todas as referências ao módulo e seus submódulos
            for key in list(sys.modules.keys()):
                if key == module_name or key.startswith(f"{module_name}."):
                    del sys.modules[key]
            
            # Importar o módulo novamente
            module = importlib.import_module(module_name)
            return module
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar módulo {module_name}: {str(e)}")
            return None
        

    def abrir_configuracoes(self):
        try:
            from configuracoes_sistema import GerenciadorConfiguracoes
            self.root.withdraw()
            app = GerenciadorConfiguracoes(parent=self.root)
            app.menu_principal = self.root  # Passa a referência correta do menu principal
            app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
            app.run()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir configurações do sistema: {str(e)}")
            self.root.deiconify()


    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        print("Finalizando janela secundária...")
        try:
            # Primeiro destruir a janela
            janela.destroy()
        except Exception as e:
            print(f"Erro ao destruir janela: {str(e)}")
        
        # Mostrar a janela principal novamente
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            print("Saída confirmada, finalizando sistema...")
            self.root.destroy()
            # Aguardar brevemente antes de forçar a saída
            self.root.after(200, force_exit)

    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()


class OutputManager:
    def __init__(self, logger=None):
        self.dev_mode = os.getenv('DEV_MODE', 'False').lower() == 'true'
        self.logger = logger
        
        # Remover redirecionamento de output que está causando problemas
        self.stdout_buffer = None
        self.stderr_buffer = None
        self.original_stdout = None
        self.original_stderr = None
    
    def start(self):
        """Método simplificado que não faz redirecionamento"""
        pass
    
    def stop(self):
        """Método simplificado que não faz redirecionamento"""
        pass
    
    def get_output(self):
        """Retorna None em vez de tentar acessar buffers"""
        return None

def main():
    # Tentar importar o logger, mas criar substituto se falhar
    try:
        from src.config.logger_config import system_logger
    except ImportError:
        # Criar logger substituto simples
        import logging
        class SimpleLogger:
            def __init__(self):
                self.logger = logging.getLogger("sistema")
                handler = logging.StreamHandler()
                self.logger.addHandler(handler)
                self.log_format = "%(asctime)s - %(levelname)s - %(message)s"
            
            def get_logger(self):
                return self.logger
                
            def set_user(self, username):
                pass
        
        system_logger = SimpleLogger()
    
    # Não usar o OutputManager para redirecionamento
    try:
        app = SistemaPrincipal()
        app.run()
    except Exception as e:
        print(f"Erro no sistema principal: {str(e)}")
        import traceback
        traceback.print_exc()

# Executar o aplicativo
if __name__ == "__main__":
    main()