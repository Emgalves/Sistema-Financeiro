# Diagnóstico imediato - coloque no início de sistema_principal.py
try:
    with open("diagnostico_sistema.log", "w") as log:
        import os, sys, platform
        from pathlib import Path
        
        log.write(f"=== Diagnóstico do Sistema ===\n")
        log.write(f"Data/Hora: {__import__('datetime').datetime.now()}\n")
        log.write(f"Sistema: {platform.system()} {platform.release()}\n")
        log.write(f"Diretório atual: {os.getcwd()}\n")
        log.write(f"SISTEMA_AMBIENTE: {os.getenv('SISTEMA_AMBIENTE', 'NÃO DEFINIDO')}\n")
        
        # Verificar caminho do Google Drive
        drive_path = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios")
        log.write(f"Caminho do Drive existe? {drive_path.exists()}\n")
        
        # Se existir, listar diretórios
        if drive_path.exists():
            log.write("Diretórios encontrados:\n")
            for item in drive_path.iterdir():
                if item.is_dir():
                    log.write(f" - {item.name}\n")
except Exception as e:
    with open("erro_diagnostico.log", "w") as err_log:
        err_log.write(f"Erro no diagnóstico: {str(e)}")

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

try:
    from src.config.window_config import configurar_janela
except ImportError:
    from config.window_config import configurar_janela

try:
    from config.logger_config import system_logger, log_action
    print("Logger importado com sucesso")
except ImportError as e:
    print(f"Erro ao importar logger: {str(e)}")

try:
    from src.config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )
except ImportError:
    from config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )

from src.gestao_taxas import GestaoTaxasAdministracao

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
        self.gestao_taxas = GestaoTaxasAdministracao(self.root)
        
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
        self.logo_path = resource_path("logo.png")
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
        
        self.create_card(grid, "Geração de Relatórios",
                        "Visualização de relatórios",
                        self.abrir_relatorios, 0, 2)
                        
        self.create_card(grid, "Configurações do Sistema",
                        "Gerenciar parâmetros básicos",
                        self.abrir_configuracoes, 1, 0)
        
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
            self.gestao_taxas.abrir_menu_taxas()
        except Exception as e:
            messagebox.showerror("Erro",
                f"Erro ao abrir gestão de taxas: {str(e)}")

    @log_action("Gerar relatório")
    def abrir_relatorios(self):
        """Abre o sistema de relatórios"""
        try:
            modulo = self.reload_module('relatorio_despesas_aprimorado')
            if not modulo:
                return

            self.root.withdraw()
            relatorio_window = tk.Toplevel(self.root)
            
            app = modulo.RelatorioUI(relatorio_window)
            app.menu_principal = self.root
            
            relatorio_window.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(relatorio_window))
            
            relatorio_window.lift()
            relatorio_window.focus_force()
            relatorio_window.mainloop()
            
        except Exception as e:
            messagebox.showerror("Erro",
                f"Erro ao abrir sistema de relatórios: {str(e)}")
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


    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            self.root.destroy()

    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        janela.destroy()
        self.root.deiconify()
        self.root.lift()

    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()


class OutputManager:
    def __init__(self, logger=None):
        self.dev_mode = os.getenv('DEV_MODE', 'False').lower() == 'true'
        self.logger = logger
        
        if not self.dev_mode:
            self.stdout_buffer = StringIO()
            self.stderr_buffer = StringIO()
            self.original_stdout = sys.stdout
            self.original_stderr = sys.stderr
            
            # Configurar logger para produção
            if self.logger:
                # Remover handlers existentes
                for handler in self.logger.logger.handlers[:]:
                    self.logger.logger.removeHandler(handler)
                
                # Adicionar FileHandler para logs
                log_dir = 'logs'
                os.makedirs(log_dir, exist_ok=True)
                log_file = os.path.join(
                    log_dir,
                    f'sistema_{datetime.now().strftime("%Y%m%d")}.log'
                )
                file_handler = logging.FileHandler(log_file, encoding='utf-8')
                file_handler.setFormatter(
                    logging.Formatter(self.logger.log_format)
                )
                self.logger.logger.addHandler(file_handler)
                
                # Em produção, só registrar logs de WARNING para cima
                self.logger.logger.setLevel(logging.WARNING)
    
    def start(self):
        """Inicia o redirecionamento da saída se não estiver em modo dev"""
        if not self.dev_mode:
            sys.stdout = self.stdout_buffer
            sys.stderr = self.stderr_buffer
    
    def stop(self):
        """Restaura a saída original"""
        if not self.dev_mode:
            sys.stdout = self.original_stdout
            sys.stderr = self.original_stderr
    
    def get_output(self):
        """Retorna o conteúdo dos buffers"""
        if not self.dev_mode:
            return {
                'stdout': self.stdout_buffer.getvalue(),
                'stderr': self.stderr_buffer.getvalue()
            }
        return None

# Modificar o sistema_principal.py para usar assim:
def main():
    from config.logger_config import system_logger
    
    output_manager = OutputManager(logger=system_logger)
    output_manager.start()
    
    try:
        app = SistemaPrincipal()
        app.run()
    finally:
        output_manager.stop()

if __name__ == '__main__':
    main()