import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import os
import sys
import importlib
from dotenv import load_dotenv
load_dotenv()
from pathlib import Path


# Adicionar o diretório pai ao sys.path
current_dir = Path(__file__).resolve().parent
parent_dir = current_dir.parent
sys.path.append(str(parent_dir))

try:
    from config.logger_config import system_logger, log_action
    print("Logger importado com sucesso")
except Exception as e:
    print(f"Erro ao importar logger: {str(e)}")
    print(f"Path atual: {sys.path}")

from config.config import (
    ARQUIVO_CLIENTES,
    ARQUIVO_MODELO,
    PASTA_CLIENTES,
    BASE_PATH
)

from gestao_taxas import GestaoTaxasAdministracao



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
        self.root.title("Sistema de Gestão Financeira")
        
        # Pega as dimensões da tela
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Define o tamanho da janela
        window_width = 900
        window_height = 900
        
        # Calcula a posição para garantir que a janela fique totalmente visível
        x = min(0, screen_width - window_width)
        y = min(0, screen_height - window_height)
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.lift()
        
        # Inicializar gerenciador de taxas
        self.gestao_taxas = GestaoTaxasAdministracao(self.root)
        
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

        # Título
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
        
        # Botão Sair
        sair_btn = ttk.Button(main_frame, 
                             text="Sair",
                             command=self.sair_sistema)
        sair_btn.pack(pady=20)

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
            print("Iniciando abertura do sistema de entrada de dados...")
            print(f"Diretório atual: {os.getcwd()}")
            print(f"ARQUIVO_CLIENTES: {ARQUIVO_CLIENTES}")
            print(f"ARQUIVO_CLIENTES existe? {os.path.exists(ARQUIVO_CLIENTES)}")


            # Mostrar o conteúdo do sys.path para diagnóstico
            print("Caminho de importação Python:")
            for path in sys.path:
                print(path)

            
            try:
                # Primeira tentativa: importar diretamente
                print("Tentando importar diretamente...")
                from Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError as e:
                print(f"Erro na importação direta: {e}")
                try:
                    # Segunda tentativa: importar de src
                    print("Tentando importar de src...")
                    from src.Sistema_Entrada_Dados import SistemaEntradaDados # type: ignore
                except ImportError as e:
                    print(f"Erro na importação de src: {e}")
                    try:
                        print("Tentando importação alternativa...")
                        import importlib.util
                        spec = importlib.util.spec_from_file_location(
                            "Sistema_Entrada_Dados", 
                            os.path.join(os.path.dirname(__file__), "Sistema_Entrada_Dados.py")
                        )
                        module = importlib.util.module_from_spec(spec)
                        spec.loader.exec_module(module)
                        SistemaEntradaDados = module.SistemaEntradaDados
                    except Exception as e:
                        raise ImportError(f"Não foi possível importar SistemaEntradaDados: {e}")

            print("Módulo importado com sucesso")
            self.root.withdraw()
            
            print("Iniciando criação da instância...")
            app = SistemaEntradaDados(parent=self.root)
            print("Instância criada com sucesso")
            
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()

        except Exception as e:
            print(f"Erro completo ao abrir sistema de entrada de dados:")
            print(f"Tipo do erro: {type(e)}")
            print(f"Erro: {str(e)}")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro",
                f"Erro ao abrir sistema de entrada de dados: {str(e)}")
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


if __name__ == '__main__':
    app = SistemaPrincipal()
    app.run()
