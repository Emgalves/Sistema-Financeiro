# sistema_principal.py - VERSÃO CORRIGIDA
import sys
import os
import traceback
import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import importlib
from datetime import datetime
from dotenv import load_dotenv

# Carregar variáveis de ambiente
load_dotenv()

def add_project_root():
    """Adiciona o diretório raiz do projeto ao sys.path"""
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Configuração de logging simplificada para PyInstaller
def setup_simple_logging():
    """Configura logging simples que funciona com PyInstaller"""
    import logging
    
    # Criar logger básico
    logger = logging.getLogger("sistema")
    logger.setLevel(logging.INFO)
    
    # Handler para arquivo (se possível)
    try:
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sistema.log")
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logger.addHandler(file_handler)
    except:
        pass  # Se não conseguir criar arquivo, continua sem
    
    # Handler para console (sempre funciona)
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(logging.Formatter('%(levelname)s - %(message)s'))
    logger.addHandler(console_handler)
    
    return logger

# Configurar logging simples
simple_logger = setup_simple_logging()

# Criar classes substitutas para o sistema de logging
class SimpleSystemLogger:
    def __init__(self):
        self.logger = simple_logger
        self.current_user = None
    
    def get_logger(self):
        return self.logger
    
    def set_user(self, username):
        self.current_user = username
        self.logger.info(f"Usuário logado: {username}")

def log_action(action_name):
    """Decorator simplificado para logging de ações"""
    def decorator(func):
        def wrapper(*args, **kwargs):
            simple_logger.info(f"Executando ação: {action_name}")
            try:
                result = func(*args, **kwargs)
                simple_logger.info(f"Ação concluída: {action_name}")
                return result
            except Exception as e:
                simple_logger.error(f"Erro na ação {action_name}: {str(e)}")
                raise
        return wrapper
    return decorator

# Instanciar o logger do sistema
system_logger = SimpleSystemLogger()

# Importações de configuração
try:
    from src.config.window_config import configurar_janela
except ImportError:
    try:
        from config.window_config import configurar_janela
    except ImportError:
        # Fallback: configuração básica de janela
        def configurar_janela(root, titulo):
            root.title(titulo)
            root.geometry("800x600")
            root.configure(bg='#f0f0f0')

try:
    from src.config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )
except ImportError:
    try:
        from config.config import (
            ARQUIVO_CLIENTES,
            ARQUIVO_MODELO,
            PASTA_CLIENTES,
            BASE_PATH
        )
    except ImportError:
        # Valores padrão se não conseguir importar
        ARQUIVO_CLIENTES = "clientes.xlsx"
        ARQUIVO_MODELO = "modelo.xlsx"
        PASTA_CLIENTES = "clientes"
        BASE_PATH = os.path.dirname(os.path.abspath(__file__))

def force_exit():
    """Força a saída do programa"""
    simple_logger.info("Forçando encerramento do programa...")
    import os
    os._exit(0)

# Importar módulo de controle de pagamentos
try:
    from src.controle_pagamentos_taxas import ControlePagamentos as ControladorTaxas
except ImportError:
    try:
        from controle_pagamentos_taxas import ControlePagamentos as ControladorTaxas
    except ImportError:
        simple_logger.warning("Módulo ControlePagamentos não encontrado, criando stub")
        
        class ControladorTaxasStub:
            def __init__(self, parent=None):
                self.parent = parent
                
            def abrir_janela_controle(self):
                messagebox.showerror("Erro", "Módulo de Controle de Pagamentos não encontrado")
                
        ControladorTaxas = ControladorTaxasStub

# Importar módulo de controle de versões
try:
    import version_control
except ImportError:
    try:
        from src import version_control
    except ImportError:
        simple_logger.warning("Módulo version_control não encontrado, criando stub")
        
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
    """Obtém o caminho correto para recursos, funciona com PyInstaller"""
    try:
        # PyInstaller cria uma pasta temporária e armazena o caminho em _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class SistemaPrincipal:

    def _configurar_paths_sistema(self):
        """Configura os paths do sistema para garantir que todos os módulos sejam encontrados"""
        import sys
        from pathlib import Path
        
        try:
            # Obter diretório atual e raiz do projeto
            current_dir = Path(__file__).resolve().parent
            project_root = current_dir.parent
            
            # Lista de diretórios para adicionar ao path
            paths_adicionar = [
                str(current_dir),      # src/
                str(project_root),     # raiz do projeto
            ]
            
            # Adicionar paths se não estiverem já incluídos
            for path in paths_adicionar:
                if path not in sys.path:
                    sys.path.insert(0, path)
                    print(f"Path adicionado: {path}")
            
            # Limpar cache de módulos problemáticos para forçar reload
            modulos_problematicos = [
                'relatorios_interface',
                'relatorio_despesas_aprimorado',
                'despesas_rateadas',
                'gestao_medicoes', 
                'configuracoes_sistema'
            ]
            
            for modulo in modulos_problematicos:
                # Remover versão direta
                if modulo in sys.modules:
                    del sys.modules[modulo]
                    print(f"Cache limpo: {modulo}")
                
                # Remover versão com src
                modulo_src = f"src.{modulo}"
                if modulo_src in sys.modules:
                    del sys.modules[modulo_src] 
                    print(f"Cache limpo: {modulo_src}")
                    
        except Exception as e:
            print(f"Erro ao configurar paths: {str(e)}")

    def __init__(self):
        # FIX: Configurar paths antes de qualquer operação
        self._configurar_paths_sistema()
        
        self.usuario_atual = None
        self.root = tk.Tk()
        
        # Configurar a janela principal
        titulo_com_versao = f"Sistema de Gestão Financeira v{version_control.get_version_string()}"
        configurar_janela(self.root, titulo_com_versao)

        # Salvar histórico de versões
        try:
            version_control.save_version_history()
        except:
            pass  # Ignorar se não conseguir salvar
        
        # Inicializar gerenciador de taxas
        self.controlador_taxas = ControladorTaxas(self.root)
        
        # Configurar estilos e conteúdo
        self.setup_style()
        self.create_main_content()
        
        simple_logger.info("Sistema principal inicializado com sucesso")
        
    def login(self, username):
        self.usuario_atual = username
        system_logger.set_user(username)

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

        # Logo - com tratamento de erro
        try:
            self.logo_path = resource_path("logo.png")
            if os.path.exists(self.logo_path):
                self.logo = PhotoImage(file=self.logo_path)
                logo_label = ttk.Label(main_frame, image=self.logo)
                logo_label.pack(pady=10)
            else:
                simple_logger.warning("Logo não encontrado, continuando sem imagem")
        except Exception as e:
            simple_logger.warning(f"Erro ao carregar logo: {str(e)}")

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
        
        # Frame para botões inferiores
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(pady=20)
        
        # Versão e botão Sobre à esquerda
        version_frame = ttk.Frame(bottom_frame)
        version_frame.pack(side='left', padx=20)
        
        version_label = ttk.Label(
            version_frame,
            text=f"Versão {version_control.get_version_string()}",
            font=('Helvetica', 9),
            foreground='#555555'
        )
        version_label.pack(pady=5)
        
        about_button = ttk.Button(
            bottom_frame,
            text="Sobre",
            command=lambda: version_control.show_version_dialog(self.root)
        )
        about_button.pack(side='left', padx=10)
        
        # Botão Sair
        sair_btn = ttk.Button(bottom_frame, text="Sair", 
                                command=self.sair_sistema,
                                style='Medium.TButton')
        sair_btn.pack(side='right', padx=5)

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
            simple_logger.info("Abrindo sistema de entrada de dados")
            
            try:
                from Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                from src.Sistema_Entrada_Dados import SistemaEntradaDados
            
            self.root.withdraw()
            app = SistemaEntradaDados(parent=self.root)
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()

        except Exception as e:
            simple_logger.error(f"Erro ao abrir sistema de entrada de dados: {str(e)}")
            messagebox.showerror("Erro", "Erro ao abrir sistema de entrada de dados.")
            self.root.deiconify()

    def abrir_gestao_taxas(self):
        """Abre o menu de gestão de taxas"""
        try:
            self.controlador_taxas.abrir_janela_controle()
        except Exception as e:
            simple_logger.error(f"Erro ao abrir gestão de taxas: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir gestão de taxas: {str(e)}")

    @log_action("Gerar relatório")
    def abrir_relatorios(self):
        """Abre o sistema integrado de relatórios"""
        try:
            modulo = self.reload_module('src.relatorios_interface')
            if not modulo:
                modulo = self.reload_module('src.relatorio_despesas_aprimorado')
                if not modulo:
                    messagebox.showerror("Erro", "Não foi possível carregar o módulo de relatórios.")
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
                return
                
            self.root.withdraw()
            app = modulo.SistemaRelatorios(parent=self.root)
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            app.root.lift()
            app.root.focus_force()
            app.run()
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir sistema de relatórios: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir sistema de relatórios: {str(e)}")
            self.root.deiconify()

    def abrir_despesas_rateadas(self):
        """Abre o sistema de despesas rateadas"""
        try:
            modulo = self.reload_module('src.despesas_rateadas')
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
            simple_logger.error(f"Erro ao abrir sistema de despesas rateadas: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir sistema de despesas rateadas: {str(e)}")
            self.root.deiconify()

    def abrir_gestao_medicoes(self):
        """Abre o sistema de gestão de medições"""
        try:
            modulo = self.reload_module('src.gestao_medicoes')
            if not modulo:
                return

            self.root.withdraw()
            app = modulo.GestaoMedicoes(parent=self.root)
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir sistema de gestão de medições: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir sistema de gestão de medições: {str(e)}")
            self.root.deiconify()

    def reload_module(self, module_name):
        """Recarrega um módulo e retorna a versão atualizada"""
        try:
            # Remover todas as referências ao módulo
            for key in list(sys.modules.keys()):
                if key == module_name or key.startswith(f"{module_name}."):
                    del sys.modules[key]
            
            module = importlib.import_module(module_name)
            return module
        except Exception as e:
            simple_logger.error(f"Erro ao carregar módulo {module_name}: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao carregar módulo {module_name}: {str(e)}")
            return None

    def abrir_configuracoes(self):
        try:
            from src.configuracoes_sistema import GerenciadorConfiguracoes
            self.root.withdraw()
            app = GerenciadorConfiguracoes(parent=self.root)
            app.menu_principal = self.root
            app.root.protocol("WM_DELETE_WINDOW", lambda: self.finalizar_sistema(app.root))
            app.run()
        except Exception as e:
            simple_logger.error(f"Erro ao abrir configurações: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir configurações do sistema: {str(e)}")
            self.root.deiconify()

    def adicionar_correcao_monetaria_ao_menu():
        """
        Função para adicionar a opção de correção monetária ao menu principal
        Adicione esta chamada ao seu menu principal
        """
        def abrir_indices_correcao():
            app = InterfaceIndicesCorrecao()
            app.root.mainloop()
        
        return abrir_indices_correcao

    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        try:
            janela.destroy()
        except Exception as e:
            simple_logger.error(f"Erro ao destruir janela: {str(e)}")
        
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def on_closing(self):
        """Manipula o fechamento da janela principal"""
        try:
            if messagebox.askyesno("Sair", "Deseja realmente sair do sistema?"):
                # Finalizar todos os subsistemas abertos
                for widget in self.root.winfo_children():
                    if isinstance(widget, tk.Toplevel):
                        try:
                            widget.destroy()
                        except:
                            pass
                
                # Salvar configurações finais
                try:
                    from src.configuracoes_sistema import GerenciadorConfiguracoes
                    # Forçar salvamento das configurações
                    config = GerenciadorConfiguracoes.carregar_configuracoes()
                    if config:
                        import json
                        with open(GerenciadorConfiguracoes.CONFIG_PATH, 'w', encoding='utf-8') as f:
                            json.dump(config, f, indent=4, ensure_ascii=False)
                except:
                    pass
                
                self.root.quit()
                self.root.destroy()
        except Exception as e:
            print(f"Erro ao fechar: {str(e)}")
            self.root.quit()

    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            simple_logger.info("Saída confirmada, finalizando sistema")
            self.root.destroy()
            self.root.after(200, force_exit)

    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()

def main():
    """Função principal"""
    try:
        simple_logger.info("=== Iniciando Sistema de Gestão Financeira ===")
        app = SistemaPrincipal()
        app.run()
    except Exception as e:
        simple_logger.error(f"Erro crítico no sistema principal: {str(e)}")
        print(f"Erro no sistema principal: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()