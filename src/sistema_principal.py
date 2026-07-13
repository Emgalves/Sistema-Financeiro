# sistema_principal.py - VERSÃO COM DIFERENCIAÇÃO DE AMBIENTES
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

# NOVO: Importar configuração de ambiente
try:
    from src.ambiente_config import (
        config_ambiente,
        criar_banner_ambiente,
        aplicar_estilo_ambiente,
        configurar_ttk_style_ambiente,
        get_cor_status
    )
except ImportError:
    try:
        from ambiente_config import (
            config_ambiente,
            criar_banner_ambiente,
            aplicar_estilo_ambiente,
            configurar_ttk_style_ambiente,
            get_cor_status
        )
    except ImportError:
        print("⚠️ Módulo ambiente_config não encontrado. Usando modo padrão.")
        # Criar fallback simples
        class ConfigAmbienteFallback:
            def eh_teste(self): return True
            def eh_producao(self): return False
            def get_nome_ambiente(self): return "TESTE"
            def get_titulo_janela(self, t): return f"[TESTE] {t}"
            def exibir_info_ambiente(self): print("Ambiente: TESTE")
        
        config_ambiente = ConfigAmbienteFallback()
        def criar_banner_ambiente(p): return None
        def aplicar_estilo_ambiente(w, t): pass
        def configurar_ttk_style_ambiente(): pass
        def get_cor_status(): return '#ff6b00'

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
        CLIENTES
    )
except ImportError:
    try:
        from config.config import (
            ARQUIVO_CLIENTES,
            ARQUIVO_MODELO,
            PASTA_CLIENTES,
            CLIENTES
        )
    except ImportError:
        # Valores padrão de fallback
        ARQUIVO_CLIENTES = "clientes.xlsx"
        ARQUIVO_MODELO = "modelo_cliente.xlsx"
        PASTA_CLIENTES = "dados_clientes"
        CLIENTES = []

try:
    from src.version_control import version_control
except ImportError:
    try:
        from version_control import version_control
    except ImportError:
        class SimpleVersionControl:
            def get_version_string(self): return "1.0.0"
            def save_version_history(self): pass
            def show_version_dialog(self, root): pass
        version_control = SimpleVersionControl()

def resource_path(relative_path):
    """Obtém o caminho absoluto do recurso"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

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

class SistemaGestaoFinanceira:
    """Sistema principal de gestão financeira"""
    
    def _configurar_paths_sistema(self):
        """Configura sys.path para imports funcionarem corretamente"""
        try:
            from pathlib import Path
            script_dir = Path(__file__).parent.absolute()
            if str(script_dir) not in sys.path:
                sys.path.insert(0, str(script_dir))
            
            modulos_problematicos = [
                'relatorio_despesas_aprimorado',
                'despesas_rateadas',
                'gestao_medicoes', 
                'configuracoes_sistema'
            ]
            
            for modulo in modulos_problematicos:
                if modulo in sys.modules:
                    del sys.modules[modulo]
                modulo_src = f"src.{modulo}"
                if modulo_src in sys.modules:
                    del sys.modules[modulo_src]
                    
        except Exception as e:
            print(f"Erro ao configurar paths: {str(e)}")

    def __init__(self):
        # Configurar paths antes de qualquer operação
        self._configurar_paths_sistema()
        
        self.usuario_atual = None
        self.root = tk.Tk()
        
        # Aplicar estilo de ambiente à janela
        aplicar_estilo_ambiente(self.root, 'janela')
        
        # Título com indicação de ambiente
        titulo_base = f"Sistema de Gestão Financeira v{version_control.get_version_string()}"
        titulo_com_ambiente = config_ambiente.get_titulo_janela(titulo_base)
        configurar_janela(self.root, titulo_com_ambiente)

        # Criar banner de ambiente (se necessário)
        self.banner_ambiente = criar_banner_ambiente(self.root)
        
        # Salvar histórico de versões
        try:
            version_control.save_version_history()
        except:
            pass
        
        # Inicializar gerenciador de taxas
        self.controlador_taxas = ControladorTaxas(self.root)
        
        # Configurar estilos TTK baseados no ambiente
        configurar_ttk_style_ambiente()
        
        # Configurar estilos e conteúdo
        self.setup_style()
        self.create_main_content()
        
        # Log com informação de ambiente
        simple_logger.info(f"Sistema iniciado em modo: {config_ambiente.get_nome_ambiente()}")
        
        # self.criar_interface()
        
        # ⭐ ADICIONAR ESTAS LINHAS AQUI:
        # Salvar geometria inicial para recuperação posterior
        self.root.update_idletasks()
        self.root._geometria_original = self.root.geometry()
        print(f"💾 Geometria inicial do menu salva: {self.root._geometria_original}")

    def login(self, username):
        self.usuario_atual = username
        system_logger.set_user(username)

    def setup_style(self):
        """Configura o estilo visual do aplicativo"""
        # A configuração já foi feita por configurar_ttk_style_ambiente()
        # Mas mantemos este método para compatibilidade
        pass

    def create_main_content(self):
        """Cria o conteúdo principal da interface"""
        # Frame principal
        main_frame = ttk.Frame(self.root)
        main_frame.pack(expand=True, fill="both", padx=20, pady=18)
        aplicar_estilo_ambiente(main_frame, 'frame')

        # Logo - com tratamento de erro
        try:
            self.logo_path = resource_path("logo3.png")
            if os.path.exists(self.logo_path):
                # logo3.png é a versão cortada (sem margem transparente
                # desperdiçada), proporção real ~2:1, por isso é redimensionada
                # com Pillow antes de ser exibida (PhotoImage puro do tkinter
                # não redimensiona).
                from PIL import Image, ImageTk
                imagem_logo = Image.open(self.logo_path).convert("RGBA")
                imagem_logo.thumbnail((340, 170), Image.LANCZOS)
                self.logo = ImageTk.PhotoImage(imagem_logo)
                logo_label = ttk.Label(main_frame, image=self.logo)
                logo_label.pack(pady=10)
            else:
                simple_logger.warning("Logo não encontrado, continuando sem imagem")
        except Exception as e:
            simple_logger.warning(f"Erro ao carregar logo: {str(e)}")

        # NOVO: Indicador visual de ambiente no título
        ambiente_emoji = "⚠️" if config_ambiente.eh_teste() else "🟢"
        ambiente_texto = config_ambiente.get_nome_ambiente()
        
        # Título com indicador de ambiente
        title_frame = tk.Frame(main_frame, bg=config_ambiente.get_config_visual()['cor_fundo'])
        title_frame.pack(pady=(0, 10))
        
        title_label = tk.Label(
            title_frame,
            text="Sistema de Gestão Financeira",
            font=('Helvetica', 24, 'bold'),
            bg=config_ambiente.get_config_visual()['cor_fundo'],
            fg=config_ambiente.get_config_visual()['cor_titulo']
        )
        title_label.pack()
        
        # Grid para cards
        grid = ttk.Frame(main_frame)
        grid.pack(expand=True, pady=16)
        aplicar_estilo_ambiente(grid, 'frame')

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
                        
        self.create_card(grid, "Gestão de Empreiteiros",
                        "Gerenciar contratos com empreiteros e por entregas",
                        self.abrir_gestao_medicoes, 2, 0)
                        
        self.create_card(grid, "Configurações do Sistema",
                        "Gerenciar parâmetros básicos",
                        self.abrir_configuracoes, 2, 1)
        
        # Frame para botões inferiores
        bottom_frame = ttk.Frame(main_frame)
        bottom_frame.pack(pady=16)
        aplicar_estilo_ambiente(bottom_frame, 'frame')
        
        # Versão e ambiente à esquerda
        info_frame = ttk.Frame(bottom_frame)
        info_frame.pack(side='left', padx=20)
        aplicar_estilo_ambiente(info_frame, 'frame')
        
        version_label = tk.Label(
            info_frame,
            text=f"Versão {version_control.get_version_string()} | {ambiente_emoji} {ambiente_texto}",
            font=('Helvetica', 9),
            fg=get_cor_status(),
            bg=config_ambiente.get_config_visual()['cor_fundo']
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
                                style='Action.TButton')
        sair_btn.pack(side='right', padx=5)

    def create_card(self, parent, title, description, command, row, col):
        """Cria um card na interface"""
        card = ttk.Frame(parent, style='Card.TFrame')
        card.grid(row=row, column=col, padx=10, pady=10, sticky='nsew')
        
        # NOVO: Aplicar cor de fundo do card baseada no ambiente
        try:
            card.configure(style='Card.TFrame')
        except:
            pass
        
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
            
            # Importar módulo
            try:
                from src.Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                from Sistema_Entrada_Dados import SistemaEntradaDados
            
            # ============================================================
            # SALVAR GEOMETRIA E OCULTAR MENU
            # ============================================================
            if not hasattr(self.root, '_geometria_original'):
                self.root.update_idletasks()
                self.root._geometria_original = self.root.geometry()
                simple_logger.info(f"💾 Geometria do menu salva: {self.root._geometria_original}")
            
            self.root.withdraw()
            self.root.update_idletasks()  # ✅ FORÇA PROCESSAMENTO
            
            # ============================================================
            # CRIAR SISTEMA DE ENTRADA DE DADOS
            # ============================================================
            app = SistemaEntradaDados(parent=self.root)
            
            # ============================================================
            # CALLBACK PARA RETORNAR AO MENU PRINCIPAL
            # ============================================================
            def retornar_ao_menu():
                """Fecha o sistema de entrada e volta ao menu"""
                try:
                    simple_logger.info("Fechando Sistema de Entrada de Dados")
                    
                    # Destruir janela do subsistema
                    if hasattr(app, 'root') and app.root.winfo_exists():
                        app.root.destroy()
                    
                    # ============================================================
                    # RESTAURAR MENU PRINCIPAL COM SEGURANÇA
                    # ============================================================
                    self.root.update_idletasks()  # ✅ FORÇA SINCRONIZAÇÃO
                    self.root.deiconify()
                    
                    # Restaurar geometria original
                    if hasattr(self.root, '_geometria_original'):
                        self.root.geometry(self.root._geometria_original)
                        simple_logger.info(f"✅ Geometria restaurada: {self.root._geometria_original}")
                    
                    self.root.update_idletasks()  # ✅ FORÇA ATUALIZAÇÃO
                    self.root.lift()
                    self.root.focus_force()
                    
                    simple_logger.info("✅ Menu principal restaurado")
                    
                except Exception as e:
                    simple_logger.error(f"❌ Erro ao retornar ao menu: {str(e)}")
                    # Em caso de erro, força exibição do menu
                    try:
                        self.root.deiconify()
                        self.root.lift()
                    except:
                        pass
            
            # ============================================================
            # CONFIGURAR PROTOCOLO DE FECHAMENTO
            # ============================================================
            app.root.protocol("WM_DELETE_WINDOW", retornar_ao_menu)
            
            # ============================================================
            # GARANTIR QUE A JANELA APAREÇA
            # ============================================================
            app.root.update_idletasks()
            app.root.lift()
            app.root.focus_force()
            
            simple_logger.info("✅ Sistema de Entrada de Dados aberto")
            
            # ❌ REMOVER ISTO: app.root.mainloop()
            # ✅ NÃO CRIAR NOVO LOOP! O loop principal já está rodando

        except Exception as e:
            simple_logger.error(f"❌ Erro ao abrir sistema de entrada de dados: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir sistema de entrada de dados:\n{str(e)}")
            
            # Restaurar menu em caso de erro
            try:
                self.root.update_idletasks()
                self.root.deiconify()
                if hasattr(self.root, '_geometria_original'):
                    self.root.geometry(self.root._geometria_original)
                self.root.update_idletasks()
                self.root.lift()
            except:
                pass

    def abrir_gestao_taxas(self):
        """Abre o sistema de gestão de taxas"""
        try:
            self.controlador_taxas.abrir_janela_controle()
        except Exception as e:
            simple_logger.error(f"Erro ao abrir gestão de taxas: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir gestão de taxas: {str(e)}")

    def abrir_despesas_rateadas(self):
        """Abre o sistema de despesas rateadas"""
        try:
            simple_logger.info("Abrindo sistema de despesas rateadas")
            
            try:
                from src.despesas_rateadas import InterfaceDespesasRateadas
            except ImportError:
                from despesas_rateadas import InterfaceDespesasRateadas
            
            janela_despesas = tk.Toplevel(self.root)
            app = InterfaceDespesasRateadas(janela_despesas)
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir despesas rateadas: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir despesas rateadas: {str(e)}")

    def abrir_relatorios(self):
        """Abre interface de relatórios"""
        try:
            simple_logger.info("Abrindo sistema de relatórios")
            
            try:
                from src.relatorios_interface import RelatoriosInterface
            except ImportError:
                from relatorios_interface import RelatoriosInterface
            
            # CORREÇÃO: Passar self.root como parent para manter referência
            app = RelatoriosInterface(parent=self.root)
            
            # Ocultar menu principal
            self.root.withdraw()
            
            # Quando fechar, voltar ao menu
            def ao_fechar():
                try:
                    app.root.destroy()
                except:
                    pass
                finally:
                    self.root.deiconify()
                    self.root.lift()
                    self.root.focus_force()
            
            if hasattr(app, 'root'):
                app.root.protocol("WM_DELETE_WINDOW", ao_fechar)
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir relatórios: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir relatórios: {str(e)}")
            self.root.deiconify()

    def abrir_gestao_medicoes(self):
        """Abre o sistema de gestão de medições"""
        try:
            simple_logger.info("Abrindo gestão de medições")
            
            try:
                from src.gestao_medicoes import GestaoMedicoes
            except ImportError:
                from gestao_medicoes import GestaoMedicoes
            
            # janela_medicoes = tk.Toplevel(self.root)
            app = GestaoMedicoes(parent=self.root)
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir gestão de medições: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir gestão de medições: {str(e)}")

    def abrir_configuracoes(self):
        """Abre as configurações do sistema"""
        try:
            simple_logger.info("Abrindo configurações do sistema")
            
            try:
                from src.configuracoes_sistema import GerenciadorConfiguracoes
            except ImportError:
                from configuracoes_sistema import GerenciadorConfiguracoes
            
            # janela_config = tk.Toplevel(self.root)
            app = GerenciadorConfiguracoes(parent=self.root)
            
        except Exception as e:
            simple_logger.error(f"Erro ao abrir configurações: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir configurações: {str(e)}")

    def sair_sistema(self):
        """Fecha o sistema"""
        resposta = messagebox.askyesno(
            "Confirmar Saída",
            "Deseja realmente sair do sistema?"
        )
        if resposta:
            simple_logger.info("Sistema encerrado pelo usuário")
            self.root.quit()
            self.root.destroy()

    def run(self):
        """Inicia o loop principal da aplicação"""
        try:
            self.root.mainloop()
        except KeyboardInterrupt:
            simple_logger.info("Sistema interrompido pelo usuário")
            self.root.quit()
        except Exception as e:
            simple_logger.error(f"Erro crítico no sistema: {str(e)}")
            messagebox.showerror(
                "Erro Crítico",
                f"Ocorreu um erro crítico:\n{str(e)}\n\nO sistema será encerrado."
            )
            sys.exit(1)


def main():
    """Função principal"""
    try:
        # Exibir informações de ambiente
        config_ambiente.exibir_info_ambiente()
        
        # Iniciar sistema
        app = SistemaGestaoFinanceira()
        app.run()
        
    except Exception as e:
        print(f"Erro crítico ao iniciar sistema: {str(e)}")
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()