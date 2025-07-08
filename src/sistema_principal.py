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
                from src.Sistema_Entrada_Dados import SistemaEntradaDados
            except ImportError:
                from Sistema_Entrada_Dados import SistemaEntradaDados
            
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
        """Abre o sistema integrado de relatórios - VERSÃO COM DIAGNÓSTICO"""
        try:
            simple_logger.info("=== INICIANDO SISTEMA DE RELATÓRIOS COM DIAGNÓSTICO ===")
            
            # === DIAGNÓSTICO INICIAL ===
            self.diagnosticar_sistema_relatorios()
            
            # Primeiro, verificar se o sistema de logging está funcionando
            try:
                from src.config.logger_config import system_logger
                system_logger.set_user('sistema_principal')
                logger = system_logger.get_logger()
                logger.info("Abrindo sistema de relatórios")
            except:
                pass  # Se falhar, continuar com simple_logger
            
            # === ESTRATÉGIA 1: BUSCA INTELIGENTE ===
            try:
                simple_logger.info("🔍 INICIANDO BUSCA INTELIGENTE POR MÓDULOS")
                
                # Obter diretório base do sistema
                import os
                from pathlib import Path
                
                if getattr(sys, 'frozen', False):
                    # Se for executável PyInstaller
                    base_dir = Path(sys._MEIPASS)
                    simple_logger.info(f"📦 Executável PyInstaller detectado: {base_dir}")
                else:
                    # Se for execução normal
                    base_dir = Path(__file__).resolve().parent
                    simple_logger.info(f"🐍 Execução Python normal: {base_dir}")
                
                # Buscar arquivo relatorios_interface.py
                arquivos_encontrados = []
                for caminho_busca in [base_dir, base_dir.parent, base_dir / "src"]:
                    arquivo_relatorios = caminho_busca / "relatorios_interface.py"
                    if arquivo_relatorios.exists():
                        arquivos_encontrados.append(str(arquivo_relatorios))
                        simple_logger.info(f"✅ ENCONTRADO: {arquivo_relatorios}")
                    else:
                        simple_logger.info(f"❌ NÃO ENCONTRADO: {arquivo_relatorios}")
                
                if not arquivos_encontrados:
                    simple_logger.error("❌ ARQUIVO relatorios_interface.py NÃO ENCONTRADO EM LUGAR NENHUM!")
                    raise Exception("Arquivo relatorios_interface.py não encontrado no sistema")
                
                # === ESTRATÉGIA 2: ADICIONAR PATHS E TENTAR IMPORTAR ===
                simple_logger.info("🔄 Configurando paths do sistema...")
                
                # Adicionar todos os paths relevantes
                paths_para_adicionar = [
                    str(base_dir),
                    str(base_dir / "src"),
                    str(base_dir.parent),
                    str(base_dir.parent / "src")
                ]
                
                for path in paths_para_adicionar:
                    if os.path.exists(path) and path not in sys.path:
                        sys.path.insert(0, path)
                        simple_logger.info(f"➕ Path adicionado: {path}")
                
                # Limpar cache de módulos relacionados
                modulos_relacionados = [
                    'relatorios_interface',
                    'src.relatorios_interface', 
                    'relatorio_despesas_service',
                    'src.relatorio_despesas_service',
                    'relatorio_despesas_aprimorado',
                    'src.relatorio_despesas_aprimorado'
                ]
                
                for module_name in modulos_relacionados:
                    if module_name in sys.modules:
                        del sys.modules[module_name]
                        simple_logger.info(f"🗑️ Cache limpo: {module_name}")
                
                # Invalidar caches do Python
                importlib.invalidate_caches()
                simple_logger.info("🔄 Cache do importlib invalidado")
                
                # === ESTRATÉGIA 3: TENTATIVA DE IMPORTAÇÃO ===
                modulo_sistema = None
                caminhos_tentativa = [
                    'relatorios_interface',           # Primeiro sem src
                    'src.relatorios_interface'        # Depois com src
                ]
                
                for caminho in caminhos_tentativa:
                    try:
                        simple_logger.info(f"🎯 Tentando importar: {caminho}")
                        modulo_sistema = importlib.import_module(caminho)
                        simple_logger.info(f"✅ SUCESSO! Módulo {caminho} importado!")
                        
                        # Verificar se tem a classe necessária
                        if hasattr(modulo_sistema, 'SistemaRelatorios'):
                            simple_logger.info(f"✅ Classe SistemaRelatorios encontrada em {caminho}")
                            break
                        else:
                            simple_logger.warning(f"⚠️ Módulo {caminho} importado mas sem classe SistemaRelatorios")
                            # Listar classes disponíveis
                            classes = [name for name in dir(modulo_sistema) if not name.startswith('_') and name[0].isupper()]
                            simple_logger.info(f"Classes disponíveis: {classes}")
                            modulo_sistema = None  # Resetar para continuar tentando
                            
                    except ImportError as e:
                        simple_logger.warning(f"❌ Falha ao importar {caminho}: {str(e)}")
                        continue
                    except Exception as e:
                        simple_logger.error(f"💥 Erro inesperado ao importar {caminho}: {str(e)}")
                        continue
                
                # === ESTRATÉGIA 4: SE ENCONTROU O MÓDULO, INICIALIZAR ===
                if modulo_sistema and hasattr(modulo_sistema, 'SistemaRelatorios'):
                    simple_logger.info("🚀 INICIALIZANDO SISTEMA INTEGRADO DE RELATÓRIOS")
                    
                    # Ocultar janela principal
                    self.root.withdraw()
                    
                    try:
                        # Criar instância do sistema integrado
                        app = modulo_sistema.SistemaRelatorios(parent=self.root)
                        
                        # Configurar referência ao menu principal
                        app.menu_principal = self.root
                        
                        # Configurar comportamento de fechamento
                        def fechar_sistema():
                            try:
                                app.root.destroy()
                            except:
                                pass
                            self.finalizar_sistema_relatorios()
                        
                        app.root.protocol("WM_DELETE_WINDOW", fechar_sistema)
                        
                        # Garantir visibilidade
                        app.root.lift()
                        app.root.focus_force()
                        
                        # Executar sistema integrado
                        simple_logger.info("🎉 SISTEMA INTEGRADO INICIADO COM SUCESSO!")
                        
                        # Usar run() se disponível, senão mainloop()
                        if hasattr(app, 'run'):
                            app.run()
                        else:
                            app.root.mainloop()
                        
                        return  # SUCESSO - sair aqui
                        
                    except Exception as e:
                        simple_logger.error(f"💥 Erro ao inicializar SistemaRelatorios: {str(e)}")
                        import traceback
                        traceback.print_exc()
                        # Restaurar janela principal em caso de erro
                        self.root.deiconify()
                        raise e
                
                else:
                    # Se não encontrou o sistema integrado
                    erro_detalhado = "MÓDULO NÃO ENCONTRADO OU SEM CLASSE CORRETA:\n"
                    erro_detalhado += f"- Arquivos encontrados: {len(arquivos_encontrados)}\n"
                    erro_detalhado += f"- Paths no sys.path: {len([p for p in sys.path if 'src' in p])}\n"
                    erro_detalhado += f"- Módulo carregado: {modulo_sistema is not None}\n"
                    
                    if modulo_sistema:
                        classes = [name for name in dir(modulo_sistema) if not name.startswith('_')]
                        erro_detalhado += f"- Classes disponíveis: {classes[:10]}..."  # Primeiras 10
                    
                    simple_logger.error(erro_detalhado)
                    raise Exception(f"Sistema integrado não pôde ser inicializado:\n{erro_detalhado}")
                    
            except Exception as e:
                simple_logger.error(f"💥 ERRO na busca inteligente: {str(e)}")
                
                # === OFERECER OPÇÕES AO USUÁRIO ===
                resposta = messagebox.askyesnocancel(
                    "Sistema de Relatórios - Diagnóstico", 
                    f"❌ O sistema integrado de relatórios não pôde ser carregado.\n\n"
                    f"🔍 DIAGNÓSTICO DETALHADO:\n"
                    f"• Arquivos encontrados: {len(getattr(self, '_arquivos_diagnostico', []))}\n"
                    f"• Erro principal: {str(e)[:100]}...\n\n"
                    f"🔄 OPÇÕES:\n"
                    f"• SIM: Ver diagnóstico completo e tentar correção\n"
                    f"• NÃO: Usar sistema básico de despesas\n"
                    f"• CANCELAR: Voltar ao menu principal"
                )
                
                if resposta is True:  # SIM - Diagnóstico completo
                    self.mostrar_diagnostico_completo(e)
                    
                elif resposta is False:  # NÃO - Sistema básico
                    try:
                        simple_logger.info("⚠️ Carregando sistema básico de despesas...")
                        self.abrir_sistema_basico_despesas()
                        return
                    except Exception as fallback_error:
                        simple_logger.error(f"Erro no sistema básico: {fallback_error}")
                        messagebox.showerror("Erro", f"Erro no sistema básico: {fallback_error}")
                
                # Se chegou aqui (CANCELAR ou erro), restaurar interface
                self.root.deiconify()
                
        except Exception as e:
            simple_logger.error(f"💥 ERRO CRÍTICO no sistema de relatórios: {str(e)}")
            import traceback
            traceback.print_exc()
            
            messagebox.showerror(
                "Erro Crítico", 
                f"Erro crítico no sistema de relatórios.\n\n"
                f"Erro: {str(e)}\n\n"
                f"O sistema será retornado ao menu principal."
            )
            self.root.deiconify()

    def diagnosticar_sistema_relatorios(self):
        """Realiza diagnóstico detalhado do sistema de relatórios"""
        try:
            simple_logger.info("🔍 INICIANDO DIAGNÓSTICO DO SISTEMA")
            
            import os
            from pathlib import Path
            
            # Obter informações do ambiente
            if getattr(sys, 'frozen', False):
                base_dir = Path(sys._MEIPASS)
                tipo_exec = "PyInstaller"
            else:
                base_dir = Path(__file__).resolve().parent
                tipo_exec = "Python Normal"
            
            simple_logger.info(f"📍 Tipo de execução: {tipo_exec}")
            simple_logger.info(f"📁 Diretório base: {base_dir}")
            
            # Buscar arquivos relevantes
            arquivos_encontrados = []
            diretorios_busca = [
                base_dir,
                base_dir.parent, 
                base_dir / "src",
                base_dir.parent / "src"
            ]
            
            for diretorio in diretorios_busca:
                if diretorio.exists():
                    simple_logger.info(f"📂 Verificando: {diretorio}")
                    
                    # Buscar arquivos Python relevantes
                    for arquivo in ['relatorios_interface.py', 'relatorio_despesas_service.py', 'relatorio_despesas_aprimorado.py']:
                        caminho_arquivo = diretorio / arquivo
                        if caminho_arquivo.exists():
                            arquivos_encontrados.append(str(caminho_arquivo))
                            simple_logger.info(f"  ✅ {arquivo}")
                        else:
                            simple_logger.info(f"  ❌ {arquivo}")
                else:
                    simple_logger.info(f"📂 NÃO EXISTE: {diretorio}")
            
            # Salvar para uso posterior
            self._arquivos_diagnostico = arquivos_encontrados
            
            simple_logger.info(f"📊 RESUMO: {len(arquivos_encontrados)} arquivos relevantes encontrados")
            
        except Exception as e:
            simple_logger.error(f"Erro no diagnóstico: {str(e)}")

    def mostrar_diagnostico_completo(self, erro_original):
        """Mostra janela com diagnóstico completo"""
        try:
            import tkinter as tk
            from tkinter import ttk, scrolledtext
            
            # Criar janela de diagnóstico
            diag_window = tk.Toplevel(self.root)
            diag_window.title("Diagnóstico Completo - Sistema de Relatórios")
            diag_window.geometry("800x600")
            diag_window.transient(self.root)
            
            # Frame principal
            main_frame = ttk.Frame(diag_window, padding=10)
            main_frame.pack(fill='both', expand=True)
            
            # Título
            ttk.Label(main_frame, text="🔍 Diagnóstico Completo do Sistema", 
                     font=('Arial', 14, 'bold')).pack(pady=(0, 10))
            
            # Área de texto com scroll para o diagnóstico
            text_area = scrolledtext.ScrolledText(main_frame, wrap=tk.WORD, height=25)
            text_area.pack(fill='both', expand=True, pady=(0, 10))
            
            # Gerar relatório de diagnóstico
            diagnostico = self.gerar_relatorio_diagnostico(erro_original)
            text_area.insert('1.0', diagnostico)
            text_area.config(state='disabled')
            
            # Botões
            btn_frame = ttk.Frame(main_frame)
            btn_frame.pack(fill='x')
            
            ttk.Button(btn_frame, text="Tentar Correção Automática", 
                      command=lambda: self.tentar_correcao_automatica(diag_window)).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Copiar Diagnóstico", 
                      command=lambda: self.copiar_diagnostico(diagnostico)).pack(side='left', padx=5)
            ttk.Button(btn_frame, text="Fechar", 
                      command=diag_window.destroy).pack(side='right', padx=5)
            
        except Exception as e:
            simple_logger.error(f"Erro ao mostrar diagnóstico: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao mostrar diagnóstico: {str(e)}")

    def gerar_relatorio_diagnostico(self, erro_original):
        """Gera relatório detalhado de diagnóstico"""
        try:
            import platform
            from pathlib import Path
            
            relatorio = []
            relatorio.append("=" * 80)
            relatorio.append("DIAGNÓSTICO COMPLETO - SISTEMA DE RELATÓRIOS")
            relatorio.append("=" * 80)
            relatorio.append("")
            
            # Informações do sistema
            relatorio.append("🖥️ INFORMAÇÕES DO SISTEMA:")
            relatorio.append(f"   Sistema Operacional: {platform.system()} {platform.release()}")
            relatorio.append(f"   Python: {platform.python_version()}")
            relatorio.append(f"   Arquitetura: {platform.architecture()[0]}")
            relatorio.append(f"   Executável PyInstaller: {'Sim' if getattr(sys, 'frozen', False) else 'Não'}")
            relatorio.append("")
            
            # Diretórios e paths
            relatorio.append("📁 DIRETÓRIOS E PATHS:")
            if getattr(sys, 'frozen', False):
                relatorio.append(f"   Diretório base (PyInstaller): {sys._MEIPASS}")
            relatorio.append(f"   Diretório do script: {Path(__file__).resolve().parent}")
            relatorio.append(f"   Diretório de trabalho: {Path.cwd()}")
            relatorio.append("")
            
            relatorio.append("   Paths no sys.path:")
            for i, path in enumerate(sys.path[:10]):  # Primeiros 10
                relatorio.append(f"   [{i+1:2d}] {path}")
            if len(sys.path) > 10:
                relatorio.append(f"   ... e mais {len(sys.path) - 10} paths")
            relatorio.append("")
            
            # Arquivos encontrados
            relatorio.append("📄 ARQUIVOS RELEVANTES ENCONTRADOS:")
            arquivos = getattr(self, '_arquivos_diagnostico', [])
            if arquivos:
                for arquivo in arquivos:
                    relatorio.append(f"   ✅ {arquivo}")
            else:
                relatorio.append("   ❌ Nenhum arquivo relevante encontrado!")
            relatorio.append("")
            
            # Módulos no cache
            relatorio.append("🗃️ MÓDULOS RELACIONADOS NO CACHE:")
            modulos_relacionados = [name for name in sys.modules.keys() 
                                   if any(termo in name.lower() for termo in ['relatorio', 'interface', 'despesa'])]
            if modulos_relacionados:
                for modulo in sorted(modulos_relacionados):
                    relatorio.append(f"   📦 {modulo}")
            else:
                relatorio.append("   📦 Nenhum módulo relacionado no cache")
            relatorio.append("")
            
            # Erro original
            relatorio.append("💥 ERRO ORIGINAL:")
            relatorio.append(f"   {str(erro_original)}")
            relatorio.append("")
            
            # Recomendações
            relatorio.append("🔧 RECOMENDAÇÕES:")
            relatorio.append("   1. Verificar se o arquivo relatorios_interface.py existe no projeto")
            relatorio.append("   2. Verificar se a estrutura de diretórios está correta")
            relatorio.append("   3. Se for executável, recompilar com todos os arquivos")
            relatorio.append("   4. Tentar a correção automática abaixo")
            relatorio.append("")
            
            relatorio.append("=" * 80)
            
            return "\n".join(relatorio)
            
        except Exception as e:
            return f"Erro ao gerar diagnóstico: {str(e)}"

    def tentar_correcao_automatica(self, diag_window):
        """Tenta correção automática do problema"""
        try:
            simple_logger.info("🔧 Tentando correção automática...")
            
            # Fechar janela de diagnóstico
            diag_window.destroy()
            
            # Tentar diferentes estratégias de correção
            
            # Estratégia 1: Busca mais ampla por arquivos
            import os
            from pathlib import Path
            
            # Buscar em todo o diretório do executável
            if getattr(sys, 'frozen', False):
                base_search = Path(sys._MEIPASS)
            else:
                base_search = Path(__file__).resolve().parent.parent
            
            # Busca recursiva
            arquivos_encontrados = list(base_search.rglob("relatorios_interface.py"))
            
            if arquivos_encontrados:
                arquivo_encontrado = arquivos_encontrados[0]
                diretorio_arquivo = arquivo_encontrado.parent
                
                # Adicionar diretório ao path
                if str(diretorio_arquivo) not in sys.path:
                    sys.path.insert(0, str(diretorio_arquivo))
                
                # Tentar importar novamente
                try:
                    if 'relatorios_interface' in sys.modules:
                        del sys.modules['relatorios_interface']
                    
                    modulo = importlib.import_module('relatorios_interface')
                    
                    if hasattr(modulo, 'SistemaRelatorios'):
                        messagebox.showinfo("Sucesso", f"Correção automática bem-sucedida!\nArquivo encontrado em: {arquivo_encontrado}")
                        
                        # Tentar abrir o sistema
                        self.root.withdraw()
                        app = modulo.SistemaRelatorios(parent=self.root)
                        app.menu_principal = self.root
                        app.root.protocol("WM_DELETE_WINDOW", 
                            lambda: self.finalizar_sistema(app.root))
                        app.root.lift()
                        app.root.focus_force()
                        if hasattr(app, 'run'):
                            app.run()
                        else:
                            app.root.mainloop()
                        return
                    
                except Exception as e:
                    simple_logger.error(f"Erro na correção automática: {str(e)}")
            
            # Se chegou aqui, correção falhou
            messagebox.showerror(
                "Correção Falhou", 
                "A correção automática não foi bem-sucedida.\n\n"
                "Recomendações:\n"
                "1. Verificar se todos os arquivos estão presentes\n"
                "2. Recompilar o executável se necessário\n"
                "3. Usar o sistema básico de despesas temporariamente"
            )
            
            self.root.deiconify()
            
        except Exception as e:
            simple_logger.error(f"Erro na correção automática: {str(e)}")
            messagebox.showerror("Erro", f"Erro na correção automática: {str(e)}")
            self.root.deiconify()

    def copiar_diagnostico(self, diagnostico):
        """Copia o diagnóstico para a área de transferência"""
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(diagnostico)
            messagebox.showinfo("Copiado", "Diagnóstico copiado para a área de transferência!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao copiar: {str(e)}")

    # Manter os métodos auxiliares existentes...
    def abrir_sistema_basico_despesas(self):
        """Sistema básico de despesas como último recurso"""
        try:
            simple_logger.info("🔧 Carregando sistema básico de despesas...")
            
            # Tentar carregar RelatorioUI do módulo de despesas
            modulos_despesas = [
                'relatorio_despesas_aprimorado',
                'src.relatorio_despesas_aprimorado'
            ]
            
            for modulo_nome in modulos_despesas:
                try:
                    # Limpar cache
                    if modulo_nome in sys.modules:
                        del sys.modules[modulo_nome]
                    
                    modulo_despesas = importlib.import_module(modulo_nome)
                    
                    if hasattr(modulo_despesas, 'RelatorioUI'):
                        self.root.withdraw()
                        despesas_window = tk.Toplevel(self.root)
                        app = modulo_despesas.RelatorioUI(despesas_window)
                        app.menu_principal = self.root
                        despesas_window.protocol("WM_DELETE_WINDOW", 
                            lambda: self.finalizar_sistema(despesas_window))
                        despesas_window.lift()
                        despesas_window.focus_force()
                        despesas_window.mainloop()
                        return
                        
                except Exception as e:
                    simple_logger.warning(f"Falha ao carregar {modulo_nome}: {str(e)}")
                    continue
            
            # Se chegou aqui, nem o sistema básico funcionou
            raise Exception("Nem o sistema básico de despesas pôde ser carregado")
            
        except Exception as e:
            simple_logger.error(f"Erro no sistema básico: {str(e)}")
            messagebox.showerror("Erro", f"Erro no sistema básico: {str(e)}")
            self.root.deiconify()

    def finalizar_sistema_relatorios(self):
        """Método específico para finalizar sistema de relatórios"""
        try:
            simple_logger.info("🔄 Finalizando sistema de relatórios e retornando ao menu")
            self.root.deiconify()
            self.root.lift()
            self.root.focus_force()
        except Exception as e:
            simple_logger.error(f"Erro ao finalizar sistema de relatórios: {str(e)}")
            # Forçar exibição do menu principal
            try:
                self.root.deiconify()
            except:
                pass

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