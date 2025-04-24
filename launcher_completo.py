"""
Launcher completo para o Sistema de Gestão Financeira
"""
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import traceback
import logging
import shutil
from pathlib import Path
import importlib

# Configurar logging
def setup_logging():
    # Criar diretório de logs se não existir
    log_dir = Path('logs')
    log_dir.mkdir(exist_ok=True)
    
    # Configurar logging
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_dir / "sistema.log", encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger("sistema")

# Configurar caminhos
def setup_paths():
    # Obter diretório base (onde o executável está)
    if getattr(sys, 'frozen', False):
        # Executando como executável
        base_dir = os.path.dirname(sys.executable)
    else:
        # Executando como script
        base_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Adicionar diretórios ao path
    sys.path.insert(0, base_dir)
    
    # Se estiver executando como script, adicionar src
    if not getattr(sys, 'frozen', False):
        src_dir = os.path.join(base_dir, 'src')
        if os.path.exists(src_dir):
            sys.path.insert(0, src_dir)
    
    return base_dir

# Stub para o logger do sistema
class SystemLogger:
    def __init__(self):
        self.logger = logging.getLogger("sistema")
        self.log_format = "%(asctime)s - %(levelname)s - %(message)s"
        
    def get_logger(self):
        return self.logger
        
    def set_user(self, username):
        pass

# Decorator para ações de log
def log_action(action_name):
    def decorator(func):
        def wrapper(*args, **kwargs):
            return func(*args, **kwargs)
        return wrapper
    return decorator

# Iniciar sistema
def iniciar_sistema(logger):
    try:
        # Instalar módulos necessários
        try:
            import babel.numbers
        except ImportError:
            logger.warning("Babel não encontrado, tentando instalar...")
            os.system("pip install babel")
            
        try:
            import reportlab
        except ImportError:
            logger.warning("ReportLab não encontrado, tentando instalar...")
            os.system("pip install reportlab")
            
        try:
            import tkcalendar
        except ImportError:
            logger.warning("tkcalendar não encontrado, tentando instalar...")
            os.system("pip install tkcalendar")
        
        # Registrar os módulos necessários no contexto global
        global system_logger, version_control
        
        # Criar stub para system_logger
        system_logger = SystemLogger()
        
        # Importar version_control
        try:
            # Criar stub para version_control se necessário
            class VersionControlStub:
                @staticmethod
                def get_version_string():
                    return "1.2.0"
                @staticmethod
                def show_version_dialog(parent):
                    messagebox.showinfo("Versão", "Sistema de Gestão Financeira v1.2.0")
                @staticmethod
                def save_version_history():
                    return []
            
            # Tentar importar version_control real
            try:
                import version_control
            except ImportError:
                try:
                    from src import version_control
                except ImportError:
                    version_control = VersionControlStub()
        except Exception as e:
            logger.error(f"Erro ao importar version_control: {str(e)}")
            # Criar stub básico
            class VersionControlStub:
                @staticmethod
                def get_version_string():
                    return "1.2.0"
                @staticmethod
                def show_version_dialog(parent):
                    messagebox.showinfo("Versão", "Sistema de Gestão Financeira v1.2.0")
                @staticmethod
                def save_version_history():
                    return []
            
            version_control = VersionControlStub()
        
        # Agora carregue o sistema principal
        class SistemaPrincipal:
            def __init__(self):
                self.root = tk.Tk()
                self.root.title(f"Sistema de Gestão Financeira v{version_control.get_version_string()}")
                self.root.geometry("800x600")
                
                # Salvar histórico de versões
                try:
                    version_control.save_version_history()
                except Exception as e:
                    logger.error(f"Erro ao salvar histórico de versões: {str(e)}")
                
                # Configurar a interface
                self.setup_ui()
            
            def setup_ui(self):
                # Frame principal
                main_frame = ttk.Frame(self.root, padding=20)
                main_frame.pack(fill="both", expand=True)
                
                # Título
                ttk.Label(
                    main_frame,
                    text="Sistema de Gestão Financeira",
                    font=('Helvetica', 24, 'bold')
                ).pack(pady=(0, 30))
                
                # Mensagem de manutenção
                ttk.Label(
                    main_frame,
                    text="O sistema está em manutenção.\nAlguns módulos podem não funcionar corretamente.",
                    font=('Helvetica', 12)
                ).pack(pady=20)
                
                # Botão para atualizar
                ttk.Button(
                    main_frame,
                    text="Verificar Atualizações",
                    command=lambda: messagebox.showinfo("Atualizações", f"Versão atual: {version_control.get_version_string()}")
                ).pack(pady=10)
                
                # Botão de versão
                ttk.Button(
                    main_frame,
                    text="Sobre",
                    command=lambda: version_control.show_version_dialog(self.root)
                ).pack(pady=10)
                
                # Botão para sair
                ttk.Button(
                    main_frame,
                    text="Sair",
                    command=self.root.destroy
                ).pack(pady=10)
            
            def run(self):
                self.root.mainloop()
        
        # Iniciar a aplicação
        app = SistemaPrincipal()
        app.run()
    
    except Exception as e:
        error_msg = f"Erro ao iniciar o sistema: {str(e)}"
        logger.error(error_msg)
        logger.error(traceback.format_exc())
        
        # Mostrar mensagem de erro
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro de Inicialização", 
                f"Ocorreu um erro ao iniciar o sistema.\n\n{str(e)}\n\n"
                f"Consulte o arquivo de log para mais detalhes.")
        except:
            print(error_msg)

if __name__ == "__main__":
    # Configurar logging
    logger = setup_logging()
    logger.info("Iniciando aplicação")
    
    # Configurar paths
    base_dir = setup_paths()
    logger.info(f"Diretório base: {base_dir}")
    
    # Iniciar sistema
    iniciar_sistema(logger)