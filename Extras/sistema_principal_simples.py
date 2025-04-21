import tkinter as tk
from tkinter import ttk, PhotoImage, messagebox
import importlib
import sys
import os
from datetime import datetime
import traceback

# Adicionar diretório raiz ao path
def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Importações básicas
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

# Função para forçar a saída
def force_exit():
    """Força a saída do programa"""
    print("Forçando encerramento do programa...")
    try:
        import os
        os._exit(0)
    except:
        sys.exit(0)

# Função para carregar/recarregar módulos
def reload_module(module_name):
    """Recarrega um módulo e retorna a versão atualizada"""
    try:
        # Remover todas as referências ao módulo e seus submódulos
        for key in list(sys.modules.keys()):
            if key == module_name or key.startswith(f"{module_name}."):
                del sys.modules[key]
        
        # Importar o módulo novamente
        module = importlib.import_module(module_name)
        return module
    except Exception as e:
        print(f"Erro ao carregar módulo {module_name}: {str(e)}")
        traceback.print_exc()
        return None

# Diretório de recursos
def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class SistemaPrincipalSimples:
    def __init__(self):
        print("Iniciando sistema principal simples...")
        
        # Criar janela principal
        self.root = tk.Tk()
        self.root.title("Sistema de Gestão Financeira")
        
        # Configurar tamanho e posição
        self.root.geometry("900x600")
        self.root.resizable(True, True)
        
        # Centralizar na tela
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        # Criar frame principal
        self.main_frame = ttk.Frame(self.root, padding=10)
        self.main_frame.pack(fill='both', expand=True)
        
        # Título
        title_label = ttk.Label(
            self.main_frame, 
            text="Sistema de Gestão Financeira", 
            font=('Helvetica', 20, 'bold')
        )
        title_label.pack(pady=20)
        
        # Botões para testes
        btn_relatorios = ttk.Button(
            self.main_frame,
            text="Abrir Relatórios",
            command=self.abrir_relatorios
        )
        btn_relatorios.pack(pady=10)
        
        btn_fornecedores = ttk.Button(
            self.main_frame,
            text="Relatório de Fornecedores",
            command=self.abrir_fornecedores
        )
        btn_fornecedores.pack(pady=10)
        
        # Botão para sair
        btn_sair = ttk.Button(
            self.main_frame,
            text="Sair",
            command=self.sair_sistema
        )
        btn_sair.pack(pady=30)
        
    def abrir_relatorios(self):
        """Abre o sistema integrado de relatórios"""
        try:
            print("Tentando abrir sistema de relatórios...")
            
            # Carregar o módulo de relatórios
            modulo = reload_module('relatorios_interface')
            if not modulo:
                messagebox.showerror("Erro", "Não foi possível carregar o módulo de relatórios.")
                return
                
            # Esconder janela principal
            self.root.withdraw()
            
            # Iniciar sistema de relatórios
            app = modulo.SistemaRelatorios(parent=self.root)
            
            # Configurar comportamento ao fechar
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            
            # Exibir janela
            app.root.lift()
            app.root.focus_force()
            app.run()
            
        except Exception as e:
            print(f"Erro ao abrir sistema de relatórios: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Erro", f"Erro ao abrir sistema de relatórios: {str(e)}")
            self.root.deiconify()
            
    def abrir_fornecedores(self):
        """Abre diretamente o relatório de fornecedores"""
        try:
            print("Tentando abrir relatório de fornecedores diretamente...")
            
            # Carregar o módulo de fornecedores
            modulo = reload_module('relatorio_fornecedores')
            if not modulo:
                messagebox.showerror("Erro", "Não foi possível carregar o módulo de relatório de fornecedores.")
                return
                
            # Esconder janela principal
            self.root.withdraw()
            
            # Iniciar relatório de fornecedores
            app = modulo.RelatorioFornecedores(parent=self.root)
            
            # Configurar comportamento ao fechar
            app.root.protocol("WM_DELETE_WINDOW", 
                lambda: self.finalizar_sistema(app.root))
            
            # Exibir janela
            app.root.lift()
            app.root.focus_force()
            app.root.mainloop()
            
        except Exception as e:
            print(f"Erro ao abrir relatório de fornecedores: {str(e)}")
            traceback.print_exc()
            messagebox.showerror("Erro", f"Erro ao abrir relatório de fornecedores: {str(e)}")
            self.root.deiconify()
            
    def finalizar_sistema(self, janela):
        """Fecha a janela do sistema e mostra a janela principal"""
        print("Finalizando janela...")
        try:
            janela.destroy()
        except Exception as e:
            print(f"Erro ao destruir janela: {str(e)}")
        
        # Mostrar janela principal novamente
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()
        
    def sair_sistema(self):
        """Fecha o sistema após confirmação"""
        if messagebox.askyesno("Confirmar Saída", "Deseja realmente sair do sistema?"):
            print("Saída confirmada, finalizando sistema...")
            self.root.destroy()
            # Forçar saída após um curto delay
            import threading
            threading.Timer(0.5, force_exit).start()
            
    def run(self):
        """Inicia a execução do sistema"""
        print("Iniciando mainloop...")
        self.root.mainloop()
        print("Mainloop encerrado")

# Função principal
def main():
    try:
        print("Iniciando aplicação...")
        app = SistemaPrincipalSimples()
        app.run()
    except Exception as e:
        print(f"Erro no sistema principal: {str(e)}")
        traceback.print_exc()
    finally:
        print("Forçando saída...")
        import threading
        threading.Timer(1.0, force_exit).start()

# Executar o aplicativo
if __name__ == "__main__":
    main()