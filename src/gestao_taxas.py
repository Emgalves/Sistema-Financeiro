import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import sys
from pathlib import Path
import os

def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

try:
    from config.utils import PASTA_CLIENTES
except ImportError:
    from src.config.utils import PASTA_CLIENTES

from src.controle_pagamentos import ControlePagamentos
from src.finalizacao_quinzena import FinalizacaoQuinzena


try:
    from config.window_config import configurar_janela
except ImportError:
    from src.config.window_config import configurar_janela

class GestaoTaxasAdministracao:
    def __init__(self, parent=None):
        self.parent = parent
        self.controle_pagamentos = None
        self.finalizacao_quinzena = None
        # Não inicializar aqui - será feito sob demanda
    
    def inicializar_controles(self):
        """Inicializa os controladores de taxa"""
        try:
            print("Preparando sistema...")
            print(f"Diretório atual: {os.getcwd()}")
            print(f"Python path: {sys.path}")
            
            # Imports locais para evitar circular imports
            print("\nTentando importar controle_pagamentos...")
            try:
                print("Tentativa 1: direto")
                from controle_pagamentos import ControlePagamentos
                print("Sucesso!")
            except ImportError as e1:
                print(f"Falha 1: {e1}")
                try:
                    print("Tentativa 2: de src")
                    from src.controle_pagamentos import ControlePagamentos # type: ignore
                    print("Sucesso!")
                except ImportError as e2:
                    print(f"Falha 2: {e2}")
                    try:
                        print("Tentativa 3: com path ajustado")
                        current_dir = Path(__file__).resolve().parent
                        root_dir = current_dir.parent
                        sys.path.insert(0, str(root_dir))
                        from controle_pagamentos import ControlePagamentos
                        print("Sucesso!")
                    except ImportError as e3:
                        print(f"Falha 3: {e3}")
                        raise ImportError("Não foi possível importar ControlePagamentos")

            print("Carregando módulo de Controle de Pagamentos...")
            self._classe_controle = ControlePagamentos

            print("\nTentando importar finalizacao_quinzena...")
            try:
                # Tentar importação absoluta primeiro
                from finalizacao_quinzena import FinalizacaoQuinzena
                print("Sucesso na importação direta!")
            except ImportError:
                try:
                    # Tentar com prefixo src
                    from src.finalizacao_quinzena import FinalizacaoQuinzena # type: ignore
                    print("Sucesso na importação via src!")
                except ImportError:
                    # Última tentativa: importar do mesmo diretório
                    current_dir = Path(__file__).resolve().parent
                    if str(current_dir) not in sys.path:
                        sys.path.append(str(current_dir))
                    from finalizacao_quinzena import FinalizacaoQuinzena
                    print("Sucesso na importação do diretório atual!")
                    
            print("Carregando módulo de Finalização de Quinzena...")
            self._classe_quinzena = FinalizacaoQuinzena
            
            # Verificar se ambos foram carregados
            if not hasattr(self, '_classe_controle') or not hasattr(self, '_classe_quinzena'):
                raise Exception("Falha ao carregar um ou mais módulos")
            
            print("Sistema preparado com sucesso!")
            return True
            
        except Exception as e:
            print(f"Erro ao preparar sistema: {str(e)}")
            print("Stack trace:")
            import traceback
            traceback.print_exc()
            messagebox.showerror("Erro", f"Erro ao preparar sistema: {str(e)}", parent=self.menu)
            return False
    
    def abrir_controle_pagamentos(self):
        """Abre a interface de controle de pagamentos"""
        if not hasattr(self, '_classe_controle'):
            if not self.inicializar_controles():
                return
        
        try:
            print("\nAbrindo Controle de Pagamentos...")
            # Criar nova instância apenas quando necessário
            self.controle_pagamentos = self._classe_controle(self.parent)
            self.parent.withdraw()
            # Fechar a janela de menu de taxas
            if hasattr(self, 'menu') and self.menu:
                self.menu.destroy()
            self.controle_pagamentos.run()
        except Exception as e:
            print(f"Erro ao abrir controle de pagamentos: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir controle de pagamentos: {str(e)}", parent=self.menu)
            self.parent.deiconify()
    
    def abrir_finalizacao_quinzena(self):
        """Abre a interface de finalização de quinzena"""
        if not hasattr(self, '_classe_quinzena'):
            if not self.inicializar_controles():
                return
        
        try:
            print("\nAbrindo Finalização de Quinzena...")
            # Criar nova instância apenas quando necessário
            self.finalizacao_quinzena = self._classe_quinzena(self.parent)
            self.parent.withdraw()
            # Fechar a janela de menu de taxas
            if hasattr(self, 'menu') and self.menu:
                self.menu.destroy()
            self.finalizacao_quinzena.run()
        except Exception as e:
            print(f"Erro ao abrir finalização de quinzena: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir finalização de quinzena: {str(e)}", parent=self.menu)
            self.parent.deiconify()
    
