"""
Launcher para o Sistema de Gestão Financeira
Este script simplificado serve como ponto de entrada para o sistema
"""
import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk
import traceback

def setup_path():
    """Configura os caminhos para que os módulos possam ser encontrados"""
    # Adicionar diretório atual e src ao path
    current_dir = os.path.dirname(os.path.abspath(__file__))
    sys.path.insert(0, current_dir)
    
    src_path = os.path.join(current_dir, 'src')
    if os.path.exists(src_path):
        sys.path.insert(0, src_path)
    
    # Adicionar diretório config ao path
    config_path = os.path.join(current_dir, 'src', 'config')
    if os.path.exists(config_path):
        sys.path.insert(0, config_path)

def iniciar_sistema():
    """Inicia o sistema principal com tratamento de erros"""
    try:
        # Primeiro tentar importar o módulo
        try:
            from src.sistema_principal import SistemaPrincipal
        except ImportError:
            # Se falhar, tentar sem o prefixo src
            from sistema_principal import SistemaPrincipal
        
        # Iniciar a aplicação
        app = SistemaPrincipal()
        app.run()
    except Exception as e:
        # Capturar e exibir qualquer erro
        erro = f"Erro ao iniciar o sistema: {str(e)}\n"
        erro += traceback.format_exc()
        
        # Salvar o erro em um arquivo
        with open("erro_inicializacao.log", "w", encoding="utf-8") as f:
            f.write(erro)
        
        # Mostrar mensagem de erro
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Erro de Inicialização", 
                f"Ocorreu um erro ao iniciar o sistema.\n\n{str(e)}\n\n"
                f"Detalhes foram salvos em 'erro_inicializacao.log'")
        except:
            print(erro)

if __name__ == "__main__":
    setup_path()
    iniciar_sistema()