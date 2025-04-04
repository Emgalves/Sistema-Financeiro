"""
Script para reiniciar completamente a aplicação.
Execute este script para garantir que todas as instâncias sejam encerradas corretamente.
"""

import os
import sys
import tkinter as tk
from tkinter import messagebox

def reiniciar_aplicacao():
    # Verificar se há alguma instância do tkinter em execução
    try:
        # Tentar criar uma root do tkinter
        root = tk.Tk()
        root.withdraw()
        
        # Exibir mensagem
        messagebox.showinfo(
            "Reinicialização", 
            "Este script irá encerrar todas as instâncias do tkinter e reiniciar a aplicação."
        )
        
        # Destruir a root
        root.destroy()
        
        # Importar e executar o módulo principal
        print("Iniciando aplicação limpa...")
        from src.controle_pagamentos import ControlePagamentos
        
        # Criar nova instância do root
        root = tk.Tk()
        root.withdraw()
        
        # Executar a aplicação
        app = ControlePagamentos(root)
        app.abrir_janela_controle()
        
        # Iniciar mainloop
        root.mainloop()
        
    except Exception as e:
        print(f"Erro ao reiniciar: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    # Garantir que estamos no diretório correto
    script_dir = os.path.dirname(os.path.abspath(__file__))
    os.chdir(script_dir)
    
    # Adicionar diretório pai ao path
    parent_dir = os.path.abspath(os.path.join(script_dir, '..'))
    if parent_dir not in sys.path:
        sys.path.insert(0, parent_dir)
    
    # Reiniciar aplicação
    reiniciar_aplicacao()