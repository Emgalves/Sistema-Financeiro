# src/config/window_config.py
import tkinter as tk

print("Carregando configurações de janela...")

def configurar_janela(janela, titulo, largura=900, altura=900):
    """
    Configura o posicionamento e dimensionamento padrão de uma janela
    
    Args:
        janela: Instância de tk.Tk ou tk.Toplevel
        titulo: Título da janela
        largura: Largura desejada (default 900)
        altura: Altura desejada (default 900)
    """
    janela.title(titulo)
    
    # Obter dimensões da tela
    screen_width = janela.winfo_screenwidth()
    screen_height = janela.winfo_screenheight()
    
    # Ajustar dimensões para não exceder o tamanho da tela
    largura = min(largura, screen_width)
    altura = min(altura, screen_height)
    
    # Definir posição (sempre no topo esquerdo)
    x = 0
    y = 0
    
    # Configurar geometria
    janela.geometry(f"{largura}x{altura}+{x}+{y}")
    
    # Permitir redimensionamento
    janela.resizable(True, True)
    
    # Configurar peso das linhas/colunas para redimensionamento proporcional
    janela.grid_rowconfigure(0, weight=1)
    janela.grid_columnconfigure(0, weight=1)
    
    # Trazer janela para frente
    janela.lift()
    janela.focus_force()