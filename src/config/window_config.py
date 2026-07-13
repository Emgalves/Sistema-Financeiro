# src/config/window_config.py
import os
import sys
import tkinter as tk


def _resource_path(relative_path):
    """
    Obtém o caminho absoluto do recurso, funcionando tanto em desenvolvimento
    quanto no .exe empacotado pelo PyInstaller (mesmo padrão usado em
    sistema_principal.py, duplicado aqui para evitar import circular).
    """
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


def configurar_janela(janela, titulo, largura=900, altura=1000):
    """
    Configura o posicionamento e dimensionamento padrão de uma janela
    
    Args:
        janela: Instância de tk.Tk ou tk.Toplevel
        titulo: Título da janela
        largura: Largura desejada (default 900)
        altura: Altura desejada (default 1000)
    """
    janela.title(titulo)

    # Ícone da janela / barra de tarefas
    try:
        icone_ico = _resource_path("logo3.ico")
        if os.path.exists(icone_ico):
            janela.iconbitmap(icone_ico)
    except Exception:
        # iconbitmap com .ico pode falhar fora do Windows; nesse caso,
        # tenta a alternativa multiplataforma iconphoto com o PNG.
        try:
            from PIL import Image, ImageTk
            icone_png = _resource_path("logo3.png")
            if os.path.exists(icone_png):
                imagem_icone = Image.open(icone_png)
                janela._icone_photo = ImageTk.PhotoImage(imagem_icone)  # manter referência
                janela.iconphoto(True, janela._icone_photo)
        except Exception:
            pass

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