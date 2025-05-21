"""
Módulo para funções de diálogo personalizadas que funcionam em qualquer contexto.
"""
import sys
import tkinter as tk
from tkinter import messagebox

# Armazenar uma referência global para a janela principal atual
_main_window = None

def set_main_window(window):
    """Define a janela principal atual do aplicativo.
    Isso deve ser chamado quando o aplicativo é iniciado."""
    global _main_window
    _main_window = window

def get_main_window():
    """Obtém a janela principal atual ou localiza uma se não estiver definida."""
    global _main_window
    
    if _main_window is not None and _main_window.winfo_exists():
        return _main_window
    
    # Procurar por janelas existentes
    if hasattr(tk, '_default_root') and tk._default_root:
        return tk._default_root
    
    # Procurar por Toplevels nas instâncias existentes
    if hasattr(tk, '_default_root') and tk._default_root:
        for widget in tk._default_root.winfo_children():
            if isinstance(widget, (tk.Toplevel, tk.Tk)):
                # Encontrada uma janela existente
                _main_window = widget
                return _main_window
    
    # Criar uma nova Tk temporária se necessário
    temp_root = tk.Tk()
    temp_root.withdraw()
    return temp_root

def _center_on_parent(dialog, parent=None):
    """Centraliza o diálogo na janela pai ou na tela"""
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    
    # Se tivermos uma janela pai válida, centralizar nela
    if parent and parent.winfo_exists():
        try:
            x = parent.winfo_x() + (parent.winfo_width() // 2) - (width // 2)
            y = parent.winfo_y() + (parent.winfo_height() // 2) - (height // 2)
        except:
            # Fallback para a posição da janela pai
            x = parent.winfo_x() + 50
            y = parent.winfo_y() + 50
    else:
        # Se não tiver parent, tenta obter a janela principal do sistema
        main_window = get_main_window()
        if main_window and main_window.winfo_exists():
            # Centralizar com base na janela principal do sistema
            x = main_window.winfo_x() + (main_window.winfo_width() // 2) - (width // 2)
            y = main_window.winfo_y() + (main_window.winfo_height() // 2) - (height // 2)
        else:
            # Último recurso: ajustar para a parte esquerda da tela (onde seu sistema fica)
            screen_width = dialog.winfo_screenwidth()
            screen_height = dialog.winfo_screenheight()
            # Considerando que seu sistema ocupa a parte esquerda da tela,
            # centralizamos apenas nessa área (metade da largura da tela)
            x = (screen_width // 4) - (width // 2)
            y = (screen_height // 2) - (height // 2)
            
    # Garantir que a janela fique dentro da tela
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    
    if x < 0: x = 0
    if y < 0: y = 0
    if x + width > screen_width: x = screen_width - width
    if y + height > screen_height: y = screen_height - height
    
    dialog.geometry(f"+{x}+{y}")

def custom_messagebox(tipo="info", titulo="Mensagem", mensagem="", opcoes=None):
    """
    Função unificada de diálogo que funciona em qualquer contexto.
    
    Args:
        tipo: string - 'info', 'warning', 'error', 'yesno'
        titulo: string - título da janela
        mensagem: string - mensagem a ser exibida
        opcoes: lista de strings - opções de botões (apenas para tipo 'yesno')
        
    Returns:
        Para 'yesno': boolean - True se "Sim", False se "Não"
        Para outros tipos: None
    """
    # Obter a janela principal atual para centralização
    parent = get_main_window()
    
    # Tente usar os diálogos padrão do Tkinter primeiro
    try:
        # Verificar se estamos em um executável
        is_executable = getattr(sys, 'frozen', False)
        
        # Usar o diálogo padrão do Tkinter
        import tkinter.messagebox as tkMessageBox
        
        # Preparar a janela para exibir o diálogo centralizado
        if parent and parent.winfo_exists():
            parent.update_idletasks()
            
            # Para executáveis, forçar diálogo a ficar visível
            if is_executable:
                parent.lift()
                parent.focus_force()
        
        # Exibir o diálogo apropriado
        if tipo == "info":
            result = tkMessageBox.showinfo(titulo, mensagem, parent=parent)
        elif tipo == "warning":
            result = tkMessageBox.showwarning(titulo, mensagem, parent=parent)
        elif tipo == "error":
            result = tkMessageBox.showerror(titulo, mensagem, parent=parent)
        elif tipo == "yesno":
            result = tkMessageBox.askyesno(titulo, mensagem, parent=parent)
        else:
            result = None
        
        return result
        
    except Exception as e:
        print(f"Erro ao mostrar diálogo padrão: {e}")
        # Se falhar, tente um diálogo personalizado como fallback
        return _custom_dialog_fallback(tipo, titulo, mensagem, parent)

def _custom_dialog_fallback(tipo, titulo, mensagem, parent=None):
    """Versão de fallback usando um diálogo personalizado."""
    try:
        # Criar janela de diálogo
        dialog = tk.Toplevel(parent)
        dialog.title(titulo)
        dialog.transient(parent)
        dialog.grab_set()
        
        # Configurar o diálogo
        tk.Label(dialog, text=mensagem, padx=20, pady=20, wraplength=300).pack()
        
        # Resultado para diálogos yesno
        result = [False]
        
        if tipo == "yesno":
            def on_yes():
                result[0] = True
                dialog.destroy()
                
            def on_no():
                result[0] = False
                dialog.destroy()
                
            # Frame para botões
            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=10)
            
            # Botões
            yes_btn = tk.Button(btn_frame, text="Sim", command=on_yes, width=10)
            yes_btn.pack(side="left", padx=5)
            
            no_btn = tk.Button(btn_frame, text="Não", command=on_no, width=10)
            no_btn.pack(side="left", padx=5)
            
            # Configurar teclas
            dialog.bind("<Return>", lambda e: on_yes())
            dialog.bind("<Escape>", lambda e: on_no())
            
            # Foco no botão Sim
            yes_btn.focus_set()
        else:
            # Botão OK
            ok_btn = tk.Button(dialog, text="OK", command=dialog.destroy, width=10)
            ok_btn.pack(pady=10)
            
            # Configurar teclas
            dialog.bind("<Return>", lambda e: dialog.destroy())
            dialog.bind("<Escape>", lambda e: dialog.destroy())
            
            # Foco no botão OK
            ok_btn.focus_set()
        
        # Centralizar diálogo na janela pai
        _center_on_parent(dialog, parent)
        
        # Levantar o diálogo para o topo
        dialog.lift()
        dialog.focus_force()
        
        # Esperar até o diálogo ser fechado
        dialog.wait_window()
        
        # Retornar resultado
        if tipo == "yesno":
            return result[0]
        return None
        
    except Exception as e:
        print(f"Erro no diálogo personalizado: {e}")
        # Se tudo falhar, use print como último recurso
        print(f"\n--- {titulo} ---\n{mensagem}\n")
        if tipo == "yesno":
            return False
        return None