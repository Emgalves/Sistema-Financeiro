"""
Módulo para funções de diálogo personalizadas que funcionam em qualquer contexto.
"""
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
            # Centralizar na tela em caso de erro
            x = (dialog.winfo_screenwidth() // 2) - (width // 2)
            y = (dialog.winfo_screenheight() // 2) - (height // 2)
    else:
        # Centralizar na tela
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
    
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
        # Forçar diálogo a ficar no topo
        if hasattr(messagebox, 'tk'):
            messagebox.tk.call('wm', 'attributes', '.', '-topmost', True)
        
        if tipo == "info":
            result = messagebox.showinfo(titulo, mensagem, parent=parent)
        elif tipo == "warning":
            result = messagebox.showwarning(titulo, mensagem, parent=parent)
        elif tipo == "error":
            result = messagebox.showerror(titulo, mensagem, parent=parent)
        elif tipo == "yesno":
            result = messagebox.askyesno(titulo, mensagem, parent=parent)
        else:
            result = None
        
        # Restaurar estado normal
        if hasattr(messagebox, 'tk'):
            messagebox.tk.call('wm', 'attributes', '.', '-topmost', False)
            
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
        dialog.attributes('-topmost', True)
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
        
        # Centralizar na janela pai
        _center_on_parent(dialog, parent)
        
        # Manter no topo - especialmente para mostrar acima do visualizador
        dialog.attributes('-topmost', True)
        dialog.update()
        
        # Função para manter a janela no topo
        def keep_on_top():
            try:
                if dialog.winfo_exists():
                    dialog.lift()
                    dialog.attributes('-topmost', True)
                    dialog.after(200, keep_on_top)
            except:
                pass
                
        # Iniciar processo para manter no topo
        dialog.after(100, keep_on_top)
        
        # Esperar até fechar
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