import tkinter as tk
from tkinter import ttk

def custom_messagebox(tipo="info", titulo="Mensagem", mensagem="", opcoes=None):
    """
    Cria uma caixa de diálogo personalizada que sempre fica visível,
    mesmo quando chamada de uma janela configurada como 'always on top'.
    
    Args:
        tipo: string - 'info', 'warning', 'error', 'yesno'
        titulo: string - título da janela
        mensagem: string - mensagem a ser exibida
        opcoes: lista de strings - opções de botões (apenas para tipo 'yesno')
        
    Returns:
        Para 'yesno': boolean - True se "Sim", False se "Não"
        Para outros tipos: None
    """
    # Encontrar a janela principal da aplicação
    root = None
    try:
        if hasattr(tk, '_default_root') and tk._default_root:
            root = tk._default_root
    except:
        pass
    
    # Se não encontrou uma janela raiz, criar uma temporária
    if root is None:
        root = tk.Tk()
        root.withdraw()  # Ocultar a janela temporária
    
    # Criar nova janela
    dialog = tk.Toplevel(root)
    dialog.title(titulo)
    dialog.geometry("450x250")
    dialog.resizable(False, False)
    
    # Tornar modal
    dialog.transient(root)
    dialog.grab_set()
    
    # Configurar para ficar sempre na frente
    dialog.attributes('-topmost', True)
    
    # Configuração de estilos e ícones baseado no tipo
    if tipo == 'info':
        icon_text = "ℹ️"
        cor_cabecalho = "#4287f5"  # Azul
    elif tipo == 'warning':
        icon_text = "⚠️"
        cor_cabecalho = "#f5a742"  # Amarelo
    elif tipo == 'error':
        icon_text = "❌"
        cor_cabecalho = "#f54242"  # Vermelho
    elif tipo == 'yesno':
        icon_text = "❓"
        cor_cabecalho = "#42f5a7"  # Verde claro
    else:
        icon_text = "ℹ️"
        cor_cabecalho = "#4287f5"  # Azul padrão
    
    # Frame para o cabeçalho colorido
    frame_cabecalho = tk.Frame(dialog, bg=cor_cabecalho, height=40)
    frame_cabecalho.pack(fill='x')
    
    # Título no cabeçalho
    tk.Label(
        frame_cabecalho, 
        text=titulo, 
        bg=cor_cabecalho, 
        fg="white", 
        font=('Arial', 12, 'bold')
    ).pack(pady=8)
    
    # Frame para o conteúdo
    frame_conteudo = tk.Frame(dialog, bg="white")
    frame_conteudo.pack(fill='both', expand=True)
    
    # Ícone e mensagem
    frame_mensagem = tk.Frame(frame_conteudo, bg="white")
    frame_mensagem.pack(fill='both', expand=True, padx=20, pady=10)
    
    tk.Label(
        frame_mensagem, 
        text=icon_text, 
        font=('Arial', 24), 
        bg="white"
    ).pack(side='left', padx=(0, 15))
    
    tk.Label(
        frame_mensagem, 
        text=mensagem, 
        justify='left', 
        wraplength=300, 
        font=('Arial', 10), 
        bg="white"
    ).pack(side='left')
    
    # Frame para botões
    frame_botoes = tk.Frame(dialog, bg="#f0f0f0", height=50)
    frame_botoes.pack(fill='x')
    
    resposta = [False]  # Para armazenar a resposta do yesno
    
    if tipo == 'yesno':
        def responder_sim():
            resposta[0] = True
            dialog.destroy()
            
        def responder_nao():
            resposta[0] = False
            dialog.destroy()
        
        # Botão "Sim"
        btn_sim = ttk.Button(
            frame_botoes, 
            text="Sim", 
            command=responder_sim, 
            width=10
        )
        btn_sim.pack(side='right', padx=10, pady=10)
        
        # Botão "Não"
        btn_nao = ttk.Button(
            frame_botoes, 
            text="Não", 
            command=responder_nao, 
            width=10
        )
        btn_nao.pack(side='right', padx=5, pady=10)
        
        # Binding para teclas
        dialog.bind('<Return>', lambda e: responder_sim())  # Enter = Sim
        dialog.bind('<Escape>', lambda e: responder_nao())  # Esc = Não
        
    else:
        # Função para fechar com OK
        def fechar_dialog(event=None):
            dialog.destroy()
        
        # Botão OK
        btn_ok = ttk.Button(
            frame_botoes, 
            text="OK", 
            command=fechar_dialog, 
            width=10
        )
        btn_ok.pack(side='right', padx=10, pady=10)
        
        # Binding para Enter e Escape
        dialog.bind('<Return>', fechar_dialog)
        dialog.bind('<Escape>', fechar_dialog)
    
    # Centralizar o diálogo na janela pai
    dialog.update_idletasks()
    dialog_width = dialog.winfo_width()
    dialog_height = dialog.winfo_height()
    
    # Tentar obter a posição da janela pai
    try:
        parent_x = root.winfo_x()
        parent_y = root.winfo_y()
        parent_width = root.winfo_width()
        parent_height = root.winfo_height()
    except:
        # Se falhar, centralizar na tela
        screen_width = dialog.winfo_screenwidth()
        screen_height = dialog.winfo_screenheight()
        parent_x = 0
        parent_y = 0
        parent_width = screen_width
        parent_height = screen_height
    
    # Calcular centro da janela pai
    center_x = parent_x + parent_width // 2
    center_y = parent_y + parent_height // 2
    
    # Centralizar a janela de diálogo nesse ponto
    x = center_x - dialog_width // 2
    y = center_y - dialog_height // 2
    
    # Garantir que a janela não fique fora da tela
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    
    if x < 0: x = 0
    if y < 0: y = 0
    if x + dialog_width > screen_width: x = screen_width - dialog_width
    if y + dialog_height > screen_height: y = screen_height - dialog_height
    
    # Definir a posição
    dialog.geometry(f"+{x}+{y}")
    
    # Importante: Mantenha topmost até que a janela tenha sido posicionada e exibida
    # para garantir que ela não fique atrás de nenhuma outra janela
    dialog.attributes('-topmost', True)
    dialog.update()  # Forçar atualização para aplicar posicionamento
    
    # Definir foco no botão apropriado
    if tipo == 'yesno':
        btn_sim.focus_set()  # Foco no botão Sim para caixas de confirmação
    else:
        btn_ok.focus_set()   # Foco no botão OK para outras caixas
    
    # Configurar evento para "Re-levantar" a janela se ela perder o foco
    def manter_na_frente():
        dialog.lift()
        dialog.attributes('-topmost', True)
        # Agendar próxima verificação
        dialog.after(100, manter_na_frente)
    
    # Iniciar o mecanismo para manter na frente
    manter_na_frente()
    
    # Aguardar o fechamento da janela
    dialog.wait_window()
    
    # Retornar resposta para 'yesno'
    if tipo == 'yesno':
        return resposta[0]
    return None