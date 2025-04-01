"""
Módulo de controle de versões para o Sistema de Gestão Financeira
"""
import os
import json
from datetime import datetime
from pathlib import Path

# Informações da versão atual
VERSION_INFO = {
    "major": 1,
    "minor": 1,
    "patch": 0,
    "release_date": "14/03/2025",
    "changes": [
        "Separação de despesas com funcionários, no Relatório, separando despesas mensais das eventuais",
        "Correção do campo percentual e inclusão de dados bancários na Entrada em Gestão de Contrato",
        "Melhoria do layout da tela de Seleção de Cliente",
        "Correção da combobox Banco em Fornecedor",
        "Reorganização da aba Fornecedor com funcionalidade de duplo clique",
        "Ajuste do layout da aba Entrada de Dados para melhor usabilidade",
        "Adição de preenchimento automático da referência baseado na especificação do fornecedor"
    ]
}

def get_version_string():
    """Retorna a string formatada da versão atual"""
    return f"{VERSION_INFO['major']}.{VERSION_INFO['minor']}.{VERSION_INFO['patch']}"

def get_version_info():
    """Retorna informações completas sobre a versão atual"""
    return {
        "version": get_version_string(),
        "release_date": VERSION_INFO["release_date"],
        "changes": VERSION_INFO["changes"]
    }

def save_version_history():
    """Salva o histórico de versões em um arquivo JSON"""
    version_file = Path("config") / "version_history.json"
    
    # Garantir que o diretório existe
    os.makedirs(version_file.parent, exist_ok=True)
    
    # Carregar histórico existente, se houver
    history = []
    if version_file.exists():
        try:
            with open(version_file, 'r', encoding='utf-8') as f:
                history = json.load(f)
        except json.JSONDecodeError:
            # Se o arquivo estiver corrompido, começar um novo
            history = []
    
    # Verificar se a versão atual já está no histórico
    current_version = get_version_string()
    if not any(entry.get('version') == current_version for entry in history):
        # Adicionar versão atual ao histórico
        version_data = get_version_info()
        version_data['timestamp'] = datetime.now().isoformat()
        history.append(version_data)
        
        # Salvar histórico atualizado
        with open(version_file, 'w', encoding='utf-8') as f:
            json.dump(history, f, indent=4, ensure_ascii=False)
    
    return history

def compare_versions(installed_version, available_version):
    """
    Compara duas versões para verificar se uma atualização está disponível
    
    Args:
        installed_version (str): Versão instalada no formato "X.Y.Z"
        available_version (str): Versão disponível no formato "X.Y.Z"
        
    Returns:
        bool: True se available_version for mais recente que installed_version
    """
    try:
        installed = [int(x) for x in installed_version.split('.')]
        available = [int(x) for x in available_version.split('.')]
        
        # Comparar componentes de versão (major, minor, patch)
        for i in range(max(len(installed), len(available))):
            # Tratamento para quando um array é mais curto que o outro
            inst_val = installed[i] if i < len(installed) else 0
            avail_val = available[i] if i < len(available) else 0
            
            if avail_val > inst_val:
                return True
            elif avail_val < inst_val:
                return False
        
        # Se chegou aqui, as versões são iguais
        return False
        
    except (ValueError, IndexError):
        # Em caso de erro de formato, retornar False por segurança
        return False

def show_version_dialog(parent):
    """
    Exibe um diálogo com informações da versão atual
    
    Args:
        parent: Widget pai para o diálogo
    """
    import tkinter as tk
    from tkinter import ttk, scrolledtext
    
    dialog = tk.Toplevel(parent)
    dialog.title(f"Sobre o Sistema - Versão {get_version_string()}")
    dialog.geometry("500x450")
    dialog.transient(parent)
    dialog.grab_set()
    
    # Centralizar a janela
    dialog.update_idletasks()
    width = dialog.winfo_width()
    height = dialog.winfo_height()
    x = (dialog.winfo_screenwidth() // 2) - (width // 2)
    y = (dialog.winfo_screenheight() // 2) - (height // 2)
    dialog.geometry(f'{width}x{height}+{x}+{y}')
    
    # Frame principal
    main_frame = ttk.Frame(dialog, padding=10)
    main_frame.pack(fill='both', expand=True)
    
    # Título
    ttk.Label(
        main_frame, 
        text=f"Sistema de Gestão Financeira", 
        font=('Helvetica', 16, 'bold')
    ).pack(pady=(0, 5))
    
    # Versão
    ttk.Label(
        main_frame, 
        text=f"Versão {get_version_string()}", 
        font=('Helvetica', 12)
    ).pack(pady=(0, 5))
    
    # Data de lançamento
    ttk.Label(
        main_frame, 
        text=f"Lançado em: {VERSION_INFO['release_date']}", 
        font=('Helvetica', 10)
    ).pack(pady=(0, 10))
    
    # Frame para mudanças nesta versão
    changes_frame = ttk.LabelFrame(main_frame, text="Mudanças nesta versão", padding=10)
    changes_frame.pack(fill='both', expand=True, pady=10)
    
    # Lista de mudanças
    changes_text = scrolledtext.ScrolledText(changes_frame, wrap=tk.WORD, height=10)
    changes_text.pack(fill='both', expand=True)
    changes_text.insert(tk.END, "\n".join(f"• {change}" for change in VERSION_INFO["changes"]))
    changes_text.config(state='disabled')  # Torna o texto somente leitura
    
    # Copyright
    ttk.Label(
        main_frame, 
        text="© 2025 Todos os direitos reservados.", 
        font=('Helvetica', 8)
    ).pack(pady=(10, 0))
    
    # Botão fechar
    ttk.Button(
        main_frame, 
        text="Fechar", 
        command=dialog.destroy
    ).pack(pady=10)
