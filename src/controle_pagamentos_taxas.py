# Arquivo: controle_pagamentos.py
# Este arquivo integrará a gestão de eventos aos módulos existentes

from pathlib import Path
import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from openpyxl import load_workbook
from dateutil.relativedelta import relativedelta

# Adicionar diretório raiz ao path
def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Importar configurações do sistema
try:
    from src.config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )
    from src.config.utils import formatar_cnpj_cpf, aplicar_formatacao_celula
    from src.config.window_config import configurar_janela
except ImportError as e:
    print(f"Erro ao importar módulos: {str(e)}")
    # Tentar importar com caminho alternativo
    try:
        from config.config import (
            ARQUIVO_CLIENTES,
            ARQUIVO_MODELO,
            PASTA_CLIENTES,
            BASE_PATH
        )
        from config.utils import formatar_cnpj_cpf, aplicar_formatacao_celula
        from config.window_config import configurar_janela
    except ImportError as e2:
        print(f"Erro ao importar módulos (caminho alternativo): {str(e2)}")
        raise

class ControlePagamentos:
    def __init__(self, parent=None):
        self.parent = parent
        self.janela = None
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.gestao_eventos = None  # Inicializar como None

    
        
    # Modificar o método abrir_janela_controle na classe ControlePagamentos
    def abrir_janela_controle(self):
        """Abre a janela principal de controle de pagamentos"""
        # Se a janela já existir, apenas traz para frente
        if self.janela and self.janela.winfo_exists():
            self.janela.lift()
            self.janela.focus_force()
            return

        # Cria nova janela
        self.janela = tk.Toplevel(self.parent)
        configurar_janela(self.janela, "Controle de Pagamentos", 900, 1000)
        # Posicionar a janela no canto superior esquerdo
        self.janela.geometry("+0+0")

        # Permitir que esta função detecte quando a aplicação é executada diretamente
        is_main_window = self.parent and self.parent.winfo_toplevel() == self.parent
        
        # Frame principal
        frame_principal = ttk.Frame(self.janela, padding=20)
        frame_principal.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            frame_principal, 
            text="Sistema de Controle de Pagamentos", 
            font=('Arial', 16, 'bold')
        ).pack(pady=15)
        
        # Frame para opções
        frame_opcoes = ttk.Frame(frame_principal)
        frame_opcoes.pack(fill='x', pady=25)
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 12), padding=(20, 10))
        
        # Botões em grade (2x2)
        frame_botoes = ttk.Frame(frame_opcoes)
        frame_botoes.pack(padx=60, pady=25, fill='both', expand=True)
        
        # Configurar o grid para distribuir espaço igualmente
        frame_botoes.columnconfigure(0, weight=1)
        frame_botoes.columnconfigure(1, weight=1)
        frame_botoes.rowconfigure(0, weight=1)
        frame_botoes.rowconfigure(1, weight=1)
        
        # Linha 1
        ttk.Button(
            frame_botoes,
            text="Pagamentos por Percentual da Quinzena",
            command=self.abrir_percentual_quinzena,
            style='Big.TButton',
            width=40
        ).grid(row=0, column=0, padx=15, pady=15, sticky='nsew')
        
        ttk.Button(
            frame_botoes,
            text="Pagamentos por Eventos",
            command=self.abrir_gestao_eventos,
            style='Big.TButton',
            width=40
        ).grid(row=0, column=1, padx=15, pady=15, sticky='nsew')
        
        # Linha 2
        ttk.Button(
            frame_botoes,
            text="Contratos de Administração",
            command=self.abrir_gestao_contratos,
            style='Big.TButton',
            width=40
        ).grid(row=1, column=0, padx=15, pady=15, sticky='nsew')
        
        ttk.Button(
            frame_botoes,
            text="Relatórios e Consultas",
            command=self.abrir_relatorios,
            style='Big.TButton',
            width=40
        ).grid(row=1, column=1, padx=15, pady=15, sticky='nsew')
        
        # Texto explicativo
        frame_info = ttk.LabelFrame(frame_principal, text="Informações")
        frame_info.pack(fill='both', expand=True, pady=40, padx=50)
        
        texto_info = """
        • Pagamentos por Percentual da Quinzena: 
        Gerencia pagamentos calculados como percentual das despesas da quinzena.
        
        • Pagamentos por Eventos: 
        Controla pagamentos vinculados à conclusão de eventos específicos definidos no contrato.
        
        • Contratos de Administração: 
        Gerencia contratos, seus administradores, eventos e parcelas.
        
        • Relatórios e Consultas: 
        Relatórios gerenciais e consultas de pagamentos por período.
        """
        
        texto = tk.Text(frame_info, wrap='word', height=10, width=90, font=('Arial', 11))
        texto.pack(padx=15, pady=15, fill='both', expand=True)
        texto.insert('1.0', texto_info)
        texto.config(state='disabled')
        
        # Frame para botões de ação
        frame_acoes = ttk.Frame(frame_principal)
        frame_acoes.pack(fill='x', pady=20)
        
        # NOVO BOTÃO: Voltar ao Menu Principal
        ttk.Button(
            frame_acoes,
            text="Voltar ao Menu Principal",
            command=self.voltar_menu_principal,
            width=25
        ).pack(side='right', padx=10)
        
        # Botão para fechar (agora apenas fecha esta janela)
        ttk.Button(
            frame_acoes,
            text="Fechar",
            command=lambda: self.fechar_janela(is_main_window),
            width=25
        ).pack(side='right', padx=10)

    # Adicionar um novo método voltar_menu_principal na classe ControlePagamentos
    def voltar_menu_principal(self):
        """Fecha a janela atual e volta para o menu principal"""
        print("Voltando ao menu principal...")
        if self.janela:
            self.janela.destroy()
        
        # Exibir novamente a janela principal se existir
        if self.parent:
            try:
                print("Mostrando janela principal")
                self.parent.deiconify()
                self.parent.lift()
                self.parent.focus_force()
            except Exception as e:
                print(f"Erro ao mostrar janela principal: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao voltar ao menu principal: {str(e)}")

    def fechar_janela(self, is_main_window):
        """Fecha a janela e encerra a aplicação apenas se for a janela principal"""
        if is_main_window and self.parent:
            # Se for executado como janela principal, encerra o aplicativo
            if messagebox.askyesno("Confirmar", "Deseja realmente sair do sistema?"):
                self.janela.destroy()
                self.parent.quit()  # Encerra o mainloop
                self.parent.destroy()  # Destrói a janela principal
        else:
            # Se não for a janela principal, apenas fecha esta janela
            self.janela.destroy()
        
    def abrir_percentual_quinzena(self):
        """Abre o módulo de percentual da quinzena"""
        try:
            # Verifica se já existe o módulo importado
            if hasattr(self.parent, 'abrir_finalizacao_quinzena'):
                self.parent.abrir_finalizacao_quinzena()
            else:
                # Tentar importar e executar
                from gestao_taxas import GestaoTaxasAdministracao
                gestao = GestaoTaxasAdministracao(self.parent)
                gestao.abrir_finalizacao_quinzena()
        except ImportError:
            messagebox.showerror("Erro", "Módulo de Finalização de Quinzena não encontrado")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir módulo: {str(e)}")
        
    def abrir_gestao_eventos(self):
        """Abre o módulo de gestão de eventos (simplificado)"""
        try:
            print("Iniciando gestão de eventos (simplificado)")
            
            # Importar o módulo localmente
            try:
                from src.controle_pagamentos import ControlePagamentos
            except ImportError:
                try:
                    from controle_pagamentos import ControlePagamentos
                except ImportError as e:
                    print(f"Erro ao importar ControlePagamentos: {str(e)}")
                    messagebox.showerror("Erro", f"Módulo de Controle de Pagamentos não encontrado: {str(e)}")
                    return
            
            # Salvar a referência para esta instância como "controlador_principal"
            # Isso será usado para voltar à janela correta
            self.janela.withdraw()  # Ocultar a janela de controle antes de abrir a nova

            # Criar instância do ControlePagamentos com referência a esta classe
            gestao = ControlePagamentos(None)  # Começar sem parent
            
            # Armazenar a referência para o controlador principal
            gestao.controlador_principal = self
            
            # Configurar um callback personalizado para quando a janela for fechada
            def on_close():
                print("Callback on_close executado")
                gestao.root.destroy()
                # Mostrar a janela do menu novamente
                if hasattr(self, 'janela') and self.janela:
                    print("Reexibindo janela do menu")
                    self.janela.deiconify()
            
            # Substituir o método voltar_menu original
            original_voltar_menu = gestao.voltar_menu
            
            def custom_voltar_menu():
                print("Custom voltar_menu executado")
                gestao.root.destroy()
                # Mostrar a janela do menu novamente
                if hasattr(self, 'janela') and self.janela:
                    print("Reexibindo janela do menu pelo custom_voltar_menu")
                    self.janela.deiconify()
            
            # Substituir o método por nossa versão personalizada
            gestao.voltar_menu = custom_voltar_menu
            
            # Configurar o protocolo para o botão de fechar (X)
            gestao.root.protocol("WM_DELETE_WINDOW", on_close)
            
        except Exception as e:
            print(f"Erro ao abrir controle de pagamentos: {str(e)}")
            messagebox.showerror("Erro", f"Erro ao abrir controle de pagamentos: {str(e)}")
            # Garantir que a janela anterior seja restaurada em caso de erro
            if hasattr(self, 'janela') and self.janela:
                self.janela.deiconify()

        
    def abrir_gestao_contratos(self):
        """Abre o módulo de gestão de contratos"""
        try:
            # Primeiro, selecionar um cliente
            if self.selecionar_cliente():
                # Importar e instanciar o módulo
                try:
                    from sistema_entrada_dados import GestaoContratos
                    gestao_contratos = GestaoContratos(self.parent)
                    
                    # Criar uma nova janela explícita para a gestão
                    janela_gestao = tk.Toplevel(self.parent)
                    janela_gestao.title(f"Gestão de Contratos - {self.cliente_atual}")
                    
                    def on_close():
                        janela_gestao.destroy()
                        self.parent.lift()
                        self.parent.focus_force()
                    
                    gestao_contratos.cliente_atual = self.cliente_atual
                    gestao_contratos.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
                    gestao_contratos.criar_interface_contratos(janela_gestao, on_close)
                    
                except ImportError as ie:
                    messagebox.showerror("Erro", f"Módulo de Gestão de Contratos não encontrado: {str(ie)}")
                except Exception as e:
                    messagebox.showerror("Erro", f"Erro ao abrir gestão de contratos: {str(e)}")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir gestão de contratos: {str(e)}")
        
    def abrir_relatorios(self):
        """Abre o módulo de relatórios"""
        messagebox.showinfo("Informação", "Módulo de Relatórios em desenvolvimento")
        
    


# Se executado diretamente, abre a janela de controle
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    app = ControlePagamentos(root)
    
    # Abrir a janela de controle
    app.abrir_janela_controle()
    
    # Definir manipulador para quando a janela de controle for fechada
    if app.janela:
        def on_control_close():
            print("Fechando aplicação")
            root.quit()
            root.destroy()
        
        app.janela.protocol("WM_DELETE_WINDOW", on_control_close)
    
    root.mainloop()