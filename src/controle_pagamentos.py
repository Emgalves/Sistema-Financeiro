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
        
    def abrir_janela_controle(self):
        """Abre a janela principal de controle de pagamentos"""
        # Se a janela já existir, apenas traz para frente
        if self.janela and self.janela.winfo_exists():
            self.janela.lift()
            self.janela.focus_force()
            return

        # Cria nova janela
        self.janela = tk.Toplevel(self.parent)
        configurar_janela(self.janela, "Controle de Pagamentos", 900, 650)
        
        # Frame principal
        frame_principal = ttk.Frame(self.janela, padding=10)
        frame_principal.pack(fill='both', expand=True)
        
        # Título
        ttk.Label(
            frame_principal, 
            text="Sistema de Controle de Pagamentos", 
            font=('Arial', 14, 'bold')
        ).pack(pady=10)
        
        # Frame para opções
        frame_opcoes = ttk.Frame(frame_principal)
        frame_opcoes.pack(fill='x', pady=20)
        
        # Estilo para botões grandes
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 12), padding=(20, 10))
        
        # Botões em grade (2x2)
        frame_botoes = ttk.Frame(frame_opcoes)
        frame_botoes.pack(padx=50, pady=20)
        
        # Linha 1
        ttk.Button(
            frame_botoes,
            text="Pagamentos por Percentual da Quinzena",
            command=self.abrir_percentual_quinzena,
            style='Big.TButton',
            width=35
        ).grid(row=0, column=0, padx=10, pady=10)
        
        ttk.Button(
            frame_botoes,
            text="Pagamentos por Eventos",
            command=self.abrir_gestao_eventos,
            style='Big.TButton',
            width=35
        ).grid(row=0, column=1, padx=10, pady=10)
        
        # Linha 2
        ttk.Button(
            frame_botoes,
            text="Contratos de Administração",
            command=self.abrir_gestao_contratos,
            style='Big.TButton',
            width=35
        ).grid(row=1, column=0, padx=10, pady=10)
        
        ttk.Button(
            frame_botoes,
            text="Relatórios e Consultas",
            command=self.abrir_relatorios,
            style='Big.TButton',
            width=35
        ).grid(row=1, column=1, padx=10, pady=10)
        
        # Texto explicativo
        frame_info = ttk.LabelFrame(frame_principal, text="Informações")
        frame_info.pack(fill='x', pady=20, padx=50)
        
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
        
        texto = tk.Text(frame_info, wrap='word', height=10, width=80)
        texto.pack(padx=10, pady=10, fill='both', expand=True)
        texto.insert('1.0', texto_info)
        texto.config(state='disabled')
        
        # Botão para fechar
        ttk.Button(
            frame_principal,
            text="Fechar",
            command=self.janela.destroy,
            width=20
        ).pack(side='right', padx=5, pady=10)
        
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
        """Abre o módulo de gestão de eventos"""
        try:
            # Primeiro, selecionar um cliente
            if self.selecionar_cliente():
                # Verificar se módulo GestaoEventos já está importado
                if self.gestao_eventos is None:
                    # Tentar importar
                    try:
                        # CORREÇÃO: Importar diretamente do arquivo atual
                        from pagamentos_eventos import GestaoEventos
                        self.gestao_eventos = GestaoEventos(self.parent)
                    except ImportError as ie:
                        messagebox.showerror("Erro", f"Não foi possível importar o módulo de Gestão de Eventos: {str(ie)}")
                        return
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao inicializar Gestão de Eventos: {str(e)}")
                        return
                
                # Abrir a janela de eventos com o cliente selecionado
                if self.gestao_eventos:
                    self.gestao_eventos.abrir_janela_eventos(self.cliente_atual)
                else:
                    messagebox.showerror("Erro", "Módulo de Gestão de Eventos não está disponível")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir gestão de eventos: {str(e)}")
        
    def abrir_gestao_contratos(self):
        """Abre o módulo de gestão de contratos"""
        try:
            # Primeiro, selecionar um cliente
            if self.selecionar_cliente():
                # Importar e instanciar o módulo
                try:
                    from sistema_entrada_dados import GestaoContratos
                    gestao_contratos = GestaoContratos(self.parent)
                    
                    # CORREÇÃO: Usar método atualizado com interface simplificada
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
        
    def selecionar_cliente(self):
        """Abre uma janela para selecionar o cliente e retorna True se selecionado"""
        selecao_janela = tk.Toplevel(self.parent)
        selecao_janela.title("Selecionar Cliente")
        selecao_janela.geometry("400x300")
        selecao_janela.transient(self.parent)
        selecao_janela.grab_set()

        frame = ttk.Frame(selecao_janela, padding=10)
        frame.pack(fill='both', expand=True)

        ttk.Label(frame, text="Selecione o Cliente:").pack(pady=10)

        # Combobox para seleção do cliente
        cliente_var = tk.StringVar()
        cliente_combo = ttk.Combobox(
            frame, 
            textvariable=cliente_var,
            width=40,
            state='readonly'
        )
        cliente_combo.pack(pady=5)

        # Carregar clientes
        clientes = self.carregar_lista_clientes()
        cliente_combo['values'] = clientes

        # Variável para controlar resultado
        self.cliente_selecionado = False

        # Botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.pack(fill='x', pady=20)

        def confirmar_selecao():
            if cliente_var.get():
                self.cliente_atual = cliente_var.get()
                self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
                self.cliente_selecionado = True
                selecao_janela.destroy()
            else:
                messagebox.showwarning("Aviso", "Selecione um cliente!")

        ttk.Button(
            frame_botoes,
            text="Confirmar",
            command=confirmar_selecao,
            width=15
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Cancelar",
            command=selecao_janela.destroy,
            width=15
        ).pack(side='right', padx=5)

        # Centralizar janela
        selecao_janela.update_idletasks()
        width = selecao_janela.winfo_width()
        height = selecao_janela.winfo_height()
        x = (selecao_janela.winfo_screenwidth() // 2) - (width // 2)
        y = (selecao_janela.winfo_screenheight() // 2) - (height // 2)
        selecao_janela.geometry(f'{width}x{height}+{x}+{y}')

        # CORREÇÃO: Esperar a janela fechar e depois retornar o resultado
        self.parent.wait_window(selecao_janela)
        return self.cliente_selecionado

    def carregar_lista_clientes(self):
        """Carrega a lista de clientes disponíveis"""
        try:
            # Verificar se arquivo existe
            if not os.path.exists(ARQUIVO_CLIENTES):
                messagebox.showwarning("Aviso", "Arquivo de clientes não encontrado!")
                return []
                
            workbook = load_workbook(ARQUIVO_CLIENTES)
            sheet = workbook['Clientes']
            clientes = []

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Nome do cliente
                    clientes.append(row[0])

            workbook.close()
            return sorted(clientes)

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
            return []


# Se executado diretamente, abre a janela de controle
if __name__ == "__main__":
    root = tk.Tk()
    root.withdraw()  # Esconde a janela principal
    
    app = ControlePagamentos(root)
    app.abrir_janela_controle()
    
    root.mainloop()