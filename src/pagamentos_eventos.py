import os
import sys
from pathlib import Path
import re
from datetime import datetime
from decimal import Decimal

import tkinter as tk
from tkinter import ttk, messagebox, StringVar
from tkcalendar import DateEntry

from dateutil.relativedelta import relativedelta

import pandas as pd
from openpyxl import load_workbook

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
        BASE_PATH,
        ARQUIVO_FORNECEDORES
    )
    from src.config.utils import formatar_cnpj_cpf, aplicar_formatacao_celula
    from src.config.window_config import configurar_janela
except ImportError as e:
    print(f"Erro ao importar configurações: {str(e)}")
    # Tentar caminho alternativo
    try:
        from config.config import (
            ARQUIVO_CLIENTES,
            ARQUIVO_MODELO,
            PASTA_CLIENTES,
            BASE_PATH,
            ARQUIVO_FORNECEDORES
        )
        from config.utils import formatar_cnpj_cpf, aplicar_formatacao_celula
        from config.window_config import configurar_janela
    except ImportError as e2:
        print(f"Erro ao importar configurações (caminho alternativo): {str(e2)}")
        raise

class GestaoEventos:
    def __init__(self, parent=None):
        self.parent = parent
        self.janela = None
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.contratos = []
        self.eventos = []
        self.administradores_contratos = {}
        self.cliente_selecionado = False  # Variável para controlar resultado da seleção

    def abrir_janela_eventos(self, cliente=None):
        """Abre a janela principal de gestão de eventos para o cliente selecionado"""
        # Se a janela já existir, apenas traz para frente
        if self.janela and self.janela.winfo_exists():
            self.janela.lift()
            self.janela.focus_force()
            return

        # Cria nova janela
        self.janela = tk.Toplevel(self.parent)
        configurar_janela(self.janela, "Gestão de Pagamentos por Eventos", 900, 700)

        # Define o cliente atual
        if cliente:
            self.cliente_atual = cliente
            self.arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
        else:
            # Se não tiver cliente selecionado, solicita seleção
            if not self.selecionar_cliente():
                self.janela.destroy()
                return

        # Frame principal
        frame_principal = ttk.Frame(self.janela, padding=10)
        frame_principal.pack(fill='both', expand=True)

        # Cabeçalho com informações do cliente
        frame_cabecalho = ttk.Frame(frame_principal)
        frame_cabecalho.pack(fill='x', pady=5)

        ttk.Label(
            frame_cabecalho, 
            text=f"Cliente: {self.cliente_atual}", 
            font=('Arial', 12, 'bold')
        ).pack(side='left')

        # Notebook para abas
        self.notebook = ttk.Notebook(frame_principal)
        self.notebook.pack(fill='both', expand=True, pady=10)

        # Aba de contratos
        self.aba_contratos = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_contratos, text="Contratos")

        # Aba de eventos
        self.aba_eventos = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_eventos, text="Eventos")

        # Aba de pagamentos
        self.aba_pagamentos = ttk.Frame(self.notebook)
        self.notebook.add(self.aba_pagamentos, text="Pagamentos")

        # Configurar cada aba
        self.configurar_aba_contratos()
        self.configurar_aba_eventos()
        self.configurar_aba_pagamentos()

        # Botões principais
        frame_botoes = ttk.Frame(frame_principal)
        frame_botoes.pack(fill='x', pady=10)

        ttk.Button(
            frame_botoes,
            text="Fechar",
            command=self.janela.destroy,
            width=15
        ).pack(side='right', padx=5)

        # Carregar dados iniciais
        self.carregar_contratos()

    def selecionar_cliente(self):
        """Abre uma janela para selecionar o cliente e retorna True se selecionado
        CORREÇÃO: Corrigido o problema de looping infinito que ocorria anteriormente"""
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

        # Botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.pack(fill='x', pady=20)

        # CORREÇÃO: Método para confirmar seleção
        def confirmar_selecao():
            if cliente_var.get():
                self.cliente_atual = cliente_var.get()
                self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
                self.cliente_selecionado = True  # Marcar como selecionado
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
            command=lambda: selecao_janela.destroy(),
            width=15
        ).pack(side='right', padx=5)

        # Centralizar janela
        selecao_janela.update_idletasks()
        width = selecao_janela.winfo_width()
        height = selecao_janela.winfo_height()
        x = (selecao_janela.winfo_screenwidth() // 2) - (width // 2)
        y = (selecao_janela.winfo_screenheight() // 2) - (height // 2)
        selecao_janela.geometry(f'{width}x{height}+{x}+{y}')

        # CORREÇÃO: Esperar a janela fechar antes de continuar
        self.parent.wait_window(selecao_janela)
        
        # Retornar se um cliente foi selecionado
        return self.cliente_selecionado

    def carregar_lista_clientes(self):
        """Carrega a lista de clientes disponíveis"""
        try:
            # Verificar se o arquivo existe
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

    def configurar_aba_contratos(self):
        """Configura a aba de gestão de contratos"""
        frame = ttk.Frame(self.aba_contratos, padding=10)
        frame.pack(fill='both', expand=True)

        # Frame para lista de contratos
        frame_lista = ttk.LabelFrame(frame, text="Contratos Disponíveis")
        frame_lista.pack(fill='both', expand=True, pady=5)

        # Treeview para contratos
        colunas = ('Nº Contrato', 'Data Início', 'Data Fim', 'Status', 'Valor Total')
        self.tree_contratos = ttk.Treeview(frame_lista, columns=colunas, show='headings')
        
        for col in colunas:
            self.tree_contratos.heading(col, text=col)
            if col == 'Valor Total':
                self.tree_contratos.column(col, width=120, anchor='e')
            else:
                self.tree_contratos.column(col, width=120)

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_lista, orient='vertical', command=self.tree_contratos.yview)
        scroll_x = ttk.Scrollbar(frame_lista, orient='horizontal', command=self.tree_contratos.xview)
        self.tree_contratos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_contratos.pack(fill='both', expand=True, padx=5, pady=5)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        # Frame para detalhes do contrato selecionado
        frame_detalhes = ttk.LabelFrame(frame, text="Detalhes do Contrato")
        frame_detalhes.pack(fill='x', pady=10)

        frame_detalhes_grid = ttk.Frame(frame_detalhes)
        frame_detalhes_grid.pack(fill='x', padx=5, pady=5)

        # Criar labels e campos para detalhes
        ttk.Label(frame_detalhes_grid, text="Contrato:").grid(row=0, column=0, sticky='w', padx=5, pady=2)
        self.lbl_contrato = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_contrato.grid(row=0, column=1, sticky='w', padx=5, pady=2)

        ttk.Label(frame_detalhes_grid, text="Período:").grid(row=1, column=0, sticky='w', padx=5, pady=2)
        self.lbl_periodo = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_periodo.grid(row=1, column=1, sticky='w', padx=5, pady=2)

        ttk.Label(frame_detalhes_grid, text="Status:").grid(row=2, column=0, sticky='w', padx=5, pady=2)
        self.lbl_status = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_status.grid(row=2, column=1, sticky='w', padx=5, pady=2)

        ttk.Label(frame_detalhes_grid, text="Valor Total:").grid(row=0, column=2, sticky='w', padx=5, pady=2)
        self.lbl_valor_total = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_valor_total.grid(row=0, column=3, sticky='w', padx=5, pady=2)

        ttk.Label(frame_detalhes_grid, text="Administradores:").grid(row=1, column=2, sticky='w', padx=5, pady=2)
        self.lbl_administradores = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_administradores.grid(row=1, column=3, sticky='w', padx=5, pady=2)

        ttk.Label(frame_detalhes_grid, text="Eventos:").grid(row=2, column=2, sticky='w', padx=5, pady=2)
        self.lbl_eventos = ttk.Label(frame_detalhes_grid, text="-")
        self.lbl_eventos.grid(row=2, column=3, sticky='w', padx=5, pady=2)

        # Frame para administradores do contrato
        frame_admins = ttk.LabelFrame(frame, text="Administradores do Contrato")
        frame_admins.pack(fill='both', expand=True, pady=5)

        # Treeview para administradores
        colunas_adm = ('CNPJ/CPF', 'Nome', 'Tipo', 'Valor/Percentual', 'Valor Total', 'Nº Parcelas')
        self.tree_adm = ttk.Treeview(frame_admins, columns=colunas_adm, show='headings', height=4)
        
        for col in colunas_adm:
            self.tree_adm.heading(col, text=col)
            if col in ['Valor/Percentual', 'Valor Total']:
                self.tree_adm.column(col, width=120, anchor='e')
            elif col == 'Nº Parcelas':
                self.tree_adm.column(col, width=80, anchor='center')
            else:
                self.tree_adm.column(col, width=120)

        # Scrollbars
        scroll_y_adm = ttk.Scrollbar(frame_admins, orient='vertical', command=self.tree_adm.yview)
        scroll_x_adm = ttk.Scrollbar(frame_admins, orient='horizontal', command=self.tree_adm.xview)
        self.tree_adm.configure(yscrollcommand=scroll_y_adm.set, xscrollcommand=scroll_x_adm.set)
        
        self.tree_adm.pack(fill='both', expand=True, padx=5, pady=5)
        scroll_y_adm.pack(side='right', fill='y')
        scroll_x_adm.pack(side='bottom', fill='x')

        # Botões de ação
        frame_botoes = ttk.Frame(frame)
        frame_botoes.pack(fill='x', pady=5)

        ttk.Button(
            frame_botoes,
            text="Definir Eventos",
            command=self.definir_eventos,
            width=20
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Atualizar Eventos",
            command=self.atualizar_eventos_contrato,
            width=20
        ).pack(side='left', padx=5)

        # Binding para mostrar detalhes ao selecionar contrato
        self.tree_contratos.bind('<<TreeviewSelect>>', self.mostrar_detalhes_contrato)

    def configurar_aba_eventos(self):
        """Configura a aba de gestão de eventos"""
        frame = ttk.Frame(self.aba_eventos, padding=10)
        frame.pack(fill='both', expand=True)

        # Frame para contrato selecionado
        frame_contrato = ttk.LabelFrame(frame, text="Contrato")
        frame_contrato.pack(fill='x', pady=5)

        ttk.Label(frame_contrato, text="Contrato Selecionado:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.contrato_selecionado = ttk.Combobox(frame_contrato, state='readonly', width=40)
        self.contrato_selecionado.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.contrato_selecionado.bind('<<ComboboxSelected>>', self.carregar_eventos_contrato)

        # Frame para botão de adicionar evento
        frame_adicionar = ttk.Frame(frame)
        frame_adicionar.pack(fill='x', pady=5)

        ttk.Button(
            frame_adicionar,
            text="Adicionar Evento",
            command=self.adicionar_evento,
            width=20
        ).pack(side='left', padx=5)

        # Frame para lista de eventos
        frame_eventos = ttk.LabelFrame(frame, text="Eventos do Contrato")
        frame_eventos.pack(fill='both', expand=True, pady=5)

        # Treeview para eventos
        colunas = ('ID', 'Descrição', 'Percentual', 'Valor', 'Status', 'Data Conclusão')
        self.tree_eventos = ttk.Treeview(frame_eventos, columns=colunas, show='headings')
        
        for col in colunas:
            self.tree_eventos.heading(col, text=col)
            if col == 'ID':
                self.tree_eventos.column(col, width=50, anchor='center')
            elif col in ['Percentual', 'Valor']:
                self.tree_eventos.column(col, width=100, anchor='e')
            elif col == 'Status':
                self.tree_eventos.column(col, width=100, anchor='center')
            else:
                self.tree_eventos.column(col, width=150)

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_eventos, orient='vertical', command=self.tree_eventos.yview)
        scroll_x = ttk.Scrollbar(frame_eventos, orient='horizontal', command=self.tree_eventos.xview)
        self.tree_eventos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_eventos.pack(fill='both', expand=True, padx=5, pady=5)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        # Frame para detalhes do evento selecionado
        frame_detalhes = ttk.LabelFrame(frame, text="Detalhes do Evento")
        frame_detalhes.pack(fill='x', pady=10)

        # Grid para formulário
        frame_form = ttk.Frame(frame_detalhes)
        frame_form.pack(fill='x', padx=5, pady=5)

        ttk.Label(frame_form, text="Descrição:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.evento_descricao = ttk.Entry(frame_form, width=50)
        self.evento_descricao.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Percentual (%):").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.evento_percentual = ttk.Entry(frame_form, width=15)
        self.evento_percentual.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Status:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.evento_status = ttk.Combobox(frame_form, values=['pendente', 'concluido'], state='readonly', width=15)
        self.evento_status.grid(row=2, column=1, padx=5, pady=2, sticky='w')
        self.evento_status.set('pendente')

        ttk.Label(frame_form, text="Data Conclusão:").grid(row=3, column=0, padx=5, pady=2, sticky='w')
        self.evento_data = DateEntry(frame_form, width=15, state='normal', date_pattern='dd/mm/yyyy')
        self.evento_data.grid(row=3, column=1, padx=5, pady=2, sticky='w')

        # Botões de ação para eventos
        frame_botoes = ttk.Frame(frame_detalhes)
        frame_botoes.pack(fill='x', pady=5)

        ttk.Button(
            frame_botoes,
            text="Salvar Evento",
            command=self.salvar_evento,
            width=15
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Marcar como Concluído",
            command=self.concluir_evento,
            width=20
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Remover Evento",
            command=self.remover_evento,
            width=15
        ).pack(side='left', padx=5)

        # Binding para seleção de evento
        self.tree_eventos.bind('<<TreeviewSelect>>', self.mostrar_detalhes_evento)

    def configurar_aba_pagamentos(self):
        """Configura a aba de pagamentos de eventos"""
        frame = ttk.Frame(self.aba_pagamentos, padding=10)
        frame.pack(fill='both', expand=True)

        # Frame para controle de filtros
        frame_filtros = ttk.LabelFrame(frame, text="Filtros")
        frame_filtros.pack(fill='x', pady=5)

        ttk.Label(frame_filtros, text="Contrato:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.pagto_contrato = ttk.Combobox(frame_filtros, state='readonly', width=40)
        self.pagto_contrato.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        self.pagto_contrato.bind('<<ComboboxSelected>>', self.filtrar_pagamentos)

        ttk.Label(frame_filtros, text="Status:").grid(row=0, column=2, padx=5, pady=5, sticky='w')
        self.pagto_status = ttk.Combobox(frame_filtros, values=['Todos', 'Pendente', 'Pago'], state='readonly', width=15)
        self.pagto_status.grid(row=0, column=3, padx=5, pady=5, sticky='w')
        self.pagto_status.set('Todos')
        self.pagto_status.bind('<<ComboboxSelected>>', self.filtrar_pagamentos)

        # Frame para lista de pagamentos
        frame_pagamentos = ttk.LabelFrame(frame, text="Pagamentos Previstos")
        frame_pagamentos.pack(fill='both', expand=True, pady=5)

        # Treeview para pagamentos
        colunas = ('Contrato', 'Evento', 'CNPJ/CPF', 'Nome', 'Data Vencimento', 'Valor', 'Status', 'Data Pagamento')
        self.tree_pagamentos = ttk.Treeview(frame_pagamentos, columns=colunas, show='headings')
        
        for col in colunas:
            self.tree_pagamentos.heading(col, text=col)
            if col in ['Data Vencimento', 'Data Pagamento']:
                self.tree_pagamentos.column(col, width=120, anchor='center')
            elif col == 'Valor':
                self.tree_pagamentos.column(col, width=100, anchor='e')
            elif col == 'Status':
                self.tree_pagamentos.column(col, width=100, anchor='center')
            elif col in ['Contrato', 'Evento']:
                self.tree_pagamentos.column(col, width=80, anchor='center')
            else:
                self.tree_pagamentos.column(col, width=150)

        # Scrollbars
        scroll_y = ttk.Scrollbar(frame_pagamentos, orient='vertical', command=self.tree_pagamentos.yview)
        scroll_x = ttk.Scrollbar(frame_pagamentos, orient='horizontal', command=self.tree_pagamentos.xview)
        self.tree_pagamentos.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        
        self.tree_pagamentos.pack(fill='both', expand=True, padx=5, pady=5)
        scroll_y.pack(side='right', fill='y')
        scroll_x.pack(side='bottom', fill='x')

        # Frame para detalhes do pagamento
        frame_detalhes = ttk.LabelFrame(frame, text="Detalhes do Pagamento")
        frame_detalhes.pack(fill='x', pady=10)

        # Grid para formulário
        frame_form = ttk.Frame(frame_detalhes)
        frame_form.pack(fill='x', padx=5, pady=5)

        ttk.Label(frame_form, text="CNPJ/CPF:").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        self.pagto_cnpj = ttk.Entry(frame_form, width=20, state='readonly')
        self.pagto_cnpj.grid(row=0, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Nome:").grid(row=0, column=2, padx=5, pady=2, sticky='w')
        self.pagto_nome = ttk.Entry(frame_form, width=40, state='readonly')
        self.pagto_nome.grid(row=0, column=3, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Valor:").grid(row=1, column=0, padx=5, pady=2, sticky='w')
        self.pagto_valor = ttk.Entry(frame_form, width=20, state='readonly')
        self.pagto_valor.grid(row=1, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Vencimento:").grid(row=1, column=2, padx=5, pady=2, sticky='w')
        self.pagto_vencimento = ttk.Entry(frame_form, width=20, state='readonly')
        self.pagto_vencimento.grid(row=1, column=3, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Status:").grid(row=2, column=0, padx=5, pady=2, sticky='w')
        self.pagto_status_atual = ttk.Combobox(frame_form, values=['Pendente', 'Pago'], state='readonly', width=15)
        self.pagto_status_atual.grid(row=2, column=1, padx=5, pady=2, sticky='w')

        ttk.Label(frame_form, text="Data Pagamento:").grid(row=2, column=2, padx=5, pady=2, sticky='w')
        self.pagto_data = DateEntry(frame_form, width=15, state='normal', date_pattern='dd/mm/yyyy')
        self.pagto_data.grid(row=2, column=3, padx=5, pady=2, sticky='w')

        # Botões de ação para pagamentos
        frame_botoes = ttk.Frame(frame_detalhes)
        frame_botoes.pack(fill='x', pady=5)

        ttk.Button(
            frame_botoes,
            text="Registrar Pagamento",
            command=self.registrar_pagamento,
            width=20
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Gerar Lançamento",
            command=self.gerar_lancamento_pagamento,
            width=20
        ).pack(side='left', padx=5)

        # Binding para seleção de pagamento
        self.tree_pagamentos.bind('<<TreeviewSelect>>', self.mostrar_detalhes_pagamento)
        
    def carregar_contratos(self):
        """Carrega contratos do cliente atual"""
        try:
            if not self.arquivo_cliente or not self.cliente_atual:
                return

            wb = load_workbook(self.arquivo_cliente, data_only=True)
            
            # Verificar se a aba Contratos_ADM existe
            if 'Contratos_ADM' not in wb.sheetnames:
                messagebox.showinfo("Informação", "Este cliente não possui contratos de administração cadastrados.")
                wb.close()
                return
                
            ws = wb['Contratos_ADM']
            
            # Limpar dados anteriores
            self.contratos = []
            self.administradores_contratos = {}
            
            for item in self.tree_contratos.get_children():
                self.tree_contratos.delete(item)

            # Processar contratos
            contratos_vistos = set()
            
            for row in ws.iter_rows(min_row=3, values_only=True):
                num_contrato = row[0]
                if num_contrato and num_contrato not in contratos_vistos:
                    contratos_vistos.add(num_contrato)
                    
                    # Formatar datas
                    data_inicio = datetime.strptime(row[1].strftime('%d/%m/%Y'), '%d/%m/%Y') if isinstance(row[1], datetime) else row[1]
                    data_fim = datetime.strptime(row[2].strftime('%d/%m/%Y'), '%d/%m/%Y') if isinstance(row[2], datetime) else row[2]
                    
                    status = row[3] or 'ATIVO'
                    
                    # Calcular valor total do contrato
                    valor_total = self.calcular_valor_total_contrato(ws, num_contrato)
                    
                    # Armazenar informações do contrato
                    contrato_info = {
                        'num_contrato': num_contrato,
                        'data_inicio': data_inicio,
                        'data_fim': data_fim,
                        'status': status,
                        'valor_total': valor_total,
                        'administradores': []
                    }
                    
                    self.contratos.append(contrato_info)
                    
                    # Adicionar à Treeview
                    data_inicio_str = data_inicio.strftime('%d/%m/%Y') if isinstance(data_inicio, datetime) else str(data_inicio)
                    data_fim_str = data_fim.strftime('%d/%m/%Y') if isinstance(data_fim, datetime) else str(data_fim)
                    
                    self.tree_contratos.insert('', 'end', values=(
                        num_contrato,
                        data_inicio_str,
                        data_fim_str,
                        status,
                        f"R$ {valor_total:,.2f}" if valor_total else "-"
                    ))

            # Processar administradores de contratos
            for row in ws.iter_rows(min_row=3, values_only=True):
                # Verificar se é um registro de administrador (coluna G em diante)
                if row[6] and row[7] and row[8]:  # Tem número contrato, CNPJ/CPF e nome
                    num_contrato = row[6]
                    
                    # Adicionamos o administrador ao dicionário
                    if num_contrato not in self.administradores_contratos:
                        self.administradores_contratos[num_contrato] = []
                    
                    # Formatar CNPJ/CPF
                    cnpj_cpf = formatar_cnpj_cpf(str(row[7]))
                    
                    # Adicionar administrador
                    self.administradores_contratos[num_contrato].append({
                        'cnpj_cpf': cnpj_cpf,
                        'nome': row[8],
                        'tipo': row[9] or 'Fixo',
                        'valor_percentual': row[10] or '0',
                        'valor_total': row[11] or '0',
                        'num_parcelas': row[12] or '1'
                    })
            
            # Atualizar comboboxes
            self.atualizar_comboboxes_contratos()
            
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar contratos: {str(e)}")
            if 'wb' in locals():
                wb.close()

    def atualizar_comboboxes_contratos(self):
        """Atualiza as comboboxes com os contratos disponíveis"""
        contratos_ativos = [c['num_contrato'] for c in self.contratos if c['status'] == 'ATIVO']
        
        # Atualizar combobox na aba de eventos
        self.contrato_selecionado['values'] = contratos_ativos
        if contratos_ativos:
            self.contrato_selecionado.set(contratos_ativos[0])
            # Carregar eventos do primeiro contrato
            self.carregar_eventos_contrato(None)
        
    def carregar_eventos_contrato(self, event):
        """Carrega os eventos do contrato selecionado"""
        num_contrato = self.contrato_selecionado.get()
        if not num_contrato:
            return
            
        try:
            wb = load_workbook(self.arquivo_cliente, data_only=True)
            
            # Verificar se existe a aba Contratos_ADM
            if 'Contratos_ADM' not in wb.sheetnames:
                messagebox.showinfo("Informação", "Este cliente não possui contratos de administração cadastrados.")
                wb.close()
                return
                
            ws = wb['Contratos_ADM']
            
            # Limpar treeview
            for item in self.tree_eventos.get_children():
                self.tree_eventos.delete(item)
                
            # Limpar lista de eventos
            self.eventos = []
            
            # Buscar valor total do contrato para cálculos percentuais
            valor_contrato = None
            for contrato in self.contratos:
                if contrato['num_contrato'] == num_contrato:
                    valor_contrato = contrato['valor_total']
                    break
            
            # Processar eventos
            for row in ws.iter_rows(min_row=3, values_only=True):
                # Verificar se é um registro de evento (coluna AE ou 31 em diante)
                if len(row) >= 33 and row[30] == num_contrato:
                    evento_id = row[31]
                    descricao = row[32]
                    percentual = row[33]
                    status = row[34] or 'pendente'
                    
                    # Calcular valor baseado no percentual
                    valor = None
                    if percentual and valor_contrato:
                        try:
                            perc = float(str(percentual).replace('%', '').replace(',', '.'))
                            valor = (perc / 100) * valor_contrato
                        except (ValueError, TypeError):
                            valor = 0
                    
                    # Formatar data de conclusão
                    data_conclusao = None
                    if row[35]:  # Data conclusão
                        if isinstance(row[35], datetime):
                            data_conclusao = row[35].strftime('%d/%m/%Y')
                        else:
                            try:
                                data_conclusao = datetime.strptime(str(row[35]), '%Y-%m-%d').strftime('%d/%m/%Y')
                            except ValueError:
                                data_conclusao = str(row[35])
                    
                    # Adicionar à lista de eventos
                    evento = {
                        'id': evento_id,
                        'descricao': descricao,
                        'percentual': percentual,
                        'valor': valor,
                        'status': status,
                        'data_conclusao': data_conclusao
                    }
                    self.eventos.append(evento)
                    
                    # Adicionar ao treeview
                    valor_fmt = f"R$ {valor:,.2f}" if valor else "-"
                    percentual_fmt = f"{percentual}" if percentual else "-"
                    
                    self.tree_eventos.insert('', 'end', values=(
                        evento_id,
                        descricao,
                        percentual_fmt,
                        valor_fmt,
                        status.capitalize(),
                        data_conclusao or "-"
                    ))
            
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar eventos: {str(e)}")
            if 'wb' in locals():
                wb.close()
                
    def adicionar_evento(self):
        """Adiciona um novo evento ao contrato selecionado"""
        num_contrato = self.contrato_selecionado.get()
        if not num_contrato:
            messagebox.showwarning("Aviso", "Selecione um contrato primeiro")
            return
            
        # Limpar campos do formulário
        self.evento_descricao.delete(0, tk.END)
        self.evento_percentual.delete(0, tk.END)
        self.evento_status.set('pendente')
        self.evento_data.set_date(datetime.now())
        
        # Definir ID do próximo evento
        proximo_id = 1
        if self.eventos:
            ids = [int(e['id']) for e in self.eventos if str(e['id']).isdigit()]
            if ids:
                proximo_id = max(ids) + 1
                
        # Mostrar janela para adicionar evento
        janela = tk.Toplevel(self.janela)
        janela.title("Adicionar Evento")
        janela.geometry("500x300")
        janela.transient(self.janela)
        janela.grab_set()
        
        frame = ttk.Frame(janela, padding=10)
        frame.pack(fill='both', expand=True)
        
        # Formulário
        ttk.Label(frame, text="Contrato:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Label(frame, text=num_contrato, font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Label(frame, text="ID do Evento:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        ttk.Label(frame, text=str(proximo_id), font=('Arial', 10, 'bold')).grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        ttk.Label(frame, text="Descrição:*").grid(row=2, column=0, padx=5, pady=5, sticky='w')
        descricao_entry = ttk.Entry(frame, width=40)
        descricao_entry.grid(row=2, column=1, padx=5, pady=5, sticky='ew')
        
        ttk.Label(frame, text="Percentual (%):*").grid(row=3, column=0, padx=5, pady=5, sticky='w')
        percentual_entry = ttk.Entry(frame, width=15)
        percentual_entry.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # Botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.grid(row=4, column=0, columnspan=2, pady=20)
        
        def confirmar():
            # Validar campos
            descricao = descricao_entry.get().strip()
            percentual_str = percentual_entry.get().strip()
            
            if not descricao:
                messagebox.showerror("Erro", "Descrição é obrigatória")
                return
                
            if not percentual_str:
                messagebox.showerror("Erro", "Percentual é obrigatório")
                return
                
            try:
                percentual = float(percentual_str.replace(',', '.'))
                if percentual <= 0 or percentual > 100:
                    messagebox.showerror("Erro", "Percentual deve estar entre 0 e 100")
                    return
            except ValueError:
                messagebox.showerror("Erro", "Percentual inválido")
                return
                
            # Calcular total de percentuais já existentes
            percentual_atual = sum([
                float(str(e['percentual']).replace('%', '').replace(',', '.'))
                for e in self.eventos
                if e['percentual'] and str(e['percentual']).replace('%', '').replace(',', '.').replace('.', '').isdigit()
            ])
            
            # Verificar se excede 100%
            if percentual_atual + percentual > 100:
                messagebox.showerror(
                    "Erro", 
                    f"Total de percentuais excede 100%! Atual: {percentual_atual:.2f}%, Novo: {percentual:.2f}%"
                )
                return
                
            # Salvar evento
            try:
                self.salvar_novo_evento(num_contrato, proximo_id, descricao, percentual)
                messagebox.showinfo("Sucesso", "Evento adicionado com sucesso!")
                janela.destroy()
                # Atualizar lista de eventos
                self.carregar_eventos_contrato(None)
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar evento: {str(e)}")
        
        ttk.Button(frame_botoes, text="Confirmar", command=confirmar, width=15).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Cancelar", command=janela.destroy, width=15).pack(side='left', padx=5)
        
        # Centralizar janela
        janela.update_idletasks()
        width = janela.winfo_width()
        height = janela.winfo_height()
        x = (janela.winfo_screenwidth() // 2) - (width // 2)
        y = (janela.winfo_screenheight() // 2) - (height // 2)
        janela.geometry(f'{width}x{height}+{x}+{y}')
        
    def salvar_novo_evento(self, num_contrato, evento_id, descricao, percentual):
        """Salva um novo evento na planilha"""
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            # Encontrar próxima linha disponível
            proxima_linha = ws.max_row + 1
            
            # Salvar evento
            ws.cell(row=proxima_linha, column=31, value=num_contrato)  # Contrato
            ws.cell(row=proxima_linha, column=32, value=evento_id)     # ID Evento
            ws.cell(row=proxima_linha, column=33, value=descricao)     # Descrição
            ws.cell(row=proxima_linha, column=34, value=f"{percentual:.2f}%")  # Percentual
            ws.cell(row=proxima_linha, column=35, value="pendente")    # Status
            
            wb.save(self.arquivo_cliente)
            
            # Atualizar lista interna de eventos
            valor_contrato = None
            for contrato in self.contratos:
                if contrato['num_contrato'] == num_contrato:
                    valor_contrato = contrato['valor_total']
                    break
                    
            valor = None
            if valor_contrato:
                valor = (percentual / 100) * valor_contrato
                
            evento = {
                'id': evento_id,
                'descricao': descricao,
                'percentual': f"{percentual:.2f}%",
                'valor': valor,
                'status': 'pendente',
                'data_conclusao': None
            }
            self.eventos.append(evento)
            
        except Exception as e:
            raise Exception(f"Erro ao salvar evento: {str(e)}")
            
    def mostrar_detalhes_evento(self, event):
        """Mostra detalhes do evento selecionado"""
        selecionado = self.tree_eventos.selection()
        if not selecionado:
            return
            
        valores = self.tree_eventos.item(selecionado)['values']
        evento_id = valores[0]
        
        # Buscar o evento na lista
        evento = next((e for e in self.eventos if str(e['id']) == str(evento_id)), None)
        if not evento:
            return
            
        # Preencher campos
        self.evento_descricao.delete(0, tk.END)
        self.evento_descricao.insert(0, evento['descricao'])
        
        self.evento_percentual.delete(0, tk.END)
        percentual_str = str(evento['percentual']).replace('%', '').strip()
        self.evento_percentual.insert(0, percentual_str)
        
        self.evento_status.set(evento['status'])
        
        if evento['data_conclusao']:
            try:
                data = datetime.strptime(evento['data_conclusao'], '%d/%m/%Y')
                self.evento_data.set_date(data)
            except ValueError:
                self.evento_data.set_date(datetime.now())

    def calcular_valor_total_contrato(self, ws, num_contrato):
        """Calcula o valor total do contrato somando os valores dos administradores"""
        valor_total = 0
        
        for row in ws.iter_rows(min_row=3, values_only=True):
            if row[6] == num_contrato and row[9] == 'Fixo' and row[11]:  # Verifica se é administrador fixo com valor total
                try:
                    valor_total += float(str(row[11]).replace('.', '').replace(',', '.'))
                except (ValueError, TypeError):
                    pass
        
        return valor_total if valor_total > 0 else None
    
    def mostrar_detalhes_pagamento(self, event):
        """Mostra detalhes do pagamento selecionado"""
        selecionado = self.tree_pagamentos.selection()
        if not selecionado:
            return
            
        valores = self.tree_pagamentos.item(selecionado)['values']
        
        # Preencher campos
        self.pagto_cnpj.config(state='normal')
        self.pagto_cnpj.delete(0, tk.END)
        self.pagto_cnpj.insert(0, valores[2])
        self.pagto_cnpj.config(state='readonly')
        
        self.pagto_nome.config(state='normal')
        self.pagto_nome.delete(0, tk.END)
        self.pagto_nome.insert(0, valores[3])
        self.pagto_nome.config(state='readonly')
        
        self.pagto_valor.config(state='normal')
        self.pagto_valor.delete(0, tk.END)
        self.pagto_valor.insert(0, valores[5])
        self.pagto_valor.config(state='readonly')
        
        self.pagto_vencimento.config(state='normal')
        self.pagto_vencimento.delete(0, tk.END)
        self.pagto_vencimento.insert(0, valores[4])
        self.pagto_vencimento.config(state='readonly')
        
        # Status
        self.pagto_status_atual.set(valores[6])
        
        # Data de pagamento
        if valores[7] != "-":
            try:
                data = datetime.strptime(valores[7], '%d/%m/%Y')
                self.pagto_data.set_date(data)
            except ValueError:
                self.pagto_data.set_date(datetime.now())

    def contar_eventos_contrato(self, num_contrato):
        """Conta o número de eventos cadastrados para o contrato"""
        try:
            wb = load_workbook(self.arquivo_cliente, data_only=True)
            ws = wb['Contratos_ADM']
            
            eventos_count = 0
            for row in ws.iter_rows(min_row=3, values_only=True):
                # Verifica registros na área de eventos (a partir da coluna AE)
                if len(row) >= 33 and row[30] == num_contrato:
                    eventos_count += 1
            
            wb.close()
            return eventos_count
            
        except Exception as e:
            print(f"Erro ao contar eventos: {str(e)}")
            if 'wb' in locals():
                wb.close()
            return 0

    def definir_eventos(self):
        """Abre a tela para definir eventos do contrato selecionado"""
        selecionado = self.tree_contratos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um contrato primeiro")
            return
            
        valores = self.tree_contratos.item(selecionado)['values']
        num_contrato = valores[0]
        
        # Verificar status do contrato
        status = valores[3]
        if status != 'ATIVO':
            messagebox.showwarning("Aviso", "Apenas contratos ativos podem ter eventos definidos")
            return
            
        # Mudar para a aba de eventos
        self.notebook.select(self.aba_eventos)
        
        # Selecionar o contrato na combobox
        self.contrato_selecionado.set(num_contrato)
        
        # Carregar eventos do contrato
        self.carregar_eventos_contrato(None)        

    def atualizar_eventos_contrato(self):
        """Atualiza a lista de eventos do contrato selecionado"""
        selecionado = self.tree_contratos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um contrato primeiro")
            return
            
        valores = self.tree_contratos.item(selecionado)['values']
        num_contrato = valores[0]
        
        # Verificar status do contrato
        status = valores[3]
        if status != 'ATIVO':
            messagebox.showwarning("Aviso", "Apenas contratos ativos podem ter eventos atualizados")
            return
            
        # Carregar eventos do contrato novamente
        self.notebook.select(self.aba_eventos)
        self.contrato_selecionado.set(num_contrato)
        self.carregar_eventos_contrato(None), '').replace('.', '').replace(',', '.').strip()
        
        # Converter valor
        try:
            valor = float(valor_str)
        except ValueError:
            messagebox.showerror("Erro", "Valor do pagamento inválido")
            return
            
        # Converter data
        try:
            data_vencto = datetime.strptime(data_vencimento, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Erro", "Data de vencimento inválida")
            return
        
        # Buscar administrador para obter categoria
        categoria = "ADM"
        
        # Verificar se o Sistema de Entrada de Dados está disponível no parent
        if hasattr(self.parent, 'dados_para_incluir'):
            # Preparar lançamento
            lancamento = {
                'data': self.calcular_data_relatorio(data_vencto),
                'tp_desp': '3',  # Tipo 3 para administração
                'cnpj_cpf': cnpj_cpf,
                'nome': nome,
                'referencia': f"ADM OBRA - EVENTO {evento_id} - CONTRATO {num_contrato}",
                'nf': '',
                'vr_unit': f"{valor:.2f}",
                'dias': 1,
                'valor': f"{valor:.2f}",
                'dt_vencto': data_vencimento,
                'categoria': categoria,
                'dados_bancarios': self.buscar_dados_bancarios(cnpj_cpf),
                'observacao': f"PAGAMENTO EVENTO {evento_id} - CONTRATO {num_contrato}",
                'forma_pagamento': 'PIX'  # Valor padrão
            }
            
            # Adicionar à lista do sistema
            self.parent.dados_para_incluir.append(lancamento)
            
            messagebox.showinfo(
                "Sucesso", 
                "Lançamento gerado com sucesso! Acesse a aba de 'Visualização de Lançamentos' para conferir."
            )
            
            # Fechar esta janela
            self.janela.destroy()
            
            # Mostrar visualizador de lançamentos
            if hasattr(self.parent, 'visualizar_lancamentos'):
                self.parent.visualizar_lancamentos()
        else:
            messagebox.showwarning(
                "Aviso", 
                "Sistema de Entrada de Dados não está disponível. O lançamento não pôde ser gerado."
            )

    def salvar_evento(self):
        """Salva as alterações em um evento existente"""
        selecionado = self.tree_eventos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um evento primeiro")
            return
            
        valores = self.tree_eventos.item(selecionado)['values']
        evento_id = valores[0]
        num_contrato = self.contrato_selecionado.get()
        
        # Validar campos
        descricao = self.evento_descricao.get().strip()
        percentual_str = self.evento_percentual.get().strip()
        status = self.evento_status.get()
        
        if not descricao:
            messagebox.showerror("Erro", "Descrição é obrigatória")
            return
            
        if not percentual_str:
            messagebox.showerror("Erro", "Percentual é obrigatório")
            return
            
        try:
            percentual = float(percentual_str.replace(',', '.'))
            if percentual <= 0 or percentual > 100:
                messagebox.showerror("Erro", "Percentual deve estar entre 0 e 100")
                return
        except ValueError:
            messagebox.showerror("Erro", "Percentual inválido")
            return
            
        # Calcular total de percentuais já existentes (excluindo o evento atual)
        percentual_atual = sum([
            float(str(e['percentual']).replace('%', '').replace(',', '.'))
            for e in self.eventos
            if e['percentual'] and str(e['id']) != str(evento_id) and 
            str(e['percentual']).replace('%', '').replace(',', '.').replace('.', '').isdigit()
        ])
        
        # Verificar se excede 100%
        if percentual_atual + percentual > 100:
            messagebox.showerror(
                "Erro", 
                f"Total de percentuais excede 100%! Outros eventos: {percentual_atual:.2f}%, Este evento: {percentual:.2f}%"
            )
            return
            
        # Atualizar evento
        try:
            # Formatar data
            data_conclusao = self.evento_data.get_date().strftime('%d/%m/%Y') if status == 'concluido' else None
            
            self.atualizar_evento(num_contrato, evento_id, descricao, percentual, status, data_conclusao)
            messagebox.showinfo("Sucesso", "Evento atualizado com sucesso!")
            
            # Atualizar lista de eventos
            self.carregar_eventos_contrato(None)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar evento: {str(e)}")
            
    def atualizar_evento(self, num_contrato, evento_id, descricao, percentual, status, data_conclusao):
        """Atualiza um evento existente na planilha"""
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            # Encontrar o evento
            evento_encontrado = False
            for row in range(3, ws.max_row + 1):
                if (ws.cell(row=row, column=31).value == num_contrato and 
                    str(ws.cell(row=row, column=32).value) == str(evento_id)):
                    
                    # Atualizar valores
                    ws.cell(row=row, column=33, value=descricao)         # Descrição
                    ws.cell(row=row, column=34, value=f"{percentual:.2f}%")  # Percentual
                    ws.cell(row=row, column=35, value=status)            # Status
                    
                    # Data de conclusão apenas se status for 'concluido'
                    if status == 'concluido' and data_conclusao:
                        data_obj = datetime.strptime(data_conclusao, '%d/%m/%Y')
                        ws.cell(row=row, column=36, value=data_obj)      # Data conclusão
                    elif status != 'concluido':
                        ws.cell(row=row, column=36, value=None)          # Limpar data
                        
                    evento_encontrado = True
                    break
                    
            if not evento_encontrado:
                raise Exception(f"Evento {evento_id} não encontrado para o contrato {num_contrato}")
                
            wb.save(self.arquivo_cliente)
            
            # Se o evento foi concluído, gerar pagamento
            if status == 'concluido':
                self.gerar_pagamento_evento(num_contrato, evento_id, descricao, percentual, data_conclusao)
                
        except Exception as e:
            raise Exception(f"Erro ao atualizar evento: {str(e)}")

    def concluir_evento(self):
        """Marca um evento como concluído e gera pagamento"""
        selecionado = self.tree_eventos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um evento primeiro")
            return
            
        valores = self.tree_eventos.item(selecionado)['values']
        evento_id = valores[0]
        num_contrato = self.contrato_selecionado.get()
        
        # Verificar se o evento já está concluído
        if valores[4] == 'Concluido':
            messagebox.showinfo("Informação", "Este evento já está concluído")
            return
            
        # Confirmar conclusão
        if not messagebox.askyesno("Confirmar", "Confirma a conclusão deste evento? Isso gerará pagamentos para os administradores."):
            return
            
        # Buscar detalhes do evento
        evento = next((e for e in self.eventos if str(e['id']) == str(evento_id)), None)
        if not evento:
            messagebox.showerror("Erro", "Evento não encontrado")
            return
            
        # Marcar como concluído
        try:
            # Obter data atual
            data_conclusao = datetime.now().strftime('%d/%m/%Y')
            
            # Atualizar campos
            self.evento_status.set('concluido')
            self.evento_data.set_date(datetime.now())
            
            # Atualizar evento
            self.atualizar_evento(
                num_contrato, 
                evento_id, 
                evento['descricao'], 
                float(str(evento['percentual']).replace('%', '').replace(',', '.')), 
                'concluido', 
                data_conclusao
            )
            
            messagebox.showinfo("Sucesso", "Evento concluído com sucesso! Pagamentos foram gerados.")
            
            # Atualizar lista de eventos
            self.carregar_eventos_contrato(None)
            
            # Ir para aba de pagamentos
            self.notebook.select(self.aba_pagamentos)
            self.pagto_contrato.set(num_contrato)
            self.filtrar_pagamentos(None)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao concluir evento: {str(e)}")

    def gerar_pagamento_evento(self, num_contrato, evento_id, descricao, percentual, data_conclusao):
        """Gera pagamentos para os administradores baseado no evento concluído"""
        try:
            # Buscar valor do contrato
            valor_contrato = None
            for contrato in self.contratos:
                if contrato['num_contrato'] == num_contrato:
                    valor_contrato = contrato['valor_total']
                    break
                    
            if not valor_contrato:
                raise Exception(f"Não foi possível encontrar o valor total do contrato {num_contrato}")
                
            # Buscar administradores do contrato
            administradores = self.administradores_contratos.get(num_contrato, [])
            if not administradores:
                raise Exception(f"Não há administradores cadastrados para o contrato {num_contrato}")
                
            # Valor do evento
            valor_evento = (percentual / 100) * valor_contrato
            
            # Determinar data de vencimento (15 dias após conclusão)
            data_conclusao_obj = datetime.strptime(data_conclusao, '%d/%m/%Y')
            data_vencimento = data_conclusao_obj + relativedelta(days=15)
            
            # Salvar pagamentos para cada administrador
            wb = load_workbook(self.arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            for admin in administradores:
                # Calcular valor do pagamento
                if admin['tipo'] == 'Percentual':
                    # Administrador com percentual sobre o contrato
                    try:
                        admin_percentual = float(str(admin['valor_percentual']).replace('%', '').replace(',', '.'))
                        valor_pagamento = (admin_percentual / 100) * valor_evento
                    except (ValueError, TypeError):
                        valor_pagamento = 0
                else:
                    # Administrador com valor fixo - calcula proporcional ao evento
                    try:
                        valor_total_admin = float(str(admin['valor_total']).replace('.', '').replace(',', '.'))
                        valor_pagamento = (percentual / 100) * valor_total_admin
                    except (ValueError, TypeError):
                        valor_pagamento = 0
                
                # Encontrar próxima linha disponível para pagamentos
                proxima_linha = ws.max_row + 1
                
                # Salvar pagamento
                ws.cell(row=proxima_linha, column=38, value=num_contrato)          # Contrato
                ws.cell(row=proxima_linha, column=39, value=evento_id)             # ID Evento
                ws.cell(row=proxima_linha, column=40, value=admin['cnpj_cpf'])     # CNPJ/CPF
                ws.cell(row=proxima_linha, column=41, value=admin['nome'])         # Nome
                ws.cell(row=proxima_linha, column=42, value=data_vencimento)       # Data Vencimento
                ws.cell(row=proxima_linha, column=43, value=valor_pagamento)       # Valor
                ws.cell(row=proxima_linha, column=44, value='pendente')            # Status
                ws.cell(row=proxima_linha, column=45, value=f"Pagamento referente ao evento: {descricao}")  # Observação
            
            wb.save(self.arquivo_cliente)
            
        except Exception as e:
            raise Exception(f"Erro ao gerar pagamentos: {str(e)}")
            
    def remover_evento(self):
        """Remove um evento não concluído"""
        selecionado = self.tree_eventos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um evento primeiro")
            return
            
        valores = self.tree_eventos.item(selecionado)['values']
        evento_id = valores[0]
        num_contrato = self.contrato_selecionado.get()
        
        # Verificar se o evento já está concluído
        if valores[4] == 'Concluido':
            messagebox.showerror("Erro", "Não é possível remover um evento já concluído")
            return
            
        # Confirmar remoção
        if not messagebox.askyesno("Confirmar", "Tem certeza que deseja remover este evento?"):
            return
            
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            # Encontrar o evento
            linha_evento = None
            for row in range(3, ws.max_row + 1):
                if (ws.cell(row=row, column=31).value == num_contrato and 
                    str(ws.cell(row=row, column=32).value) == str(evento_id)):
                    linha_evento = row
                    break
                    
            if not linha_evento:
                raise Exception(f"Evento {evento_id} não encontrado para o contrato {num_contrato}")
                
            # Remover linha
            ws.delete_rows(linha_evento)
            
            wb.save(self.arquivo_cliente)
            
            messagebox.showinfo("Sucesso", "Evento removido com sucesso!")
            
            # Atualizar lista de eventos
            self.carregar_eventos_contrato(None)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao remover evento: {str(e)}")

    def filtrar_pagamentos(self, event):
        """Filtra pagamentos conforme seleção"""
        contrato = self.pagto_contrato.get()
        status = self.pagto_status.get()
        
        try:
            self.carregar_pagamentos(contrato, status)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao filtrar pagamentos: {str(e)}")

    def carregar_pagamentos(self, contrato_filtro, status_filtro):
        """Carrega os pagamentos de eventos"""
        try:
            wb = load_workbook(self.arquivo_cliente, data_only=True)
            
            # Verificar se existe a aba
            if 'Contratos_ADM' not in wb.sheetnames:
                wb.close()
                return
                
            ws = wb['Contratos_ADM']
            
            # Limpar treeview
            for item in self.tree_pagamentos.get_children():
                self.tree_pagamentos.delete(item)
                
            # Processar pagamentos
            for row in ws.iter_rows(min_row=3, values_only=True):
                # Verificar se é um registro de pagamento (coluna AM ou 38 em diante)
                if len(row) >= 44 and row[37] and row[38]:
                    num_contrato = row[37]
                    evento_id = row[38]
                    cnpj_cpf = row[39]
                    nome = row[40]
                    
                    # Aplicar filtro de contrato
                    if contrato_filtro != 'Todos' and num_contrato != contrato_filtro:
                        continue
                    
                    # Formatar data de vencimento
                    data_vencimento = None
                    if row[41]:
                        if isinstance(row[41], datetime):
                            data_vencimento = row[41].strftime('%d/%m/%Y')
                        else:
                            try:
                                data_vencimento = datetime.strptime(str(row[41]), '%Y-%m-%d').strftime('%d/%m/%Y')
                            except ValueError:
                                data_vencimento = str(row[41])
                    
                    valor = row[42]
                    status = row[43] or 'pendente'
                    
                    # Aplicar filtro de status
                    if status_filtro != 'Todos' and status.capitalize() != status_filtro:
                        continue
                    
                    # Formatar data de pagamento
                    data_pagamento = None
                    if row[44]:
                        if isinstance(row[44], datetime):
                            data_pagamento = row[44].strftime('%d/%m/%Y')
                        else:
                            try:
                                data_pagamento = datetime.strptime(str(row[44]), '%Y-%m-%d').strftime('%d/%m/%Y')
                            except ValueError:
                                data_pagamento = str(row[44])
                    
                    # Adicionar ao treeview
                    valor_fmt = f"R$ {valor:,.2f}" if isinstance(valor, (int, float)) else valor
                    
                    self.tree_pagamentos.insert('', 'end', values=(
                        num_contrato,
                        evento_id,
                        str(cnpj_cpf),  # Garantir que seja string
                        str(nome),      # Garantir que seja string
                        data_vencimento or "-",
                        valor_fmt,
                        status.capitalize(),
                        data_pagamento or "-"
                    ))
            
            wb.close()
            
        except Exception as e:
            if 'wb' in locals():
                wb.close()
            raise Exception(f"Erro ao carregar pagamentos: {str(e)}")

    
            
    def registrar_pagamento(self):
        """Registra um pagamento como efetuado"""
        selecionado = self.tree_pagamentos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um pagamento primeiro")
            return
            
        valores = self.tree_pagamentos.item(selecionado)['values']
        num_contrato = valores[0]
        evento_id = valores[1]
        cnpj_cpf = valores[2]
        
        # Verificar se o pagamento já está como pago
        if valores[6] == 'Pago':
            messagebox.showinfo("Informação", "Este pagamento já está registrado como pago")
            return
            
        # Confirmar registro
        if not messagebox.askyesno("Confirmar", "Confirma o registro deste pagamento?"):
            return
            
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            # Encontrar o pagamento
            pagamento_encontrado = False
            for row in range(3, ws.max_row + 1):
                if (ws.cell(row=row, column=38).value == num_contrato and 
                    str(ws.cell(row=row, column=39).value) == str(evento_id) and
                    str(ws.cell(row=row, column=40).value) == str(cnpj_cpf)):
                    
                    # Atualizar status
                    ws.cell(row=row, column=44, value='pago')
                    
                    # Atualizar data de pagamento
                    data_pagamento = self.pagto_data.get_date()
                    ws.cell(row=row, column=45, value=data_pagamento)
                    
                    pagamento_encontrado = True
                    break
                    
            if not pagamento_encontrado:
                raise Exception("Pagamento não encontrado na planilha")
                
            wb.save(self.arquivo_cliente)
            
            messagebox.showinfo("Sucesso", "Pagamento registrado com sucesso!")
            
            # Atualizar lista de pagamentos
            self.filtrar_pagamentos(None)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao registrar pagamento: {str(e)}")
            if 'wb' in locals():
                wb.close()

    def gerar_lancamento_pagamento(self):
        """Gera um lançamento para o sistema de entrada de dados"""
        selecionado = self.tree_pagamentos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um pagamento primeiro")
            return
            
        valores = self.tree_pagamentos.item(selecionado)['values']
        num_contrato = valores[0]
        evento_id = valores[1]
        cnpj_cpf = valores[2]
        nome = valores[3]
        data_vencimento = valores[4]
        valor_str = valores[5].replace('R', '').replace('.', '').replace(',', '.').strip()
        
        # Converter valor
        try:
            valor = float(valor_str)
        except ValueError:
            messagebox.showerror("Erro", "Valor do pagamento inválido")
            return
            
        # Converter data
        try:
            data_vencto = datetime.strptime(data_vencimento, '%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Erro", "Data de vencimento inválida")
            return
        
        # Buscar administrador para obter categoria
        categoria = "ADM"
        
        # Verificar se o Sistema de Entrada de Dados está disponível no parent
        if hasattr(self.parent, 'dados_para_incluir'):
            # Preparar lançamento
            lancamento = {
                'data': self.calcular_data_relatorio(data_vencto),
                'tp_desp': '3',  # Tipo 3 para administração
                'cnpj_cpf': cnpj_cpf,
                'nome': nome,
                'referencia': f"ADM OBRA - EVENTO {evento_id} - CONTRATO {num_contrato}",
                'nf': '',
                'vr_unit': f"{valor:.2f}",
                'dias': 1,
                'valor': f"{valor:.2f}",
                'dt_vencto': data_vencimento,
                'categoria': categoria,
                'dados_bancarios': self.buscar_dados_bancarios(cnpj_cpf),
                'observacao': f"PAGAMENTO EVENTO {evento_id} - CONTRATO {num_contrato}",
                'forma_pagamento': 'PIX'  # Valor padrão
            }
            
            # Adicionar à lista do sistema
            self.parent.dados_para_incluir.append(lancamento)
            
            messagebox.showinfo(
                "Sucesso", 
                "Lançamento gerado com sucesso! Acesse a aba de 'Visualização de Lançamentos' para conferir."
            )
            
            # Fechar esta janela
            self.janela.destroy()
            
            # Mostrar visualizador de lançamentos
            if hasattr(self.parent, 'visualizar_lancamentos'):
                self.parent.visualizar_lancamentos()
        else:
            messagebox.showwarning(
                "Aviso", 
                "Sistema de Entrada de Dados não está disponível. O lançamento não pôde ser gerado."
            )

    def buscar_dados_bancarios(self, cnpj_cpf):
        """Busca dados bancários do fornecedor"""
        try:
            wb = load_workbook(ARQUIVO_FORNECEDORES, data_only=True)
            ws = wb['Fornecedores']
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if str(row[0]) == str(cnpj_cpf):
                    # Verificar se tem chave PIX
                    if row[10]:  # Coluna K - Chave PIX
                        dados = f"PIX: {row[10]}"
                    else:
                        # Construir dados baseados nas informações bancárias
                        partes = []
                        if row[6]: partes.append(str(row[6]))  # Banco
                        if row[7]: partes.append(str(row[7]))  # OP
                        if row[8]: partes.append(str(row[8]))  # Agência
                        if row[9]: partes.append(str(row[9]))  # Conta
                        if row[0]: partes.append(str(row[0]))  # CNPJ/CPF
                        
                        dados = ' - '.join(partes)
                    
                    wb.close()
                    return dados or "DADOS BANCÁRIOS NÃO CADASTRADOS"
                    
            wb.close()
            return "DADOS BANCÁRIOS NÃO CADASTRADOS"
            
        except Exception as e:
            print(f"Erro ao buscar dados bancários: {str(e)}")
            if 'wb' in locals():
                wb.close()
            return "ERRO AO BUSCAR DADOS BANCÁRIOS"

    def calcular_data_relatorio(self, data_vencimento):
        """Calcula a data do relatório com base na data de vencimento"""
        try:
            hoje = datetime.now()
            
            # Para as demais parcelas, manter a lógica existente
            if data_vencimento.day <= 5:
                # Se vence até dia 5, relatório é dia 20 do mês anterior
                data_rel = (data_vencimento - relativedelta(months=1)).replace(day=20)
            elif data_vencimento.day <= 20:
                # Se vence até dia 20, relatório é dia 5 do mesmo mês
                data_rel = data_vencimento.replace(day=5)
            else:
                # Se vence após dia 20, relatório é dia 20 do mesmo mês
                data_rel = data_vencimento.replace(day=20)
                
            # Garantir que a data do relatório não seja anterior à data atual
            if data_rel < hoje:
                if hoje.day <= 5:
                    data_rel = hoje.replace(day=5)
                elif hoje.day <= 20:
                    data_rel = hoje.replace(day=20)
                else:
                    proximo_mes = hoje + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
                    
            return data_rel.strftime('%d/%m/%Y')
            
        except Exception as e:
            print(f"Erro ao calcular data do relatório: {str(e)}")
            return hoje.strftime('%d/%m/%Y')
        

    