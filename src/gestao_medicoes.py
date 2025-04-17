import os
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from dateutil.relativedelta import relativedelta
from tkcalendar import DateEntry
import pandas as pd
from pathlib import Path
import xlwings as xw
from openpyxl import load_workbook
import openpyxl
import babel

# Adicionar diretório raiz ao path para importar módulos corretamente
def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

# Importar configurações e utilitários
try:
    from config.logger_config import system_logger, log_action
    logger = system_logger.get_logger()
    logger.info("Logger importado com sucesso em gestao_medicoes.py")
except Exception as e:
    print(f"Erro ao importar logger: {str(e)}")
    
try:
    from config.config import (
        ARQUIVO_CLIENTES,
        ARQUIVO_FORNECEDORES,
        ARQUIVO_MODELO,
        PASTA_CLIENTES,
        BASE_PATH
    )
    print("Configurações importadas com sucesso")
except ImportError as e:
    print(f"Erro ao importar configurações: {str(e)}")
    raise

try:
    from config.window_config import configurar_janela
    print("window_config importado com sucesso")
except ImportError as e:
    from src.config.window_config import configurar_janela
    print("window_config importado pelo caminho alternativo")
    
# Importar funções auxiliares
from config.utils import (
    formatar_cnpj_cpf,
    buscar_dados_bancarios_fornecedor,
    validar_cnpj_cpf,
    formatar_cnpj_cpf,
    validar_data,
    aplicar_formatacao_celula,
    formatar_moeda_br
)

class GestaoMedicoes:
    """Classe principal para gestão de medições"""
    def __init__(self, parent=None):
        """Inicializa a interface de gestão de medições"""
        self.parent = parent
        
        if parent:
            self.root = tk.Toplevel(parent)
            self.menu_principal = parent
        else:
            self.root = tk.Tk()
            self.menu_principal = None
            
        configurar_janela(self.root, "Gestão de Medições")
        
        # Configuração de variáveis
        self.cliente_atual = None
        self.arquivo_cliente = None
        self.contrato_atual = None
        self.fornecedor_atual = None
        self.dados_para_incluir = []
        
        # Configurar interface
        self.setup_gui()
        
    def setup_gui(self):
        """Configuração da interface gráfica"""
        # Notebook (abas)
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Criar abas
        self.aba_selecao = ttk.Frame(self.notebook)
        self.aba_contratos = ttk.Frame(self.notebook)
        self.aba_medicoes = ttk.Frame(self.notebook)
        
        self.notebook.add(self.aba_selecao, text='Seleção de Cliente')
        self.notebook.add(self.aba_contratos, text='Contratos')
        self.notebook.add(self.aba_medicoes, text='Medições')
        
        # Configurar cada aba
        self.setup_aba_selecao()
        self.setup_aba_contratos()
        self.setup_aba_medicoes()
        
    def setup_aba_selecao(self):
        """Configura a aba de seleção de cliente"""
        # Frame principal
        frame_principal = ttk.Frame(self.aba_selecao)
        frame_principal.pack(expand=True, fill='both', padx=10, pady=5)

        # Frame para seleção de cliente
        frame_selecao = ttk.LabelFrame(frame_principal, text="Seleção do Cliente")
        frame_selecao.pack(fill='x', pady=10)

        # Container para label e combobox
        frame_cliente = ttk.Frame(frame_selecao)
        frame_cliente.pack(fill='x', padx=10, pady=10)

        # Label alinhado à esquerda
        ttk.Label(frame_cliente, text="Selecione o Cliente:", font=('Arial', 11)).pack(side='left', pady=5)
        
        # Combobox com largura aumentada
        self.cliente_combobox = ttk.Combobox(frame_cliente, width=60, font=('Arial', 11))
        self.cliente_combobox.pack(side='left', padx=5, fill='x', expand=True)
        
        # Frame para botões de gerenciamento de clientes
        frame_gerenciar = ttk.Frame(frame_principal)
        frame_gerenciar.pack(pady=15)
        
        # Estilo para botões maiores
        style = ttk.Style()
        style.configure('Big.TButton', font=('Arial', 12, 'bold'), padding=(15, 10))
        
        # Botões
        ttk.Button(frame_gerenciar, 
                text="Continuar →",
                command=self.continuar_para_contratos,
                style='Big.TButton').pack(side='right', padx=10)
        
        # Frame de botões
        frame_botoes_selecao = ttk.Frame(frame_principal)
        frame_botoes_selecao.pack(fill='x', side='bottom', pady=10)

        ttk.Button(frame_botoes_selecao, 
                text="Voltar ao Menu", 
                command=self.voltar_menu,
                style='Big.TButton').pack(side='left', padx=10)
        
        # Carregar clientes existentes
        self.atualizar_lista_clientes()
        
        # Binding para seleção de cliente
        self.cliente_combobox.bind('<<ComboboxSelected>>', self.selecionar_cliente)
        
    def setup_aba_contratos(self):
        """Configura a aba de contratos"""
        # Frame principal
        frame_principal = ttk.Frame(self.aba_contratos)
        frame_principal.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Label para cliente atual
        self.lbl_cliente_contratos = ttk.Label(
            frame_principal, 
            text="Cliente: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_cliente_contratos.pack(anchor='w', padx=5, pady=5)
        
        # Frame para lista de contratos
        frame_contratos = ttk.LabelFrame(frame_principal, text="Contratos Registrados")
        frame_contratos.pack(fill='both', expand=True, pady=5)
        
        # Treeview para contratos
        colunas = ('ID', 'Fornecedor', 'Descrição', 'Data Início', 'Valor Global', 'Valor Pago', 'Saldo')
        self.tree_contratos = ttk.Treeview(frame_contratos, columns=colunas, show='headings', height=10)
        
        # Configurar colunas
        self.tree_contratos.heading('ID', text='ID')
        self.tree_contratos.heading('Fornecedor', text='Fornecedor')
        self.tree_contratos.heading('Descrição', text='Descrição')
        self.tree_contratos.heading('Data Início', text='Data Início')
        self.tree_contratos.heading('Valor Global', text='Valor Global')
        self.tree_contratos.heading('Valor Pago', text='Valor Pago')
        self.tree_contratos.heading('Saldo', text='Saldo')
        
        # Ajustar larguras das colunas
        self.tree_contratos.column('ID', width=50, anchor='center')
        self.tree_contratos.column('Fornecedor', width=150)
        self.tree_contratos.column('Descrição', width=200)
        self.tree_contratos.column('Data Início', width=100, anchor='center')
        self.tree_contratos.column('Valor Global', width=100, anchor='e')
        self.tree_contratos.column('Valor Pago', width=100, anchor='e')
        self.tree_contratos.column('Saldo', width=100, anchor='e')
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_contratos, orient='vertical', command=self.tree_contratos.yview)
        scrollx = ttk.Scrollbar(frame_contratos, orient='horizontal', command=self.tree_contratos.xview)
        self.tree_contratos.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_contratos.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
        # Frame para botões
        frame_botoes = ttk.Frame(frame_principal)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, text="Novo Contrato", 
                  command=self.novo_contrato).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Editar Contrato", 
                  command=self.editar_contrato).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Selecionar", 
                  command=self.selecionar_contrato).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Voltar", 
                  command=lambda: self.notebook.select(0)).pack(side='right', padx=5)
        
        # Binding para duplo clique
        self.tree_contratos.bind('<Double-1>', lambda e: self.selecionar_contrato())

    def setup_aba_medicoes(self):
        """Configura a aba de medições"""
        # Frame principal
        frame_principal = ttk.Frame(self.aba_medicoes)
        frame_principal.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Labels para cliente e contrato atuais
        frame_info = ttk.Frame(frame_principal)
        frame_info.pack(fill='x', pady=5)
        
        self.lbl_cliente_medicoes = ttk.Label(
            frame_info, 
            text="Cliente: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_cliente_medicoes.pack(side='left', padx=10)
        
        self.lbl_contrato_medicoes = ttk.Label(
            frame_info, 
            text="Contrato: Nenhum selecionado", 
            font=('Arial', 12, 'bold'),
            foreground='#0056b3'
        )
        self.lbl_contrato_medicoes.pack(side='left', padx=10)
        
        # Frame para lista de medições
        frame_medicoes = ttk.LabelFrame(frame_principal, text="Medições Registradas")
        frame_medicoes.pack(fill='both', expand=True, pady=5)
        
        # Treeview para medições
        colunas = ('ID', 'Data Medição', 'Data Pagamento', 'Referência', 'Valor', 'Status')
        self.tree_medicoes = ttk.Treeview(frame_medicoes, columns=colunas, show='headings', height=10)
        
        # Configurar colunas
        self.tree_medicoes.heading('ID', text='ID')
        self.tree_medicoes.heading('Data Medição', text='Data Medição')
        self.tree_medicoes.heading('Data Pagamento', text='Data Pagamento')
        self.tree_medicoes.heading('Referência', text='Referência')
        self.tree_medicoes.heading('Valor', text='Valor')
        self.tree_medicoes.heading('Status', text='Status')
        
        # Ajustar larguras das colunas
        self.tree_medicoes.column('ID', width=50, anchor='center')
        self.tree_medicoes.column('Data Medição', width=100, anchor='center')
        self.tree_medicoes.column('Data Pagamento', width=100, anchor='center')
        self.tree_medicoes.column('Referência', width=300)
        self.tree_medicoes.column('Valor', width=100, anchor='e')
        self.tree_medicoes.column('Status', width=100, anchor='center')
        
        # Scrollbars
        scrolly = ttk.Scrollbar(frame_medicoes, orient='vertical', command=self.tree_medicoes.yview)
        scrollx = ttk.Scrollbar(frame_medicoes, orient='horizontal', command=self.tree_medicoes.xview)
        self.tree_medicoes.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
        
        # Posicionamento
        self.tree_medicoes.pack(side='left', fill='both', expand=True)
        scrolly.pack(side='right', fill='y')
        scrollx.pack(side='bottom', fill='x')
        
        # Frame para botões
        frame_botoes = ttk.Frame(frame_principal)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, text="Nova Medição", 
                  command=self.nova_medicao).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Editar Medição", 
                  command=self.editar_medicao).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Lançar no Cliente", 
                  command=self.lancar_medicao).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Voltar", 
                  command=lambda: self.notebook.select(1)).pack(side='right', padx=5)
        
        # Botão para voltar ao menu principal
        ttk.Button(frame_principal, text="Voltar ao Menu Principal", 
                 command=self.voltar_menu).pack(side='bottom', pady=10)

    def centralizar_janela(self, janela, largura=600, altura=400):
        """Centraliza a janela em relação à janela principal"""
        # Atualizar a geometria para aplicar dimensões
        janela.geometry(f"{largura}x{altura}")
        janela.update_idletasks()
        
        # Obter as dimensões da janela principal
        if self.root and self.root.winfo_exists():
            main_x = self.root.winfo_x()
            main_y = self.root.winfo_y()
            main_width = self.root.winfo_width()
            main_height = self.root.winfo_height()
            
            # Calcular posição centralizada
            x = main_x + (main_width - largura) // 2
            y = main_y + (main_height - altura) // 2
        else:
            # Centralizar na tela se não houver janela principal
            x = (janela.winfo_screenwidth() - largura) // 2
            y = (janela.winfo_screenheight() - altura) // 2
            
        # Aplicar posição
        janela.geometry(f"{largura}x{altura}+{x}+{y}")
        
        # Tornar a janela modal
        janela.transient(self.root)
        janela.grab_set()
        janela.focus_force()
        # Atualizar a geometria para aplicar dimensões
        janela.geometry(f"{largura}x{altura}")
        janela.update_idletasks()
        
        # Obter as dimensões da janela principal
        if self.root and self.root.winfo_exists():
            main_x = self.root.winfo_x()
            main_y = self.root.winfo_y()
            main_width = self.root.winfo_width()
            main_height = self.root.winfo_height()
            
            # Calcular posição centralizada
            x = main_x + (main_width - largura) // 2
            y = main_y + (main_height - altura) // 2
        else:
            # Centralizar na tela se não houver janela principal
            x = (janela.winfo_screenwidth() - largura) // 2
            y = (janela.winfo_screenheight() - altura) // 2
            
        # Aplicar posição
        janela.geometry(f"{largura}x{altura}+{x}+{y}")
        
        # Tornar a janela modal
        janela.transient(self.root)
        janela.grab_set()
        janela.focus_force()
        try:
            # Obter dados do contrato
            contrato = self.obter_dados_contrato(self.contrato_atual)
            if not contrato:
                return 0
                
            return float(contrato['saldo'])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar saldo: {str(e)}")
            return 0

    def formatar_documento(self, valor):
        """Formata um documento (CNPJ/CPF) preservando zeros à esquerda e adicionando pontuação"""
        # Garantir que estamos trabalhando com string
        valor_str = str(valor)
        
        # Limpar a string, removendo caracteres não numéricos
        valor_limpo = ''.join(filter(str.isdigit, valor_str))
        
        # Determinar se é CPF (11 dígitos) ou CNPJ (14 dígitos)
        if len(valor_limpo) <= 11:
            # É um CPF, garantir 11 dígitos
            documento = valor_limpo.zfill(11)
            # Formatar como XXX.XXX.XXX-XX
            return f"{documento[:3]}.{documento[3:6]}.{documento[6:9]}-{documento[9:]}"
        else:
            # É um CNPJ, garantir 14 dígitos
            documento = valor_limpo.zfill(14)
            # Formatar como XX.XXX.XXX/XXXX-XX
            return f"{documento[:2]}.{documento[2:5]}.{documento[5:8]}/{documento[8:12]}-{documento[12:]}"

    # Funções da aba Cliente
    def atualizar_lista_clientes(self):
        """Atualiza a lista de clientes no combobox"""
        try:
            # Carregar arquivo de clientes
            workbook = load_workbook(ARQUIVO_CLIENTES)
            sheet = workbook['Clientes']  # Assumindo que existe uma aba chamada 'Clientes'
            
            # Limpar lista atual
            self.cliente_combobox['values'] = []
            
            # Pegar todos os clientes (pulando o cabeçalho)
            clientes = []
            for row in sheet.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Nome do cliente está na primeira coluna
                    clientes.append(row[0])
            
            # Atualizar combobox
            self.cliente_combobox['values'] = sorted(clientes)
            workbook.close()
            
        except FileNotFoundError:
            messagebox.showerror("Erro", "Arquivo de clientes não encontrado.")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
    
    def selecionar_cliente(self, event=None):
        """Atualiza seleção de cliente e habilita botão de continuar"""
        self.cliente_atual = self.cliente_combobox.get()
        
        # Atualiza label em todas as abas
        if self.cliente_atual:
            # Atualiza labels nas abas
            self.lbl_cliente_contratos.config(text=f"Cliente: {self.cliente_atual}")
            self.lbl_cliente_medicoes.config(text=f"Cliente: {self.cliente_atual}")
            
            # Define o caminho do arquivo
            self.arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
            
            # Verifica se arquivo existe e cria a aba de medições se necessário
            self.verificar_aba_medicoes()
    
    def continuar_para_contratos(self):
        """Avança para a aba de contratos após confirmar seleção"""
        if self.cliente_atual:
            self.notebook.select(1)  # Vai para aba de contratos
            self.carregar_contratos()
        else:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            
    def verificar_aba_medicoes(self):
        """Verifica se a aba de medições existe na planilha do cliente e cria se necessário"""
        try:
            if not os.path.exists(self.arquivo_cliente):
                messagebox.showerror("Erro", f"Arquivo do cliente '{self.cliente_atual}' não encontrado!")
                return False
                
            wb = load_workbook(self.arquivo_cliente)
            
            # Verificar se a aba de medições já existe
            if "Medicoes" not in wb.sheetnames:
                # Criar aba de medições
                ws = wb.create_sheet("Medicoes")
                
                # Definir cabeçalhos
                headers = [
                    "ID_Contrato", "ID_Medicao", "CNPJ_Fornecedor", "Nome_Fornecedor", 
                    "Data_Medicao", "Data_Pagamento", "Referencia", "Valor", 
                    "Status", "Data_Lancamento", "Observacao"
                ]
                
                for col, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col, value=header)
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center')
                
                # Ajustar largura das colunas
                for col in range(1, len(headers) + 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
                
                wb.save(self.arquivo_cliente)
                return True
            
            return True
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar aba de medições: {str(e)}")
            return False
    
    # Funções da aba Contratos
    def carregar_contratos(self):
        """Carrega os contratos do cliente atual"""
        try:
            if not self.arquivo_cliente:
                messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
                return
                
            # Limpar treeview
            for item in self.tree_contratos.get_children():
                self.tree_contratos.delete(item)
                
            # Verificar existência da aba de medições
            self.verificar_aba_medicoes()
                
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            
            # Verificar se a aba de contratos existe
            if "Contratos_Medicao" not in wb.sheetnames:
                # Criar aba de contratos
                ws = wb.create_sheet("Contratos_Medicao")
                
                # Definir cabeçalhos
                headers = [
                    "ID_Contrato", "CNPJ_Fornecedor", "Nome_Fornecedor", "Descricao", 
                    "Data_Inicio", "Valor_Global", "Valor_Pago", "Saldo", "Status", "Observacao"
                ]
                
                for col, header in enumerate(headers, 1):
                    cell = ws.cell(row=1, column=col, value=header)
                    cell.font = openpyxl.styles.Font(bold=True)
                    cell.alignment = openpyxl.styles.Alignment(horizontal='center')
                
                # Ajustar largura das colunas
                for col in range(1, len(headers) + 1):
                    ws.column_dimensions[openpyxl.utils.get_column_letter(col)].width = 15
                
                wb.save(self.arquivo_cliente)
                return
                
            # Carregar dados da aba Contratos_Medicao
            ws = wb["Contratos_Medicao"]
            
            # Percorrer as linhas (pulando o cabeçalho)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0]:  # Se tem ID_Contrato
                    # Formatação de dados
                    try:
                        data_inicio = row[4].strftime('%d/%m/%Y') if isinstance(row[4], datetime) else row[4]
                        valor_global = formatar_moeda_br(row[5]) if row[5] else "R$ 0,00"
                        valor_pago = formatar_moeda_br(row[6]) if row[6] else "R$ 0,00"
                        saldo = formatar_moeda_br(row[7]) if row[7] else "R$ 0,00"
                    except (ValueError, TypeError, AttributeError) as e:
                        data_inicio = str(row[4]) if row[4] else ""
                        valor_global = str(row[5]) if row[5] else "R$ 0,00"
                        valor_pago = str(row[6]) if row[6] else "R$ 0,00"
                        saldo = str(row[7]) if row[7] else "R$ 0,00"
                    
                    # Adicionar à treeview
                    self.tree_contratos.insert('', 'end', values=(
                        row[0],             # ID
                        row[2],             # Nome Fornecedor
                        row[3],             # Descrição
                        data_inicio,        # Data Início
                        valor_global,       # Valor Global
                        valor_pago,         # Valor Pago
                        saldo               # Saldo
                    ))
                
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar contratos: {str(e)}")
    
    def novo_contrato(self):
        """Abre janela para cadastro de novo contrato"""
        if not self.cliente_atual:
            messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
            return
            
        # Criar janela de cadastro
        janela = tk.Toplevel(self.root)
        janela.title("Novo Contrato")
        
        # Centralizar a janela em relação à janela principal
        self.centralizar_janela(janela, 700, 600)
        
        # Frame principal com scroll
        canvas = tk.Canvas(janela)
        scrollbar = ttk.Scrollbar(janela, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(
                scrollregion=canvas.bbox("all")
            )
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Frame principal dentro do scroll
        frame = ttk.Frame(scrollable_frame, padding="10")
        frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Frame para busca de fornecedor
        frame_busca = ttk.LabelFrame(frame, text="Buscar Fornecedor")
        frame_busca.pack(fill='x', pady=5)
        
        ttk.Label(frame_busca, text="Nome:").grid(row=0, column=0, padx=5, pady=5)
        busca_entry = ttk.Entry(frame_busca, width=30)
        busca_entry.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_busca, text="Buscar", 
                  command=lambda: self.buscar_fornecedor(tree_fornecedores, busca_entry.get())
                 ).grid(row=0, column=2, padx=5, pady=5)
                 
        ttk.Button(frame_busca, text="Novo Fornecedor", 
                  command=self.novo_fornecedor
                 ).grid(row=0, column=3, padx=5, pady=5)
        
        # Frame para lista de fornecedores
        frame_fornecedores = ttk.LabelFrame(frame, text="Fornecedores")
        frame_fornecedores.pack(fill='x', pady=5)
        
        # Treeview para fornecedores
        tree_fornecedores = ttk.Treeview(
            frame_fornecedores, 
            columns=('CNPJ/CPF', 'Nome', 'Categoria'),
            show='headings',
            height=4
        )
        tree_fornecedores.heading('CNPJ/CPF', text='CNPJ/CPF')
        tree_fornecedores.heading('Nome', text='Nome')
        tree_fornecedores.heading('Categoria', text='Categoria')
        
        tree_fornecedores.column('CNPJ/CPF', width=150)
        tree_fornecedores.column('Nome', width=250)
        tree_fornecedores.column('Categoria', width=100)
        
        tree_fornecedores.pack(fill='x', padx=5, pady=5)
        
        # Binding para selecionar fornecedor
        tree_fornecedores.bind('<Double-1>', lambda e: self.selecionar_fornecedor_contrato(
            tree_fornecedores, cnpj_entry, nome_entry
        ))
        
        # Frame para dados do contrato
        frame_contrato = ttk.LabelFrame(frame, text="Dados do Contrato")
        frame_contrato.pack(fill='x', pady=10)
        
        # CNPJ Fornecedor
        ttk.Label(frame_contrato, text="CNPJ/CPF:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        cnpj_entry = ttk.Entry(frame_contrato, width=30)
        cnpj_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        cnpj_entry.config(state='readonly')
        
        # Nome Fornecedor
        ttk.Label(frame_contrato, text="Nome:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        nome_entry = ttk.Entry(frame_contrato, width=50)
        nome_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        nome_entry.config(state='readonly')
        
        # Descrição do Contrato
        ttk.Label(frame_contrato, text="Descrição:*").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        descricao_entry = ttk.Entry(frame_contrato, width=90)
        descricao_entry.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        # Data de Início
        ttk.Label(frame_contrato, text="Data de Início:*").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        data_inicio = DateEntry(frame_contrato, width=12, background='darkblue',
                              foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        data_inicio.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # Valor Global
        ttk.Label(frame_contrato, text="Valor Global (R$):*").grid(row=4, column=0, padx=5, pady=5, sticky='e')
        valor_global = ttk.Entry(frame_contrato, width=20)
        valor_global.grid(row=4, column=1, padx=5, pady=5, sticky='w')
        
        # Observações
        ttk.Label(frame_contrato, text="Observações:").grid(row=5, column=0, padx=5, pady=5, sticky='ne')
        observacoes = tk.Text(frame_contrato, width=40, height=4)
        observacoes.grid(row=5, column=1, padx=5, pady=5, sticky='w')
        
        # Frame para botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, 
                 text="Salvar", 
                 command=lambda: self.salvar_contrato(
                     janela,
                     cnpj_entry.get(),
                     nome_entry.get(),
                     descricao_entry.get(),
                     data_inicio.get(),
                     valor_global.get(),
                     observacoes.get("1.0", "end-1c")
                 )).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                 text="Cancelar", 
                 command=janela.destroy).pack(side='left', padx=5)
                 
    def selecionar_fornecedor_contrato(self, tree, cnpj_entry, nome_entry):
        """Seleciona um fornecedor da lista para o contrato"""
        selecionado = tree.selection()
        if not selecionado:
            return
            
        # Obter valores selecionados
        valores = tree.item(selecionado)['values']
        
        # Preencher os campos
        cnpj_entry.config(state='normal')
        cnpj_entry.delete(0, tk.END)
        cnpj_entry.insert(0, valores[0])
        cnpj_entry.config(state='readonly')
        
        nome_entry.config(state='normal')
        nome_entry.delete(0, tk.END)
        nome_entry.insert(0, valores[1])
        nome_entry.config(state='readonly')
    
    def buscar_fornecedor(self, tree, termo):
        """Busca fornecedores na base com tratamento melhorado para CNPJ/CPF"""
        try:
            # Limpar treeview
            for item in tree.get_children():
                tree.delete(item)
                    
            if not termo:
                return
                    
            # Carregar planilha de fornecedores
            wb = load_workbook(ARQUIVO_FORNECEDORES)
            ws = wb['Fornecedores']
            
            # Buscar fornecedores que contenham o termo
            termo = termo.upper()
            # Verificar se o termo pode ser um CNPJ/CPF (apenas dígitos)
            termo_numerico = ''.join(filter(str.isdigit, termo))
            
            encontrados = []
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # Pular linhas sem CNPJ/CPF
                    continue
                    
                # Converter CNPJ/CPF da linha para comparação
                row_cnpj = ''.join(filter(str.isdigit, str(row[0])))
                
                # Verificar se termo está no CNPJ/CPF, nome ou razão social
                if (termo_numerico and termo_numerico in row_cnpj) or \
                (termo in str(row[3]).upper()) or \
                (termo in str(row[2]).upper()):
                    # Formatar o CNPJ/CPF corretamente
                    cnpj_formatado = self.formatar_documento(row[0])
                    encontrados.append((cnpj_formatado, row[3], row[11]))
                        
            # Adicionar à treeview
            for fornecedor in encontrados:
                tree.insert('', 'end', values=fornecedor)
                    
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao buscar fornecedores: {str(e)}")
    
    def novo_fornecedor(self):
        """Abre janela para cadastro de novo fornecedor"""
        try:
            # Importar funcionalidade de cadastro de fornecedor
            from Sistema_Entrada_Dados import SistemaEntradaDados
            
            sistema = SistemaEntradaDados(self.root)
            sistema.novo_fornecedor()
            
        except ImportError:
            messagebox.showerror("Erro", "Módulo de cadastro de fornecedor não encontrado")
            
    def salvar_contrato(self, janela, cnpj, nome, descricao, data_inicio, valor_global, observacoes):
        """Salva um novo contrato"""
        try:
            # Validar campos obrigatórios
            if not cnpj or not nome or not descricao or not data_inicio or not valor_global:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
            
            # Validar valor global
            try:
                valor = float(valor_global.replace(',', '.'))
                if valor <= 0:
                    messagebox.showerror("Erro", "Valor global deve ser maior que zero!")
                    return
            except ValueError:
                messagebox.showerror("Erro", "Valor global inválido!")
                return
            
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            
            # Verificar se a aba de contratos existe
            if "Contratos_Medicao" not in wb.sheetnames:
                messagebox.showerror("Erro", "Aba de contratos não encontrada!")
                wb.close()
                return
            
            ws = wb["Contratos_Medicao"]
            
            # Gerar ID do contrato (próximo número sequencial)
            next_id = 1
            for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                if row[0] and isinstance(row[0], int) and row[0] >= next_id:
                    next_id = row[0] + 1
            
            # Converter data
            try:
                data = datetime.strptime(data_inicio, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida! Use o formato dd/mm/aaaa")
                wb.close()
                return
            
            # Adicionar novo contrato
            proxima_linha = ws.max_row + 1
            ws.cell(row=proxima_linha, column=1, value=next_id)                # ID_Contrato
            ws.cell(row=proxima_linha, column=2, value=cnpj)                   # CNPJ_Fornecedor
            ws.cell(row=proxima_linha, column=3, value=nome)                   # Nome_Fornecedor
            ws.cell(row=proxima_linha, column=4, value=descricao.upper())      # Descricao
            ws.cell(row=proxima_linha, column=5, value=data)                   # Data_Inicio
            
            # Formatação para células de valor
            valor_cell = ws.cell(row=proxima_linha, column=6, value=valor)     # Valor_Global
            valor_cell.number_format = '#.##0,00'
            
            zero_cell_1 = ws.cell(row=proxima_linha, column=7, value=0)        # Valor_Pago
            zero_cell_1.number_format = '#.##0,00'
            
            saldo_cell = ws.cell(row=proxima_linha, column=8, value=valor)     # Saldo
            saldo_cell.number_format = '#.##0,00'
            
            ws.cell(row=proxima_linha, column=9, value="ATIVO")                # Status
            ws.cell(row=proxima_linha, column=10, value=observacoes.upper())   # Observacao
            
            # Salvar arquivo
            wb.save(self.arquivo_cliente)
            wb.close()
            
            messagebox.showinfo("Sucesso", "Contrato cadastrado com sucesso!")
            janela.destroy()
            
            # Atualizar lista de contratos
            self.carregar_contratos()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar contrato: {str(e)}")
            try:
                wb.close()
            except:
                pass
    
    def editar_contrato(self):
        """Edita o contrato selecionado"""
        selecionado = self.tree_contratos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um contrato para editar")
            return
            
        # Obter ID do contrato selecionado
        valores = self.tree_contratos.item(selecionado)['values']
        id_contrato = valores[0]
        
        # Obter dados completos do contrato
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb["Contratos_Medicao"]
            
            dados_contrato = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == id_contrato:
                    dados_contrato = {
                        'id': row[0],
                        'cnpj': row[1],
                        'nome': row[2],
                        'descricao': row[3],
                        'data_inicio': row[4],
                        'valor_global': row[5],
                        'valor_pago': row[6],
                        'saldo': row[7],
                        'status': row[8],
                        'observacao': row[9]
                    }
                    break
                    
            wb.close()
            
            if not dados_contrato:
                messagebox.showerror("Erro", "Contrato não encontrado!")
                return
                
            # Criar janela de edição
            janela = tk.Toplevel(self.root)
            janela.title("Editar Contrato")
            
            # Centralizar a janela em relação à janela principal
            self.centralizar_janela(janela, 700, 500)
            
            # Frame principal
            frame = ttk.Frame(janela, padding="10")
            frame.pack(fill='both', expand=True)

            # Frame para dados do fornecedor
            frame_fornecedor = ttk.LabelFrame(frame, text="Dados do Fornecedor")
            frame_fornecedor.pack(fill='x', pady=5)
            
            # CNPJ/CPF
            ttk.Label(frame_fornecedor, text="CNPJ/CPF:").grid(row=0, column=0, padx=5, pady=5, sticky='e')
            cnpj_entry = ttk.Entry(frame_fornecedor, width=30)
            cnpj_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
            cnpj_entry.insert(0, dados_contrato['cnpj'])
            cnpj_entry.config(state='readonly')
            
            # Nome
            ttk.Label(frame_fornecedor, text="Nome:").grid(row=1, column=0, padx=5, pady=5, sticky='e')
            nome_entry = ttk.Entry(frame_fornecedor, width=50)
            nome_entry.grid(row=1, column=1, padx=5, pady=5, sticky='w')
            nome_entry.insert(0, dados_contrato['nome'])
            nome_entry.config(state='readonly')
            
            # Frame para dados do contrato
            frame_contrato = ttk.LabelFrame(frame, text="Dados do Contrato")
            frame_contrato.pack(fill='x', pady=10)
            
            # Descrição
            ttk.Label(frame_contrato, text="Descrição:*").grid(row=0, column=0, padx=5, pady=5, sticky='e')
            descricao_entry = ttk.Entry(frame_contrato, width=80)
            descricao_entry.grid(row=0, column=1, padx=5, pady=5, sticky='w')
            descricao_entry.insert(0, dados_contrato['descricao'])
            
            # Data de Início
            ttk.Label(frame_contrato, text="Data de Início:*").grid(row=1, column=0, padx=5, pady=5, sticky='e')
            data_inicio = DateEntry(frame_contrato, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
            data_inicio.grid(row=1, column=1, padx=5, pady=5, sticky='w')
            if isinstance(dados_contrato['data_inicio'], datetime):
                data_inicio.set_date(dados_contrato['data_inicio'])
            
            # Valor Global
            ttk.Label(frame_contrato, text="Valor Global (R$):*").grid(row=2, column=0, padx=5, pady=5, sticky='e')
            valor_global = ttk.Entry(frame_contrato, width=20)
            valor_global.grid(row=2, column=1, padx=5, pady=5, sticky='w')
            valor_global.insert(0, str(dados_contrato['valor_global']))
            
            # Valor Pago (somente leitura)
            ttk.Label(frame_contrato, text="Valor Pago (R$):").grid(row=3, column=0, padx=5, pady=5, sticky='e')
            valor_pago = ttk.Entry(frame_contrato, width=20)
            valor_pago.grid(row=3, column=1, padx=5, pady=5, sticky='w')
            valor_pago.insert(0, str(dados_contrato['valor_pago']))
            valor_pago.config(state='readonly')
            
            # Saldo (somente leitura)
            ttk.Label(frame_contrato, text="Saldo (R$):").grid(row=4, column=0, padx=5, pady=5, sticky='e')
            saldo = ttk.Entry(frame_contrato, width=20)
            saldo.grid(row=4, column=1, padx=5, pady=5, sticky='w')
            saldo.insert(0, str(dados_contrato['saldo']))
            saldo.config(state='readonly')
            
            # Status
            ttk.Label(frame_contrato, text="Status:").grid(row=5, column=0, padx=5, pady=5, sticky='e')
            status = ttk.Combobox(frame_contrato, values=["ATIVO", "CONCLUÍDO", "CANCELADO"], width=15)
            status.grid(row=5, column=1, padx=5, pady=5, sticky='w')
            status.set(dados_contrato['status'] if dados_contrato['status'] else "ATIVO")
            
            # Observações
            ttk.Label(frame_contrato, text="Observações:").grid(row=6, column=0, padx=5, pady=5, sticky='ne')
            observacoes = tk.Text(frame_contrato, width=40, height=4)
            observacoes.grid(row=6, column=1, padx=5, pady=5, sticky='w')
            observacoes.insert("1.0", dados_contrato['observacao'] if dados_contrato['observacao'] else "")
            
            # Frame para botões
            frame_botoes = ttk.Frame(frame)
            frame_botoes.pack(fill='x', pady=10)
            
            ttk.Button(frame_botoes, 
                     text="Salvar", 
                     command=lambda: self.atualizar_contrato(
                         janela,
                         id_contrato,
                         descricao_entry.get(),
                         data_inicio.get(),
                         valor_global.get(),
                         status.get(),
                         observacoes.get("1.0", "end-1c")
                     )).pack(side='left', padx=5)
            
            ttk.Button(frame_botoes, 
                     text="Cancelar", 
                     command=janela.destroy).pack(side='left', padx=5)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao editar contrato: {str(e)}")
    
    def atualizar_contrato(self, janela, id_contrato, descricao, data_inicio, valor_global, status, observacoes):
        """Atualiza os dados de um contrato"""
        try:
            # Validar campos obrigatórios
            if not descricao or not data_inicio or not valor_global:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
            
            # Validar valor global
            try:
                valor = float(valor_global.replace(',', '.'))
                if valor <= 0:
                    messagebox.showerror("Erro", "Valor global deve ser maior que zero!")
                    return
            except ValueError:
                messagebox.showerror("Erro", "Valor global inválido!")
                return
            
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            ws = wb["Contratos_Medicao"]
            
            # Buscar contrato pelo ID
            valor_pago = 0
            row_index = None
            
            for idx, row in enumerate(ws.iter_rows(min_row=2, max_col=1, values_only=True), 2):
                if row[0] == id_contrato:
                    row_index = idx
                    # Obter valor pago atual
                    valor_pago = ws.cell(row=row_index, column=7).value or 0
                    break
                    
            if not row_index:
                messagebox.showerror("Erro", "Contrato não encontrado!")
                wb.close()
                return
                
            # Converter data
            try:
                data = datetime.strptime(data_inicio, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida! Use o formato dd/mm/aaaa")
                wb.close()
                return
                
            # Calcular novo saldo
            saldo = valor - valor_pago
            
            # Atualizar dados
            ws.cell(row=row_index, column=4, value=descricao.upper())     # Descricao
            ws.cell(row=row_index, column=5, value=data)                   # Data_Inicio
            
            # Valor Global
            valor_cell = ws.cell(row=row_index, column=6, value=valor)     # Valor_Global
            valor_cell.number_format = '#.##0,00'
            
            # Saldo
            saldo_cell = ws.cell(row=row_index, column=8, value=saldo)     # Saldo
            saldo_cell.number_format = '#.##0,00'
            
            ws.cell(row=row_index, column=9, value=status)                # Status
            ws.cell(row=row_index, column=10, value=observacoes.upper())   # Observacao
            
            # Salvar arquivo
            wb.save(self.arquivo_cliente)
            wb.close()
            
            messagebox.showinfo("Sucesso", "Contrato atualizado com sucesso!")
            janela.destroy()
            
            # Atualizar lista de contratos
            self.carregar_contratos()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar contrato: {str(e)}")
            try:
                wb.close()
            except:
                pass
    
    def selecionar_contrato(self):
        """Seleciona um contrato para visualizar/editar medições"""
        selecionado = self.tree_contratos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um contrato")
            return
            
        # Obter ID e nome do fornecedor do contrato selecionado
        valores = self.tree_contratos.item(selecionado)['values']
        id_contrato = valores[0]
        nome_fornecedor = valores[1]
        
        # Armazenar contrato atual
        self.contrato_atual = id_contrato
        self.fornecedor_atual = nome_fornecedor
        
        # Atualizar label na aba de medições
        self.lbl_contrato_medicoes.config(text=f"Contrato: {id_contrato} - {nome_fornecedor}")
        
        # Carregar medições do contrato
        self.carregar_medicoes()
        
        # Mudar para a aba de medições
        self.notebook.select(2)  # Índice da aba de medições
    
    # Funções da aba Medições
    def carregar_medicoes(self):
        """Carrega as medições do contrato selecionado"""
        try:
            if not self.contrato_atual:
                messagebox.showwarning("Aviso", "Selecione um contrato primeiro!")
                return
                
            # Limpar treeview
            for item in self.tree_medicoes.get_children():
                self.tree_medicoes.delete(item)
                
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            
            # Verificar se a aba de medições existe
            if "Medicoes" not in wb.sheetnames:
                messagebox.showerror("Erro", "Aba de medições não encontrada!")
                wb.close()
                return
                
            ws = wb["Medicoes"]
            
            # Percorrer as linhas (pulando o cabeçalho)
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == self.contrato_atual:  # Filtrar pelo ID do contrato atual
                    # Formatação de dados
                    try:
                        data_medicao = row[4].strftime('%d/%m/%Y') if isinstance(row[4], datetime) else row[4]
                        data_pagamento = row[5].strftime('%d/%m/%Y') if isinstance(row[5], datetime) else row[5]
                        valor = f"R$ {float(row[7]):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') if row[5] else "R$ 0,00"
                    except (ValueError, TypeError, AttributeError):
                        data_medicao = str(row[4] or "")
                        data_pagamento = str(row[5] or "")
                        valor = str(row[7] or "R$ #.##0,00")
                    
                    # Adicionar à treeview
                    self.tree_medicoes.insert('', 'end', values=(
                        row[1],             # ID Medição
                        data_medicao,       # Data Medição
                        data_pagamento,     # Data Pagamento
                        row[6],             # Referência
                        valor,              # Valor
                        row[8]              # Status
                    ))
                
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar medições: {str(e)}")
    
    def nova_medicao(self):
        """Abre janela para cadastro de nova medição"""
        if not self.contrato_atual:
            messagebox.showwarning("Aviso", "Selecione um contrato primeiro!")
            return
        
        # Verificar se o contrato tem saldo disponível
        saldo = self.verificar_saldo_contrato()
        if saldo <= 0:
            messagebox.showwarning("Aviso", "Este contrato não possui saldo disponível para novas medições!")
            return
            
        # Criar janela de cadastro
        janela = tk.Toplevel(self.root)
        janela.title("Nova Medição")
        
        # Centralizar a janela em relação à janela principal
        self.centralizar_janela(janela, 600, 400)
        
        # Frame principal
        frame = ttk.Frame(janela, padding="10")
        frame.pack(fill='both', expand=True)
        
        # Informações do contrato
        frame_info = ttk.LabelFrame(frame, text="Informações do Contrato")
        frame_info.pack(fill='x', pady=5)
        
        # Obter informações do contrato
        try:
            contrato = self.obter_dados_contrato(self.contrato_atual)
            
            if contrato:
                ttk.Label(frame_info, text=f"ID: {contrato['id']}").grid(row=0, column=0, padx=5, pady=2, sticky='w')
                ttk.Label(frame_info, text=f"Fornecedor: {contrato['nome']}").grid(row=0, column=1, padx=5, pady=2, sticky='w')
                ttk.Label(frame_info, text=f"Valor Global: {formatar_moeda_br(contrato['valor_global'])}").grid(row=1, column=0, padx=5, pady=2, sticky='w')
                ttk.Label(frame_info, text=f"Saldo: {formatar_moeda_br(contrato['saldo'])}").grid(row=1, column=1, padx=5, pady=2, sticky='w')
        except Exception as e:
            ttk.Label(frame_info, text=f"Erro ao carregar dados: {str(e)}").grid(row=0, column=0, padx=5, pady=2, sticky='w')
        
        # Frame para dados da medição
        frame_medicao = ttk.LabelFrame(frame, text="Dados da Medição")
        frame_medicao.pack(fill='x', pady=10)
        
        # Data da Medição
        ttk.Label(frame_medicao, text="Data da Medição:*").grid(row=0, column=0, padx=5, pady=5, sticky='e')
        data_medicao = DateEntry(frame_medicao, width=12, background='darkblue',
                              foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        data_medicao.grid(row=0, column=1, padx=5, pady=5, sticky='w')
        data_medicao.set_date(datetime.now())
        
        # Data de Pagamento
        ttk.Label(frame_medicao, text="Data de Pagamento:*").grid(row=1, column=0, padx=5, pady=5, sticky='e')
        data_pagamento = DateEntry(frame_medicao, width=12, background='darkblue',
                                foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
        data_pagamento.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Calcular data de pagamento padrão (30 dias após a medição)
        data_pag_default = datetime.now() + relativedelta(days=30)
        data_pagamento.set_date(data_pag_default)
        
        # Referência
        ttk.Label(frame_medicao, text="Referência:*").grid(row=2, column=0, padx=5, pady=5, sticky='e')
        referencia = ttk.Entry(frame_medicao, width=40)
        referencia.grid(row=2, column=1, padx=5, pady=5, sticky='w')
        
        # Valor
        ttk.Label(frame_medicao, text="Valor (R$):*").grid(row=3, column=0, padx=5, pady=5, sticky='e')
        valor = ttk.Entry(frame_medicao, width=20)
        valor.grid(row=3, column=1, padx=5, pady=5, sticky='w')
        
        # Saldo Máximo (como referência)
        ttk.Label(frame_medicao, text=f"Saldo disponível: {formatar_moeda_br(contrato['saldo'])}").grid(row=3, column=2, padx=5, sticky='w')

        # Observações
        ttk.Label(frame_medicao, text="Observações:").grid(row=4, column=0, padx=5, pady=5, sticky='ne')
        observacoes = tk.Text(frame_medicao, width=40, height=4)
        observacoes.grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky='w')
        
        # Frame para botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, 
                 text="Salvar", 
                 command=lambda: self.salvar_medicao(
                     janela,
                     self.contrato_atual,
                     data_medicao.get(),
                     data_pagamento.get(),
                     referencia.get(),
                     valor.get(),
                     observacoes.get("1.0", "end-1c")
                 )).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                 text="Cancelar", 
                 command=janela.destroy).pack(side='left', padx=5)
        
    def verificar_saldo_contrato(self):
        """Verifica o saldo disponível do contrato atual"""
        try:
            # Obter dados do contrato
            contrato = self.obter_dados_contrato(self.contrato_atual)
            if not contrato:
                return 0
                
            return float(contrato['saldo'])
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao verificar saldo: {str(e)}")
            return 0
            
    def obter_dados_contrato(self, id_contrato):
        """Obtém os dados completos de um contrato pelo ID"""
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb["Contratos_Medicao"]
            
            dados_contrato = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == id_contrato:
                    dados_contrato = {
                        'id': row[0],
                        'cnpj': row[1],
                        'nome': row[2],
                        'descricao': row[3],
                        'data_inicio': row[4],
                        'valor_global': row[5],
                        'valor_pago': row[6],
                        'saldo': row[7],
                        'status': row[8],
                        'observacao': row[9]
                    }
                    break
                    
            wb.close()
            return dados_contrato
            
        except Exception as e:
            logger.error(f"Erro ao obter dados do contrato: {str(e)}")
            return None
    
    def salvar_medicao(self, janela, id_contrato, data_medicao, data_pagamento, referencia, valor, observacoes):
        """Salva uma nova medição"""
        try:
            # Validar campos obrigatórios
            if not data_medicao or not data_pagamento or not referencia or not valor:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
                
            # Validar valor
            try:
                valor_float = float(valor.replace(',', '.'))
                if valor_float <= 0:
                    messagebox.showerror("Erro", "Valor deve ser maior que zero!")
                    return
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!")
                return
                
            # Verificar saldo disponível
            saldo = self.verificar_saldo_contrato()
            if valor_float > saldo:
                messagebox.showerror("Erro", f"Valor excede o saldo disponível de R$ {saldo:.2f}!")
                return
                
            # Converter datas
            try:
                data_med = datetime.strptime(data_medicao, '%d/%m/%Y')
                data_pag = datetime.strptime(data_pagamento, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida! Use o formato dd/mm/aaaa")
                return
                
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            ws_medicoes = wb["Medicoes"]
            
            # Gerar ID da medição (próximo número sequencial para o contrato)
            next_id = 1
            for row in ws_medicoes.iter_rows(min_row=2, values_only=True):
                if row[0] == id_contrato and row[1] and isinstance(row[1], int) and row[1] >= next_id:
                    next_id = row[1] + 1
            
            # Obter dados do fornecedor
            contrato = self.obter_dados_contrato(id_contrato)
            if not contrato:
                messagebox.showerror("Erro", "Não foi possível obter dados do contrato!")
                wb.close()
                return
                
            # Adicionar nova medição
            proxima_linha = ws_medicoes.max_row + 1
            ws_medicoes.cell(row=proxima_linha, column=1, value=id_contrato)  # ID_Contrato
            ws_medicoes.cell(row=proxima_linha, column=2, value=next_id)      # ID_Medicao

            # Salvar CNPJ como texto para preservar zeros à esquerda
            cnpj_limpo = ''.join(filter(str.isdigit, str(contrato['cnpj'])))

            # Determinar se é CPF ou CNPJ e garantir formatação adequada
            if len(cnpj_limpo) <= 11:
                # É um CPF
                cnpj_formatado = cnpj_limpo.zfill(11)
                cnpj_formatado = f"{cnpj_formatado[:3]}.{cnpj_formatado[3:6]}.{cnpj_formatado[6:9]}-{cnpj_formatado[9:]}"
            else:
                # É um CNPJ
                cnpj_formatado = cnpj_limpo.zfill(14)
                cnpj_formatado = f"{cnpj_formatado[:2]}.{cnpj_formatado[2:5]}.{cnpj_formatado[5:8]}/{cnpj_formatado[8:12]}-{cnpj_formatado[12:]}"

            # Garantir que o CNPJ seja armazenado como texto na planilha
            cnpj_cell = ws_medicoes.cell(row=proxima_linha, column=3, value=cnpj_formatado)

            ws_medicoes.cell(row=proxima_linha, column=4, value=contrato['nome'])  # Nome_Fornecedor
            ws_medicoes.cell(row=proxima_linha, column=5, value=data_med)     # Data_Medicao
            ws_medicoes.cell(row=proxima_linha, column=6, value=data_pag)     # Data_Pagamento
            ws_medicoes.cell(row=proxima_linha, column=7, value=referencia.upper())  # Referencia
            
            # Formatação para valor
            valor_cell = ws_medicoes.cell(row=proxima_linha, column=8, value=valor_float)  # Valor
            valor_cell.number_format = '#.##0,00'
            
            ws_medicoes.cell(row=proxima_linha, column=9, value="PENDENTE")   # Status
            ws_medicoes.cell(row=proxima_linha, column=10, value=None)        # Data_Lancamento
            ws_medicoes.cell(row=proxima_linha, column=11, value=observacoes.upper())  # Observacao
            
            # Atualizar saldo do contrato na aba Contratos_Medicao
            ws_contratos = wb["Contratos_Medicao"]
            for idx, row in enumerate(ws_contratos.iter_rows(min_row=2, max_col=1, values_only=True), 2):
                if row[0] == id_contrato:
                    # Obter valores atuais
                    valor_global = ws_contratos.cell(row=idx, column=6).value or 0
                    valor_pago = ws_contratos.cell(row=idx, column=7).value or 0
                    
                    # Atualizar valor pago
                    novo_valor_pago = valor_pago + valor_float
                    valor_pago_cell = ws_contratos.cell(row=idx, column=7, value=novo_valor_pago)
                    valor_pago_cell.number_format = '#.##0,00'
                    
                    # Atualizar saldo
                    novo_saldo = valor_global - novo_valor_pago
                    saldo_cell = ws_contratos.cell(row=idx, column=8, value=novo_saldo)
                    saldo_cell.number_format = '#.##0,00'
                    
                    # Se saldo zerou, atualizar status para CONCLUÍDO
                    if novo_saldo <= 0:
                        ws_contratos.cell(row=idx, column=9, value="CONCLUÍDO")
                    
                    break
            
            # Salvar arquivo
            wb.save(self.arquivo_cliente)
            wb.close()
            
            messagebox.showinfo("Sucesso", "Medição cadastrada com sucesso!")
            janela.destroy()
            
            # Atualizar lista de medições
            self.carregar_medicoes()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar medição: {str(e)}")
            try:
                wb.close()
            except:
                pass
    
    def editar_medicao(self):
        """Edita a medição selecionada"""
        selecionado = self.tree_medicoes.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma medição para editar")
            return
            
        # Obter ID da medição selecionada
        valores = self.tree_medicoes.item(selecionado)['values']
        id_medicao = valores[0]
        
        # Obter dados completos da medição
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb["Medicoes"]
            
            dados_medicao = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == self.contrato_atual and row[1] == id_medicao:
                    dados_medicao = {
                        'id_contrato': row[0],
                        'id_medicao': row[1],
                        'cnpj': row[2],
                        'nome': row[3],
                        'data_medicao': row[4],
                        'data_pagamento': row[5],
                        'referencia': row[6],
                        'valor': row[7],
                        'status': row[8],
                        'data_lancamento': row[9],
                        'observacao': row[10]
                    }
                    break
                    
            wb.close()
            
            if not dados_medicao:
                messagebox.showerror("Erro", "Medição não encontrada!")
                return
                
            # Verificar se já foi lançada
            if dados_medicao['status'] != 'PENDENTE':
                messagebox.showwarning("Aviso", "Esta medição já foi lançada e não pode ser editada!")
                return
                
            # Criar janela de edição
            janela = tk.Toplevel(self.root)
            janela.title("Editar Medição")
            
            # Centralizar a janela em relação à janela principal
            self.centralizar_janela(janela, 650, 430)
            
            # Frame principal
            frame = ttk.Frame(janela, padding="10")
            frame.pack(fill='both', expand=True)

            # Informações do contrato
            frame_info = ttk.LabelFrame(frame, text="Informações da Medição")
            frame_info.pack(fill='x', pady=5)
            
            ttk.Label(frame_info, text=f"Contrato: {dados_medicao['id_contrato']}").grid(row=0, column=0, padx=5, pady=2, sticky='w')
            ttk.Label(frame_info, text=f"Medição: {dados_medicao['id_medicao']}").grid(row=0, column=1, padx=5, pady=2, sticky='w')
            ttk.Label(frame_info, text=f"Fornecedor: {dados_medicao['nome']}").grid(row=1, column=0, columnspan=2, padx=5, pady=2, sticky='w')
            
            # Frame para dados da medição
            frame_medicao = ttk.LabelFrame(frame, text="Dados da Medição")
            frame_medicao.pack(fill='x', pady=10)
            
            # Data da Medição
            ttk.Label(frame_medicao, text="Data da Medição:*").grid(row=0, column=0, padx=5, pady=5, sticky='e')
            data_medicao = DateEntry(frame_medicao, width=12, background='darkblue',
                                  foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
            data_medicao.grid(row=0, column=1, padx=5, pady=5, sticky='w')
            if isinstance(dados_medicao['data_medicao'], datetime):
                data_medicao.set_date(dados_medicao['data_medicao'])
            
            # Data de Pagamento
            ttk.Label(frame_medicao, text="Data de Pagamento:*").grid(row=1, column=0, padx=5, pady=5, sticky='e')
            data_pagamento = DateEntry(frame_medicao, width=12, background='darkblue',
                                    foreground='white', borderwidth=2, date_pattern='dd/mm/yyyy')
            data_pagamento.grid(row=1, column=1, padx=5, pady=5, sticky='w')
            if isinstance(dados_medicao['data_pagamento'], datetime):
                data_pagamento.set_date(dados_medicao['data_pagamento'])
            
            # Referência
            ttk.Label(frame_medicao, text="Referência:*").grid(row=2, column=0, padx=5, pady=5, sticky='e')
            referencia = ttk.Entry(frame_medicao, width=80)
            referencia.grid(row=2, column=1, padx=5, pady=5, sticky='w')
            referencia.insert(0, dados_medicao['referencia'] if dados_medicao['referencia'] else "")
            
            # Valor original (somente leitura)
            ttk.Label(frame_medicao, text="Valor Original (R$):").grid(row=3, column=0, padx=5, pady=5, sticky='e')
            valor_original = ttk.Entry(frame_medicao, width=15)
            valor_original.grid(row=3, column=1, padx=5, pady=5, sticky='w')

            # Formatar o valor usando a função antes de inserir no campo
            valor_formatado = formatar_moeda_br(dados_medicao['valor']).replace('R$ ', '')  # Remove o "R$ " para ficar só o número
            valor_original.insert(0, valor_formatado)
            valor_original.config(state='readonly')
            
            # Valor novo
            ttk.Label(frame_medicao, text="Novo Valor (R$):*").grid(row=4, column=0, padx=5, pady=5, sticky='e')
            valor_novo = ttk.Entry(frame_medicao, width=15)
            valor_novo.grid(row=4, column=1, padx=5, pady=5, sticky='w')
            valor_novo.insert(0, str(dados_medicao['valor']))
            
            # Observações
            ttk.Label(frame_medicao, text="Observações:").grid(row=5, column=0, padx=5, pady=5, sticky='ne')
            observacoes = tk.Text(frame_medicao, width=40, height=4)
            observacoes.grid(row=5, column=1, padx=5, pady=5, sticky='w')
            observacoes.insert("1.0", dados_medicao['observacao'] if dados_medicao['observacao'] else "")
            
            # Frame para botões
            frame_botoes = ttk.Frame(frame)
            frame_botoes.pack(fill='x', pady=10)
            
            ttk.Button(frame_botoes, 
                     text="Salvar", 
                     command=lambda: self.atualizar_medicao(
                         janela,
                         self.contrato_atual,
                         id_medicao,
                         data_medicao.get(),
                         data_pagamento.get(),
                         referencia.get(),
                         valor_original.get(),
                         valor_novo.get(),
                         observacoes.get("1.0", "end-1c")
                     )).pack(side='left', padx=5)
            
            ttk.Button(frame_botoes, 
                     text="Cancelar", 
                     command=janela.destroy).pack(side='left', padx=5)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao editar medição: {str(e)}")
    
    def atualizar_medicao(self, janela, id_contrato, id_medicao, data_medicao, data_pagamento, 
                        referencia, valor_original, valor_novo, observacoes):
        """Atualiza os dados de uma medição"""
        try:
            # Validar campos obrigatórios
            if not data_medicao or not data_pagamento or not referencia or not valor_novo:
                messagebox.showerror("Erro", "Preencha todos os campos obrigatórios!")
                return
                
            # Validar valor
            try:
                valor_org = float(valor_original.replace(',', '.'))
                valor_novo_float = float(valor_novo.replace(',', '.'))
                if valor_novo_float <= 0:
                    messagebox.showerror("Erro", "Valor deve ser maior que zero!")
                    return
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!")
                return
                
            # Se valor mudou, verificar saldo do contrato
            if valor_novo_float > valor_org:
                # Calcular a diferença
                diferenca = valor_novo_float - valor_org
                
                # Verificar se o contrato tem saldo para a diferença
                saldo = self.verificar_saldo_contrato()
                if diferenca > saldo:
                    messagebox.showerror("Erro", 
                                       f"O aumento de R$ {diferenca:.2f} excede o saldo disponível de R$ {saldo:.2f}!")
                    return
                
            # Converter datas
            try:
                data_med = datetime.strptime(data_medicao, '%d/%m/%Y')
                data_pag = datetime.strptime(data_pagamento, '%d/%m/%Y')
            except ValueError:
                messagebox.showerror("Erro", "Data inválida! Use o formato dd/mm/aaaa")
                return
                
            # Abrir arquivo do cliente
            wb = load_workbook(self.arquivo_cliente)
            ws_medicoes = wb["Medicoes"]
            
            # Buscar medição pelo ID
            medicao_row = None
            for idx, row in enumerate(ws_medicoes.iter_rows(min_row=2, values_only=True), 2):
                if row[0] == id_contrato and row[1] == id_medicao:
                    medicao_row = idx
                    break
                    
            if not medicao_row:
                messagebox.showerror("Erro", "Medição não encontrada!")
                wb.close()
                return
                
            # Atualizar dados da medição
            ws_medicoes.cell(row=medicao_row, column=5, value=data_med)  # Data_Medicao
            ws_medicoes.cell(row=medicao_row, column=6, value=data_pag)  # Data_Pagamento
            ws_medicoes.cell(row=medicao_row, column=7, value=referencia.upper())  # Referencia
            
            # Atualizar valor apenas se mudou
            if valor_novo_float != valor_org:
                valor_cell = ws_medicoes.cell(row=medicao_row, column=8, value=valor_novo_float)  # Valor
                valor_cell.number_format = '#.##0,00'
                
                # Atualizar valores do contrato
                ws_contratos = wb["Contratos_Medicao"]
                for idx, row in enumerate(ws_contratos.iter_rows(min_row=2, max_col=1, values_only=True), 2):
                    if row[0] == id_contrato:
                        # Obter valores atuais
                        valor_global = ws_contratos.cell(row=idx, column=6).value or 0
                        valor_pago = ws_contratos.cell(row=idx, column=7).value or 0
                        
                        # Ajustar com a diferença
                        novo_valor_pago = valor_pago + (valor_novo_float - valor_org)
                        valor_pago_cell = ws_contratos.cell(row=idx, column=7, value=novo_valor_pago)
                        valor_pago_cell.number_format = '#.##0,00'
                        
                        # Atualizar saldo
                        novo_saldo = valor_global - novo_valor_pago
                        saldo_cell = ws_contratos.cell(row=idx, column=8, value=novo_saldo)
                        saldo_cell.number_format = '#.##0,00'
                        
                        # Se saldo zerou, atualizar status para CONCLUÍDO
                        if novo_saldo <= 0:
                            ws_contratos.cell(row=idx, column=9, value="CONCLUÍDO")
                        
                        break
            
            # Atualizar observações
            ws_medicoes.cell(row=medicao_row, column=11, value=observacoes.upper())  # Observacao
            
            # Salvar arquivo
            wb.save(self.arquivo_cliente)
            wb.close()
            
            messagebox.showinfo("Sucesso", "Medição atualizada com sucesso!")
            janela.destroy()
            
            # Atualizar lista de medições
            self.carregar_medicoes()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao atualizar medição: {str(e)}")
            try:
                wb.close()
            except:
                pass
    
    def lancar_medicao(self):
        """Lança a medição selecionada na planilha do cliente"""
        # Verificar se há seleção
        selecionado = self.tree_medicoes.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma medição para lançar")
            return
            
        # Obter ID da medição selecionada
        valores = self.tree_medicoes.item(selecionado)['values']
        id_medicao = valores[0]
        
        # Verificar se já foi lançada
        if valores[5] != "PENDENTE":
            messagebox.showwarning("Aviso", "Esta medição já foi lançada!")
            return
            
        # Obter dados completos da medição
        try:
            wb = load_workbook(self.arquivo_cliente)
            ws = wb["Medicoes"]
            
            dados_medicao = None
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row[0] == self.contrato_atual and row[1] == id_medicao:
                    dados_medicao = {
                        'id_contrato': row[0],
                        'id_medicao': row[1],
                        'cnpj': row[2],
                        'nome': row[3],
                        'data_medicao': row[4],
                        'data_pagamento': row[5],
                        'referencia': row[6],
                        'valor': row[7],
                        'status': row[8],
                        'data_lancamento': row[9],
                        'observacao': row[10]
                    }
                    break
                    
            if not dados_medicao:
                messagebox.showerror("Erro", "Medição não encontrada!")
                wb.close()
                return
                
            # Calcular data de relatório
            data_pagamento = dados_medicao['data_pagamento']
            dt_vencto = data_pagamento
            
            # Estimar data para relatório
            hoje = datetime.now()
            data_rel = hoje
            
            # Ajustar para dia 5 ou 20 mais próximo
            if hoje.day < 5:
                data_rel = hoje.replace(day=5)
            elif hoje.day < 20:
                data_rel = hoje.replace(day=20)
            else:
                # Próximo mês, dia 5
                if hoje.month == 12:
                    data_rel = hoje.replace(year=hoje.year+1, month=1, day=5)
                else:
                    data_rel = hoje.replace(month=hoje.month+1, day=5)
                    
            # Verificar se já existe dados_para_incluir e criar se não existir
            if not hasattr(self, 'dados_para_incluir'):
                self.dados_para_incluir = []
                
            # Preparar dados do lançamento
            dados_lancamento = {
                'data': data_rel.strftime('%d/%m/%Y'),  # Data do relatório
                'tp_desp': '2',  # Tipo de despesa (serviço/material)
                'cnpj_cpf': dados_medicao['cnpj'],
                'nome': dados_medicao['nome'],
                'categoria': 'SERV',  # Categoria padrão (ajustar conforme necessário)
                'referencia': dados_medicao['referencia'],
                'nf': '',  # NF em branco (pode ser ajustado)
                'vr_unit': str(dados_medicao['valor']),
                'dias': 1,
                'valor': str(dados_medicao['valor']),
                'dt_vencto': dt_vencto.strftime('%d/%m/%Y'),
                'dados_bancarios': self.obter_dados_bancarios(dados_medicao['cnpj']),
                'observacao': f"MEDIÇÃO {id_medicao} - {dados_medicao['observacao'] or ''}",
                'forma_pagamento': 'PIX'  # Forma padrão
            }
            
            # Adicionar à lista de dados para incluir
            self.dados_para_incluir.append(dados_lancamento)
            
            # Atualizar status da medição na planilha
            for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
                if row[0] == self.contrato_atual and row[1] == id_medicao:
                    ws.cell(row=idx, column=9, value="LANÇADO")  # Status
                    ws.cell(row=idx, column=10, value=hoje)      # Data_Lancamento
                    break
                    
            # Salvar alterações
            wb.save(self.arquivo_cliente)
            wb.close()
            
            # Confirmar ao usuário
            messagebox.showinfo("Sucesso", 
                             f"Medição lançada com sucesso!\n\n"
                             f"Fornecedor: {dados_medicao['nome']}\n"
                             f"Valor: R$ {float(dados_medicao['valor']):.2f}\n"
                             f"Vencimento: {dt_vencto.strftime('%d/%m/%Y')}")
            
            # Atualizar lista de medições
            self.carregar_medicoes()
            
            # Perguntar se deseja enviar os dados imediatamente
            if messagebox.askyesno("Enviar Dados", 
                                 "Deseja enviar os dados para a planilha do cliente agora?"):
                self.enviar_dados()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao lançar medição: {str(e)}")
            try:
                wb.close()
            except:
                pass
    
    def obter_dados_bancarios(self, cnpj):
        """Obtém os dados bancários do fornecedor com tratamento robusto para CNPJ/CPF"""
        try:
            # Garantir que estamos trabalhando com string
            cnpj_str = str(cnpj)
            
            # Limpar a string, removendo caracteres não numéricos
            cnpj_limpo = ''.join(filter(str.isdigit, cnpj_str))
            
            # Determinar se é CPF (11 dígitos) ou CNPJ (14 dígitos) e garantir o preenchimento com zeros
            if len(cnpj_limpo) <= 11:
                # É um CPF, garantir 11 dígitos
                cnpj_formatado = cnpj_limpo.zfill(11)
                # Formatar como XXX.XXX.XXX-XX
                cnpj_formatado = f"{cnpj_formatado[:3]}.{cnpj_formatado[3:6]}.{cnpj_formatado[6:9]}-{cnpj_formatado[9:]}"
            else:
                # É um CNPJ, garantir 14 dígitos
                cnpj_formatado = cnpj_limpo.zfill(14)
                # Formatar como XX.XXX.XXX/XXXX-XX
                cnpj_formatado = f"{cnpj_formatado[:2]}.{cnpj_formatado[2:5]}.{cnpj_formatado[5:8]}/{cnpj_formatado[8:12]}-{cnpj_formatado[12:]}"
            
            # Usar a função auxiliar importada
            dados_bancarios = buscar_dados_bancarios_fornecedor(cnpj_formatado, "PIX")
            
            # Verificar se obteve dados bancários
            if not dados_bancarios or dados_bancarios == "DADOS BANCÁRIOS NÃO CADASTRADOS":
                # Tentar sem formatação
                dados_bancarios = buscar_dados_bancarios_fornecedor(cnpj_limpo, "PIX")
                
            return dados_bancarios if dados_bancarios else "DADOS BANCÁRIOS NÃO CADASTRADOS"
        except Exception as e:
            logger.error(f"Erro ao obter dados bancários: {str(e)}")
            return "DADOS BANCÁRIOS NÃO CADASTRADOS"
    
    def enviar_dados(self):
        """Envia os dados para a planilha do cliente"""
        try:
            if not self.dados_para_incluir:
                messagebox.showwarning("Aviso", "Não há dados para enviar!")
                return
                
            # Verificar arquivo do cliente
            arquivo_cliente = PASTA_CLIENTES / f"{self.cliente_atual}.xlsx"
            
            try:
                workbook = load_workbook(arquivo_cliente)
            except PermissionError:
                messagebox.showerror(
                    "Erro", 
                    f"A planilha '{self.cliente_atual}.xlsx' está aberta!\n\n"
                    "Por favor:\n"
                    "1. Feche a planilha\n"
                    "2. Clique em OK\n"
                    "3. Tente enviar novamente"
                )
                return
            
            sheet = workbook["Dados"]
            
            # Processar registros
            for dados in self.dados_para_incluir:
                proxima_linha = sheet.max_row + 1
                
                # Converter e salvar data de referência
                data_rel = datetime.strptime(dados['data'], '%d/%m/%Y')
                data_cell = sheet.cell(row=proxima_linha, column=1, value=data_rel)
                data_cell.number_format = 'DD/MM/YYYY'

                # Converter tipo de despesa para número
                tp_desp_cell = sheet.cell(row=proxima_linha, column=2, value=int(dados['tp_desp']))
                tp_desp_cell.number_format = '0'

                sheet.cell(row=proxima_linha, column=3, value=dados['cnpj_cpf'])
                sheet.cell(row=proxima_linha, column=4, value=dados['nome'])
                sheet.cell(row=proxima_linha, column=5, value=dados['referencia'])
                sheet.cell(row=proxima_linha, column=6, value=dados['nf'])

                # Valores financeiros
                vr_unit = float(dados['vr_unit'].replace(',', '.'))
                vr_unit_cell = sheet.cell(row=proxima_linha, column=7, value=vr_unit)
                aplicar_formatacao_celula(vr_unit_cell)

                sheet.cell(row=proxima_linha, column=8, value=int(dados.get('dias', 1)))

                valor = float(dados['valor'].replace(',', '.'))
                valor_cell = sheet.cell(row=proxima_linha, column=9, value=valor)
                aplicar_formatacao_celula(valor_cell)

                dt_vencto = datetime.strptime(dados['dt_vencto'], '%d/%m/%Y')
                dt_vencto_cell = sheet.cell(row=proxima_linha, column=10, value=dt_vencto)
                dt_vencto_cell.number_format = 'DD/MM/YYYY'

                sheet.cell(row=proxima_linha, column=11, value=dados['categoria'])
                sheet.cell(row=proxima_linha, column=12, value=dados['dados_bancarios'])
                sheet.cell(row=proxima_linha, column=13, value=dados['observacao'])

            try:
                # Tentar salvar o arquivo
                workbook.save(arquivo_cliente)
                messagebox.showinfo("Sucesso", "Dados salvos com sucesso na planilha do cliente!")
                    
                # Limpar após salvar
                self.dados_para_incluir.clear()
                
            except PermissionError:
                messagebox.showerror(
                    "Erro", 
                    f"Não foi possível salvar! A planilha '{self.cliente_atual}.xlsx' está aberta.\n\n"
                    "Por favor:\n"
                    "1. Feche a planilha\n"
                    "2. Clique em OK\n"
                    "3. Tente enviar novamente"
                )
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar arquivo: {str(e)}")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar dados: {str(e)}")
    
    def voltar_menu(self):
        """Volta ao menu principal"""
        if hasattr(self, 'dados_para_incluir') and self.dados_para_incluir:
            if messagebox.askyesno("Aviso", "Existem dados não enviados. Deseja enviá-los antes de sair?"):
                self.enviar_dados()
                
        # Fechar a janela atual
        self.root.destroy()
        
        # Mostrar janela principal
        if self.menu_principal:
            self.menu_principal.deiconify()
            self.menu_principal.lift()
            self.menu_principal.focus_force()


def main():
    """Função principal para executar o módulo de forma independente"""
    app = GestaoMedicoes()
    app.root.mainloop()
    
if __name__ == "__main__":
    main()