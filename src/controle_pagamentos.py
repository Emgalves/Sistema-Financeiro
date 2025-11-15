import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook
from datetime import datetime, date
from dateutil.relativedelta import relativedelta
import os
import sys
from pathlib import Path

# No início de controle_pagamentos.py
from src.correcao_monetaria import GerenciadorCorrecaoMonetaria
from src.config.utils import (
    PASTA_CLIENTES,
    formatar_moeda,
    validar_data,
    formatar_valor_excel,
    aplicar_formatacao_celula,
    buscar_dados_bancarios_fornecedor
)

class ControlePagamentos:
    def __init__(self, parent=None):
        self.parent = parent
        # Se parent não for especificado, criar uma janela Tk, caso contrário criar uma Toplevel
        if parent is None:
            self.root = tk.Tk()
            self.is_independent = True  # Marcar como janela independente
        else:
            self.root = tk.Toplevel(parent)
            self.is_independent = False  # Marcar como janela secundária
        
        self.root.title("Controle de Pagamentos de Taxas")
        self.root.geometry("1400x900+50+50")  # Aumentado e reposicionado
        
        # Forçar a janela para frente
        self.root.lift()
        self.root.attributes('-topmost', True)
        self.root.after(100, lambda: self.root.attributes('-topmost', False))
        
        # Variáveis de controle
        self.cliente_selecionado = None
        self.parcelas_selecionadas = []
        self.scrollbar_y = None
        self.scrollbar_x = None
        self.valor_editado = {}  # Dicionário para armazenar valores editados
        
        self.setup_gui()
    
    def setup_gui(self):
        # Frame principal SEM scroll para melhor controle do layout
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Frame seleção de cliente
        frame_cliente = ttk.LabelFrame(main_frame, text="Selecione o Cliente")
        frame_cliente.pack(fill='x', pady=5)
        
        self.cliente_combo = ttk.Combobox(frame_cliente, state='readonly', width=50)
        self.cliente_combo.pack(side='left', padx=5, pady=5)
        self.cliente_combo.bind('<<ComboboxSelected>>', self.carregar_parcelas)
        
        # Frame lista de parcelas - EXPANDIDO para usar todo espaço disponível
        self.frame_parcelas = ttk.LabelFrame(main_frame, text="Parcelas Pendentes")
        self.frame_parcelas.pack(fill='both', expand=True, pady=5)
        
        # Container para treeview e scrollbars
        self.tree_container = ttk.Frame(self.frame_parcelas)
        self.tree_container.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Treeview com coluna adicional para valor editado
        colunas = ('Nº Contrato', 'Nº Parcela', 'CNPJ', 'Adm', 'Eventos/Fases', 
                   'Valor Original', 'Valor a Pagar', 'Status', 'Data Pagamento')
        self.tree_parcelas = ttk.Treeview(self.tree_container, columns=colunas, 
                                         show='headings')
        
        # Configurar colunas com larguras proporcionais ao espaço disponível
        larguras_iniciais = {
            'Nº Contrato': 90,
            'Nº Parcela': 80,
            'CNPJ': 130,
            'Adm': 200,
            'Eventos/Fases': 350,
            'Valor Original': 120,
            'Valor a Pagar': 120,
            'Status': 100,
            'Data Pagamento': 120
        }

        for col in colunas:
            self.tree_parcelas.heading(col, text=col)
            self.tree_parcelas.column(col, width=larguras_iniciais.get(col, 100), minwidth=50)
        
        # Scrollbars
        self.scrollbar_y = ttk.Scrollbar(self.tree_container, orient='vertical',
                                       command=self.tree_parcelas.yview)
        self.scrollbar_x = ttk.Scrollbar(self.tree_container, orient='horizontal',
                                       command=self.tree_parcelas.xview)
        
        # Configurar treeview
        self.tree_parcelas.configure(yscrollcommand=self.scrollbar_y.set,
                                   xscrollcommand=self.scrollbar_x.set)
        
        # Grid layout para treeview e scrollbars
        self.tree_parcelas.grid(row=0, column=0, sticky='nsew')
        self.scrollbar_y.grid(row=0, column=1, sticky='ns')
        self.scrollbar_x.grid(row=1, column=0, sticky='ew')
        
        # Configurar grid weights
        self.tree_container.grid_rowconfigure(0, weight=1)
        self.tree_container.grid_columnconfigure(0, weight=1)
        
        # Bind para duplo clique editar valor
        self.tree_parcelas.bind('<Double-Button-1>', self.editar_valor_parcela)
        
        # Frame para edição de valores selecionados
        frame_edicao = ttk.LabelFrame(main_frame, text="Edição de Valores")
        frame_edicao.pack(fill='x', pady=5)
        
        # Subframe para organizar componentes
        subframe_edicao = ttk.Frame(frame_edicao)
        subframe_edicao.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(subframe_edicao, text="Parcela Selecionada:").grid(row=0, column=0, padx=5, sticky='w')
        self.label_parcela_selecionada = ttk.Label(subframe_edicao, text="Nenhuma", font=('Arial', 9, 'bold'))
        self.label_parcela_selecionada.grid(row=0, column=1, padx=5, sticky='w')
        
        ttk.Label(subframe_edicao, text="Valor Original:").grid(row=0, column=2, padx=5, sticky='w')
        self.label_valor_original = ttk.Label(subframe_edicao, text="R$ 0,00", font=('Arial', 9))
        self.label_valor_original.grid(row=0, column=3, padx=5, sticky='w')
        
        ttk.Label(subframe_edicao, text="Novo Valor:").grid(row=0, column=4, padx=5, sticky='w')
        self.entry_novo_valor = ttk.Entry(subframe_edicao, width=15)
        self.entry_novo_valor.grid(row=0, column=5, padx=5)
        
        ttk.Button(subframe_edicao, text="Aplicar", 
                  command=self.aplicar_valor_editado).grid(row=0, column=6, padx=5)
        
        ttk.Button(subframe_edicao, text="Resetar Todos", 
                  command=self.resetar_valores).grid(row=0, column=7, padx=5)
        
        # Frame para registrar pagamento - POSICIONADO MAIS ACIMA
        frame_pagamento = ttk.LabelFrame(main_frame, text="Registrar Pagamento")
        frame_pagamento.pack(fill='x', pady=5)
        
        # Container interno para melhor organização
        container_pagamento = ttk.Frame(frame_pagamento)
        container_pagamento.pack(fill='x', padx=5, pady=10)
        
        ttk.Label(container_pagamento, text="Data do Pagamento:").pack(side='left', padx=5)
        
        # DateEntry com configuração para garantir visibilidade
        self.data_pagamento = DateEntry(container_pagamento, width=12, locale='pt_BR',
                                      background='darkblue', foreground='white',
                                      borderwidth=2, date_pattern='dd/mm/yyyy',
                                      showweeknumbers=False)
        self.data_pagamento.pack(side='left', padx=5)
        
        ttk.Button(container_pagamento, text="Registrar Pagamento das Parcelas Selecionadas",
                  command=self.registrar_pagamento).pack(side='left', padx=20)
        
        # NOVO: Botão de Vincular
        ttk.Button(container_pagamento, text="Vincular a Lançamento Existente",
                  command=self.vincular_parcela).pack(side='left', padx=5)
        
        # Label informativo
        self.label_info = ttk.Label(container_pagamento, 
                                   text="(Dica: Dê duplo clique em uma parcela para editar o valor)",
                                   foreground='#666')
        self.label_info.pack(side='left', padx=20)
        
        # Frame de botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, text="Voltar ao Menu", 
                  command=self.voltar_menu).pack(side='right', padx=5)
        
        self.carregar_clientes()
        self.verificar_correcoes_pendentes()
    
    def carregar_clientes(self):
        """Carrega lista de clientes no combo"""
        try:
            clientes_pasta = Path(PASTA_CLIENTES)
            if not clientes_pasta.exists():
                messagebox.showerror("Erro", f"Pasta de clientes não encontrada: {PASTA_CLIENTES}")
                return
            
            arquivos_excel = list(clientes_pasta.glob("*.xlsx"))
            arquivos_excel = [f for f in arquivos_excel if not f.name.startswith('~$')]
            
            if not arquivos_excel:
                messagebox.showinfo("Aviso", "Nenhum arquivo de cliente encontrado")
                return
            
            nomes_clientes = [f.stem for f in arquivos_excel]
            self.cliente_combo['values'] = sorted(nomes_clientes)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
    
    def carregar_parcelas(self, event=None):
        """Carrega as parcelas pendentes do cliente selecionado"""
        self.cliente_selecionado = self.cliente_combo.get()
        
        if not self.cliente_selecionado:
            return
        
        # Limpar valores editados
        self.valor_editado.clear()
        
        # Limpar treeview
        for item in self.tree_parcelas.get_children():
            self.tree_parcelas.delete(item)
        
        try:
            arquivo_cliente = Path(PASTA_CLIENTES) / f"{self.cliente_selecionado}.xlsx"
            
            if not arquivo_cliente.exists():
                messagebox.showerror("Erro", f"Arquivo não encontrado: {arquivo_cliente}")
                return
            
            wb = load_workbook(arquivo_cliente, data_only=True)
            
            if 'Contratos_ADM' not in wb.sheetnames:
                messagebox.showinfo("Aviso", "Planilha 'Contratos_ADM' não encontrada")
                wb.close()
                return
            
            ws = wb['Contratos_ADM']
            
            parcelas_encontradas = 0
            
            for row in ws.iter_rows(min_row=3, values_only=True):
                # Colunas esperadas (ajustar conforme estrutura real)
                # Y=25: Num_Contrato, Z=26: Num_Parcela, AA=27: CNPJ_CPF, AB=28: Nome, 
                # AC=29: Eventos_Fases, AD=30: Valor, AE=31: Status, AF=32: Data_Pagamento
                
                if row[24] is None:  # Se Num_Contrato está vazio, linha vazia
                    continue
                
                num_contrato = str(row[24])
                num_parcela = row[25]
                cnpj_cpf = row[26]
                nome = row[27]
                eventos_fases = row[32] if row[32] else ""
                valor = row[29]
                status = row[30] if row[30] else "PENDENTE"
                data_pagamento = row[31]
                
                # Filtrar apenas parcelas PENDENTES
                if status and status.upper() == "PENDENTE":
                    valor_str = formatar_moeda(valor)
                    data_str = data_pagamento.strftime('%d/%m/%Y') if isinstance(data_pagamento, datetime) else ""
                    
                    # Inserir na treeview (valor original = valor a pagar inicialmente)
                    self.tree_parcelas.insert('', 'end', values=(
                        num_contrato,
                        num_parcela,
                        cnpj_cpf,
                        nome,
                        eventos_fases,
                        valor_str,
                        valor_str,  # Valor a pagar começa igual ao original
                        status,
                        data_str
                    ))
                    
                    parcelas_encontradas += 1
            
            wb.close()
            
            if parcelas_encontradas == 0:
                messagebox.showinfo("Informação", "Não há parcelas pendentes para este cliente")
            else:
                self.frame_parcelas.config(text=f"Parcelas Pendentes ({parcelas_encontradas} encontradas)")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar parcelas: {str(e)}")
    
    def editar_valor_parcela(self, event):
        """Permite editar o valor da parcela selecionada"""
        selecao = self.tree_parcelas.selection()
        if not selecao:
            return
        
        item = selecao[0]
        valores = self.tree_parcelas.item(item)['values']
        
        # Atualizar labels
        self.label_parcela_selecionada.config(
            text=f"Contrato {valores[0]} - Parcela {valores[1]}"
        )
        self.label_valor_original.config(text=valores[5])
        
        # Limpar e focar no entry
        self.entry_novo_valor.delete(0, tk.END)
        self.entry_novo_valor.insert(0, valores[6].replace('R$ ', ''))
        self.entry_novo_valor.focus()
    
    def aplicar_valor_editado(self):
        """Aplica o valor editado à parcela selecionada"""
        selecao = self.tree_parcelas.selection()
        if not selecao:
            messagebox.showwarning("Aviso", "Selecione uma parcela primeiro")
            return
        
        try:
            # Obter novo valor
            novo_valor_str = self.entry_novo_valor.get().strip()
            if not novo_valor_str:
                messagebox.showwarning("Aviso", "Informe o novo valor")
                return
            
            # Limpar e converter
            novo_valor_str = novo_valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
            novo_valor = float(novo_valor_str)
            
            if novo_valor <= 0:
                messagebox.showerror("Erro", "Valor deve ser maior que zero")
                return
            
            # Atualizar na treeview
            item = selecao[0]
            valores = list(self.tree_parcelas.item(item)['values'])
            valores[6] = formatar_moeda(novo_valor)  # Coluna "Valor a Pagar"
            
            self.tree_parcelas.item(item, values=valores)
            
            # Armazenar no dicionário
            chave = f"{valores[0]}_{valores[1]}_{valores[2]}"
            self.valor_editado[chave] = novo_valor
            
            messagebox.showinfo("Sucesso", f"Valor atualizado para {formatar_moeda(novo_valor)}")
            
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido. Use formato: 1234.56 ou 1234,56")
    
    def resetar_valores(self):
        """Reseta todos os valores editados para os valores originais"""
        if not messagebox.askyesno("Confirmar", "Resetar todos os valores para os originais?"):
            return
        
        # Limpar dicionário
        self.valor_editado.clear()
        
        # Atualizar treeview
        for item in self.tree_parcelas.get_children():
            valores = list(self.tree_parcelas.item(item)['values'])
            valores[6] = valores[5]  # Valor a pagar = Valor original
            self.tree_parcelas.item(item, values=valores)
        
        messagebox.showinfo("Sucesso", "Valores resetados")
    
    def registrar_pagamento(self):
        """Registra o pagamento das parcelas selecionadas"""
        selecionados = self.tree_parcelas.selection()
        
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos uma parcela para pagar")
            return
        
        if not self.cliente_selecionado:
            messagebox.showwarning("Aviso", "Nenhum cliente selecionado")
            return
        
        try:
            data_pagto = self.data_pagamento.get_date()
            
            arquivo_cliente = Path(PASTA_CLIENTES) / f"{self.cliente_selecionado}.xlsx"
            wb = load_workbook(arquivo_cliente)
            ws_contratos = wb['Contratos_ADM']
            ws_dados = wb['Dados']
            
            parcelas_processadas = []
            
            # Para cada parcela selecionada
            for item in selecionados:
                valores = self.tree_parcelas.item(item)['values']
                num_contrato = str(valores[0])
                num_parcela = int(valores[1])
                cnpj_cpf = str(valores[2])
                nome = str(valores[3])
                
                # Pegar o valor a pagar (editado ou original)
                valor_pagar_str = str(valores[6]).replace('.', '').replace(',', '.')
                valor_pagar = float(valor_pagar_str)
                
                # Buscar total de parcelas para este contrato
                total_parcelas = 0
                for row in ws_contratos.iter_rows(min_row=3, values_only=True):
                    if str(row[24]) == num_contrato:
                        total_parcelas += 1
                
                # Atualizar na aba Contratos_ADM
                for row_idx, row in enumerate(ws_contratos.iter_rows(min_row=3), start=3):
                    if (str(row[24].value) == num_contrato and
                        int(row[25].value) == num_parcela and
                        str(row[26].value) == cnpj_cpf):
                        
                        # Atualizar status e data de pagamento
                        ws_contratos.cell(row=row_idx, column=31, value='PAGO')
                        ws_contratos.cell(row=row_idx, column=32, value=data_pagto)
                        
                        # Se o valor foi editado, registrar em colunas disponíveis após AH
                        if valor_pagar != float(row[29].value):
                            # Usar coluna AI (35) para valor efetivamente pago
                            ws_contratos.cell(row=row_idx, column=35, value=valor_pagar)
                            # Usar coluna AJ (36) para observação sobre diferença
                            ws_contratos.cell(row=row_idx, column=36, 
                                            value=f"Valor pago diferente do original: R$ {valor_pagar:,.2f}")
                        
                        # Registrar na aba Dados
                        proxima_linha = ws_dados.max_row + 1
                        
                        # Calcular data de referência
                        data_pagto_informada = data_pagto
                        if data_pagto_informada.day <= 5:
                            data_ref = data_pagto_informada.replace(day=5)
                        elif data_pagto_informada.day <= 20:
                            data_ref = data_pagto_informada.replace(day=20)
                        else:
                            if data_pagto_informada.month == 12:
                                data_ref = data_pagto_informada.replace(year=data_pagto_informada.year + 1, month=1, day=5)
                            else:
                                data_ref = data_pagto_informada.replace(month=data_pagto_informada.month + 1, day=5)

                        ws_dados.cell(row=proxima_linha, column=1, value=data_ref)
                        ws_dados.cell(row=proxima_linha, column=1).number_format = 'DD/MM/YYYY'
                        
                        # Tipo e dados
                        ws_dados.cell(row=proxima_linha, column=2, value=2)
                        ws_dados.cell(row=proxima_linha, column=3, value=cnpj_cpf)
                        ws_dados.cell(row=proxima_linha, column=4, value=nome)
                        ws_dados.cell(row=proxima_linha, column=5, value=f"ADM OBRA - PARC. {num_parcela}/{total_parcelas}")
                        
                        # Usar o valor editado/pago
                        valor_formatado = formatar_valor_excel(valor_pagar)

                        # Valores com formato brasileiro
                        cell_vr_unit = ws_dados.cell(row=proxima_linha, column=7, value=valor_formatado)
                        cell_vr_unit.number_format = '#,##0.00'
                        cell_vr_unit = aplicar_formatacao_celula(cell_vr_unit)

                        ws_dados.cell(row=proxima_linha, column=8, value=1)

                        cell_valor = ws_dados.cell(row=proxima_linha, column=9, value=valor_formatado)
                        cell_valor.number_format = '#,##0.00'
                        cell_valor = aplicar_formatacao_celula(cell_valor)
                        
                        # Data de pagamento
                        ws_dados.cell(row=proxima_linha, column=10, value=data_pagto)
                        ws_dados.cell(row=proxima_linha, column=10).number_format = 'DD/MM/YYYY'
                        
                        ws_dados.cell(row=proxima_linha, column=11, value='TAX')

                        # Buscar dados bancários do fornecedor
                        dados_bancarios = buscar_dados_bancarios_fornecedor(cnpj_cpf)
                        ws_dados.cell(row=proxima_linha, column=12, value=dados_bancarios)
                        
                        # Se houve edição de valor, adicionar observação
                        if valor_pagar != float(row[29].value):
                            ws_dados.cell(row=proxima_linha, column=13, 
                                        value=f'LANÇAMENTO AUTOMÁTICO - Valor ajustado (Original: R$ {float(row[29].value):,.2f})')
                        else:
                            ws_dados.cell(row=proxima_linha, column=13, value='LANÇAMENTO AUTOMÁTICO')
                        
                        valor_info = f"R$ {valor_pagar:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
                        parcelas_processadas.append(f"Contrato {num_contrato} - Parcela {num_parcela}/{total_parcelas} - {valor_info}")
                        break
            
            wb.save(arquivo_cliente)
            
            # Limpar valores editados após pagamento
            self.valor_editado.clear()
            
            self.carregar_parcelas()
            
            if parcelas_processadas:
                mensagem = "Pagamentos registrados com sucesso:\n\n" + "\n".join(parcelas_processadas)
                messagebox.showinfo("Sucesso", mensagem)
            else:
                messagebox.showwarning("Aviso", "Nenhuma parcela foi processada!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao registrar pagamento: {str(e)}")
            if 'wb' in locals():
                wb.close()

    # ============= NOVAS FUNÇÕES DE VINCULAÇÃO =============
    
    def vincular_parcela(self):
        """Vincula uma parcela a um lançamento existente na aba Dados"""
        try:
            # Verificar se há parcela selecionada
            selecao = self.tree_parcelas.selection()
            if not selecao:
                self.root.lift()
                messagebox.showwarning("Aviso", "Selecione uma parcela para vincular!", parent=self.root)
                return
            
            if len(selecao) > 1:
                self.root.lift()
                messagebox.showwarning("Aviso", "Selecione apenas uma parcela por vez para vincular!", parent=self.root)
                return
            
            # Obter dados da parcela selecionada
            item = self.tree_parcelas.item(selecao[0])
            valores = item['values']
            
            # Preparar dados da parcela
            dados_parcela = {
                'num_contrato': str(valores[0]),
                'num_parcela': int(valores[1]),
                'cnpj': str(valores[2]),
                'nome': str(valores[3]),
                'eventos_fases': str(valores[4]),
                'valor_original': str(valores[5]),
                'valor_pagar': str(valores[6]),
                'status': str(valores[7])
            }
            
            # Verificar se já está paga ou vinculada
            if dados_parcela['status'].upper() in ['PAGO', 'VINCULADO']:
                self.root.lift()
                messagebox.showwarning(
                    "Aviso", 
                    f"Esta parcela já está {dados_parcela['status']}!\n\n"
                    "Não é possível vincular novamente.",
                    parent=self.root
                )
                return
            
            # Abrir janela de seleção de lançamento
            self.abrir_janela_selecao_lancamento(dados_parcela)
            
        except Exception as e:
            self.root.lift()
            messagebox.showerror("Erro", f"Erro ao vincular parcela: {str(e)}", parent=self.root)
    
    def abrir_janela_selecao_lancamento(self, dados_parcela):
        """Abre janela para seleção de lançamento existente"""
        try:
            # Criar janela modal
            janela = tk.Toplevel(self.root)
            janela.title("Vincular a Lançamento Existente")
            janela.geometry("1100x700")
            
            # Centralizar janela
            janela.update_idletasks()
            x = (janela.winfo_screenwidth() // 2) - (1100 // 2)
            y = (janela.winfo_screenheight() // 2) - (700 // 2)
            janela.geometry(f"+{x}+{y}")
            
            # Configurações para manter janela no topo
            janela.transient(self.root)
            janela.grab_set()
            janela.lift()
            janela.attributes('-topmost', True)
            janela.after(100, lambda: janela.attributes('-topmost', False))
            janela.focus_force()
            
            # Frame de informações da parcela
            frame_info = ttk.LabelFrame(janela, text="Dados da Parcela", padding=10)
            frame_info.pack(fill='x', padx=10, pady=5)
            
            # Converter valor para float para formatação
            valor_float = float(dados_parcela['valor_pagar'].replace('R$', '').replace('.', '').replace(',', '.').strip())
            
            info_text = f"""Contrato: {dados_parcela['num_contrato']} - Parcela: {dados_parcela['num_parcela']}
Fornecedor: {dados_parcela['nome']}
CNPJ/CPF: {dados_parcela['cnpj']}
Valor: R$ {valor_float:,.2f}
Eventos/Fases: {dados_parcela['eventos_fases']}"""
            
            ttk.Label(frame_info, text=info_text, justify='left').pack()
            
            # Frame de filtros
            frame_filtros = ttk.LabelFrame(janela, text="Filtros de Busca", padding=10)
            frame_filtros.pack(fill='x', padx=10, pady=5)
            
            # Filtro por nome
            ttk.Label(frame_filtros, text="Buscar por Nome:").grid(row=0, column=0, sticky='w', padx=5)
            var_filtro_nome = tk.StringVar(value=dados_parcela['nome'])
            entry_filtro = ttk.Entry(frame_filtros, textvariable=var_filtro_nome, width=40)
            entry_filtro.grid(row=0, column=1, sticky='ew', padx=5)
            
            # Filtro por valor aproximado
            var_valor_aprox = tk.BooleanVar(value=True)
            ttk.Checkbutton(
                frame_filtros, 
                text="Buscar valor aproximado (±10%)", 
                variable=var_valor_aprox
            ).grid(row=0, column=2, sticky='w', padx=10)
            
            # Botão de buscar
            btn_buscar = ttk.Button(
                frame_filtros, 
                text="Buscar",
                command=lambda: self.buscar_lancamentos_existentes(
                    tree_lancamentos, 
                    dados_parcela, 
                    var_filtro_nome.get(),
                    var_valor_aprox.get()
                )
            )
            btn_buscar.grid(row=0, column=3, padx=5)
            
            frame_filtros.columnconfigure(1, weight=1)
            
            # Frame para lista de lançamentos
            frame_lancamentos = ttk.LabelFrame(janela, text="Lançamentos Encontrados na Aba 'Dados'", padding=5)
            frame_lancamentos.pack(fill='both', expand=True, padx=10, pady=5)
            
            # Treeview para lançamentos
            colunas = ('Linha', 'Data', 'Nome', 'CNPJ/CPF', 'Valor', 'Vencimento', 'Referência', 'Observação')
            tree_lancamentos = ttk.Treeview(frame_lancamentos, columns=colunas, show='headings', height=15)
            
            # Configurar colunas
            larguras = {'Linha': 60, 'Data': 90, 'Nome': 200, 'CNPJ/CPF': 130, 
                    'Valor': 100, 'Vencimento': 90, 'Referência': 150, 'Observação': 200}
            
            for col in colunas:
                tree_lancamentos.heading(col, text=col)
                tree_lancamentos.column(col, width=larguras.get(col, 100))
            
            # Scrollbars
            scrolly = ttk.Scrollbar(frame_lancamentos, orient='vertical', command=tree_lancamentos.yview)
            scrollx = ttk.Scrollbar(frame_lancamentos, orient='horizontal', command=tree_lancamentos.xview)
            tree_lancamentos.configure(yscrollcommand=scrolly.set, xscrollcommand=scrollx.set)
            
            tree_lancamentos.grid(row=0, column=0, sticky='nsew')
            scrolly.grid(row=0, column=1, sticky='ns')
            scrollx.grid(row=1, column=0, sticky='ew')
            
            frame_lancamentos.grid_rowconfigure(0, weight=1)
            frame_lancamentos.grid_columnconfigure(0, weight=1)
            
            # Buscar lançamentos automaticamente ao abrir
            self.buscar_lancamentos_existentes(tree_lancamentos, dados_parcela, 
                                            var_filtro_nome.get(), var_valor_aprox.get())
            
            # Frame para botões de ação
            frame_botoes = ttk.Frame(janela)
            frame_botoes.pack(fill='x', padx=10, pady=10)
            
            ttk.Button(
                frame_botoes, 
                text="Vincular Selecionado",
                command=lambda: self.confirmar_vinculacao(
                    janela, dados_parcela, tree_lancamentos
                )
            ).pack(side='left', padx=5)
            
            ttk.Button(
                frame_botoes, 
                text="Cancelar",
                command=janela.destroy
            ).pack(side='right', padx=5)
            
            # Label de instruções
            ttk.Label(
                janela, 
                text="Dica: Selecione o lançamento correspondente e clique em 'Vincular Selecionado'",
                foreground='#666'
            ).pack(pady=5)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir janela de seleção: {str(e)}")
    
    def buscar_lancamentos_existentes(self, tree, dados_parcela, filtro_nome, valor_aproximado):
        """Busca lançamentos existentes que podem corresponder à parcela"""
        try:
            # Limpar treeview
            for item in tree.get_children():
                tree.delete(item)
            
            # Carregar planilha
            arquivo_cliente = Path(PASTA_CLIENTES) / f"{self.cliente_selecionado}.xlsx"
            wb = load_workbook(arquivo_cliente, data_only=True)
            ws = wb['Dados']
            
            # Valor da parcela
            valor_str = dados_parcela['valor_pagar'].replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor_parcela = float(valor_str)
            
            # Calcular margem de valor (±10%)
            margem = valor_parcela * 0.10
            valor_min = valor_parcela - margem
            valor_max = valor_parcela + margem
            
            # Buscar lançamentos
            encontrados = 0
            for idx, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
                # Extrair dados
                data_rel = row[0]
                cnpj_cpf = str(row[2]) if row[2] else ""
                nome = str(row[3]) if row[3] else ""
                referencia = str(row[4]) if row[4] else ""
                valor = row[8] if row[8] else 0
                dt_vencto = row[9]
                observacao = str(row[12]) if row[12] else ""
                
                # Aplicar filtros
                # 1. Filtro de nome (case insensitive, busca parcial)
                if filtro_nome:
                    filtro_lower = filtro_nome.lower()
                    nome_lower = nome.lower()
                    
                    # Verificar se há correspondência parcial
                    if filtro_lower not in nome_lower:
                        continue
                
                # 2. Filtro de valor
                try:
                    valor_float = float(valor)
                    if valor_aproximado:
                        # Buscar valor aproximado (±10%)
                        if not (valor_min <= valor_float <= valor_max):
                            continue
                    else:
                        # Buscar valor exato
                        if abs(valor_float - valor_parcela) > 0.01:
                            continue
                except:
                    continue
                
                # Formatar dados para exibição
                data_formatada = data_rel.strftime('%d/%m/%Y') if isinstance(data_rel, datetime) else str(data_rel)
                vencto_formatado = dt_vencto.strftime('%d/%m/%Y') if isinstance(dt_vencto, datetime) else str(dt_vencto)
                valor_formatado = f"R$ {valor_float:,.2f}"
                
                # Adicionar ao treeview
                tree.insert('', 'end', values=(
                    idx,  # Número da linha
                    data_formatada,
                    nome,
                    cnpj_cpf,
                    valor_formatado,
                    vencto_formatado,
                    referencia,
                    observacao
                ))
                encontrados += 1
            
            wb.close()
            
            # Mensagem se nada foi encontrado
            if encontrados == 0:
                # Temporariamente desabilitar topmost da janela pai para messagebox aparecer
                janela_pai = tree.master
                while janela_pai and not isinstance(janela_pai, tk.Toplevel):
                    janela_pai = janela_pai.master
                
                if janela_pai:
                    janela_pai.attributes('-topmost', False)
                
                messagebox.showinfo(
                    "Busca", 
                    "Nenhum lançamento encontrado com os critérios especificados.\n\n"
                    "Dicas:\n"
                    "• Experimente remover parte do nome\n"
                    "• Verifique se marcou 'valor aproximado'\n"
                    "• O fornecedor pode estar com nome diferente (PF/PJ)",
                    parent=janela_pai if janela_pai else None
                )
                
                if janela_pai:
                    janela_pai.attributes('-topmost', True)
                    janela_pai.lift()
            
        except Exception as e:
            # Encontrar janela pai
            janela_pai = tree.master
            while janela_pai and not isinstance(janela_pai, tk.Toplevel):
                janela_pai = janela_pai.master
            
            if janela_pai:
                janela_pai.attributes('-topmost', False)
                messagebox.showerror("Erro", f"Erro ao buscar lançamentos: {str(e)}", parent=janela_pai)
                janela_pai.attributes('-topmost', True)
                janela_pai.lift()
            else:
                messagebox.showerror("Erro", f"Erro ao buscar lançamentos: {str(e)}")
    
    def confirmar_vinculacao(self, janela, dados_parcela, tree):
        """Confirma e executa a vinculação da parcela ao lançamento selecionado"""
        try:
            # Verificar seleção
            selecao = tree.selection()
            if not selecao:
                # Garantir que messagebox apareça no topo
                janela.attributes('-topmost', False)
                messagebox.showwarning("Aviso", "Selecione um lançamento para vincular!", parent=janela)
                janela.attributes('-topmost', True)
                janela.lift()
                return
            
            # Obter dados do lançamento selecionado
            item = tree.item(selecao[0])
            valores = item['values']
            linha_lancamento = valores[0]
            nome_lancamento = valores[2]
            valor_lancamento = valores[4]
            
            # Garantir que o diálogo de confirmação apareça no topo
            janela.attributes('-topmost', False)
            
            # Confirmar com usuário
            resposta = messagebox.askyesno(
                "Confirmar Vinculação",
                f"Confirma a vinculação?\n\n"
                f"PARCELA:\n"
                f"Contrato: {dados_parcela['num_contrato']} - Parcela: {dados_parcela['num_parcela']}\n"
                f"Fornecedor: {dados_parcela['nome']}\n"
                f"Valor: {dados_parcela['valor_pagar']}\n\n"
                f"SERÁ VINCULADA AO LANÇAMENTO:\n"
                f"Linha: {linha_lancamento}\n"
                f"Nome: {nome_lancamento}\n"
                f"Valor: {valor_lancamento}\n\n"
                f"Esta ação marcará a parcela como 'VINCULADO'.",
                parent=janela
            )
            
            if not resposta:
                janela.attributes('-topmost', True)
                janela.lift()
                return
            
            # Executar vinculação
            arquivo_cliente = Path(PASTA_CLIENTES) / f"{self.cliente_selecionado}.xlsx"
            wb = load_workbook(arquivo_cliente)
            ws_contratos = wb['Contratos_ADM']
            
            # Atualizar status e dados da parcela
            hoje = datetime.now()
            
            for row_idx, row in enumerate(ws_contratos.iter_rows(min_row=3), start=3):
                if (str(row[24].value) == dados_parcela['num_contrato'] and
                    int(row[25].value) == dados_parcela['num_parcela'] and
                    str(row[26].value) == dados_parcela['cnpj']):
                    
                    # Atualizar status para VINCULADO
                    ws_contratos.cell(row=row_idx, column=31, value="VINCULADO")  # Status
                    ws_contratos.cell(row=row_idx, column=32, value=hoje)         # Data_Pagamento
                    
                    # Adicionar observação sobre a vinculação na coluna AJ (36)
                    obs_atual = ws_contratos.cell(row=row_idx, column=36).value or ""
                    nova_obs = f"{obs_atual} [VINCULADO À DESPESA DA LINHA {linha_lancamento} DE DADOS]".strip()
                    ws_contratos.cell(row=row_idx, column=36, value=nova_obs)
                    break
            
            # Salvar alterações
            wb.save(arquivo_cliente)
            wb.close()
            
            # Mensagem de sucesso
            messagebox.showinfo(
                "Sucesso", 
                f"Parcela vinculada com sucesso!\n\n"
                f"Contrato: {dados_parcela['num_contrato']} - Parcela: {dados_parcela['num_parcela']}\n"
                f"Status: VINCULADO\n"
                f"Linha do lançamento: {linha_lancamento}",
                parent=janela
            )
            
            # Fechar janela e atualizar lista
            janela.destroy()
            self.carregar_parcelas()
            
        except Exception as e:
            if 'janela' in locals() and janela.winfo_exists():
                janela.attributes('-topmost', False)
                messagebox.showerror("Erro", f"Erro ao confirmar vinculação: {str(e)}", parent=janela)
                janela.attributes('-topmost', True)
                janela.lift()
            else:
                messagebox.showerror("Erro", f"Erro ao confirmar vinculação: {str(e)}")
            if 'wb' in locals():
                wb.close()

    def verificar_correcoes_pendentes(self):
        """Verifica se há correções monetárias pendentes"""
        try:
            gerenciador_correcao = GerenciadorCorrecaoMonetaria()
            
            # Verificar se está na época de correção
            hoje = date.today()
            config_correcao = gerenciador_correcao.config.get('correcao_automatica', {})
            
            if (config_correcao.get('ativa', False) and 
                hoje.day == config_correcao.get('dia_calculo', 15)):
                
                if messagebox.askyesno("Correção Monetária", 
                                    "Detectamos que hoje pode ser o dia de aplicar correção monetária nos contratos.\n\n"
                                    "Deseja abrir o gerenciador de correção?"):
                    from src.correcao_monetaria import InterfaceIndicesCorrecao
                    interface = InterfaceIndicesCorrecao(self.root)
                    
        except Exception as e:
            print(f"Erro ao verificar correções: {str(e)}")

    def voltar_menu(self):
        """Fecha a janela e retorna ao menu principal"""
        # Verificar se temos uma referência para o controlador principal
        if hasattr(self, 'controlador_principal') and self.controlador_principal:
            self.root.destroy()
            
            if hasattr(self.controlador_principal, 'janela') and self.controlador_principal.janela:
                self.controlador_principal.janela.deiconify()
                return
        
        # Se não temos controlador_principal, usar o comportamento padrão
        self.root.destroy()
        
        if self.parent:
            if hasattr(self.parent, 'janela') and self.parent.janela:
                self.parent.janela.deiconify()
            else:
                self.parent.deiconify()

    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ControlePagamentos()
    app.run()