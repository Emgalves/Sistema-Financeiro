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
        self.root.geometry("1400x850+50+50")  # Aumentado e reposicionado
        
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
        
        # Label informativo
        self.label_info = ttk.Label(container_pagamento, 
                                   text="(Dica: Dê duplo clique em uma parcela para editar o valor)",
                                   font=('Arial', 8, 'italic'))
        self.label_info.pack(side='left', padx=10)
        
        
        # Frame de botões no final - COMPACTADO
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill='x', pady=5)
        
        ttk.Button(frame_botoes, text="Voltar ao Menu",
                  command=self.voltar_menu).pack(side='right', padx=5)
        
        ttk.Button(frame_botoes, text="Atualizar Lista",
                  command=self.carregar_parcelas).pack(side='right', padx=5)
        
        # Carregar lista de clientes
        self.carregar_clientes()
        self.verificar_correcoes_pendentes()
    
    def editar_valor_parcela(self, event):
        """Permite editar o valor de uma parcela com duplo clique"""
        item = self.tree_parcelas.selection()
        if not item:
            return
        
        valores = self.tree_parcelas.item(item[0])['values']
        
        # Atualizar labels de informação
        self.label_parcela_selecionada.config(text=f"Contrato {valores[0]} - Parcela {valores[1]}")
        self.label_valor_original.config(text=f"R$ {valores[5]}")
        
        # Preencher campo com valor atual (editado ou original)
        valor_atual = valores[6] if valores[6] != valores[5] else valores[5]
        self.entry_novo_valor.delete(0, tk.END)
        self.entry_novo_valor.insert(0, valor_atual)
        self.entry_novo_valor.focus()
    
    def aplicar_valor_editado(self):
        """Aplica o valor editado à parcela selecionada"""
        item = self.tree_parcelas.selection()
        if not item:
            messagebox.showwarning("Aviso", "Selecione uma parcela para editar!")
            return
        
        novo_valor_str = self.entry_novo_valor.get().strip()
        if not novo_valor_str:
            messagebox.showwarning("Aviso", "Informe o novo valor!")
            return
        
        try:
            # Converter valor para float (aceita vírgula ou ponto)
            novo_valor_str = novo_valor_str.replace('.', '').replace(',', '.')
            novo_valor = float(novo_valor_str)
            
            # Formatar valor para exibição
            valor_formatado = f"{novo_valor:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
            
            # Atualizar treeview
            valores = list(self.tree_parcelas.item(item[0])['values'])
            
            # Criar chave única para identificar a parcela
            chave_parcela = f"{valores[0]}_{valores[1]}_{valores[2]}"
            
            # Armazenar valor editado
            self.valor_editado[chave_parcela] = novo_valor
            
            # Atualizar coluna "Valor a Pagar"
            valores[6] = valor_formatado
            
            # Se o valor editado for diferente do original, destacar
            if valor_formatado != valores[5]:
                self.tree_parcelas.item(item[0], values=valores, tags=('editado',))
                self.tree_parcelas.tag_configure('editado', foreground='blue')
            else:
                self.tree_parcelas.item(item[0], values=valores, tags=())
            
            messagebox.showinfo("Sucesso", f"Valor alterado para R$ {valor_formatado}")
            
            # Limpar campos
            self.entry_novo_valor.delete(0, tk.END)
            self.label_parcela_selecionada.config(text="Nenhuma")
            self.label_valor_original.config(text="R$ 0,00")
            
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido! Use apenas números e vírgula/ponto decimal.")
    
    def resetar_valores(self):
        """Reseta todos os valores editados para os originais"""
        if messagebox.askyesno("Confirmar", "Deseja resetar todos os valores editados?"):
            self.valor_editado.clear()
            self.carregar_parcelas()
            messagebox.showinfo("Sucesso", "Todos os valores foram resetados!")
    
    def tem_taxa_fixa(self, arquivo_cliente):
        """
        Verifica se o cliente possui contratos com taxa fixa.
        """
        try:
            wb = load_workbook(arquivo_cliente)
            if 'Contratos_ADM' not in wb.sheetnames:
                wb.close()
                return False
                
            ws = wb['Contratos_ADM']
            
            # Converter todas as linhas em lista para facilitar a navegação
            rows = list(ws.iter_rows(min_row=3, values_only=True))
            
            for i in range(len(rows) - 1):
                row = rows[i]
                
                # Se encontrou um contrato com status ATIVO
                if row[0] and row[3] == 'ATIVO':
                    num_contrato = row[0]
                    
                    # Verificar a próxima linha para o tipo de taxa
                    if i + 1 < len(rows):
                        next_row = rows[i + 1]
                        # Verifica se é a linha de administrador
                        if next_row[6] == num_contrato and next_row[9] == 'Fixo':
                            wb.close()
                            return True
            
            wb.close()
            return False
            
        except Exception as e:
            if 'wb' in locals():
                wb.close()
            return False

    def carregar_clientes(self):
        """Carrega a lista de clientes que possuem contratos com taxa fixa ativa"""
        clientes = []
        
        for arquivo in os.listdir(PASTA_CLIENTES):
            if arquivo.endswith('.xlsx'):
                try:
                    arquivo_path = PASTA_CLIENTES / arquivo
                    if self.tem_taxa_fixa(arquivo_path):
                        nome_cliente = arquivo.replace('.xlsx', '')
                        clientes.append(nome_cliente)
                        
                except Exception as e:
                    print(f"Erro ao verificar arquivo {arquivo}: {str(e)}")
        
        self.cliente_combo['values'] = sorted(clientes)

    def carregar_parcelas(self, event=None):
        """Carrega as parcelas do cliente selecionado"""
        cliente = self.cliente_combo.get()
        if not cliente:
            return
            
        try:
            # Limpar lista atual
            for item in self.tree_parcelas.get_children():
                self.tree_parcelas.delete(item)
            
            # Limpar valores editados anteriores
            self.valor_editado.clear()
                
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            
            wb = load_workbook(arquivo_cliente)
            ws = wb['Contratos_ADM']
            
            # Primeiro criar um dicionário de status dos contratos
            contratos_ativos = {}
            for row in ws.iter_rows(min_row=3, values_only=True):
                if row[0]:  # Se tem número de contrato
                    contratos_ativos[str(row[0])] = row[3] == 'ATIVO'
            
            # Buscar parcelas apenas de contratos ativos
            for row in ws.iter_rows(min_row=3, values_only=True):
                num_contrato = row[24]  # Coluna Y - Número do contrato
                
                # Pular se não tem contrato ou se o contrato está inativo
                if not num_contrato or not contratos_ativos.get(str(num_contrato)):
                    continue
                    
                num_parcela = row[25]   # Coluna Z - Número da parcela
                cnpj_cpf = row[26]      # Coluna AA - CNPJ/CPF
                nome = row[27]          # Coluna AB - Nome
                descricao = row[32]     # Coluna AG - Eventos/Fases
                valor = row[29]         # Coluna AD - Valor
                status = row[30] if len(row) > 30 else "PENDENTE"  # Coluna AE - Status
                dt_pagto = row[31] if len(row) > 31 else None      # Coluna AF - Data pagamento
                
                # Só exibir parcelas PENDENTES
                if status == "PAGO":
                    continue
                
                # Formatar datas
                dt_pagto_str = dt_pagto.strftime('%d/%m/%Y') if isinstance(dt_pagto, datetime) else ""
                
                # Formatar valor usando formato brasileiro
                valor_str = f"{float(valor):,.2f}".replace(',', '_').replace('.', ',').replace('_', '.') if valor else ""
                
                # Verificar se há valor editado para esta parcela
                chave_parcela = f"{num_contrato}_{num_parcela}_{cnpj_cpf}"
                if chave_parcela in self.valor_editado:
                    valor_pagar = self.valor_editado[chave_parcela]
                    valor_pagar_str = f"{valor_pagar:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')
                    tags = ('editado',)
                else:
                    valor_pagar_str = valor_str
                    tags = ()
                
                # Inserir na treeview
                item = self.tree_parcelas.insert('', 'end', values=(
                    num_contrato,
                    num_parcela,
                    cnpj_cpf,
                    nome,           
                    descricao,
                    valor_str,
                    valor_pagar_str,  # Nova coluna com valor a pagar
                    status or "PENDENTE",
                    dt_pagto_str                    
                ), tags=tags)
                
                # Configurar cor para itens editados
                if tags:
                    self.tree_parcelas.tag_configure('editado', foreground='blue')
            
            wb.close()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar parcelas: {str(e)}")
            if 'wb' in locals():
                wb.close()

    def registrar_pagamento(self):
        """Registra o pagamento das parcelas selecionadas"""
        selecionados = self.tree_parcelas.selection()
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione as parcelas para pagamento!")
            return
        
        data_pagto = self.data_pagamento.get_date()
        
        if not validar_data(data_pagto.strftime('%d/%m/%Y')):
            messagebox.showerror("Erro", "Data de pagamento inválida!")
            return
            
        cliente = self.cliente_combo.get()
        if not cliente:
            return
        
        # Confirmar pagamentos com valores editados
        parcelas_editadas = []
        for item in selecionados:
            valores = self.tree_parcelas.item(item)['values']
            if valores[5] != valores[6]:  # Valor original != Valor a pagar
                parcelas_editadas.append(f"Contrato {valores[0]} - Parcela {valores[1]}: "
                                        f"R$ {valores[5]} → R$ {valores[6]}")
        
        if parcelas_editadas:
            msg = "As seguintes parcelas têm valores editados:\n\n"
            msg += "\n".join(parcelas_editadas)
            msg += "\n\nDeseja continuar com o pagamento?"
            
            if not messagebox.askyesno("Confirmação", msg):
                return
        
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            
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