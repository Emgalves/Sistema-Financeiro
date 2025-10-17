def add_project_root():
    import sys
    from pathlib import Path
    current_dir = Path(__file__).resolve().parent
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.append(str(project_root))

add_project_root()

import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
from openpyxl import load_workbook, Workbook
from datetime import datetime
from dateutil.relativedelta import relativedelta
import pandas as pd
import os

from src.config.utils import (
    validar_data,
    ARQUIVO_CLIENTES,
    PASTA_CLIENTES,
    buscar_dados_bancarios_fornecedor
)


class FinalizacaoQuinzena:
    def __init__(self, parent=None):
        self.parent = parent
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Finalização de Quinzena com Compensação Automática")
        self.root.geometry("1100x650")
        
        self.data_ref_entry = None
        self.tree_clientes = None
        self._detalhes_divergencia = []
        
        self.setup_gui()
        
    def run(self):
        """Inicia a execução do sistema"""
        self.root.protocol("WM_DELETE_WINDOW", self.voltar_menu)
        
        # Centralizar janela
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
        
        self.root.lift()
        self.root.focus_force()
        self.root.mainloop()

    def setup_gui(self):
        """Configura a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Frame superior: Data e busca
        frame_topo = ttk.LabelFrame(main_frame, text="Período de Referência", padding="10")
        frame_topo.pack(fill='x', pady=(0, 10))

        ttk.Label(frame_topo, text="Data de Referência:").pack(side='left', padx=5)
        self.data_ref_entry = DateEntry(
            frame_topo,
            format='dd/mm/yyyy',
            locale='pt_BR'
        )
        self.data_ref_entry.pack(side='left', padx=5)

        ttk.Button(
            frame_topo, 
            text="Buscar Clientes",
            command=self.carregar_clientes
        ).pack(side='left', padx=10)

        # Frame de clientes
        frame_lista = ttk.LabelFrame(main_frame, text="Clientes Pendentes", padding="5")
        frame_lista.pack(fill='both', expand=True, pady=(0, 10))

        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_lista)
        scrollbar.pack(side='right', fill='y')

        # Treeview com colunas expandidas
        self.tree_clientes = ttk.Treeview(
            frame_lista,
            columns=('Cliente', 'Base Atual', 'Taxa %', 'Taxa Calculada', 'Compensação', 'Valor Final'),
            show='headings',
            selectmode='extended',
            yscrollcommand=scrollbar.set
        )

        scrollbar.config(command=self.tree_clientes.yview)

        # Configurar colunas
        self.tree_clientes.heading('Cliente', text='Cliente')
        self.tree_clientes.heading('Base Atual', text='Base Quinzena')
        self.tree_clientes.heading('Taxa %', text='Taxa %')
        self.tree_clientes.heading('Taxa Calculada', text='Taxa Quinzena')
        self.tree_clientes.heading('Compensação', text='Compensação')
        self.tree_clientes.heading('Valor Final', text='Valor Final')

        self.tree_clientes.column('Cliente', width=200)
        self.tree_clientes.column('Base Atual', width=120)
        self.tree_clientes.column('Taxa %', width=80)
        self.tree_clientes.column('Taxa Calculada', width=120)
        self.tree_clientes.column('Compensação', width=120)
        self.tree_clientes.column('Valor Final', width=120)

        self.tree_clientes.pack(fill='both', expand=True)

        # Frame de botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill='x', pady=5)

        ttk.Button(
            frame_botoes, 
            text="Processar Selecionados",
            command=self.processar_clientes_selecionados
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Ver Detalhes Divergências",
            command=self.ver_detalhes_divergencias
        ).pack(side='left', padx=5)

        ttk.Button(
            frame_botoes,
            text="Voltar ao Menu",
            command=self.voltar_menu
        ).pack(side='right', padx=5)

    def carregar_clientes(self):
        """Carrega clientes com cálculo de taxas e compensações"""
        data_ref = self.data_ref_entry.get()
        if not validar_data(data_ref):
            messagebox.showerror("Erro", "Data inválida!")
            return

        try:
            # Limpar tree
            for item in self.tree_clientes.get_children():
                self.tree_clientes.delete(item)

            data_ref_dt = datetime.strptime(data_ref, '%d/%m/%Y')
            wb_clientes = load_workbook(ARQUIVO_CLIENTES)
            ws_clientes = wb_clientes['Clientes']

            print(f"\n{'='*60}")
            print(f"BUSCANDO CLIENTES PARA {data_ref}")
            print(f"{'='*60}\n")

            # Verificar se coluna Tipo Taxa existe
            headers = [cell.value for cell in ws_clientes[1]]
            tem_coluna_tipo = 'Tipo Taxa' in headers or 'TIPO TAXA' in [str(h).upper() for h in headers if h]
            
            if not tem_coluna_tipo:
                print("⚠️  Coluna 'Tipo Taxa' não encontrada na planilha Clientes.xlsx")
                print("💡 Criando coluna automaticamente...\n")
                self._adicionar_coluna_tipo_taxa(ws_clientes, wb_clientes)
                # Recarregar após adicionar coluna
                wb_clientes = load_workbook(ARQUIVO_CLIENTES)
                ws_clientes = wb_clientes['Clientes']

            # Identificar índice da coluna Tipo Taxa
            headers = [cell.value for cell in ws_clientes[1]]
            idx_tipo_taxa = None
            for idx, header in enumerate(headers):
                if header and 'TIPO' in str(header).upper() and 'TAXA' in str(header).upper():
                    idx_tipo_taxa = idx
                    break

            clientes_processados = 0
            clientes_com_percentual = 0

            for row in ws_clientes.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue

                nome_cliente = row[0]
                
                # Verificar tipo de taxa (se coluna existir)
                tipo_taxa = row[idx_tipo_taxa] if idx_tipo_taxa and len(row) > idx_tipo_taxa else None
                
                # FILTRO: Só processar clientes com taxa Percentual
                if tipo_taxa and str(tipo_taxa).upper() != 'PERCENTUAL':
                    print(f"⏭️  {nome_cliente}: Taxa {tipo_taxa} - IGNORADO")
                    continue

                arquivo_cliente = PASTA_CLIENTES / f"{nome_cliente}.xlsx"
                if not os.path.exists(arquivo_cliente):
                    print(f"⚠️  {nome_cliente}: Arquivo não encontrado")
                    continue

                try:
                    clientes_processados += 1
                    print(f"\n📋 Processando: {nome_cliente}")

                    # Verificar se já tem taxa lançada
                    if self._existe_taxa_na_quinzena(nome_cliente, data_ref_dt):
                        print(f"  ⏭️  Já tem taxa lançada - IGNORADO")
                        continue

                    # Obter percentual
                    percentual = self._obter_percentual_cliente(nome_cliente)
                    if percentual == 0:
                        print(f"  ⚠️  Sem taxa percentual configurada - IGNORADO")
                        continue

                    # Calcular base da quinzena
                    base_quinzena = self._calcular_base_quinzena(nome_cliente, data_ref_dt)
                    if base_quinzena == 0:
                        print(f"  ⚠️  Base zero - IGNORADO")
                        continue

                    # Calcular taxa da quinzena
                    taxa_quinzena = base_quinzena * (percentual / 100)

                    # Calcular divergência histórica
                    divergencia = self._calcular_divergencia_historica_total(nome_cliente)

                    # Valor final
                    valor_final = taxa_quinzena + divergencia

                    print(f"  ✅ Base: R$ {base_quinzena:.2f} | Taxa: {percentual}% | "
                          f"Quinzena: R$ {taxa_quinzena:.2f} | "
                          f"Comp: R$ {divergencia:+.2f} | Final: R$ {valor_final:.2f}")

                    # Adicionar na tree
                    self.tree_clientes.insert('', 'end', values=(
                        nome_cliente,
                        self._formatar_moeda(base_quinzena),
                        f"{percentual:.1f}%",
                        self._formatar_moeda(taxa_quinzena),
                        self._formatar_moeda_com_sinal(divergencia),
                        self._formatar_moeda(valor_final)
                    ), tags=(nome_cliente,))
                    
                    clientes_com_percentual += 1

                except Exception as e:
                    print(f"  ❌ Erro: {str(e)}")
                    import traceback
                    print(traceback.format_exc())
                    continue

            wb_clientes.close()
            
            print(f"\n{'='*60}")
            print(f"BUSCA CONCLUÍDA")
            print(f"Clientes verificados: {clientes_processados}")
            print(f"Clientes com taxa percentual encontrados: {clientes_com_percentual}")
            print(f"{'='*60}\n")
            
            if clientes_com_percentual == 0:
                messagebox.showinfo("Aviso", 
                    "Nenhum cliente com taxa percentual pendente foi encontrado.\n\n"
                    "Verifique:\n"
                    "• Se a coluna 'Tipo Taxa' está preenchida com 'Percentual'\n"
                    "• Se os clientes têm lançamentos na data selecionada\n"
                    "• Se já não existe taxa lançada para esta quinzena")

        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar clientes: {str(e)}")
            import traceback
            print(traceback.format_exc())
            if 'wb_clientes' in locals():
                wb_clientes.close()

    def processar_clientes_selecionados(self):
        """Processa os clientes selecionados com confirmação individual"""
        selecionados = self.tree_clientes.selection()
        
        if not selecionados:
            messagebox.showwarning("Aviso", "Selecione pelo menos um cliente!")
            return

        data_ref = datetime.strptime(self.data_ref_entry.get(), '%d/%m/%Y')
        processados = []
        erros = []

        for item in selecionados:
            valores = self.tree_clientes.item(item)['values']
            cliente = valores[0]

            try:
                # Recalcular para ter valores precisos
                percentual = self._obter_percentual_cliente(cliente)
                base_quinzena = self._calcular_base_quinzena(cliente, data_ref)
                taxa_quinzena = base_quinzena * (percentual / 100)
                divergencia = self._calcular_divergencia_historica_total(cliente)
                valor_final = taxa_quinzena + divergencia

                # Preparar mensagem de confirmação
                mensagem = self._preparar_mensagem_confirmacao(
                    cliente, data_ref, base_quinzena, percentual,
                    taxa_quinzena, divergencia, valor_final
                )

                # Confirmar com usuário
                if messagebox.askyesno("Confirmar Lançamento", mensagem):
                    self._criar_lancamento_taxa(cliente, data_ref, valor_final, divergencia)
                    processados.append(f"✅ {cliente} - R$ {valor_final:.2f}")
                    print(f"✅ {cliente} processado com sucesso!")
                else:
                    print(f"⏭️  {cliente} ignorado pelo usuário")

            except Exception as e:
                erro_msg = f"❌ {cliente} - {str(e)}"
                erros.append(erro_msg)
                print(erro_msg)

        # Mostrar resultado
        self._mostrar_resultado_processamento(processados, erros)
        
        # Recarregar lista
        self.carregar_clientes()

    # ============================================
    # MÉTODOS DE CÁLCULO
    # ============================================

    def _adicionar_coluna_tipo_taxa(self, ws_clientes, wb_clientes):
        """Adiciona coluna 'Tipo Taxa' na planilha Clientes.xlsx e detecta automaticamente o tipo"""
        try:
            # Adicionar header na coluna F (após Data Final que está em E)
            ws_clientes.cell(row=1, column=6, value='Tipo Taxa')
            
            print("📝 Analisando tipo de taxa de cada cliente...\n")
            
            # Para cada cliente, detectar o tipo
            for row_idx in range(2, ws_clientes.max_row + 1):
                nome_cliente = ws_clientes.cell(row=row_idx, column=1).value
                
                if not nome_cliente:
                    continue
                
                # Verificar no arquivo do cliente
                tipo_detectado = self._detectar_tipo_taxa_cliente(nome_cliente)
                ws_clientes.cell(row=row_idx, column=6, value=tipo_detectado)
                
                print(f"  • {nome_cliente}: {tipo_detectado}")
            
            # Salvar
            wb_clientes.save(ARQUIVO_CLIENTES)
            print(f"\n✅ Coluna 'Tipo Taxa' adicionada com sucesso!\n")
            
        except Exception as e:
            print(f"❌ Erro ao adicionar coluna: {e}")

    def _detectar_tipo_taxa_cliente(self, cliente):
        """Detecta automaticamente se cliente tem taxa Fixa ou Percentual"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            
            if not os.path.exists(arquivo_cliente):
                return "N/A"
            
            wb = load_workbook(arquivo_cliente, data_only=True)
            
            if 'Contratos_ADM' not in wb.sheetnames:
                wb.close()
                return "N/A"
            
            ws = wb['Contratos_ADM']
            
            # Buscar contratos ATIVOS
            contratos_ativos = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
                if not row or len(row) < 4:
                    continue
                
                status = row[3]
                num_contrato = row[0]
                
                if status == 'ATIVO' and num_contrato:
                    contratos_ativos.append(row_idx)
            
            if not contratos_ativos:
                wb.close()
                return "Sem Taxa"
            
            # Para cada contrato ativo, verificar tipo nas linhas seguintes
            tem_fixo = False
            tem_percentual = False
            
            for linha_contrato in contratos_ativos:
                for offset in range(1, 11):
                    linha_atual = linha_contrato + offset
                    
                    if linha_atual > ws.max_row:
                        break
                    
                    row = list(ws.iter_rows(min_row=linha_atual, max_row=linha_atual, values_only=True))[0]
                    
                    if not row or len(row) < 11:
                        continue
                    
                    tipo = row[9]  # Coluna J
                    
                    if tipo == 'Fixo':
                        tem_fixo = True
                        break
                    elif tipo == 'Percentual':
                        tem_percentual = True
                        break
                    
                    # Se encontrou outro contrato, parar
                    status_linha = row[3] if len(row) > 3 else None
                    if status_linha in ['ATIVO', 'INATIVO']:
                        break
            
            wb.close()
            
            # Priorizar Percentual
            if tem_percentual:
                return "Percentual"
            elif tem_fixo:
                return "Fixo"
            else:
                return "Sem Taxa"
                
        except Exception as e:
            print(f"    ⚠️  Erro ao detectar tipo: {e}")
            return "Erro"

    def _diagnosticar_estrutura_contratos(self, cliente):
        """Método de diagnóstico para entender a estrutura da aba Contratos_ADM"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            wb = load_workbook(arquivo_cliente, data_only=True)
            
            if 'Contratos_ADM' not in wb.sheetnames:
                print(f"  ⚠️  Aba 'Contratos_ADM' não existe!")
                wb.close()
                return
            
            ws = wb['Contratos_ADM']
            
            # Mostrar cabeçalhos (linha 1 ou 2)
            print(f"  📊 Estrutura da aba Contratos_ADM:")
            headers_row1 = [cell.value for cell in ws[1]]
            headers_row2 = [cell.value for cell in ws[2]] if ws.max_row > 1 else []
            
            print(f"     Linha 1: {headers_row1[:15]}")  # Primeiras 15 colunas
            if headers_row2:
                print(f"     Linha 2: {headers_row2[:15]}")
            
            # Mostrar primeira linha de dados (linha 3)
            if ws.max_row >= 3:
                primeira_linha = [cell.value for cell in ws[3]]
                print(f"     Dados linha 3: {primeira_linha[:15]}")
            
            # Tentar identificar a coluna "Tipo"
            for idx, header in enumerate(headers_row1):
                if header and 'tipo' in str(header).lower():
                    print(f"  ✅ Coluna 'Tipo' encontrada no índice {idx} (coluna {chr(65+idx)})")
            
            for idx, header in enumerate(headers_row2):
                if header and 'tipo' in str(header).lower():
                    print(f"  ✅ Coluna 'Tipo' encontrada na linha 2, índice {idx} (coluna {chr(65+idx)})")
            
            wb.close()
            
        except Exception as e:
            print(f"  ❌ Erro no diagnóstico: {e}")
            if 'wb' in locals():
                wb.close()

    def _existe_taxa_na_quinzena(self, cliente, data_quinzena):
        """Verifica se já existe taxa na quinzena"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            df = pd.read_excel(arquivo_cliente, sheet_name='Dados')
            
            # Verificar se coluna DATA_REL existe
            if 'DATA_REL' not in df.columns:
                # Tentar primeira coluna
                df.columns.values[0] = 'DATA_REL'
            
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')
            
            # Verificar se coluna STATUS existe, senão usar filtro alternativo
            if 'STATUS' in df.columns:
                taxas = df[
                    (df['DATA_REL'].dt.date == data_quinzena.date()) & 
                    (df['TP_DESP'] == 7) & 
                    (df['STATUS'] == 'ATIVO')
                ]
            else:
                # Sem coluna STATUS, apenas verificar data e tipo
                taxas = df[
                    (df['DATA_REL'].dt.date == data_quinzena.date()) & 
                    (df['TP_DESP'] == 7)
                ]
            
            return not taxas.empty
            
        except Exception as e:
            print(f"  ⚠️  Erro ao verificar taxa existente: {e}")
            return False

    def _obter_percentual_cliente(self, cliente):
        """Obtém percentual de taxa do cliente"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            wb = load_workbook(arquivo_cliente, data_only=True)
            
            if 'Contratos_ADM' not in wb.sheetnames:
                wb.close()
                return 0
            
            ws = wb['Contratos_ADM']
            
            # NOVA LÓGICA: Buscar contratos ATIVOS e depois verificar linha seguinte
            contratos_ativos = []
            
            # Primeiro pass: identificar linhas com contratos ATIVOS
            for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
                if not row or len(row) < 4:
                    continue
                
                # Coluna D (índice 3) = Status
                # Coluna A (índice 0) = Número do Contrato
                status = row[3] if len(row) > 3 else None
                num_contrato = row[0] if len(row) > 0 else None
                
                if status == 'ATIVO' and num_contrato:
                    contratos_ativos.append(row_idx)
                    print(f"    ✓ Contrato ATIVO encontrado na linha {row_idx}: {num_contrato}")
            
            if not contratos_ativos:
                print(f"  ℹ️  Nenhum contrato ATIVO encontrado")
                wb.close()
                return 0
            
            # Segundo pass: para cada contrato ativo, verificar linhas seguintes
            percentual_total = 0
            
            for linha_contrato in contratos_ativos:
                # Verificar as próximas 10 linhas após o contrato
                for offset in range(1, 11):
                    linha_atual = linha_contrato + offset
                    
                    if linha_atual > ws.max_row:
                        break
                    
                    row = list(ws.iter_rows(min_row=linha_atual, max_row=linha_atual, values_only=True))[0]
                    
                    if not row or len(row) < 11:
                        continue
                    
                    # Coluna J (índice 9) = Tipo
                    # Coluna K (índice 10) = Valor/Percentual
                    tipo = row[9] if len(row) > 9 else None
                    valor_taxa = row[10] if len(row) > 10 else None
                    
                    # Se encontrou um Tipo=Percentual com valor
                    if tipo == 'Percentual' and valor_taxa:
                        try:
                            percentual_str = str(valor_taxa).replace('%', '').replace(',', '.').strip()
                            percentual = float(percentual_str)
                            print(f"    ✅ Percentual encontrado na linha {linha_atual}: {percentual}%")
                            percentual_total += percentual
                            break  # Encontrou para este contrato, próximo contrato
                        except (ValueError, TypeError) as e:
                            print(f"    ⚠️ Erro ao converter '{valor_taxa}': {e}")
                            continue
                    
                    # Se encontrou outro contrato (linha com Status), parar busca
                    status_linha = row[3] if len(row) > 3 else None
                    if status_linha in ['ATIVO', 'INATIVO']:
                        break
            
            wb.close()
            
            if percentual_total > 0:
                print(f"  ✅ Percentual TOTAL do cliente: {percentual_total}%")
                return percentual_total
            else:
                print(f"  ℹ️  Nenhum percentual encontrado para contratos ativos")
                return 0
            
        except Exception as e:
            print(f"  ❌ Erro ao obter percentual: {e}")
            if 'wb' in locals():
                wb.close()
            return 0

    def _calcular_base_quinzena(self, cliente, data_quinzena):
        """Calcula base de cálculo da quinzena"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            df = pd.read_excel(arquivo_cliente, sheet_name='Dados')
            
            # Verificar se coluna DATA_REL existe
            if 'DATA_REL' not in df.columns:
                df.columns.values[0] = 'DATA_REL'
            
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')
            
            # Filtrar pela data e tipo diferente de 7 (taxa)
            if 'STATUS' in df.columns:
                df_quinzena = df[
                    (df['DATA_REL'].dt.date == data_quinzena.date()) & 
                    (df['TP_DESP'] != 7) & 
                    (df['STATUS'] == 'ATIVO')
                ]
            else:
                # Sem coluna STATUS
                df_quinzena = df[
                    (df['DATA_REL'].dt.date == data_quinzena.date()) & 
                    (df['TP_DESP'] != 7)
                ]
            
            base = df_quinzena['VALOR'].apply(
                lambda x: float(str(x).replace(',', '.')) if pd.notna(x) else 0
            ).sum()
            
            return base
            
        except Exception as e:
            print(f"  ❌ Erro ao calcular base: {e}")
            return 0

    def _calcular_divergencia_historica_total(self, cliente):
        """Calcula divergência acumulada de todas as quinzenas anteriores"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            df = pd.read_excel(arquivo_cliente, sheet_name='Dados')
            
            # Verificar se coluna DATA_REL existe
            if 'DATA_REL' not in df.columns:
                df.columns.values[0] = 'DATA_REL'
            
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')
            
            datas_unicas = sorted(df['DATA_REL'].dt.date.dropna().unique())
            
            divergencia_total = 0
            detalhes = []
            percentual = self._obter_percentual_cliente(cliente)
            
            if percentual == 0:
                return 0
            
            tem_status = 'STATUS' in df.columns
            
            for data in datas_unicas:
                df_data = df[df['DATA_REL'].dt.date == data]
                
                # Base: tudo exceto taxa
                if tem_status:
                    base_data = df_data[
                        (df_data['TP_DESP'] != 7) & 
                        (df_data['STATUS'] == 'ATIVO')
                    ]['VALOR'].apply(lambda x: float(str(x).replace(',', '.')) if pd.notna(x) else 0).sum()
                    
                    taxa_cobrada = df_data[
                        (df_data['TP_DESP'] == 7) & 
                        (df_data['STATUS'] == 'ATIVO')
                    ]['VALOR'].apply(lambda x: float(str(x).replace(',', '.')) if pd.notna(x) else 0).sum()
                else:
                    base_data = df_data[
                        df_data['TP_DESP'] != 7
                    ]['VALOR'].apply(lambda x: float(str(x).replace(',', '.')) if pd.notna(x) else 0).sum()
                    
                    taxa_cobrada = df_data[
                        df_data['TP_DESP'] == 7
                    ]['VALOR'].apply(lambda x: float(str(x).replace(',', '.')) if pd.notna(x) else 0).sum()
                
                if taxa_cobrada > 0:
                    taxa_devida = base_data * (percentual / 100)
                    divergencia = taxa_devida - taxa_cobrada
                    
                    divergencia_total += divergencia
                    
                    if abs(divergencia) > 0.01:
                        detalhes.append({
                            'data': data.strftime('%d/%m/%Y'),
                            'base': base_data,
                            'taxa_devida': taxa_devida,
                            'taxa_cobrada': taxa_cobrada,
                            'divergencia': divergencia
                        })
            
            # Salvar detalhes
            if not hasattr(self, '_cache_divergencias'):
                self._cache_divergencias = {}
            self._cache_divergencias[cliente] = detalhes
            
            return divergencia_total
            
        except Exception as e:
            print(f"  ❌ Erro ao calcular divergência: {e}")
            return 0

    def _criar_lancamento_taxa(self, cliente, data_quinzena, valor_taxa, compensacao):
        """Cria lançamento da taxa na planilha do cliente"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            wb = load_workbook(arquivo_cliente)
            ws = wb["Dados"]
            
            # Obter próximo ID
            max_id = 0
            for row in range(2, ws.max_row + 1):
                id_val = ws.cell(row=row, column=15).value
                if id_val:
                    try:
                        max_id = max(max_id, int(float(id_val)))
                    except:
                        pass
            
            novo_id = max_id + 1
            proxima_linha = ws.max_row + 1
            
            # Obter administrador
            administrador = self._obter_administrador_principal(cliente)
            
            # Preparar dados
            quinzena = "1ª" if data_quinzena.day == 5 else "2ª"
            referencia = f"TAXA ADM - {quinzena} QUINZ. {data_quinzena.strftime('%m/%Y')}"
            
            # Observação com compensação
            observacao = ""
            if abs(compensacao) > 0.01:
                if compensacao > 0:
                    observacao = f"Inclui compensação de R$ {compensacao:.2f}"
                else:
                    observacao = f"Com desconto de R$ {abs(compensacao):.2f}"
            
            # Gravar lançamento
            ws.cell(row=proxima_linha, column=1, value=data_quinzena)
            ws.cell(row=proxima_linha, column=2, value=7)
            ws.cell(row=proxima_linha, column=3, value=administrador['cnpj_cpf'])
            ws.cell(row=proxima_linha, column=4, value=administrador['nome'])
            ws.cell(row=proxima_linha, column=5, value=referencia)
            ws.cell(row=proxima_linha, column=6, value='')
            ws.cell(row=proxima_linha, column=7, value=valor_taxa)
            ws.cell(row=proxima_linha, column=8, value=1)
            ws.cell(row=proxima_linha, column=9, value=valor_taxa)
            ws.cell(row=proxima_linha, column=10, value=data_quinzena)
            ws.cell(row=proxima_linha, column=11, value='TAX')
            ws.cell(row=proxima_linha, column=12, value=administrador.get('dados_bancarios', 'PIX'))
            ws.cell(row=proxima_linha, column=13, value=observacao)
            ws.cell(row=proxima_linha, column=14, value='ATIVO')
            ws.cell(row=proxima_linha, column=15, value=novo_id)
            ws.cell(row=proxima_linha, column=16, value=f"Criado em {datetime.now().strftime('%d/%m/%Y %H:%M')}")
            
            # Formatar
            ws.cell(row=proxima_linha, column=1).number_format = 'DD/MM/YYYY'
            ws.cell(row=proxima_linha, column=7).number_format = '#,##0.00'
            ws.cell(row=proxima_linha, column=9).number_format = '#,##0.00'
            ws.cell(row=proxima_linha, column=10).number_format = 'DD/MM/YYYY'
            
            wb.save(arquivo_cliente)
            wb.close()
            
        except Exception as e:
            if 'wb' in locals():
                wb.close()
            raise Exception(f"Erro ao criar lançamento: {str(e)}")

    def _obter_administrador_principal(self, cliente):
        """Obtém dados do administrador principal seguindo estrutura de linhas separadas"""
        try:
            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            wb = load_workbook(arquivo_cliente, data_only=True)
            
            if 'Contratos_ADM' not in wb.sheetnames:
                wb.close()
                return self._administrador_padrao()
            
            ws = wb['Contratos_ADM']
            
            # Buscar contratos ATIVOS
            contratos_ativos = []
            for row_idx, row in enumerate(ws.iter_rows(min_row=3, values_only=True), start=3):
                if not row or len(row) < 4:
                    continue
                
                status = row[3]
                num_contrato = row[0]
                
                if status == 'ATIVO' and num_contrato:
                    contratos_ativos.append(row_idx)
            
            # Para cada contrato ativo, buscar dados do administrador
            for linha_contrato in contratos_ativos:
                for offset in range(1, 11):
                    linha_atual = linha_contrato + offset
                    
                    if linha_atual > ws.max_row:
                        break
                    
                    row = list(ws.iter_rows(min_row=linha_atual, max_row=linha_atual, values_only=True))[0]
                    
                    if not row or len(row) < 11:
                        continue
                    
                    tipo = row[9]  # Coluna J
                    
                    if tipo == 'Percentual':
                        # Colunas: H=CNPJ/CPF, I=Nome
                        cnpj_cpf = row[7] if len(row) > 7 else None
                        nome = row[8] if len(row) > 8 else None
                        
                        if cnpj_cpf and nome:
                            administrador = {
                                'cnpj_cpf': cnpj_cpf,
                                'nome': nome,
                                'dados_bancarios': buscar_dados_bancarios_fornecedor(cnpj_cpf) or 'PIX'
                            }
                            wb.close()
                            return administrador
                        break
                    
                    # Se encontrou outro contrato, parar
                    status_linha = row[3] if len(row) > 3 else None
                    if status_linha in ['ATIVO', 'INATIVO']:
                        break
            
            wb.close()
            return self._administrador_padrao()
            
        except Exception as e:
            print(f"Erro ao obter administrador: {e}")
            if 'wb' in locals():
                wb.close()
            return self._administrador_padrao()
    
    def _administrador_padrao(self):
        """Retorna dados padrão de administrador"""
        return {
            'cnpj_cpf': '00000000000',
            'nome': 'ADMINISTRADOR',
            'dados_bancarios': 'PIX'
        }

    # ============================================
    # MÉTODOS DE INTERFACE
    # ============================================

    def _preparar_mensagem_confirmacao(self, cliente, data_ref, base, percentual, 
                                       taxa, divergencia, valor_final):
        """Prepara mensagem de confirmação detalhada"""
        mensagem = f"📊 FINALIZAÇÃO DE QUINZENA\n"
        mensagem += f"{'='*40}\n\n"
        mensagem += f"Cliente: {cliente}\n"
        mensagem += f"Data: {data_ref.strftime('%d/%m/%Y')}\n\n"
        
        mensagem += f"📅 QUINZENA ATUAL:\n"
        mensagem += f"   Base: R$ {base:,.2f}\n"
        mensagem += f"   Taxa: {percentual}%\n"
        mensagem += f"   Valor: R$ {taxa:,.2f}\n\n"
        
        if abs(divergencia) > 0.01:
            mensagem += f"⚠️ COMPENSAÇÃO:\n"
            mensagem += f"   {('Cobrar' if divergencia > 0 else 'Creditar')}: R$ {abs(divergencia):,.2f}\n\n"
        
        mensagem += f"💰 VALOR FINAL: R$ {valor_final:,.2f}\n\n"
        mensagem += "Confirma o lançamento?"
        
        return mensagem

    def ver_detalhes_divergencias(self):
        """Mostra janela com detalhes das divergências"""
        selecionados = self.tree_clientes.selection()
        
        if not selecionados:
            messagebox.showinfo("Info", "Selecione um cliente para ver detalhes")
            return
        
        cliente = self.tree_clientes.item(selecionados[0])['values'][0]
        
        if not hasattr(self, '_cache_divergencias') or cliente not in self._cache_divergencias:
            messagebox.showinfo("Info", "Nenhuma divergência encontrada")
            return
        
        detalhes = self._cache_divergencias[cliente]
        
        if not detalhes:
            messagebox.showinfo("Info", f"{cliente}: Todas as taxas estão corretas!")
            return
        
        # Janela de detalhes
        janela = tk.Toplevel(self.root)
        janela.title(f"Divergências Históricas - {cliente}")
        janela.geometry("700x400")
        
        frame = ttk.Frame(janela, padding="10")
        frame.pack(fill='both', expand=True)
        
        tree = ttk.Treeview(
            frame,
            columns=('Data', 'Base', 'Devida', 'Cobrada', 'Diferença'),
            show='headings'
        )
        
        tree.heading('Data', text='Data')
        tree.heading('Base', text='Base')
        tree.heading('Devida', text='Taxa Devida')
        tree.heading('Cobrada', text='Taxa Cobrada')
        tree.heading('Diferença', text='Diferença')
        
        for detalhe in detalhes:
            tree.insert('', 'end', values=(
                detalhe['data'],
                self._formatar_moeda(detalhe['base']),
                self._formatar_moeda(detalhe['taxa_devida']),
                self._formatar_moeda(detalhe['taxa_cobrada']),
                self._formatar_moeda_com_sinal(detalhe['divergencia'])
            ))
        
        tree.pack(fill='both', expand=True)
        
        ttk.Button(frame, text="Fechar", command=janela.destroy).pack(pady=10)

    def _mostrar_resultado_processamento(self, processados, erros):
        """Mostra resultado do processamento"""
        janela = tk.Toplevel(self.root)
        janela.title("Resultado do Processamento")
        janela.geometry("600x400")
        
        frame = ttk.Frame(janela, padding="10")
        frame.pack(fill='both', expand=True)
        
        if processados:
            frame_proc = ttk.LabelFrame(frame, text="Processados com Sucesso", padding="5")
            frame_proc.pack(fill='both', expand=True, pady=5)
            
            texto_proc = tk.Text(frame_proc, height=10)
            texto_proc.pack(fill='both', expand=True)
            
            for msg in processados:
                texto_proc.insert(tk.END, f"{msg}\n")
            texto_proc.config(state='disabled')
        
        if erros:
            frame_erros = ttk.LabelFrame(frame, text="Erros", padding="5")
            frame_erros.pack(fill='both', expand=True, pady=5)
            
            texto_erros = tk.Text(frame_erros, height=10)
            texto_erros.pack(fill='both', expand=True)
            
            for msg in erros:
                texto_erros.insert(tk.END, f"{msg}\n")
            texto_erros.config(state='disabled')
        
        ttk.Button(frame, text="Fechar", command=janela.destroy).pack(pady=10)

    def _formatar_moeda(self, valor):
        """Formata valor em moeda brasileira"""
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def _formatar_moeda_com_sinal(self, valor):
        """Formata valor com sinal + ou -"""
        sinal = "+" if valor >= 0 else ""
        return f"{sinal}R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

    def voltar_menu(self):
        """Volta ao menu principal"""
        self.root.destroy()
        if self.parent:
            self.parent.deiconify()
            self.parent.lift()


# ============================================
# EXECUÇÃO
# ============================================

if __name__ == "__main__":
    app = FinalizacaoQuinzena()
    app.run()