"""
Configuração do Relatório Quinzenal de Medições para integração com relatorios_interface.py
Versão modificada com:
- Combobox para seleção de clientes (ao invés de seleção de arquivo)
- Data de referência automática (dias 5 e 20)
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pathlib import Path
import sys
import os
import logging

# Configurar logger
logger = logging.getLogger(__name__)

# Importar DateEntry se disponível
try:
    from tkcalendar import DateEntry
    tem_tkcalendar = True
except ImportError:
    tem_tkcalendar = False

# Importar configurações do sistema
try:
    from src.config.config import ARQUIVO_CLIENTES, PASTA_CLIENTES
    usa_config_sistema = True
except ImportError:
    ARQUIVO_CLIENTES = None
    PASTA_CLIENTES = None
    usa_config_sistema = False

# Importar gerador de PDF
try:
    from gerar_relatorio_quinzenal_pdf import RelatorioQuinzenalPDF
except ImportError:
    try:
        from src.gerar_relatorio_quinzenal_pdf import RelatorioQuinzenalPDF
    except ImportError:
        RelatorioQuinzenalPDF = None


def _estado_sessao():
    """Estado compartilhado entre módulos (último cliente usado).
    Retorna None se o módulo não estiver disponível."""
    try:
        from src.estado_compartilhado import estado_sessao
        return estado_sessao
    except ImportError:
        try:
            from estado_compartilhado import estado_sessao
            return estado_sessao
        except ImportError:
            return None


def _lembrar_cliente(nome_cliente):
    """Guarda o cliente escolhido aqui para as outras telas."""
    try:
        estado = _estado_sessao()
        if estado is not None and nome_cliente and nome_cliente != 'Todos os Clientes':
            estado.ultimo_cliente = nome_cliente
    except Exception as e:
        logger.warning(f"Não foi possível lembrar o cliente: {e}")


def _selecionar_cliente_preferido(combobox, nome_preferido=None):
    """Seleciona, nesta ordem: nome_preferido (se estiver na lista), o último
    cliente usado em outra tela (estado_sessao) e, por fim, o 1º item."""
    valores = list(combobox['values'])
    if not valores:
        return
    candidatos = [nome_preferido]
    try:
        estado = _estado_sessao()
        if estado is not None:
            candidatos.append(estado.ultimo_cliente)
    except Exception:
        pass
    for nome in candidatos:
        if nome and nome in valores:
            combobox.set(nome)
            return
    combobox.current(0)


def escolher_contratos_dialog(parent, contratos, titulo_cliente=''):
    """
    Janela modal para escolher quais contratos entram no relatório quinzenal.

    Args:
        parent:    janela pai
        contratos: lista de RelatorioQuinzenalPDF.listar_contratos()
                   (= conteúdo do relatório completo)
        titulo_cliente: nome do cliente, só para exibição

    Returns:
        set com os ID_Contrato DESMARCADOS (vazio = todos marcados),
        ou None se o usuário cancelar.
    """
    def _moeda(v):
        return "{:,.2f}".format(float(v or 0)).replace(',', 'X').replace('.', ',').replace('X', '.')

    marcados = {c['id']: True for c in contratos}
    resultado = {'valor': None}

    win = tk.Toplevel(parent)
    win.title("Escolher contratos do relatório")
    win.transient(parent)
    win.geometry("980x420")

    frame = ttk.Frame(win, padding=10)
    frame.pack(fill='both', expand=True)

    ttk.Label(
        frame,
        text=(f"Cliente: {titulo_cliente}\n" if titulo_cliente else "") +
             "Desmarque os contratos que NÃO devem sair no PDF (clique na linha ou tecle Espaço).\n"
             "Totais e % Executado do PDF consideram apenas os contratos marcados.",
        justify='left'
    ).pack(anchor='w', pady=(0, 8))

    colunas = ('sel', 'fornecedor', 'num', 'descricao', 'status', 'global', 'saldo', 'medq')
    tree_frame = ttk.Frame(frame)
    tree_frame.pack(fill='both', expand=True)
    tree = ttk.Treeview(tree_frame, columns=colunas, show='headings', selectmode='browse')
    cab = {'sel': ('Incluir', 60, 'center'), 'fornecedor': ('Empreiteiro', 200, 'w'),
           'num': ('Nº/Emp', 55, 'center'), 'descricao': ('Descrição', 300, 'w'),
           'status': ('Status', 85, 'center'), 'global': ('Valor Global', 90, 'e'),
           'saldo': ('Saldo', 80, 'e'), 'medq': ('Medição na quinzena', 120, 'e')}
    for col in colunas:
        txt, larg, anc = cab[col]
        tree.heading(col, text=txt)
        tree.column(col, width=larg, anchor=anc, stretch=(col == 'descricao'))
    sb = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side='left', fill='both', expand=True)
    sb.pack(side='right', fill='y')
    tree.tag_configure('desmarcado', foreground='#999999')

    lbl_resumo = ttk.Label(frame, text='')
    lbl_resumo.pack(anchor='w', pady=(6, 0))

    def _atualizar_linha(id_c):
        iid = str(id_c)
        vals = list(tree.item(iid, 'values'))
        vals[0] = '☑' if marcados[id_c] else '☐'
        tree.item(iid, values=vals, tags=() if marcados[id_c] else ('desmarcado',))

    def _atualizar_resumo():
        n = sum(marcados.values())
        lbl_resumo.config(text=f"{n} de {len(marcados)} contrato(s) marcado(s)")

    for c in contratos:
        tree.insert('', 'end', iid=str(c['id']), values=(
            '☑', c['fornecedor'], c['num_emp'], c['descricao'][:120], c['status'],
            _moeda(c['valor_global']), _moeda(c['saldo']),
            _moeda(c['med_quinzena']) if c['med_quinzena'] else '—'))
    _atualizar_resumo()

    def _alternar(iid):
        if not iid:
            return
        id_c = int(iid)
        marcados[id_c] = not marcados[id_c]
        _atualizar_linha(id_c)
        _atualizar_resumo()

    tree.bind('<ButtonRelease-1>', lambda e: _alternar(tree.identify_row(e.y)))
    tree.bind('<space>', lambda e: _alternar(tree.focus()))

    def _todos(valor):
        for id_c in marcados:
            marcados[id_c] = valor
            _atualizar_linha(id_c)
        _atualizar_resumo()

    def _ok():
        if not any(marcados.values()):
            messagebox.showwarning("Aviso", "Marque ao menos um contrato.", parent=win)
            return
        resultado['valor'] = {id_c for id_c, m in marcados.items() if not m}
        win.destroy()

    def _cancelar():
        resultado['valor'] = None
        win.destroy()

    botoes = ttk.Frame(frame)
    botoes.pack(fill='x', pady=(8, 0))
    ttk.Button(botoes, text="Marcar todos", command=lambda: _todos(True)).pack(side='left')
    ttk.Button(botoes, text="Desmarcar todos", command=lambda: _todos(False)).pack(side='left', padx=6)
    ttk.Button(botoes, text="Cancelar", command=_cancelar).pack(side='right')
    ttk.Button(botoes, text="OK — Gerar com os marcados", command=_ok).pack(side='right', padx=6)
    win.protocol("WM_DELETE_WINDOW", _cancelar)

    win.grab_set()
    tree.focus_set()
    parent.wait_window(win)
    return resultado['valor']


def carregar_clientes():
    """Carrega a lista de clientes ativos do arquivo de clientes"""
    try:
        # Importar bibliotecas necessárias
        import pandas as pd
        from openpyxl import load_workbook
        
        # Caminho para o arquivo de clientes
        try:
            from src.config.config import ARQUIVO_CLIENTES
            logger.info(f"Carregando clientes de: {ARQUIVO_CLIENTES}")
        except ImportError:
            # Caminho padrão se não conseguir importar das configurações
            ARQUIVO_CLIENTES = "dados/clientes.xlsx"
            logger.warning(f"Usando caminho padrão para clientes: {ARQUIVO_CLIENTES}")
        
        # Verificar se o arquivo existe
        if not os.path.exists(ARQUIVO_CLIENTES):
            logger.warning(f"Arquivo de clientes não encontrado: {ARQUIVO_CLIENTES}")
            return ['Todos os Clientes']
        
        # Carregar o arquivo usando pandas
        try:
            # Ler o arquivo Excel
            df = pd.read_excel(ARQUIVO_CLIENTES, sheet_name='Clientes')
            
            # Debug: mostrar as colunas disponíveis
            logger.info(f"Colunas disponíveis: {df.columns.tolist()}")
            
            # Verificar se a coluna E existe (coluna 4 em índice baseado em 0)
            # Ou verificar pelo nome da coluna se existir
            if len(df.columns) >= 5:  # Verifica se tem pelo menos 5 colunas (A-E)
                # Filtrar clientes ativos (coluna E vazia)
                coluna_status = df.columns[4]  # Coluna E (índice 4)
                logger.info(f"Coluna de status: {coluna_status}")
                
                # Considera como vazio: None, NaN, '', etc.
                df_ativos = df[df[coluna_status].isna() | (df[coluna_status] == '')]
                
                # Verificar se a primeira coluna contém os nomes dos clientes
                coluna_nome = df.columns[0]  # Coluna A
                logger.info(f"Coluna de nome: {coluna_nome}")
                
                # Extrair nomes dos clientes ativos (assumindo que estão na primeira coluna)
                clientes_ativos = df_ativos[coluna_nome].dropna().tolist()
                
                logger.info(f"Total de clientes ativos encontrados: {len(clientes_ativos)}")
                
                # Ordenar alfabeticamente
                clientes_ativos.sort()
                
                # Adicionar "Todos os Clientes" no início
                clientes = ['Todos os Clientes'] + clientes_ativos
                
                return clientes
            else:
                logger.warning("Arquivo não tem colunas suficientes (precisa de pelo menos 5 colunas - A até E)")
                return ['Todos os Clientes']
            
        except Exception as e:
            logger.error(f"Erro ao ler arquivo Excel com pandas: {str(e)}")
            # Tentar com openpyxl como fallback
            try:
                workbook = load_workbook(ARQUIVO_CLIENTES)
                sheet = workbook['Clientes']
                
                clientes = ['Todos os Clientes']
                for row in sheet.iter_rows(min_row=2, values_only=True):
                    # Verifica se a coluna E (índice 4) está vazia
                    if row[0] and (len(row) < 5 or not row[4]):
                        clientes.append(row[0])
                
                workbook.close()
                clientes.sort()  # Ordenar alfabeticamente (mantendo "Todos os Clientes" primeiro)
                return clientes
                
            except Exception as inner_e:
                logger.error(f"Erro ao ler arquivo Excel com openpyxl: {str(inner_e)}")
                return ['Todos os Clientes']
            
    except Exception as e:
        logger.error(f"Erro ao carregar clientes: {str(e)}", exc_info=True)
        return ['Todos os Clientes']


def calcular_data_rel_automatica():
    """Calcula automaticamente a data do relatório baseado na regra dos dias 5 e 20"""
    try:
        hoje = datetime.now()
        
        if 6 <= hoje.day <= 20:
            # Entre dia 6 e 20: relatório do dia 20 do mês atual
            data_rel = hoje.replace(day=20)
        else:
            if hoje.day > 20:
                # Após dia 20: relatório do dia 5 do próximo mês
                data_rel = (hoje + relativedelta(months=1)).replace(day=5)
            else:
                # Antes do dia 6: relatório do dia 5 do mês atual
                data_rel = hoje.replace(day=5)
        
        logger.info(f"Data calculada automaticamente: {data_rel.strftime('%d/%m/%Y')}")
        return data_rel
        
    except Exception as e:
        logger.error(f"Erro ao calcular data automática: {str(e)}")
        # Fallback: retorna data atual
        return datetime.now()


def obter_caminho_arquivo_cliente(nome_cliente):
    """Obtém o caminho do arquivo Excel do cliente baseado no nome"""
    try:
        if not usa_config_sistema or not PASTA_CLIENTES:
            return None
        
        # Normalizar o nome do cliente para o formato do arquivo
        # Geralmente: nome_cliente.xlsx ou NOME_CLIENTE.xlsx
        nome_normalizado = nome_cliente.replace(' ', '_')
        
        # Tentar diferentes variações do nome
        variações = [
            f"{nome_normalizado}.xlsx",
            f"{nome_normalizado.upper()}.xlsx",
            f"{nome_normalizado.lower()}.xlsx",
        ]
        
        for variacao in variações:
            caminho = Path(PASTA_CLIENTES) / variacao
            if caminho.exists():
                logger.info(f"Arquivo encontrado: {caminho}")
                return str(caminho)
        
        # Se não encontrou, procurar na pasta
        pasta = Path(PASTA_CLIENTES)
        if pasta.exists():
            for arquivo in pasta.glob("*.xlsx"):
                # Comparar sem considerar case e underscores
                nome_arquivo = arquivo.stem.replace('_', ' ').lower()
                nome_busca = nome_cliente.replace('_', ' ').lower()
                if nome_arquivo == nome_busca:
                    logger.info(f"Arquivo encontrado por busca: {arquivo}")
                    return str(arquivo)
        
        logger.warning(f"Arquivo não encontrado para cliente: {nome_cliente}")
        return None
        
    except Exception as e:
        logger.error(f"Erro ao obter caminho do arquivo do cliente: {str(e)}")
        return None


def configurar_relatorio_quinzenal(parent_frame, sistema_relatorios):
    """
    Cria a interface de configuração do relatório quinzenal no painel direito
    
    Args:
        parent_frame: Frame onde as configurações serão adicionadas
        sistema_relatorios: Instância do SistemaRelatorios para acesso a métodos
    """
    
    # Limpar frame anterior
    for widget in parent_frame.winfo_children():
        widget.destroy()
    
    # Variáveis de controle
    cliente_selecionado_var = tk.StringVar()
    arquivo_clientes_var = tk.StringVar()
    
    # Pré-carregar Clientes.xlsx se disponível
    if usa_config_sistema and ARQUIVO_CLIENTES and ARQUIVO_CLIENTES.exists():
        arquivo_clientes_var.set(str(ARQUIVO_CLIENTES))
    
    # ========== TÍTULO ==========
    titulo_frame = ttk.Frame(parent_frame)
    titulo_frame.pack(fill='x', pady=(0, 15))
    
    ttk.Label(
        titulo_frame,
        text="Relatório Quinzenal de Medições (PDF)",
        font=('Arial', 14, 'bold')
    ).pack(anchor='w')
    
    ttk.Label(
        titulo_frame,
        text="Relatório PDF de medições da quinzena (dias 5 e 20)",
        font=('Arial', 9),
        foreground='gray'
    ).pack(anchor='w')
    
    # ========== DESCRIÇÃO ==========
    desc_frame = ttk.LabelFrame(parent_frame, text="ℹ️ Sobre este Relatório", padding=10)
    desc_frame.pack(fill='x', pady=(0, 15))
    
    desc_text = """✓ Contratos com medições na quinzena (dias 5 e 20)
✓ Histórico completo de medições por contrato
✓ Medições da quinzena destacadas em amarelo
✓ Resumo financeiro completo
✓ Opcional: escolher quais contratos entram no PDF"""
    
    ttk.Label(desc_frame, text=desc_text, justify='left', font=('Arial', 9)).pack(anchor='w')
    
    # ========== SELEÇÃO DE CLIENTE ==========
    cliente_frame = ttk.LabelFrame(parent_frame, text="👤 Seleção de Cliente", padding=10)
    cliente_frame.pack(fill='x', pady=(0, 15))
    
    ttk.Label(
        cliente_frame,
        text="Selecione o cliente:",
        font=('Arial', 10, 'bold')
    ).pack(anchor='w', pady=(0, 5))
    
    # Carregar lista de clientes
    lista_clientes = carregar_clientes()
    
    # Combobox de clientes
    cliente_combobox = ttk.Combobox(
        cliente_frame,
        textvariable=cliente_selecionado_var,
        values=lista_clientes,
        state='readonly',
        width=40,
        font=('Arial', 10)
    )
    cliente_combobox.pack(anchor='w', pady=(0, 5))
    
    # Pré-selecionar o último cliente usado em outra tela; se não houver,
    # fica o 1º item ("Todos os Clientes"), como antes.
    _selecionar_cliente_preferido(cliente_combobox)

    # Cliente escolhido aqui passa a valer para as outras telas
    cliente_combobox.bind(
        '<<ComboboxSelected>>',
        lambda e: _lembrar_cliente(cliente_selecionado_var.get())
    )
    
    # Informação sobre a pasta
    if usa_config_sistema and PASTA_CLIENTES:
        ttk.Label(
            cliente_frame,
            text=f"📂 Pasta de clientes: {PASTA_CLIENTES}",
            font=('Arial', 8),
            foreground='blue'
        ).pack(anchor='w', pady=(5, 0))
    
    # Botão para atualizar lista de clientes
    def atualizar_clientes():
        selecionado_antes = cliente_selecionado_var.get()
        nova_lista = carregar_clientes()
        cliente_combobox['values'] = nova_lista
        # Mantém o cliente que estava selecionado, se continuar na lista
        _selecionar_cliente_preferido(cliente_combobox, selecionado_antes)
        messagebox.showinfo("Info", f"Lista atualizada com {len(nova_lista)} clientes", parent=parent_frame.winfo_toplevel())
    
    ttk.Button(
        cliente_frame,
        text="🔄 Atualizar Lista",
        command=atualizar_clientes
    ).pack(anchor='w', pady=(8, 0))
    
    # ========== ARQUIVO CLIENTES.XLSX (Oculto, mas mantido) ==========
    # Mantido para compatibilidade com o código de geração
    
    # ========== DATA DE REFERÊNCIA AUTOMÁTICA ==========
    data_frame = ttk.LabelFrame(parent_frame, text="📅 Data de Referência", padding=10)
    data_frame.pack(fill='x', pady=(0, 15))
    
    # Calcular data automática
    data_automatica = calcular_data_rel_automatica()
    
    ttk.Label(
        data_frame,
        text="Data do relatório (calculada automaticamente):",
        font=('Arial', 10)
    ).pack(anchor='w', pady=(0, 5))
    
    # Frame para exibir a data
    frame_data_display = ttk.Frame(data_frame)
    frame_data_display.pack(anchor='w', pady=(0, 10))
    
    # Label com a data calculada
    data_label = ttk.Label(
        frame_data_display,
        text=data_automatica.strftime('%d/%m/%Y'),
        font=('Arial', 12, 'bold'),
        foreground='#006600'
    )
    data_label.pack(side='left', padx=(0, 10))
    
    # Ícone de calendário
    ttk.Label(
        frame_data_display,
        text="📅",
        font=('Arial', 14)
    ).pack(side='left')
    
    # Explicação da lógica
#     info_logica = ttk.Frame(data_frame)
#     info_logica.pack(anchor='w', fill='x')
    
#     ttk.Label(
#         info_logica,
#         text="ℹ️ Lógica de cálculo:",
#         font=('Arial', 9, 'bold'),
#         foreground='#0066CC'
#     ).pack(anchor='w')
    
#     texto_logica = """• Dias 6 a 20: Relatório do dia 20 do mês atual
# • Dias 21 a 31: Relatório do dia 5 do próximo mês
# • Dias 1 a 5: Relatório do dia 5 do mês atual"""
    
#     ttk.Label(
#         info_logica,
#         text=texto_logica,
#         font=('Arial', 8),
#         foreground='#666666',
#         justify='left'
#     ).pack(anchor='w', pady=(3, 8))
    
    # Opção para alterar data manualmente (se necessário)
    var_alterar_data = tk.BooleanVar(value=False)
    
    def toggle_alteracao_data():
        if var_alterar_data.get():
            frame_data_manual.pack(anchor='w', pady=(5, 0))
        else:
            frame_data_manual.pack_forget()
            data_label.config(text=data_automatica.strftime('%d/%m/%Y'))
    
    check_alterar = ttk.Checkbutton(
        data_frame,
        text="Alterar data manualmente",
        variable=var_alterar_data,
        command=toggle_alteracao_data
    )
    check_alterar.pack(anchor='w')
    
    # Frame para alteração manual (inicialmente oculto)
    frame_data_manual = ttk.Frame(data_frame)
    
    ttk.Label(
        frame_data_manual,
        text="Nova data:",
        font=('Arial', 9)
    ).pack(side='left', padx=(0, 5))
    
    if tem_tkcalendar:
        data_entry = DateEntry(
            frame_data_manual,
            width=12,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR',
            font=('Arial', 10)
        )
        data_entry.pack(side='left', padx=(0, 5))
        data_entry.set_date(data_automatica)
    else:
        data_entry = ttk.Entry(frame_data_manual, width=12, font=('Arial', 10))
        data_entry.pack(side='left', padx=(0, 5))
        data_entry.insert(0, data_automatica.strftime('%d/%m/%Y'))
    
    def aplicar_data_manual():
        try:
            if tem_tkcalendar:
                nova_data = data_entry.get_date()
            else:
                data_str = data_entry.get()
                nova_data = datetime.strptime(data_str, '%d/%m/%Y')
            
            data_label.config(text=nova_data.strftime('%d/%m/%Y'))
            messagebox.showinfo("Sucesso", f"Data alterada para {nova_data.strftime('%d/%m/%Y')}", parent=parent_frame.winfo_toplevel())
        except Exception as e:
            messagebox.showerror("Erro", f"Data inválida: {str(e)}", parent=parent_frame.winfo_toplevel())
    
    ttk.Button(
        frame_data_manual,
        text="✓ Aplicar",
        command=aplicar_data_manual
    ).pack(side='left')
    
    # Informação sobre quinzenas
    ttk.Separator(data_frame, orient='horizontal').pack(fill='x', pady=(10, 8))
    
    info_quinzena = ttk.Frame(data_frame)
    info_quinzena.pack(anchor='w', fill='x')
    
    ttk.Label(
        info_quinzena,
        text="📊 Quinzenas:",
        font=('Arial', 9, 'bold')
    ).pack(anchor='w')
    
    ttk.Label(
        info_quinzena,
        text="1ª Quinzena: dia 21 ao dia 5  |  2ª Quinzena: dia 6 ao dia 20",
        font=('Arial', 8),
        foreground='#666666'
    ).pack(anchor='w', pady=(2, 0))
    
    # ========== CONTRATOS INCLUÍDOS ==========
    contratos_frame = ttk.LabelFrame(parent_frame, text="📑 Contratos", padding=10)
    contratos_frame.pack(fill='x', pady=(0, 15))

    # Padrão desmarcado = relatório completo (comportamento original)
    var_escolher_contratos = tk.BooleanVar(value=False)
    ttk.Checkbutton(
        contratos_frame,
        text="Escolher contratos antes de gerar",
        variable=var_escolher_contratos
    ).pack(anchor='w')
    ttk.Label(
        contratos_frame,
        text="Ao gerar, abre a lista dos contratos do relatório (todos marcados)\n"
             "para desmarcar os que não devem sair no PDF.",
        font=('Arial', 8),
        foreground='#666666',
        justify='left'
    ).pack(anchor='w', pady=(3, 0))

    # ========== SEPARADOR ==========
    ttk.Separator(parent_frame, orient='horizontal').pack(fill='x', pady=15)
    
    # ========== BOTÃO GERAR (DESTACADO) ==========
    btn_frame = ttk.Frame(parent_frame)
    btn_frame.pack(fill='x', pady=(5, 15))
    
    # Instrução
    ttk.Label(
        btn_frame,
        text="📌 Pronto para gerar! Clique no botão abaixo:",
        font=('Arial', 9, 'bold'),
        foreground='#006600'
    ).pack(anchor='w', pady=(0, 8))
    
    def gerar_relatorio():
        """Gera o relatório quinzenal"""
        # Validar seleção de cliente
        cliente = cliente_selecionado_var.get()
        
        if not cliente or cliente == "Todos os Clientes":
            messagebox.showwarning(
                "Aviso", 
                "Por favor, selecione um cliente específico.\n\n"
                "A opção 'Todos os Clientes' não está disponível para este relatório."
            , parent=parent_frame.winfo_toplevel())
            return
        
        # Validar arquivo Clientes.xlsx
        if not arquivo_clientes_var.get():
            if usa_config_sistema and ARQUIVO_CLIENTES and ARQUIVO_CLIENTES.exists():
                arquivo_clientes_var.set(str(ARQUIVO_CLIENTES))
            else:
                messagebox.showwarning(
                    "Aviso", 
                    "Arquivo Clientes.xlsx não encontrado.\n\n"
                    "Verifique a configuração do sistema."
                , parent=parent_frame.winfo_toplevel())
                return
        
        # Obter caminho do arquivo do cliente
        arquivo_cliente = obter_caminho_arquivo_cliente(cliente)
        
        if not arquivo_cliente:
            # Se não encontrou automaticamente, perguntar ao usuário
            resposta = messagebox.askyesno(
                "Arquivo não encontrado",
                f"O arquivo para o cliente '{cliente}' não foi encontrado automaticamente.\n\n"
                f"Deseja selecionar o arquivo manualmente?"
            , parent=parent_frame.winfo_toplevel())
            
            if resposta:
                pasta_inicial = str(PASTA_CLIENTES) if usa_config_sistema and PASTA_CLIENTES and PASTA_CLIENTES.exists() else None
                
                arquivo_cliente = filedialog.askopenfilename(
                    title=f"Selecionar Arquivo do Cliente: {cliente}",
                    initialdir=pasta_inicial,
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                
                if not arquivo_cliente:
                    return
            else:
                return
        
        try:
            # Obter data de referência
            if var_alterar_data.get():
                # Data foi alterada manualmente
                if tem_tkcalendar:
                    data_ref = data_entry.get_date()
                else:
                    try:
                        data_str = data_entry.get()
                        data_ref = datetime.strptime(data_str, '%d/%m/%Y')
                    except ValueError:
                        messagebox.showerror("Erro", "Data inválida! Use o formato DD/MM/AAAA", parent=parent_frame.winfo_toplevel())
                        return
            else:
                # Usar data automática
                data_ref = data_automatica

            if RelatorioQuinzenalPDF is None:
                raise ImportError("Módulo gerar_relatorio_quinzenal_pdf não encontrado")

            _lembrar_cliente(cliente)

            # === SELEÇÃO MANUAL DE CONTRATOS (opcional) ===
            contratos_excluidos = None
            if var_escolher_contratos.get():
                gerador_lista = RelatorioQuinzenalPDF(arquivo_cliente, arquivo_clientes_var.get())
                contratos = gerador_lista.listar_contratos(data_ref)
                if not contratos:
                    messagebox.showinfo(
                        "Aviso",
                        f"Nenhuma medição foi encontrada na quinzena especificada.\n\n"
                        f"Cliente: {cliente}\n"
                        f"Data de referência: {data_ref.strftime('%d/%m/%Y')}"
                    , parent=parent_frame.winfo_toplevel())
                    return
                contratos_excluidos = escolher_contratos_dialog(
                    parent_frame.winfo_toplevel(), contratos, cliente)
                if contratos_excluidos is None:
                    return  # usuário cancelou

            # Sugerir nome do arquivo e pedir local de salvamento
            nome_cliente_arquivo = cliente.replace(' ', '_').upper()
            sufixo = '_PARCIAL' if contratos_excluidos else ''
            nome_sugerido = f"REL_MEDICOES_{nome_cliente_arquivo}_{data_ref.strftime('%d-%m-%Y')}{sufixo}.pdf"

            arquivo_saida = filedialog.asksaveasfilename(
                title="Salvar Relatório PDF",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=nome_sugerido
            )

            if not arquivo_saida:
                return

            # === GERAR O PDF ===
            # gerador.gerar_pdf() já faz todo o trabalho internamente (carrega
            # o cliente, filtra a quinzena, monta o PDF) e retorna None se não
            # encontrar nenhuma medição no período — não precisamos (e a
            # classe atual não tem) métodos separados de pré-verificação.
            gerador = RelatorioQuinzenalPDF(
                arquivo_cliente,
                arquivo_clientes_var.get()
            )

            progress = tk.Toplevel(sistema_relatorios.root)
            progress.title("Gerando PDF")
            progress.geometry("400x120")
            progress.transient(sistema_relatorios.root)
            progress.grab_set()

            # Centralizar
            progress.update_idletasks()
            x = (progress.winfo_screenwidth() // 2) - 200
            y = (progress.winfo_screenheight() // 2) - 60
            progress.geometry(f"+{x}+{y}")

            frame_prog = ttk.Frame(progress, padding=20)
            frame_prog.pack(fill='both', expand=True)

            ttk.Label(frame_prog, text="Gerando relatório...", font=('Arial', 11)).pack(pady=10)

            progress_bar = ttk.Progressbar(frame_prog, mode='indeterminate', length=300)
            progress_bar.pack()
            progress_bar.start()

            progress.update()

            try:
                resultado = gerador.gerar_pdf(data_ref, arquivo_saida,
                                              contratos_excluidos=contratos_excluidos)
            finally:
                # Garante que a janela de progresso feche mesmo se gerar_pdf
                # levantar uma exceção — antes, um erro aqui deixava a janela
                # "Gerando PDF" presa na tela até o usuário fechar manualmente.
                progress_bar.stop()
                progress.destroy()

            # === MOSTRAR RESULTADO ===
            if resultado:
                resposta = messagebox.askyesno(
                    "Sucesso!",
                    f"✅ Relatório gerado com sucesso!\n\n"
                    f"Cliente: {cliente}\n"
                    f"Arquivo: {Path(resultado).name}\n\n"
                    f"Deseja abrir o PDF?"
                , parent=parent_frame.winfo_toplevel())
                
                if resposta:
                    sistema_relatorios.abrir_arquivo(resultado)
            else:
                # gerar_pdf() retorna None quando não encontra nenhuma
                # medição na quinzena da data informada
                messagebox.showinfo(
                    "Aviso",
                    f"Nenhuma medição foi encontrada na quinzena especificada.\n\n"
                    f"Cliente: {cliente}\n"
                    f"Data de referência: {data_ref.strftime('%d/%m/%Y')}\n\n"
                    f"Verifique a data de referência."
                , parent=parent_frame.winfo_toplevel())
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{str(e)}", parent=parent_frame.winfo_toplevel())
            import traceback
            traceback.print_exc()
    
    # Botão grande e destacado
    btn_gerar = ttk.Button(
        btn_frame,
        text="🚀 Gerar Relatório PDF",
        command=gerar_relatorio,
        style='Accentuated.TButton'
    )
    btn_gerar.pack(fill='x', ipady=8)
    
    # ========== DICA FINAL ==========
    dica_frame = ttk.Frame(parent_frame, relief='solid', borderwidth=1)
    dica_frame.pack(fill='x', pady=(10, 0))
    
    # Fundo cinza claro
    dica_frame.configure(style='Info.TFrame')
    
    dica_content = ttk.Frame(dica_frame, padding=10)
    dica_content.pack(fill='x')
    
    ttk.Label(
        dica_content,
        text="💡 Dica",
        font=('Arial', 9, 'bold'),
        foreground='#666666'
    ).pack(anchor='w')
    
    ttk.Label(
        dica_content,
        text="O relatório mostra o histórico completo de cada contrato,\ncom destaque visual para as medições da quinzena atual.",
        font=('Arial', 8),
        foreground='#666666',
        justify='left'
    ).pack(anchor='w', pady=(3, 0))


# Para compatibilidade com o sistema de relatórios
