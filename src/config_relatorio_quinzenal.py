"""
Configuração do Relatório Quinzenal de Medições para integração com relatorios_interface.py
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from datetime import datetime
from pathlib import Path
import sys

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
    arquivo_cliente_var = tk.StringVar()
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
✓ Resumo financeiro completo"""
    
    ttk.Label(desc_frame, text=desc_text, justify='left', font=('Arial', 9)).pack(anchor='w')
    
    # ========== SELEÇÃO DE ARQUIVOS ==========
    files_frame = ttk.LabelFrame(parent_frame, text="📁 Arquivos Necessários", padding=10)
    files_frame.pack(fill='x', pady=(0, 15))
    
    # Arquivo do Cliente
    ttk.Label(files_frame, text="1. Arquivo do Cliente:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0, 5))
    
    if usa_config_sistema and PASTA_CLIENTES:
        ttk.Label(
            files_frame,
            text=f"📂 {PASTA_CLIENTES}",
            font=('Arial', 8),
            foreground='blue'
        ).pack(anchor='w', pady=(0, 5))
    
    frame_cliente = ttk.Frame(files_frame)
    frame_cliente.pack(fill='x', pady=(0, 10))
    
    lbl_arquivo_cliente = ttk.Label(
        frame_cliente,
        text="Nenhum arquivo selecionado",
        foreground='gray'
    )
    lbl_arquivo_cliente.pack(side='left', fill='x', expand=True)
    
    def selecionar_cliente():
        pasta_inicial = str(PASTA_CLIENTES) if usa_config_sistema and PASTA_CLIENTES and PASTA_CLIENTES.exists() else None
        
        arquivo = filedialog.askopenfilename(
            title="Selecionar Arquivo do Cliente",
            initialdir=pasta_inicial,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if arquivo:
            arquivo_cliente_var.set(arquivo)
            nome = Path(arquivo).name
            lbl_arquivo_cliente.config(text=f"✅ {nome}", foreground='green')
    
    ttk.Button(
        frame_cliente,
        text="📁 Selecionar",
        command=selecionar_cliente
    ).pack(side='right')
    
    # Arquivo Clientes.xlsx
    ttk.Label(files_frame, text="2. Arquivo de Clientes:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(0, 5))
    
    frame_clientes = ttk.Frame(files_frame)
    frame_clientes.pack(fill='x')
    
    if arquivo_clientes_var.get():
        texto_inicial = f"✅ {Path(arquivo_clientes_var.get()).name}"
        cor_inicial = 'green'
    else:
        texto_inicial = "Nenhum arquivo selecionado"
        cor_inicial = 'gray'
    
    lbl_arquivo_clientes = ttk.Label(
        frame_clientes,
        text=texto_inicial,
        foreground=cor_inicial
    )
    lbl_arquivo_clientes.pack(side='left', fill='x', expand=True)
    
    def selecionar_clientes():
        try:
            from src.config.config import BASE_PATH
            pasta_inicial = str(BASE_PATH) if BASE_PATH.exists() else None
        except:
            pasta_inicial = None
        
        arquivo = filedialog.askopenfilename(
            title="Selecionar Clientes.xlsx",
            initialdir=pasta_inicial,
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        
        if arquivo:
            arquivo_clientes_var.set(arquivo)
            nome = Path(arquivo).name
            lbl_arquivo_clientes.config(text=f"✅ {nome}", foreground='green')
    
    # ttk.Button(
    #     frame_clientes,
    #     text="📁 Selecionar",
    #     command=selecionar_clientes
    # ).pack(side='right')
    
    # ========== DATA DE REFERÊNCIA ==========
    data_frame = ttk.LabelFrame(parent_frame, text="📅 Data de Referência", padding=10)
    data_frame.pack(fill='x', pady=(0, 15))
    
    ttk.Label(
        data_frame,
        text="Selecione a data da quinzena:",
        font=('Arial', 10)
    ).pack(anchor='w', pady=(0, 5))
    
    if tem_tkcalendar:
        data_entry = DateEntry(
            data_frame,
            width=15,
            background='darkblue',
            foreground='white',
            borderwidth=2,
            date_pattern='dd/mm/yyyy',
            locale='pt_BR',
            font=('Arial', 11)
        )
        data_entry.pack(anchor='w')
        data_entry.set_date(datetime.now())
    else:
        frame_data_input = ttk.Frame(data_frame)
        frame_data_input.pack(anchor='w')
        
        ttk.Label(frame_data_input, text="Data (DD/MM/AAAA):").pack(side='left', padx=(0, 5))
        data_entry = ttk.Entry(frame_data_input, width=15, font=('Arial', 11))
        data_entry.pack(side='left')
        data_entry.insert(0, datetime.now().strftime('%d/%m/%Y'))
    
    # Informação sobre quinzenas
    info_quinzena = ttk.Frame(data_frame)
    info_quinzena.pack(anchor='w', pady=(8, 0), fill='x')
    
    ttk.Label(
        info_quinzena,
        text="ℹ️",
        font=('Arial', 10)
    ).pack(side='left', padx=(0, 5))
    
    ttk.Label(
        info_quinzena,
        text="1ª Quinzena: dia 21 ao dia 5  |  2ª Quinzena: dia 6 ao dia 20",
        font=('Arial', 9),
        foreground='#0066CC'
    ).pack(side='left')
    
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
        # Validar
        if not arquivo_cliente_var.get():
            messagebox.showwarning("Aviso", "Por favor, selecione o arquivo do cliente.")
            return
        
        if not arquivo_clientes_var.get():
            messagebox.showwarning("Aviso", "Por favor, selecione o arquivo Clientes.xlsx.")
            return
        
        try:
            # Obter data
            if tem_tkcalendar:
                data_ref = data_entry.get_date()
            else:
                data_str = data_entry.get()
                try:
                    data_ref = datetime.strptime(data_str, '%d/%m/%Y')
                except ValueError:
                    messagebox.showerror("Erro", "Data inválida! Use o formato DD/MM/AAAA")
                    return
            
            # === PASSO 1: VERIFICAR SE HÁ DADOS ANTES DE PEDIR ARQUIVO ===
            # Criar janela de verificação
            verificacao = tk.Toplevel(sistema_relatorios.root)
            verificacao.title("Verificando dados")
            verificacao.geometry("350x100")
            verificacao.transient(sistema_relatorios.root)
            verificacao.grab_set()
            
            # Centralizar
            verificacao.update_idletasks()
            x = (verificacao.winfo_screenwidth() // 2) - 175
            y = (verificacao.winfo_screenheight() // 2) - 50
            verificacao.geometry(f"+{x}+{y}")
            
            frame_verif = ttk.Frame(verificacao, padding=20)
            frame_verif.pack(fill='both', expand=True)
            
            ttk.Label(frame_verif, text="Verificando medições...", font=('Arial', 11)).pack(pady=10)
            
            progress_bar_verif = ttk.Progressbar(frame_verif, mode='indeterminate', length=250)
            progress_bar_verif.pack()
            progress_bar_verif.start()
            
            verificacao.update()
            
            # Criar gerador e verificar dados
            if RelatorioQuinzenalPDF is None:
                raise ImportError("Módulo gerar_relatorio_quinzenal_pdf não encontrado")
            
            gerador = RelatorioQuinzenalPDF(
                arquivo_cliente_var.get(),
                arquivo_clientes_var.get()
            )
            
            # Carregar dados do cliente
            gerador.carregar_dados_cliente()
            
            # Verificar se há medições na quinzena
            contratos_encontrados = gerador.filtrar_medicoes_quinzena(data_ref)
            
            progress_bar_verif.stop()
            verificacao.destroy()
            
            # Se não encontrou medições, avisar e parar
            if not contratos_encontrados:
                data_inicio, data_fim = gerador.identificar_quinzena(data_ref)
                messagebox.showinfo(
                    "Aviso",
                    f"Nenhuma medição foi encontrada na quinzena especificada.\n\n"
                    f"Período verificado:\n"
                    f"  • De: {data_inicio.strftime('%d/%m/%Y')}\n"
                    f"  • Até: {data_fim.strftime('%d/%m/%Y')}\n\n"
                    f"Verifique a data de referência."
                )
                return
            
            # === PASSO 2: HÁ DADOS! AGORA PODE PEDIR ARQUIVO DE SAÍDA ===
            # Sugerir nome do arquivo
            nome_cliente = Path(arquivo_cliente_var.get()).stem.replace('_', ' ').upper()
            nome_sugerido = f"REL_MEDICOES_{nome_cliente}_{data_ref.strftime('%d-%m-%Y')}.pdf"
            
            # Solicitar local de salvamento
            arquivo_saida = filedialog.asksaveasfilename(
                title="Salvar Relatório PDF",
                defaultextension=".pdf",
                filetypes=[("PDF files", "*.pdf")],
                initialfile=nome_sugerido
            )
            
            if not arquivo_saida:
                return
            
            # === PASSO 3: GERAR O PDF ===
            # Criar janela de progresso
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
            
            # Gerar PDF (já temos os dados carregados no gerador)
            resultado = gerador.gerar_pdf(data_ref, arquivo_saida)
            
            progress_bar.stop()
            progress.destroy()
            
            # === PASSO 4: MOSTRAR RESULTADO ===
            if resultado:
                # Perguntar se quer abrir
                total_contratos = len(gerador.contratos_quinzena) if hasattr(gerador, 'contratos_quinzena') else 0
                
                resposta = messagebox.askyesno(
                    "Sucesso!",
                    f"✅ Relatório gerado com sucesso!\n\n"
                    f"Arquivo: {Path(resultado).name}\n"
                    f"Total de contratos: {total_contratos}\n\n"
                    f"Deseja abrir o PDF?"
                )
                
                if resposta:
                    sistema_relatorios.abrir_arquivo(resultado)
            else:
                # Não deveria chegar aqui, pois já verificamos antes
                messagebox.showwarning(
                    "Aviso",
                    "Não foi possível gerar o relatório."
                )
        
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{str(e)}")
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
class ConfiguracaoRelatorioQuinzenal:
    """Classe wrapper para manter compatibilidade"""
    
    def __init__(self, parent_frame, sistema_relatorios):
        configurar_relatorio_quinzenal(parent_frame, sistema_relatorios)
