"""
Parcelamento genérico de despesas (fatura, compra no cartão, valor fixo
com entrada, datas específicas, ou parcelas totalmente personalizadas).

Usada tanto na Entrada de Dados (parcelamento de fornecedor comum) quanto
dentro de GestaoTaxasFixas (parcelamento de taxa de administração fixa) —
por isso vive em seu próprio módulo genérico, fora de taxas_administracao/.

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].
Nenhuma alteração de lógica foi feita nesta extração — apenas mudança de
localização e ajuste de imports.

ATENÇÃO — pendências conhecidas, não corrigidas nesta extração:
    1) O método run() no final da classe (self.root.mainloop()) não
       corresponde a nenhum atributo real desta classe (que usa
       self.parent, não self.root) — parece resíduo colado de outra
       classe (provavelmente SistemaEntradaDados). Mantido por
       fidelidade ao original; aparenta ser código morto (nenhuma
       chamada a .run() encontrada nas buscas feitas até agora), mas
       vale confirmar antes de remover.
    2) O import de ComboboxAutocompletar, que no arquivo original vivia
       comentado no corpo da classe (funcionava só porque o nome já
       existia no namespace global do Sistema_Entrada_Dados.py), foi
       tornado explícito abaixo — necessário para a classe funcionar
       fora daquele arquivo.
"""
import logging
from datetime import datetime

import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from dateutil.relativedelta import relativedelta

from src.config.dialogs import custom_messagebox
from src.widgets.combobox_autocompletar import ComboboxAutocompletar
from src.configuracoes_sistema import GerenciadorConfiguracoes

# Mesmo logger usado no restante do sistema (Sistema_Entrada_Dados.py).
logger = logging.getLogger("sistema")


class GestorParcelas:
    # Mantido como atributo de classe (padrão do arquivo original) —
    # redundante com o import local em abrir_janela_parcelas, mas
    # preservado por fidelidade à extração.
    GerenciadorConfiguracoes = GerenciadorConfiguracoes

    def __init__(self, parent):
        logger.debug("Inicializando GestorParcelas")  # Debug
        self.parent = parent
        self.parcelas = []
        self.tipo_despesa_valor = '3'
        self.janela_parcelas = None
        self._var_tem_entrada = None  # Inicializa como None
        # Limpar referências de widgets
        self.frame_modalidade = None
        self.frame_valor_entrada = None
        self.lbl_entrada = None
        self.valor_entrada = None
        self.modalidade_entrada = None
        self.parcelas_personalizadas = []
        self.frame_parcelas_personalizadas = None
        self.valor_total_personalizado = 0.0
        self.canvas_parcelas = None
        self.scrollbar_parcelas = None
        self.campos_parcelas = []

    @property
    def tem_entrada(self):
        """Getter para tem_entrada - cria apenas quando necessário"""
        if self._var_tem_entrada is None:
            self._var_tem_entrada = tk.BooleanVar(master=self.parent.root, value=False)
        return self._var_tem_entrada

    # Interface e Controles
    def abrir_janela_parcelas(self):
        logger.debug("Abrindo janela de parcelas")  # Debug
        # Criar janela como Toplevel do parent
        self.janela_parcelas = tk.Toplevel(self.parent.root)
        self.janela_parcelas.title("Configuração de Parcelas")
        self.janela_parcelas.geometry("700x800")  # Aumentado para acomodar nova opção

        # Garantir que a janela seja modal
        self.janela_parcelas.transient(self.parent.root)
        self.janela_parcelas.grab_set()

        frame = ttk.Frame(self.janela_parcelas, padding="10")
        frame.pack(fill='both', expand=True)

        # Frame para entrada
        frame_entrada = ttk.LabelFrame(frame, text="Entrada")
        frame_entrada.grid(row=0, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        logger.debug("Criando Checkbutton")  # Debug
        check = ttk.Checkbutton(
            frame_entrada,
            text="Possui entrada?",
            variable=self.tem_entrada,
            command=self.atualizar_campos_entrada
        )
        check.grid(row=0, column=0, padx=5, pady=5)

        # Frame para modalidades de entrada
        logger.debug("Criando frame modalidade")  # Debug
        self.frame_modalidade = ttk.Frame(frame_entrada)
        self.frame_modalidade.grid(row=1, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        ttk.Label(self.frame_modalidade, text="Modalidade de Entrada:").grid(row=0, column=0, padx=5, pady=2)
        self.modalidade_entrada = ttk.Combobox(self.frame_modalidade, state='readonly', width=40)
        self.modalidade_entrada['values'] = [
            "Percentual do valor total na primeira parcela",
            "Primeira parcela igual às demais (arredonda no final)",
            "Valor específico na primeira parcela"
        ]
        self.modalidade_entrada.grid(row=0, column=1, padx=5, pady=2)

        # Garantir que o frame modalidade começa oculto
        logger.debug("Ocultando frame modalidade inicialmente")  # Debug
        self.frame_modalidade.grid_remove()

        # Frame para valor da entrada (dinâmico baseado na modalidade)
        self.frame_valor_entrada = ttk.Frame(frame_entrada)
        self.frame_valor_entrada.grid(row=2, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

        # Ocultar frames inicialmente
        self.frame_modalidade.grid_remove()
        self.frame_valor_entrada.grid_remove()

        # Tipo de Despesa
        ttk.Label(frame, text="Tipo de Despesa:").grid(row=1, column=0, padx=5, pady=5)
        self.tipo_despesa = ttk.Combobox(frame, values=['2', '3', '5', '6'], state='readonly', width=5)
        self.tipo_despesa.grid(row=1, column=1, sticky='w', padx=5, pady=5)
        self.tipo_despesa.set('3')  # Tipo 3 como padrão

        # Tipo de Parcelamento
        ttk.Label(frame, text="Tipo de Parcelamento:").grid(row=2, column=0, padx=5, pady=5)
        self.tipo_parcelamento = ttk.Combobox(frame, values=[
            "Prazo Fixo em Dias",
            "Datas Específicas",
            "Cartão de Crédito",
            "Parcelas Personalizadas"
        ], state="readonly")
        self.tipo_parcelamento.grid(row=2, column=1, padx=5, pady=5)
        self.tipo_parcelamento.set("Prazo Fixo em Dias")
        self.tipo_parcelamento.bind('<<ComboboxSelected>>', self.atualizar_campos_parcelamento)

        # Frame para campos dinâmicos
        self.frame_dinamico = ttk.Frame(frame)
        self.frame_dinamico.grid(row=3, column=0, columnspan=2, pady=10, sticky='ew')

        # Campos comuns
        ttk.Label(frame, text="Data da Despesa:").grid(row=4, column=0, padx=5, pady=5)
        self.data_despesa = DateEntry(
            frame,
            format='dd/mm/yyyy',
            locale='pt_BR',
            background='darkblue',
            foreground='white',
            borderwidth=2
        )

        self.data_despesa.grid(row=4, column=1, padx=5, pady=5)
        self.data_despesa.configure(state='normal')
        self.configurar_calendario(self.data_despesa)

        ttk.Label(frame, text="Valor Original:").grid(row=5, column=0, padx=5, pady=5)
        self.valor_original = ttk.Entry(frame)
        self.valor_original.grid(row=5, column=1, padx=5, pady=5)
        self.valor_original.bind('<KeyPress>', self.on_valor_original_manual_edit)

        # Alterar o label do número de parcelas para ser mais claro
        self.lbl_num_parcelas = ttk.Label(frame, text="Número de Parcelas:")
        self.lbl_num_parcelas.grid(row=6, column=0, padx=5, pady=5)
        self.num_parcelas = ttk.Entry(frame)
        self.num_parcelas.grid(row=6, column=1, padx=5, pady=5)

        # Frame específico para informação sobre parcelas
        frame_info_parcelas = ttk.Frame(frame)
        frame_info_parcelas.grid(row=7, column=0, columnspan=2, padx=5, pady=5, sticky='ew')

        self.lbl_info_parcelas = ttk.Label(
            frame_info_parcelas,
            text="",
            wraplength=500,  # Permitir quebra de linha se necessário
            justify='center'
        )
        self.lbl_info_parcelas.pack(fill='x', padx=5)

        # Referência Base
        ttk.Label(frame, text="Referência Base:").grid(row=8, column=0, padx=5, pady=5)
        self.referencia_base = ttk.Entry(frame)
        self.referencia_base.grid(row=8, column=1, padx=5, pady=5, sticky='ew')

        # Campo NF
        ttk.Label(frame, text="NF:").grid(row=9, column=0, padx=5, pady=5)
        self.campos_nf = ttk.Entry(frame)
        self.campos_nf.grid(row=9, column=1, padx=5, pady=5, sticky='ew')

        # Campos Etapa da Obra e Insumos
        ttk.Label(frame, text="Etapa da Obra:").grid(row=10, column=0, padx=5, pady=5)

        etapas_obra = GerenciadorConfiguracoes.get_etapas_obra()

        self.etapa_obra = ComboboxAutocompletar(
            frame,
            values=etapas_obra,
            config_key='etapas_obra',
            config_manager=GerenciadorConfiguracoes,
            width=37,  # Ajustado para combinar com outros campos
            state='normal'
        )
        self.etapa_obra.grid(row=10, column=1, padx=5, pady=5, sticky='ew')

        # Campo Insumos - USANDO COMBOBOX AUTOCOMPLETAR
        ttk.Label(frame, text="Insumos:").grid(row=11, column=0, padx=5, pady=5)

        insumos = GerenciadorConfiguracoes.get_insumos()

        self.insumo = ComboboxAutocompletar(
            frame,
            values=insumos,
            config_key='insumos',
            config_manager=GerenciadorConfiguracoes,
            width=37,
            state='normal'
        )
        self.insumo.grid(row=11, column=1, padx=5, pady=5, sticky='ew')

        # Botões
        frame_botoes = ttk.Frame(frame)
        frame_botoes.grid(row=12, column=0, columnspan=2, pady=30)

        ttk.Button(frame_botoes,
                  text="Gerar Parcelas",
                  command=self.gerar_parcelas).pack(side='left', padx=5)
        ttk.Button(frame_botoes,
                  text="Cancelar",
                  command=self.cancelar_parcelamento).pack(side='left', padx=5)

        # Inicializar campos do tipo padrão
        self.atualizar_campos_parcelamento(None)

        # Fazer a janela modal
        self.janela_parcelas.transient(self.parent.root)
        self.janela_parcelas.grab_set()

        # Centralizar a janela
        self.janela_parcelas.update_idletasks()
        width = self.janela_parcelas.winfo_width()
        height = self.janela_parcelas.winfo_height()
        x = (self.janela_parcelas.winfo_screenwidth() // 2) - (width // 2)
        y = (self.janela_parcelas.winfo_screenheight() // 2) - (height // 2)
        self.janela_parcelas.geometry(f'{width}x{height}+{x}+{y}')

    def atualizar_campos_entrada(self):
        """Mostra/oculta campos relacionados à entrada e atualiza labels"""
        if self.tem_entrada.get():
            # Mostrar frame modalidade
            if self.frame_modalidade:
                self.frame_modalidade.grid()

                # Criar campos se não existirem
                if not hasattr(self, 'valor_entrada') or not self.valor_entrada:
                    if not self.frame_valor_entrada:
                        self.frame_valor_entrada = ttk.Frame(self.frame_modalidade)
                        self.frame_valor_entrada.grid(row=1, column=0, columnspan=2, sticky='ew', padx=5, pady=5)

                    self.lbl_entrada = ttk.Label(self.frame_valor_entrada, text="Valor:")
                    self.lbl_entrada.grid(row=0, column=0, padx=5, pady=2)

                    self.valor_entrada = ttk.Entry(self.frame_valor_entrada)
                    self.valor_entrada.grid(row=0, column=1, padx=5, pady=2)

                if self.frame_valor_entrada:
                    self.frame_valor_entrada.grid()
        else:
            # Ocultar frames
            if self.frame_modalidade:
                self.frame_modalidade.grid_remove()
            if self.frame_valor_entrada:
                self.frame_valor_entrada.grid_remove()

            # Restaurar label original
            for widget in self.janela_parcelas.winfo_children():
                if isinstance(widget, ttk.Label) and widget.cget("text").startswith("Número de Parcelas"):
                    widget.config(text="Número de Parcelas:")
            self.lbl_info_parcelas.config(text="")

    def atualizar_campos_modalidade(self, event=None):
        """Atualiza campos baseado na modalidade selecionada"""
        modalidade = self.modalidade_entrada.get()

        if not hasattr(self, 'frame_valor_entrada') or not hasattr(self, 'lbl_entrada'):
            return

        self.frame_valor_entrada.grid()

        if modalidade == "Percentual do valor total na primeira parcela":
            self.lbl_entrada.config(text="Percentual (%): ")
            self.valor_entrada.delete(0, tk.END)
        elif modalidade == "Primeira parcela igual às demais (arredonda no final)":
            self.frame_valor_entrada.grid_remove()
        elif modalidade == "Valor específico na primeira parcela":
            self.lbl_entrada.config(text="Valor (R$): ")
            self.valor_entrada.delete(0, tk.END)

    def atualizar_campos_parcelamento(self, event):
        # Limpar frame dinâmico
        for widget in self.frame_dinamico.winfo_children():
            widget.destroy()

        tipo = self.tipo_parcelamento.get()

        # Ocultar/mostrar campos baseado no tipo
        if tipo == "Parcelas Personalizadas":
            # Ocultar campos tradicionais
            self.lbl_num_parcelas.grid_remove()
            self.num_parcelas.grid_remove()
            self.lbl_info_parcelas.config(text="Configure cada parcela individualmente")

            # Criar interface para parcelas personalizadas
            self.criar_interface_parcelas_personalizadas()
        else:
            # Mostrar campos tradicionais
            self.lbl_num_parcelas.grid()
            self.num_parcelas.grid()

            # Lógica existente para outros tipos
            if tipo == "Prazo Fixo em Dias":
                ttk.Label(self.frame_dinamico, text="Prazo entre Parcelas (dias):").grid(row=0, column=0, padx=5, pady=5)
                self.prazo_dias = ttk.Entry(self.frame_dinamico)
                self.prazo_dias.grid(row=0, column=1, padx=5, pady=5)
                self.prazo_dias.insert(0, "30")  # Valor padrão

            elif tipo == "Datas Específicas":
                num_parcelas_txt = "parcelas após a entrada" if self.tem_entrada.get() else "parcelas"

                ttk.Label(self.frame_dinamico,
                         text=f"Informe as datas de vencimento das {num_parcelas_txt}:").grid(
                             row=0, column=0, columnspan=2, padx=5, pady=5)

                self.texto_datas = tk.Text(self.frame_dinamico, height=4, width=30)
                self.texto_datas.grid(row=1, column=0, columnspan=2, padx=5, pady=5)

                ttk.Label(self.frame_dinamico,
                         text="Digite uma data por linha no formato dd/mm/aaaa\n"
                              "(não inclua a data da entrada)").grid(
                             row=2, column=0, columnspan=2, padx=5, pady=5)

            elif tipo == "Cartão de Crédito":
                ttk.Label(self.frame_dinamico, text="Dia do Vencimento:").grid(row=0, column=0, padx=5, pady=5)
                self.dia_vencimento = ttk.Entry(self.frame_dinamico, width=5)
                self.dia_vencimento.grid(row=0, column=1, padx=5, pady=5)
                self.dia_vencimento.insert(0, "10")  # Valor padrão

    def criar_interface_parcelas_personalizadas(self):
        """Cria interface específica para parcelas personalizadas"""

        # Frame para controles básicos
        frame_controles = ttk.Frame(self.frame_dinamico)
        frame_controles.pack(fill='x', pady=5)

        # Número total de parcelas
        ttk.Label(frame_controles, text="Número de Parcelas:").pack(side='left', padx=5)
        self.num_parcelas_personalizado = tk.IntVar(value=2)
        spin_parcelas = ttk.Spinbox(frame_controles, from_=2, to=12,
                                   textvariable=self.num_parcelas_personalizado,
                                   width=5, command=self.atualizar_grid_parcelas)
        spin_parcelas.pack(side='left', padx=5)

        # Botão para gerar grid
        ttk.Button(frame_controles, text="Gerar Grid",
                  command=self.atualizar_grid_parcelas).pack(side='left', padx=10)

        # Frame scrollável para o grid de parcelas
        self.criar_frame_scrollavel_parcelas()

        # Frame para controles da última parcela
        frame_ultima_parcela = ttk.LabelFrame(self.frame_dinamico,
                                            text="Configuração da Última Parcela")
        frame_ultima_parcela.pack(fill='x', pady=10)

        self.condicao_ultima_parcela = tk.BooleanVar()
        ttk.Checkbutton(frame_ultima_parcela,
                       text="Última parcela depende de entrega/condição específica",
                       variable=self.condicao_ultima_parcela,
                       command=self.toggle_condicao_ultima_parcela).pack(anchor='w', padx=5, pady=5)

        # Frame para data condicional (inicialmente oculto)
        self.frame_data_condicional = ttk.Frame(frame_ultima_parcela)

        ttk.Label(self.frame_data_condicional,
                 text="Data estimada:").pack(side='left', padx=5)

        self.data_condicional = DateEntry(
            self.frame_data_condicional,
            format='dd/mm/yyyy',
            locale='pt_BR',
            background='darkblue',
            foreground='white',
            borderwidth=2,
            width=12
        )
        self.data_condicional.pack(side='left', padx=5)

        ttk.Label(self.frame_data_condicional,
                 text="Observação:").pack(side='left', padx=(20, 5))

        self.obs_condicional = ttk.Entry(self.frame_data_condicional, width=25)
        self.obs_condicional.pack(side='left', padx=5)

    def criar_frame_scrollavel_parcelas(self):
        """Cria um frame scrollável para edição das parcelas"""

        # Frame container para o canvas e scrollbar
        container_frame = ttk.Frame(self.frame_dinamico)
        container_frame.pack(fill='both', expand=True, pady=10)

        self.canvas_parcelas = tk.Canvas(container_frame, height=200, width=650)  # Largura aumentada
        self.scrollbar_parcelas = ttk.Scrollbar(container_frame, orient="vertical", command=self.canvas_parcelas.yview)
        self.frame_parcelas_personalizadas = ttk.Frame(self.canvas_parcelas)

        self.frame_parcelas_personalizadas.bind(
            "<Configure>",
            lambda e: self.canvas_parcelas.configure(scrollregion=self.canvas_parcelas.bbox("all"))
        )

        def configure_canvas_width(event):
            canvas_width = event.width
            self.canvas_parcelas.itemconfig(self.canvas_window_id, width=canvas_width)

        self.canvas_parcelas.bind('<Configure>', configure_canvas_width)

        # Criar a janela no canvas
        self.canvas_window_id = self.canvas_parcelas.create_window((0, 0), window=self.frame_parcelas_personalizadas, anchor="nw")
        self.canvas_parcelas.configure(yscrollcommand=self.scrollbar_parcelas.set)

        self.canvas_parcelas.pack(side="left", fill="both", expand=True)
        self.scrollbar_parcelas.pack(side="right", fill="y")

        # Bind mouse wheel para scroll
        def _on_mousewheel(event):
            self.canvas_parcelas.yview_scroll(int(-1*(event.delta/120)), "units")
        self.canvas_parcelas.bind_all("<MouseWheel>", _on_mousewheel)

    def atualizar_grid_parcelas(self):
        """Atualiza o grid de parcelas baseado no número selecionado"""

        if not hasattr(self, 'frame_parcelas_personalizadas'):
            return

        # Limpar grid existente
        for widget in self.frame_parcelas_personalizadas.winfo_children():
            widget.destroy()

        num_parcelas = self.num_parcelas_personalizado.get()

        self.frame_parcelas_personalizadas.columnconfigure(0, weight=0, minsize=60)   # Parcela - fixo
        self.frame_parcelas_personalizadas.columnconfigure(1, weight=1, minsize=120)  # Valor - expansível
        self.frame_parcelas_personalizadas.columnconfigure(2, weight=1, minsize=130)  # Data - expansível
        self.frame_parcelas_personalizadas.columnconfigure(3, weight=2, minsize=200)  # Observação - maior espaço

        # Cabeçalho do grid
        ttk.Label(self.frame_parcelas_personalizadas, text="Parcela",
                font=('Arial', 10, 'bold')).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        ttk.Label(self.frame_parcelas_personalizadas, text="Valor (R$)",
                font=('Arial', 10, 'bold')).grid(row=0, column=1, padx=5, pady=5, sticky='w')
        ttk.Label(self.frame_parcelas_personalizadas, text="Data Vencimento",
                font=('Arial', 10, 'bold')).grid(row=0, column=2, padx=5, pady=5, sticky='w')
        ttk.Label(self.frame_parcelas_personalizadas, text="Observação",
                font=('Arial', 10, 'bold')).grid(row=0, column=3, padx=5, pady=5, sticky='w')

        # Criar campos para cada parcela
        self.campos_parcelas = []

        for i in range(num_parcelas):
            parcela_num = i + 1

            # Número da parcela
            ttk.Label(self.frame_parcelas_personalizadas,
                    text=f"{parcela_num}ª").grid(row=parcela_num, column=0, padx=5, pady=2, sticky='w')

            # Campo valor
            valor_var = tk.StringVar()
            entry_valor = ttk.Entry(self.frame_parcelas_personalizadas,
                                textvariable=valor_var, width=15)
            entry_valor.grid(row=parcela_num, column=1, padx=5, pady=2, sticky='ew')
            entry_valor.bind('<KeyRelease>', self.calcular_total_personalizado)

            # Campo data
            data_entry = DateEntry(
                self.frame_parcelas_personalizadas,
                format='dd/mm/yyyy',
                locale='pt_BR',
                background='darkblue',
                foreground='white',
                borderwidth=1,
                width=12
            )
            data_entry.grid(row=parcela_num, column=2, padx=5, pady=2, sticky='ew')

            # Campo observação
            obs_var = tk.StringVar()
            entry_obs = ttk.Entry(self.frame_parcelas_personalizadas,
                                textvariable=obs_var)  # Removido width fixo
            entry_obs.grid(row=parcela_num, column=3, padx=5, pady=2, sticky='ew')  # sticky='ew' para expandir

            self.campos_parcelas.append({
                'valor': valor_var,
                'data': data_entry,
                'observacao': obs_var,
                'entry_valor': entry_valor
            })

        # Frame para total
        frame_total = ttk.Frame(self.frame_parcelas_personalizadas)
        frame_total.grid(row=num_parcelas + 1, column=0, columnspan=4, pady=10, sticky='ew')

        ttk.Label(frame_total, text="TOTAL:",
                font=('Arial', 10, 'bold')).pack(side='left', padx=5)

        self.label_total_personalizado = ttk.Label(frame_total, text="R$ 0,00",
                                                font=('Arial', 10, 'bold'),
                                                foreground='red')
        self.label_total_personalizado.pack(side='left', padx=5)

        # Forçar atualização do layout
        self.frame_parcelas_personalizadas.update_idletasks()

        # Forçar o canvas a reconhecer o novo tamanho
        self.canvas_parcelas.after(100, self._update_canvas_scroll)

    def _update_canvas_scroll(self):
        """Método auxiliar para atualizar o scroll do canvas após mudanças no grid"""
        if hasattr(self, 'canvas_parcelas') and hasattr(self, 'frame_parcelas_personalizadas'):
            # Atualizar a scroll region
            self.canvas_parcelas.configure(scrollregion=self.canvas_parcelas.bbox("all"))

            # Forçar atualização da largura se necessário
            canvas_width = self.canvas_parcelas.winfo_width()
            if canvas_width > 1:  # Verificar se o canvas já foi renderizado
                self.canvas_parcelas.itemconfig(self.canvas_window_id, width=canvas_width)

    def calcular_total_personalizado(self, event=None):
        """Calcula o total das parcelas personalizadas e atualiza o campo Valor Original"""
        total = 0.0

        if hasattr(self, 'campos_parcelas'):
            for campo in self.campos_parcelas:
                try:
                    valor_str = campo['valor'].get().replace(',', '.')
                    if valor_str:
                        total += float(valor_str)
                except ValueError:
                    pass

        self.valor_total_personalizado = total

        # Atualizar label do total
        if hasattr(self, 'label_total_personalizado'):
            self.label_total_personalizado.config(text=f"R$ {total:,.2f}")

        # Auto-preencher o campo Valor Original
        if hasattr(self, 'valor_original') and total > 0:
            # Limpar o campo atual
            self.valor_original.delete(0, tk.END)
            # Inserir o novo valor formatado
            self.valor_original.insert(0, f"{total:.2f}".replace('.', ','))

            # Alterar cor do campo para indicar que foi preenchido automaticamente
            self.valor_original.configure(style='AutoFilled.TEntry')

            # Criar o estilo se não existir
            try:
                style = ttk.Style()
                style.configure('AutoFilled.TEntry',
                            fieldbackground='#E8F5E8',  # Verde claro
                            bordercolor='#4CAF50')       # Verde
            except:
                pass  # Se o estilo já existir ou houver erro, ignorar

    def on_valor_original_manual_edit(self, event=None):
        """Reseta o estilo quando o usuário edita manualmente o Valor Original"""
        if hasattr(self, 'valor_original'):
            try:
                self.valor_original.configure(style='TEntry')  # Estilo padrão
            except:
                pass

    def toggle_condicao_ultima_parcela(self):
        """Mostra/oculta configurações da última parcela condicional"""
        if self.condicao_ultima_parcela.get():
            self.frame_data_condicional.pack(fill='x', padx=5, pady=5)
        else:
            self.frame_data_condicional.pack_forget()

    # Métodos de geração e validação de parcelas
    def validar_dados_entrada(self, valor_original, num_parcelas, referencia_base, tipo):
        """Valida os dados básicos antes de gerar parcelas"""
        if not referencia_base or num_parcelas <= 0:
            custom_messagebox("error", "Erro", "Preencha todos os campos obrigatórios!")
            return False

        # Validações específicas por tipo de parcelamento
        if tipo == "Prazo Fixo em Dias":
            if not hasattr(self, 'prazo_dias') or not self.prazo_dias.get():
                custom_messagebox("error", "Erro", "Informe o prazo entre as parcelas!")
                return False
        elif tipo == "Datas Específicas":
            if not hasattr(self, 'texto_datas'):
                custom_messagebox("error", "Erro", "Configure as datas específicas!")
                return False
        elif tipo == "Cartão de Crédito":
            if not hasattr(self, 'dia_vencimento') or not self.dia_vencimento.get():
                custom_messagebox("error", "Erro", "Informe o dia do vencimento!")
                return False
            try:
                dia_vencimento = int(self.dia_vencimento.get())
                if not (1 <= dia_vencimento <= 31):
                    custom_messagebox("error", "Erro", "Dia de vencimento deve estar entre 1 e 31!")
                    return False
            except ValueError:
                custom_messagebox("error", "Erro", "Dia de vencimento inválido!")
                return False

        return True

    def gerar_parcelas(self):
        """Método principal para gerar parcelas"""
        try:
            tipo = self.tipo_parcelamento.get()

            # Se for parcelas personalizadas, usar lógica específica
            if tipo == "Parcelas Personalizadas":
                return self.gerar_parcelas_personalizadas()

            # Lógica existente para outros tipos
            self.tipo_despesa_valor = self.tipo_despesa.get()
            valor_original = float(self.valor_original.get().replace(',', '.'))
            num_parcelas = int(self.num_parcelas.get())
            referencia_base = self.referencia_base.get().strip()
            nf = self.campos_nf.get().strip()

            # Validar dados
            if not self.validar_dados_entrada(valor_original, num_parcelas, referencia_base, tipo):
                return False

            # Data base é a data da despesa
            data_base = datetime.strptime(self.data_despesa.get(), '%d/%m/%Y')

            # Limpar lista de parcelas anterior
            self.parcelas = []

            # Calcular valores das parcelas
            valores_parcelas = self.calcular_valores_parcelas(valor_original, num_parcelas)
            if not valores_parcelas:
                return False

            # Gerar parcelas conforme o tipo
            if tipo == "Prazo Fixo em Dias":
                self.gerar_parcelas_prazo_fixo(data_base, valores_parcelas, referencia_base, num_parcelas, nf)
            elif tipo == "Datas Específicas":
                self.gerar_parcelas_datas_especificas(data_base, valores_parcelas, referencia_base, num_parcelas, nf)
            elif tipo == "Cartão de Crédito":
                self.gerar_parcelas_cartao(data_base, valores_parcelas, referencia_base, num_parcelas, nf)

            if self.parcelas:
                custom_messagebox("info", "Sucesso", f"{len(self.parcelas)} parcela(s) gerada(s) com sucesso!")
                self.limpar_campos()
                return True

        except Exception as e:
            custom_messagebox("error", "Erro", f"Erro ao gerar parcelas: {str(e)}")
            return False

    def gerar_parcelas_personalizadas(self):
        """Gera as parcelas personalizadas baseadas nos dados inseridos"""

        if not hasattr(self, 'campos_parcelas') or not self.campos_parcelas:
            custom_messagebox("error", "Erro", "Configure as parcelas primeiro!")
            return False

        try:
            # Obter dados básicos
            self.tipo_despesa_valor = self.tipo_despesa.get()
            referencia_base = self.referencia_base.get().strip()
            nf_base = self.campos_nf.get().strip()
            data_base = datetime.strptime(self.data_despesa.get(), '%d/%m/%Y')

            if not referencia_base:
                custom_messagebox("error", "Erro", "Referência base é obrigatória!")
                return False

            self.parcelas = []

            for i, campo in enumerate(self.campos_parcelas):
                parcela_num = i + 1

                try:
                    valor_str = campo['valor'].get().replace(',', '.')
                    if not valor_str:
                        custom_messagebox("error", "Erro", f"Valor da {parcela_num}ª parcela não informado!")
                        return False

                    valor = float(valor_str)
                    dt_vencto_obj = campo['data'].get_date()
                    data_vencto = dt_vencto_obj.strftime('%d/%m/%Y')
                    observacao = campo['observacao'].get().strip()

                    # Montar referência da parcela
                    if observacao:
                        referencia = f"{referencia_base} - {parcela_num}ª PARCELA - {observacao}"
                    else:
                        referencia = f"{referencia_base} - {parcela_num}ª PARCELA"

                    # NF da parcela
                    if nf_base:
                        nf_parcela = f"{nf_base}-{parcela_num:02d}"
                    else:
                        nf_parcela = f"PARC-{parcela_num:02d}"

                    # Verificar se é a última parcela com condição especial
                    if (parcela_num == len(self.campos_parcelas) and
                        hasattr(self, 'condicao_ultima_parcela') and
                        self.condicao_ultima_parcela.get()):

                        if hasattr(self, 'obs_condicional'):
                            obs_condicional = self.obs_condicional.get().strip()
                            if obs_condicional:
                                referencia += f" - CONDICIONADA: {obs_condicional}"
                            else:
                                referencia += " - CONDICIONADA À ENTREGA"

                        # Usar data condicional se especificada
                        if hasattr(self, 'data_condicional'):
                            data_condicional_obj = self.data_condicional.get_date()
                            if data_condicional_obj != dt_vencto_obj:
                                dt_vencto_obj = data_condicional_obj
                                data_vencto = data_condicional_obj.strftime('%d/%m/%Y')

                    # Calcular data do relatório específica para cada parcela
                    logger.debug(f"DEBUG - Parcela {parcela_num}:")
                    logger.debug(f"  Data vencimento: {dt_vencto_obj}")
                    logger.debug(f"  Tipo despesa: {self.tipo_despesa_valor}")

                    # Para parcelas personalizadas, nenhuma é considerada "primeira parcela" com entrada
                    eh_primeira_parcela = False
                    data_rel_obj = self.calcular_data_rel_personalizada(dt_vencto_obj)
                    data_rel = data_rel_obj.strftime('%d/%m/%Y')

                    logger.debug(f"  Data relatório calculada: {data_rel}")
                    logger.debug("---")

                    parcela = {
                        'data_rel': data_rel,
                        'nf': nf_parcela,
                        'referencia': referencia,
                        'valor': valor,
                        'dt_vencto': data_vencto,
                        'etapa_obra': self.etapa_obra.get().strip(),
                        'insumo': self.insumo.get().strip(),
                        'observacao': observacao
                    }

                    self.parcelas.append(parcela)

                except ValueError:
                    custom_messagebox("error", "Erro",
                                    f"Valor inválido na {parcela_num}ª parcela!")
                    return False

            if self.parcelas:
                # Mostrar resumo antes de confirmar
                self.mostrar_resumo_parcelas_personalizadas()
                return True
            else:
                custom_messagebox("error", "Erro", "Nenhuma parcela foi gerada!")
                return False

        except Exception as e:
            custom_messagebox("error", "Erro", f"Erro ao gerar parcelas personalizadas: {str(e)}")
            return False

    def mostrar_resumo_parcelas_personalizadas(self):
        """Mostra resumo das parcelas personalizadas antes da confirmação"""

        resumo_window = tk.Toplevel(self.janela_parcelas)
        resumo_window.title("Resumo das Parcelas Personalizadas")
        resumo_window.geometry("900x600")
        resumo_window.transient(self.janela_parcelas)
        resumo_window.grab_set()

        # Frame principal
        main_frame = ttk.Frame(resumo_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)

        # Título
        ttk.Label(main_frame, text="Resumo das Parcelas Personalizadas",
                 font=('Arial', 14, 'bold')).pack(pady=10)

        # Frame para Treeview com scrollbar
        tree_frame = ttk.Frame(main_frame)
        tree_frame.pack(fill='both', expand=True, pady=10)

        # Treeview para mostrar as parcelas
        columns = ('Parcela', 'Valor', 'Vencimento', 'Referência', 'Etapa', 'Insumo')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=15)

        # Configurar colunas
        tree.column('Parcela', width=80, anchor='center')
        tree.column('Valor', width=120, anchor='e')
        tree.column('Vencimento', width=100, anchor='center')
        tree.column('Referência', width=500, anchor='w')
        tree.column('Etapa', width=150, anchor='w')
        tree.column('Insumo', width=150, anchor='w')

        for col in columns:
            tree.heading(col, text=col)

        # Scrollbar para o Treeview
        scrollbar_tree = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar_tree.set)

        tree.pack(side="left", fill="both", expand=True)
        scrollbar_tree.pack(side="right", fill="y")

        # Inserir dados
        total_resumo = 0
        for i, parcela in enumerate(self.parcelas, 1):
            tree.insert('', 'end', values=(
                f"{i}ª",
                f"R$ {parcela['valor']:,.2f}",
                parcela['dt_vencto'],
                parcela['referencia'],
                parcela.get('etapa_obra', ''),
                parcela.get('insumo', '')
            ))
            total_resumo += parcela['valor']

        # Frame para informações finais
        info_frame = ttk.Frame(main_frame)
        info_frame.pack(fill='x', pady=10)

        # Total
        ttk.Label(info_frame, text=f"TOTAL: R$ {total_resumo:,.2f}",
                 font=('Arial', 12, 'bold')).pack()

        # Botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.pack(fill='x', pady=10)

        ttk.Button(frame_botoes, text="Voltar para Edição",
                  command=resumo_window.destroy).pack(side='left', padx=5)

        ttk.Button(frame_botoes, text="Confirmar e Processar Parcelas",
                  command=lambda: self.finalizar_confirmacao_personalizada(resumo_window)).pack(side='right', padx=5)

    def calcular_data_rel_personalizada(self, dt_vencto):
        """
        Calcula a data do relatório para parcelas personalizadas.
        Sempre retorna dia 5 ou 20, anterior à data de vencimento.
        """
        try:
            logger.debug(f"  calcular_data_rel_personalizada chamado com: {dt_vencto}")

            hoje = datetime.now().date()
            tp_desp = self.tipo_despesa_valor

            logger.debug(f"  Hoje: {hoje}")
            logger.debug(f"  Tipo despesa: {tp_desp}")
            logger.debug(f"  Dia do vencimento: {dt_vencto.day}")

            # Lógica principal baseada na data de vencimento
            if dt_vencto.day == 5:
                # Se vence dia 5, relatório é dia 20 do mês anterior
                data_rel = (dt_vencto - relativedelta(months=1)).replace(day=20)
                logger.debug(f"  Vence dia 5 -> Relatório dia 20 mês anterior: {data_rel}")
            elif dt_vencto.day == 20:
                # Se vence dia 20, relatório é dia 5 do mesmo mês
                data_rel = dt_vencto.replace(day=5)
                logger.debug(f"  Vence dia 20 -> Relatório dia 5 mesmo mês: {data_rel}")
            elif tp_desp == '5':
                # Para tipo 5, usar período mais próximo
                if dt_vencto.day <= 5:
                    data_rel = dt_vencto.replace(day=5)
                    logger.debug(f"  Tipo 5, vence <= 5 -> Relatório dia 5: {data_rel}")
                elif dt_vencto.day <= 20:
                    data_rel = dt_vencto.replace(day=20)
                    logger.debug(f"  Tipo 5, vence <= 20 -> Relatório dia 20: {data_rel}")
                else:
                    proximo_mes = dt_vencto + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
                    logger.debug(f"  Tipo 5, vence > 20 -> Relatório dia 5 próximo mês: {data_rel}")
            else:
                # Para outros tipos (2, 3, 6), usar período anterior ao vencimento
                if dt_vencto.day <= 5:
                    data_rel = (dt_vencto - relativedelta(months=1)).replace(day=20)
                    logger.debug(f"  Outros tipos, vence <= 5 -> Relatório dia 20 mês anterior: {data_rel}")
                elif dt_vencto.day <= 20:
                    data_rel = dt_vencto.replace(day=5)
                    logger.debug(f"  Outros tipos, vence <= 20 -> Relatório dia 5 mesmo mês: {data_rel}")
                else:
                    data_rel = dt_vencto.replace(day=20)
                    logger.debug(f"  Outros tipos, vence > 20 -> Relatório dia 20 mesmo mês: {data_rel}")

            logger.debug(f"  Data relatório antes da verificação: {data_rel}")

            # Garantir que a data do relatório não seja anterior à data atual
            if data_rel < hoje:
                logger.debug(f"  Data relatório {data_rel} é anterior a hoje {hoje}, ajustando...")
                if hoje.day <= 5:
                    data_rel = hoje.replace(day=5)
                    logger.debug(f"  Hoje <= 5 -> Ajustado para dia 5: {data_rel}")
                elif hoje.day <= 20:
                    data_rel = hoje.replace(day=20)
                    logger.debug(f"  Hoje <= 20 -> Ajustado para dia 20: {data_rel}")
                else:
                    proximo_mes = hoje + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
                    logger.debug(f"  Hoje > 20 -> Ajustado para dia 5 próximo mês: {data_rel}")

            logger.debug(f"  Data relatório final: {data_rel}")

            # Retornar como datetime para manter consistência
            return datetime.combine(data_rel, datetime.min.time())

        except Exception as e:
            logger.debug(f"ERRO ao calcular data do relatório personalizada: {str(e)}")
            import traceback
            logger.debug(traceback.format_exc())
            # Em caso de erro, retornar uma data válida baseada em hoje
            hoje = datetime.now().date()
            if hoje.day <= 5:
                data_fallback = hoje.replace(day=5)
            elif hoje.day <= 20:
                data_fallback = hoje.replace(day=20)
            else:
                proximo_mes = hoje + relativedelta(months=1)
                data_fallback = proximo_mes.replace(day=5)

            return datetime.combine(data_fallback, datetime.min.time())

    def finalizar_confirmacao_personalizada(self, resumo_window):
        """Finaliza a confirmação das parcelas personalizadas"""
        resumo_window.destroy()

        # Mostrar mensagem de sucesso
        custom_messagebox("info", "Sucesso",
                        f"{len(self.parcelas)} parcela(s) personalizada(s) gerada(s) com sucesso!")

        # Limpar campos e fechar janela
        self.limpar_campos()

    def adicionar_parcela(self, data_rel, dt_vencto, valor_parcela, referencia_base, i, num_parcelas, eh_primeira_parcela, nf):
        """Método auxiliar para criar uma parcela com todos os dados necessários"""
        parcela = {
            'data_rel': data_rel.strftime('%d/%m/%Y'),
            'dt_vencto': dt_vencto.strftime('%d/%m/%Y'),
            'valor': valor_parcela,
            'referencia': self.gerar_referencia_parcela(referencia_base, i, num_parcelas, eh_primeira_parcela),
            'nf': nf,
            'etapa_obra': self.etapa_obra.get().strip(),
            'insumo': self.insumo.get().strip()
        }
        self.parcelas.append(parcela)

    def gerar_parcelas_prazo_fixo(self, data_base, valores_parcelas, referencia_base, num_parcelas, nf):
        """Gera parcelas com prazo fixo em dias"""
        prazo_dias = int(self.prazo_dias.get())

        for i, valor_parcela in enumerate(valores_parcelas):
            eh_primeira_parcela = (i == 0)

            if eh_primeira_parcela and self.tem_entrada.get():
                dt_vencto = data_base
                data_rel = self.calcular_data_rel(data_base, dt_vencto, True)
            else:
                dt_vencto = data_base + relativedelta(days=prazo_dias * (i + (0 if self.tem_entrada.get() else 1)))
                dt_vencto = self.proximo_dia_util(dt_vencto)
                data_rel = self.calcular_data_rel(data_base, dt_vencto, eh_primeira_parcela)

            self.adicionar_parcela(
                data_rel,
                dt_vencto,
                valor_parcela,
                referencia_base,
                i,
                num_parcelas,
                eh_primeira_parcela,
                nf
            )

    def gerar_parcelas_datas_especificas(self, data_base, valores_parcelas, referencia_base, num_parcelas, nf):
        """Gera parcelas com datas específicas"""
        datas_texto = self.texto_datas.get("1.0", tk.END).strip().split('\n')
        datas_texto = [d.strip() for d in datas_texto if d.strip()]

        num_datas_esperado = num_parcelas
        if len(datas_texto) != num_datas_esperado:
            custom_messagebox("error",
                "Erro",
                f"Para {num_parcelas} {'parcelas após a entrada' if self.tem_entrada.get() else 'parcelas'}, "
                f"é necessário informar {num_datas_esperado} data(s) de vencimento."
            )
            return

        for i, valor_parcela in enumerate(valores_parcelas):
            eh_primeira_parcela = (i == 0)

            try:
                if eh_primeira_parcela and self.tem_entrada.get():
                    dt_vencto = data_base
                    data_rel = self.calcular_data_rel(data_base, dt_vencto, True)
                else:
                    idx_data = i - 1 if self.tem_entrada.get() else i
                    if 0 <= idx_data < len(datas_texto):
                        dt_vencto = datetime.strptime(datas_texto[idx_data], '%d/%m/%Y')
                        dt_vencto = self.proximo_dia_util(dt_vencto)
                        data_rel = self.calcular_data_rel(data_base, dt_vencto, eh_primeira_parcela)
                    else:
                        raise ValueError(f"Índice de data inválido: {idx_data}")

                self.adicionar_parcela(
                    data_rel,
                    dt_vencto,
                    valor_parcela,
                    referencia_base,
                    i,
                    num_parcelas,
                    eh_primeira_parcela,
                    nf
                )

            except ValueError as e:
                custom_messagebox("error", "Erro", f"Erro ao processar data: {str(e)}")
                return
            except IndexError:
                custom_messagebox("error", "Erro", "Número insuficiente de datas fornecidas")
                return

    def gerar_parcelas_cartao(self, data_base, valores_parcelas, referencia_base, num_parcelas, nf):
        """Gera parcelas para pagamento com cartão"""
        dia_vencimento = int(self.dia_vencimento.get())

        for i, valor_parcela in enumerate(valores_parcelas):
            eh_primeira_parcela = (i == 0)

            if eh_primeira_parcela:
                data_atual = data_base + relativedelta(months=1)
            else:
                data_atual = data_base + relativedelta(months=i + 1)

            try:
                dt_vencto = data_atual.replace(day=dia_vencimento)
            except ValueError:
                dt_vencto = data_atual + relativedelta(day=31)

            dt_vencto = self.proximo_dia_util(dt_vencto)

            if eh_primeira_parcela:
                hoje = datetime.now()
                if hoje.day <= 5:
                    data_rel = hoje.replace(day=5)
                elif hoje.day <= 20:
                    data_rel = hoje.replace(day=20)
                else:
                    proximo_mes = hoje + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
            else:
                data_rel = self.calcular_data_rel(data_base, dt_vencto, False)

            self.adicionar_parcela(
                data_rel,
                dt_vencto,
                valor_parcela,
                referencia_base,
                i,
                num_parcelas,
                eh_primeira_parcela,
                nf
            )

    # Métodos de cálculo e utilitários
    def calcular_valores_parcelas(self, valor_original, num_parcelas):
        """Calcula os valores das parcelas considerando entrada se houver"""
        try:
            if self.tem_entrada.get():
                if not self.modalidade_entrada.get():
                    custom_messagebox("error", "Erro", "Selecione a modalidade de entrada!")
                    return None
                valores_parcelas = self.calcular_parcelas_entrada(valor_original, num_parcelas)
            else:
                valores_parcelas = self.calcular_parcelas_ajustadas(valor_original, num_parcelas)

            # Verificar se a soma está correta
            soma_parcelas = sum(valores_parcelas)
            if abs(soma_parcelas - valor_original) > 0.01:
                custom_messagebox("error",
                    "Erro",
                    f"Erro no cálculo das parcelas: soma ({soma_parcelas:.2f}) "
                    f"diferente do valor original ({valor_original:.2f})!"
                )
                return None

            return valores_parcelas
        except Exception as e:
            custom_messagebox("error", "Erro", f"Erro ao calcular valores: {str(e)}")
            return None

    def calcular_parcelas_entrada(self, valor_total, num_parcelas):
        """Calcula valores das parcelas considerando a modalidade de entrada"""
        modalidade = self.modalidade_entrada.get()
        valores_parcelas = []

        # Se tem entrada, o número de parcelas informado é adicional à entrada
        num_parcelas_real = num_parcelas + 1 if self.tem_entrada.get() else num_parcelas

        if modalidade == "Percentual do valor total na primeira parcela":
            try:
                percentual = float(self.valor_entrada.get().replace(',', '.'))
                if not (0 < percentual < 100):
                    raise ValueError("Percentual deve estar entre 0 e 100")

                valor_entrada = (percentual / 100) * valor_total
                valor_restante = valor_total - valor_entrada

                valores_parcelas = [valor_entrada]  # Primeira parcela (entrada)
                # Distribuir o valor restante no número de parcelas informado
                demais_parcelas = self.calcular_parcelas_ajustadas(valor_restante, num_parcelas)
                valores_parcelas.extend(demais_parcelas)

            except ValueError as e:
                raise ValueError(f"Erro no percentual de entrada: {str(e)}")

        elif modalidade == "Primeira parcela igual às demais (arredonda no final)":
            # Dividir o valor total pelo número total de parcelas (incluindo entrada)
            valores_parcelas = self.calcular_parcelas_ajustadas(valor_total, num_parcelas_real)

        elif modalidade == "Valor específico na primeira parcela":
            try:
                valor_entrada = float(self.valor_entrada.get().replace(',', '.'))
                if valor_entrada >= valor_total:
                    raise ValueError("Valor da entrada não pode ser maior ou igual ao valor total")

                valor_restante = valor_total - valor_entrada
                valores_parcelas = [valor_entrada]  # Primeira parcela (entrada)
                # Distribuir o valor restante no número de parcelas informado
                demais_parcelas = self.calcular_parcelas_ajustadas(valor_restante, num_parcelas)
                valores_parcelas.extend(demais_parcelas)

            except ValueError as e:
                raise ValueError(f"Erro no valor da entrada: {str(e)}")

        return valores_parcelas

    def calcular_parcelas_ajustadas(self, valor_total, num_parcelas):
        """Calcula valores das parcelas garantindo que a soma seja igual ao valor total"""
        valor_parcela_base = valor_total / num_parcelas
        valor_parcela_round = round(valor_parcela_base, 2)

        # Calcular diferença total devido aos arredondamentos
        diferenca = valor_total - (valor_parcela_round * num_parcelas)

        # Distribuir a diferença na última parcela
        parcelas = [valor_parcela_round] * (num_parcelas - 1)
        ultima_parcela = valor_parcela_round + round(diferenca, 2)
        parcelas.append(ultima_parcela)

        return parcelas

    def calcular_data_rel(self, data_base, dt_vencto, eh_primeira_parcela):
        """
        Calcula a data do relatório com base na data de vencimento e tipo de despesa.
        Considera a data atual para não retroagir em períodos fechados.
        """
        try:
            hoje = datetime.now()

            # Se for entrada, calcula a partir da data atual
            if eh_primeira_parcela and self.tem_entrada.get():
                if hoje.day <= 5:
                    data_rel = hoje.replace(day=5)
                elif hoje.day <= 20:
                    data_rel = hoje.replace(day=20)
                else:
                    proximo_mes = hoje + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
                return data_rel

            # Para as demais parcelas, manter a lógica existente
            tp_desp = self.tipo_despesa_valor

            if dt_vencto.day == 5:
                # Se vence dia 5, relatório é dia 20 do mês anterior
                data_rel = (dt_vencto - relativedelta(months=1)).replace(day=20)
            elif dt_vencto.day == 20:
                # Se vence dia 20, relatório é dia 5 do mesmo mês
                data_rel = dt_vencto.replace(day=5)
            elif tp_desp == '5':
                if dt_vencto.day <= 5:
                    data_rel = dt_vencto.replace(day=5)
                elif dt_vencto.day <= 20:
                    data_rel = dt_vencto.replace(day=20)
                else:
                    proximo_mes = dt_vencto + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)
            else:
                if dt_vencto.day <= 5:
                    data_rel = (dt_vencto - relativedelta(months=1)).replace(day=20)
                elif dt_vencto.day <= 20:
                    data_rel = dt_vencto.replace(day=5)
                else:
                    data_rel = dt_vencto.replace(day=20)

            # Garantir que a data do relatório não seja anterior à data atual
            if data_rel < hoje:
                if hoje.day <= 5:
                    data_rel = hoje.replace(day=5)
                elif hoje.day <= 20:
                    data_rel = hoje.replace(day=20)
                else:
                    proximo_mes = hoje + relativedelta(months=1)
                    data_rel = proximo_mes.replace(day=5)

            return data_rel
        except Exception as e:
            logger.debug(f"Erro ao calcular data do relatório: {str(e)}")
            return dt_vencto

    def configurar_calendario(self, dateentry):
        """Configura o comportamento do calendário"""
        def on_calendar_click(event):
            # Permite cliques no calendário
            return True

        def on_calendar_select(event):
            dateentry._top_cal.withdraw()  # Fecha o calendário
            self.janela_parcelas.after(100, lambda: self.janela_parcelas.focus_set())  # Retorna foco

        def on_calendar_focus(event):
            # Mantém o foco quando o calendário está aberto
            if dateentry._top_cal:
                dateentry._top_cal.focus_set()
            return True

        # Configurar bindings
        dateentry.bind('<<DateEntrySelected>>', on_calendar_select)
        dateentry.bind('<FocusIn>', on_calendar_focus)

        if hasattr(dateentry, '_top_cal'):
            cal = dateentry._top_cal
            cal.bind('<Button-1>', on_calendar_click)
            for w in cal.winfo_children():
                w.bind('<Button-1>', on_calendar_click)

    def proximo_dia_util(self, data):
        """
        Ajusta a data para o próximo dia útil se cair em fim de semana ou feriado
        """
        # Lista de feriados nacionais fixos
        feriados_fixos = [
            (1, 1),   # Ano Novo
            (21, 4),  # Tiradentes
            (1, 5),   # Dia do Trabalho
            (7, 9),   # Independência
            (12, 10), # Nossa Senhora
            (2, 11),  # Finados
            (15, 11), # Proclamação da República
            (25, 12), # Natal
        ]

        while True:
            # Verifica se é fim de semana
            if data.weekday() >= 5:  # 5 = Sábado, 6 = Domingo
                data = data + relativedelta(days=1)
                continue

            # Verifica se é feriado fixo
            if (data.day, data.month) in feriados_fixos:
                data = data + relativedelta(days=1)
                continue

            # Se não é fim de semana nem feriado, é dia útil
            break

        return data

    def gerar_referencia_parcela(self, referencia_base, indice, num_parcelas, eh_primeira_parcela):
        """Gera a referência apropriada para a parcela"""
        if eh_primeira_parcela and self.tem_entrada.get():
            return f"{referencia_base} - ENTRADA"
        else:
            if self.tem_entrada.get():
                # Para as parcelas após a entrada
                return f"{referencia_base} - PARC. {indice}/{num_parcelas}"
            else:
                # Para parcelamento sem entrada
                return f"{referencia_base} - PARC. {indice + 1}/{num_parcelas}"

    # Métodos de limpeza e finalização
    def limpar_campos(self):
        """Limpa todos os campos após sucesso"""
        # Limpar referências de widgets existentes
        self.frame_modalidade = None
        self.frame_valor_entrada = None
        self.lbl_entrada = None
        self.valor_entrada = None
        self.modalidade_entrada = None

        # Limpar campos de parcelas personalizadas
        self.parcelas_personalizadas = []
        self.frame_parcelas_personalizadas = None
        self.valor_total_personalizado = 0.0
        self.canvas_parcelas = None
        self.scrollbar_parcelas = None
        self.campos_parcelas = []

        # Resetar checkbox
        if self._var_tem_entrada:
            self._var_tem_entrada.set(False)

        # Fechar janela
        if self.janela_parcelas:
            self.janela_parcelas.destroy()
            self.janela_parcelas = None

    def cancelar_parcelamento(self):
        """Cancela o parcelamento e limpa todos os campos"""
        self.parcelas = []

        # Limpar referências de widgets existentes
        self.frame_modalidade = None
        self.frame_valor_entrada = None
        self.lbl_entrada = None
        self.valor_entrada = None
        self.modalidade_entrada = None

        # Limpar campos de parcelas personalizadas
        self.parcelas_personalizadas = []
        self.frame_parcelas_personalizadas = None
        self.valor_total_personalizado = 0.0
        self.canvas_parcelas = None
        self.scrollbar_parcelas = None
        self.campos_parcelas = []

        # Resetar variável de entrada
        if self._var_tem_entrada:
            self._var_tem_entrada.set(False)

        if self.janela_parcelas:
            self.janela_parcelas.destroy()
            self.janela_parcelas = None

    # NOTA DE EXTRAÇÃO: método suspeito de ser resíduo colado de outra
    # classe — GestorParcelas não define self.root em nenhum lugar (usa
    # self.parent). Mantido por fidelidade ao original. Ver docstring do
    # módulo, pendência (1).
    def run(self):
        """Inicia a execução do sistema"""
        self.root.mainloop()
