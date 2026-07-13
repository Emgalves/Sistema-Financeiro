"""
Widget de Combobox com autocompletar e adição dinâmica de itens novos
(usado para campos como Etapa da Obra e Insumo).

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].

Atualizado em [DATA_DA_ATUALIZACAO] para adotar um comportamento híbrido
de filtragem, combinando o autocomplete original desta classe com o
padrão usado em ImportadorRH.solicitar_etapa_obra:
  - Busca por prefixo tem prioridade e autocompleta com o primeiro
    resultado da lista filtrada (não exige mais resultado único).
  - Quando não há match por prefixo, cai para busca por substring em
    qualquer parte do texto, sem forçar preenchimento automático.
  - Backspace/Delete apenas filtram a lista, sem re-forçar o
    autocomplete — evita o campo "grudar" no mesmo valor ao apagar.
"""
import json
import tkinter as tk
from tkinter import ttk, messagebox


class ComboboxAutocompletar(ttk.Combobox):
    """Combobox personalizada com funcionalidade de autocompletar e adição dinâmica"""

    def __init__(self, parent, values=None, config_key=None, config_manager=None, **kwargs):
        """
        Inicializa o Combobox com autocompletar

        Args:
            parent: Widget pai
            values: Lista inicial de valores
            config_key: Chave no arquivo de configuração ('etapas_obra' ou 'insumos')
            config_manager: Instância do GerenciadorConfiguracoes
            **kwargs: Argumentos adicionais para ttk.Combobox
        """
        super().__init__(parent, **kwargs)

        self.config_key = config_key
        self.config_manager = config_manager
        self.valores_originais = values or []

        # Configurar valores iniciais
        self['values'] = self.valores_originais

        # Bind de eventos
        self.bind('<KeyRelease>', self.on_keyrelease)
        self.bind('<Button-1>', self.on_click)
        self.bind('<FocusOut>', self.on_focus_out)
        self.bind('<Return>', self.on_enter)

        # Variável para controlar se está filtrando
        self.filtrando = False

    # Teclas que não devem disparar filtragem/autocomplete
    _TECLAS_NAVEGACAO = {
        'Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape',
        'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'
    }
    # Teclas de edição que filtram a lista, mas não devem forçar
    # o preenchimento automático (senão o usuário nunca consegue apagar)
    _TECLAS_APAGANDO = {'BackSpace', 'Delete'}

    def on_keyrelease(self, event):
        """Evento chamado quando uma tecla é liberada — filtra e autocompleta.

        Comportamento híbrido:
        1) Tenta por PREFIXO primeiro. Se houver resultado(s), mostra no
           dropdown e autocompleta com o primeiro deles (exceto ao apagar).
        2) Sem match por prefixo, cai para busca por SUBSTRING em qualquer
           parte do texto — não força preenchimento, só filtra a lista,
           já que o termo digitado pode não ser o início do item.
        """
        if event.keysym in self._TECLAS_NAVEGACAO:
            return

        apagando = event.keysym in self._TECLAS_APAGANDO
        texto_digitado = self.get().upper()

        if not texto_digitado:
            # Se não há texto, mostrar todos os valores
            self['values'] = self.valores_originais
            self.filtrando = False
            return

        # 1) Busca por prefixo (comportamento principal)
        valores_por_prefixo = [v for v in self.valores_originais if v.upper().startswith(texto_digitado)]

        if valores_por_prefixo:
            self['values'] = valores_por_prefixo
            self.filtrando = True

            # Autocompletar com o primeiro resultado — mas só ao digitar
            # para frente. Ao apagar (Backspace/Delete), deixa o texto
            # do jeito que o usuário digitou, sem re-preencher.
            if not apagando:
                valor_completo = valores_por_prefixo[0]
                self.delete(0, tk.END)
                self.insert(0, valor_completo)
                self.selection_range(len(texto_digitado), len(valor_completo))
            return

        # 2) Fallback: busca por substring em qualquer parte do texto
        valores_por_substring = [v for v in self.valores_originais if texto_digitado in v.upper()]

        if valores_por_substring:
            # Apenas restringe o dropdown. Não preenche automaticamente,
            # pois o termo digitado não corresponde ao início do item.
            self['values'] = valores_por_substring
            self.filtrando = True
        else:
            # Nenhuma correspondência: dropdown fica vazio, texto do
            # usuário permanece livre (permite o fluxo de "adicionar novo item").
            self['values'] = []
            self.filtrando = True

    def on_click(self, event):
        """Evento de clique - mostrar todos os valores"""
        if not self.filtrando:
            self['values'] = self.valores_originais

    def on_focus_out(self, event):
        """Evento quando perde o foco - verificar se precisa adicionar novo item"""
        texto_digitado = self.get().strip().upper()

        if texto_digitado and texto_digitado not in [v.upper() for v in self.valores_originais]:
            # Perguntar se quer adicionar o novo item
            self.confirmar_adicao_item(texto_digitado)

        # Restaurar lista completa
        self['values'] = self.valores_originais
        self.filtrando = False

    def on_enter(self, event):
        """Evento Enter - confirmar seleção ou adicionar novo item"""
        texto_digitado = self.get().strip().upper()

        if texto_digitado and texto_digitado not in [v.upper() for v in self.valores_originais]:
            self.confirmar_adicao_item(texto_digitado)

        # Restaurar lista completa
        self['values'] = self.valores_originais
        self.filtrando = False

    def confirmar_adicao_item(self, novo_item):
        """Confirma se o usuário quer adicionar um novo item"""
        tipo_item = "Etapa da Obra" if self.config_key == "etapas_obra" else "Insumo"

        resposta = messagebox.askyesno(
            "Adicionar Novo Item",
            f"O {tipo_item.lower()} '{novo_item}' não existe na lista.\n\n"
            f"Deseja adicioná-lo aos parâmetros do sistema?",
            parent=self.winfo_toplevel()
        )

        if resposta:
            self.adicionar_novo_item(novo_item)

    def adicionar_novo_item(self, novo_item):
        """Adiciona um novo item ao arquivo de configurações"""
        try:
            # Carregar configurações atuais
            # Import local mantido de propósito (evita import circular na
            # inicialização do módulo — GerenciadorConfiguracoes só é
            # necessário quando o usuário efetivamente adiciona um item novo).
            from src.configuracoes_sistema import GerenciadorConfiguracoes
            config = GerenciadorConfiguracoes.carregar_configuracoes()

            if not config:
                messagebox.showerror("Erro", "Não foi possível carregar as configurações.")
                return

            # Verificar se a seção existe
            if self.config_key not in config:
                config[self.config_key] = {'lista': [], 'historico_alteracoes': []}

            # Adicionar o novo item
            if novo_item not in config[self.config_key]['lista']:
                config[self.config_key]['lista'].append(novo_item)
                config[self.config_key]['lista'].sort()

                # Salvar no arquivo
                config_path = GerenciadorConfiguracoes.CONFIG_PATH
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)

                # Atualizar cache
                GerenciadorConfiguracoes._atualizar_cache(config)

                # Atualizar valores do combobox
                self.valores_originais = config[self.config_key]['lista']
                self['values'] = self.valores_originais

                # Definir o novo valor no combobox
                self.set(novo_item)

                tipo_item = "Etapa da Obra" if self.config_key == "etapas_obra" else "Insumo"
                messagebox.showinfo(
                    "Sucesso",
                    f"{tipo_item} '{novo_item}' adicionado com sucesso!",
                    parent=self.winfo_toplevel()
                )

        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"Erro ao adicionar item: {str(e)}",
                parent=self.winfo_toplevel()
            )

    def atualizar_valores(self, novos_valores):
        """Atualiza a lista de valores do combobox"""
        self.valores_originais = novos_valores
        self['values'] = self.valores_originais
