"""
Widget de Combobox com autocompletar e adição dinâmica de itens novos
(usado para campos como Etapa da Obra e Insumo).

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].
Nenhuma alteração de lógica foi feita nesta extração — apenas mudança
de localização e ajuste de imports.
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

    def on_keyrelease(self, event):
        """Evento chamado quando uma tecla é liberada"""
        if event.keysym in ['Up', 'Down', 'Left', 'Right', 'Return', 'Tab']:
            return

        texto_digitado = self.get().upper()

        if not texto_digitado:
            # Se não há texto, mostrar todos os valores
            self['values'] = self.valores_originais
            self.filtrando = False
            return

        # Filtrar valores que começam com o texto digitado
        valores_filtrados = [v for v in self.valores_originais if v.upper().startswith(texto_digitado)]

        if valores_filtrados:
            # Atualizar lista com valores filtrados
            self['values'] = valores_filtrados
            self.filtrando = True

            # Autocompletar com o primeiro resultado
            if len(valores_filtrados) == 1:
                valor_completo = valores_filtrados[0]
                self.delete(0, tk.END)
                self.insert(0, valor_completo)
                self.selection_range(len(texto_digitado), len(valor_completo))
        else:
            # Nenhuma correspondência encontrada
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
