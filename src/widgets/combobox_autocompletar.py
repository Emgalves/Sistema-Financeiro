"""
Widget de Combobox com autocompletar e adição dinâmica de itens novos
(usado para campos como Etapa da Obra e Insumo).

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].

HISTÓRICO DE REVISÕES (mantido porque houve idas e vindas - registrar aqui
evita repetir os mesmos erros de novo):

1ª revisão: tentou corrigir "seleciono um item e aparece outro" adiando a
checagem de FocusOut. Insuficiente - o atraso fixo não cobre todos os
tempos de interação do usuário.

2ª revisão: parou de filtrar self['values'] durante a digitação (lista
suspensa sempre completa) e passou a só autocompletar em correspondência
ÚNICA de prefixo. Eliminava de fato qualquer risco de seleção trocada, mas
foi uma mudança mais agressiva do que o necessário: o código original
(recuperado do GitHub pelo usuário, comprovadamente estável em uso real)
já filtrava self['values'] ao vivo e já só autocompletava em
correspondência única - ou seja, esse padrão nunca foi, na prática, a
causa do bug relatado. Um efeito colateral dessa 2ª revisão apareceu:
como um prefixo ambíguo (ex.: "INSTALAÇ", que bate com 3 itens) nunca se
tornava um valor válido nem era filtrado, qualquer oscilação de foco
(inclusive ao interagir com o próprio dropdown) reabria a pergunta
"Adicionar Novo Item" repetidamente, mesmo depois do usuário responder
"Não".

3ª revisão: voltou à base do código original/GitHub (filtragem ao vivo de
self['values'], autocomplete só em correspondência única), mantendo as
correções de tecla morta e memória de recusa. Insuficiente: testes
mostraram o bug de seleção trocada voltando a acontecer mesmo com esse
código "original comprovado", e revelaram a causa mais profunda.

4ª revisão (ESTA) - causa raiz confirmada: a caixa de diálogo "Adicionar
Novo Item" (messagebox.askyesno) NÃO está bloqueando de fato a interação
com o campo no ambiente do usuário - evidência direta: um teste mostrou o
campo com "DEMOLIÇÃO" enquanto a caixa ainda perguntava sobre "INST", ou
seja, o usuário conseguiu clicar no dropdown e disparar uma nova seleção
COM a caixa de diálogo anterior ainda aberta por baixo. Presumir que "uma
vez a caixa aberta, nada mais acontece até a resposta" era falso, e isso
invalidava várias das proteções anteriores.

Duas mudanças estruturais para fechar isso de vez:
  1. self['values'] volta a ser estático (nunca filtrado durante a
     digitação, como na 2ª revisão) - com evidência agora dupla de que a
     filtragem ao vivo é, sim, uma fonte real do bug de seleção trocada
     (não uma falsa hipótese).
  2. O próprio campo é desabilitado enquanto a caixa de diálogo está
     aberta, e há uma trava (self._dialogo_aberto) impedindo uma segunda
     pergunta simultânea - isso garante bloqueio de interação mesmo que o
     grab nativo do Tk não esteja funcionando como esperado no ambiente
     do usuário.

5ª revisão: tentativa de melhoria cosmética (pedida pelo usuário) para
rolar/realçar a lista suspensa até a região do item correspondente ao
prefixo digitado, via comando interno ttk::combobox::PopdownWindow, sem
tocar em self['values']. REVERTIDA: testado sem gerar erro, mas também
sem efeito visível algum - o ttk::combobox provavelmente resincroniza o
destaque da lista com o texto atual do campo no momento em que o dropdown
é de fato aberto, sobrescrevendo a seleção manual feita antes disso.
Não valia o risco de mexer mais fundo em comportamento interno não
documentado do Tk por um ganho puramente cosmético - removida.

6ª revisão: ajuste de fluxo pedido pelo usuário. Ao responder "Não" na
pergunta de "Adicionar Novo Item", o campo agora é limpo automaticamente
(em vez de manter o texto inválido parado ali) - a recusa não deve travar
o fluxo, só descartar o texto e liberar o campo para seguir em frente.
Texto da pergunta também ajustado para ficar mais direto.

7ª revisão - causa raiz confirmada: abrir a lista suspensa (clique na
seta, sem escolher nada ainda) já dispara <FocusOut> por si só, ANTES de
qualquer clique em item. O adiamento por after_idle (6ª revisão e
anteriores) presumia que só valia a pena esperar quando havia um clique
"em andamento" para a fila de eventos escoar - não havia nada em
andamento nesse caso, só a abertura em si, então a checagem ainda
disparava cedo demais (relatado pelo usuário: "digito 1-2 letras, clico
na seta, e a mensagem já aparece"). Corrigido consultando ativamente se o
dropdown está aberto (winfo ismapped no popdown, só leitura) antes de
perguntar; se estiver, a checagem se reagenda e tenta de novo mais
adiante, só perguntando de fato quando a lista realmente fechar (por
seleção ou por o usuário clicar fora).
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
        self.bind('<FocusOut>', self.on_focus_out)
        self.bind('<Return>', self.on_enter)
        # Disparado quando o usuário escolhe um item existente na lista
        # suspensa (clique do mouse ou teclado). Reforço de baixo risco:
        # trata a escolha como definitiva imediatamente, sem depender só
        # do FocusOut/Return para resolver o estado do campo.
        self.bind('<<ComboboxSelected>>', self.on_select)
        # Ao reganhar o foco, esquece qualquer recusa anterior - se o
        # usuário voltar ao campo depois, tem o direito de ser perguntado
        # de novo caso ainda queira digitar o mesmo texto.
        self.bind('<FocusIn>', self.on_focus_in)

        # Variável para controlar se está filtrando
        self.filtrando = False

        # Último texto (maiúsculo) já processado por on_keyrelease. Serve
        # para detectar teclas que disparam KeyRelease sem mudar o texto -
        # o caso mais comum é uma tecla morta de acentuação (ex.: '~'
        # isolado, aguardando o próximo caractere para formar 'ã'/'õ').
        self._ultimo_texto_processado = ''

        # Texto (maiúsculo) para o qual o usuário já respondeu "Não" na
        # pergunta de "Adicionar Novo Item". Evita repetir a pergunta
        # sobre o mesmo texto a cada oscilação de foco, até que o texto
        # realmente mude.
        self._texto_recusado = None

        # Job do after_idle() usado para adiar a pergunta de "Adicionar
        # Novo Item" (ver on_focus_out) até que a fila de eventos do
        # Tkinter esteja livre - ou seja, até que um eventual clique em
        # andamento no próprio dropdown já tenha terminado de ser
        # processado.
        self._verificacao_job = None

        # Trava contra duas perguntas de "Adicionar Novo Item" abertas ao
        # mesmo tempo. Necessária porque a caixa de diálogo, no ambiente
        # testado, não bloqueia de forma confiável a interação com este
        # campo (ver docstring do módulo, 4ª revisão) - por isso o próprio
        # campo também é desabilitado enquanto ela está na tela (ver
        # confirmar_adicao_item).
        self._dialogo_aberto = False

    # Teclas que não devem disparar filtragem/autocomplete
    _TECLAS_NAVEGACAO = {
        'Up', 'Down', 'Left', 'Right', 'Return', 'Tab', 'Escape',
        'Shift_L', 'Shift_R', 'Control_L', 'Control_R', 'Alt_L', 'Alt_R'
    }

    def on_keyrelease(self, event):
        """Evento chamado quando uma tecla é liberada — autocompleta o texto
        digitado quando o prefixo já digitado identifica um único item, e
        posiciona a lista suspensa próxima da região correspondente (sem
        nunca alterar self['values'] - ver docstring do módulo).
        """
        if event.keysym in self._TECLAS_NAVEGACAO:
            return

        # Teclas mortas (acentos compostos) disparam KeyRelease sem
        # inserir nenhum caractere - a composição só se resolve na tecla
        # seguinte. Processá-las aqui destruiria a seleção da sugestão de
        # autocomplete no meio da digitação de uma palavra acentuada.
        if event.keysym.startswith('dead_'):
            return

        texto_digitado = self.get().upper()

        # Reforço do filtro acima: cobre qualquer tecla que dispare
        # KeyRelease sem mudar o texto de fato.
        if texto_digitado == self._ultimo_texto_processado:
            return
        self._ultimo_texto_processado = texto_digitado

        # O texto mudou de verdade: uma eventual recusa anterior não vale
        # mais para este novo texto.
        self._texto_recusado = None

        if not texto_digitado:
            self.filtrando = False
            return

        self.filtrando = True

        # Busca por prefixo. Autocompleta somente quando ele já identifica
        # um único item - evita "atropelar" o usuário enquanto ele ainda
        # está diferenciando entre itens com prefixo em comum (inclusive
        # no meio de acentos compostos).
        correspondencias = [v for v in self.valores_originais if v.upper().startswith(texto_digitado)]

        if len(correspondencias) == 1:
            valor_completo = correspondencias[0]
            self.delete(0, tk.END)
            self.insert(0, valor_completo)
            self.selection_range(len(texto_digitado), len(valor_completo))
            self._ultimo_texto_processado = valor_completo.upper()

    def on_select(self, event):
        """Evento <<ComboboxSelected>> - usuário escolheu um item existente
        na lista suspensa (mouse ou teclado). Escolha sempre válida e final.
        """
        if self._verificacao_job is not None:
            self.after_cancel(self._verificacao_job)
            self._verificacao_job = None

        self.filtrando = False
        self._texto_recusado = None
        self._ultimo_texto_processado = self.get().upper()
        self.icursor(tk.END)
        self.selection_clear()

    def on_focus_in(self, event):
        """Ao reganhar o foco, esquece uma eventual recusa anterior e
        cancela uma verificação de texto livre que ainda estivesse
        pendente - se o foco voltou, a saída anterior não foi definitiva.
        """
        self._texto_recusado = None
        if self._verificacao_job is not None:
            self.after_cancel(self._verificacao_job)
            self._verificacao_job = None

    def on_focus_out(self, event):
        """Evento quando perde o foco - adia a pergunta de "Adicionar Novo
        Item" (ver _agendar_verificacao_texto_livre) para não abrir a
        caixa de diálogo no meio de um clique em andamento no próprio
        dropdown.
        """
        self.filtrando = False
        self._agendar_verificacao_texto_livre()

    def on_enter(self, event):
        """Evento Enter - confirmar seleção ou adicionar novo item.

        Diferente do FocusOut, o Enter é uma ação deliberada do usuário
        (não pode disparar no meio de um clique no dropdown - nesse caso
        quem processa a tecla é a lista suspensa, que gera
        <<ComboboxSelected>>, não este evento). Por isso a checagem aqui
        continua síncrona.
        """
        if self._verificacao_job is not None:
            self.after_cancel(self._verificacao_job)
            self._verificacao_job = None

        self._verificar_texto_livre()
        self.filtrando = False

    def _agendar_verificacao_texto_livre(self):
        """Agenda a checagem de texto livre para rodar quando a lista
        suspensa não estiver mais aberta.

        Só adiar até a fila de eventos esvaziar (after_idle sozinho) não
        basta: abrir o dropdown com a seta já dispara FocusOut por si só,
        antes de qualquer clique em item acontecer - não há nada "em
        andamento" para a fila esvaziar. Por isso, se o dropdown ainda
        estiver aberto quando a checagem for tentada, ela se reagenda e
        tenta de novo mais adiante, em vez de perguntar imediatamente.
        Se uma seleção real acontecer nesse meio tempo, on_select cancela
        este job antes que ele chegue a perguntar qualquer coisa.
        """
        if self._verificacao_job is not None:
            self.after_cancel(self._verificacao_job)
        self._verificacao_job = self.after_idle(self._executar_verificacao_adiada)

    def _executar_verificacao_adiada(self):
        if self._dropdown_esta_aberto():
            # Ainda aberto - o FocusOut foi causado por abrir a própria
            # lista, não por o usuário ter saído do campo de verdade.
            # Tenta de novo em breve, sem perguntar nada agora.
            self._verificacao_job = self.after(150, self._executar_verificacao_adiada)
            return

        self._verificacao_job = None
        self._verificar_texto_livre()

    def _dropdown_esta_aberto(self):
        """Consulta (somente leitura) se a lista suspensa deste combobox
        está aberta na tela no momento. Usa o comando interno
        ttk::combobox::PopdownWindow só para localizar o widget do
        dropdown - diferente da tentativa anterior de "pular" para um
        item (que chegou a escrever no estado do dropdown e não
        funcionou de forma confiável), aqui é só uma leitura de estado
        (winfo ismapped), bem mais simples e robusta entre versões do Tk.
        Qualquer falha é tratada como "não está aberto", para nunca travar
        a checagem indefinidamente.
        """
        try:
            popdown = self.tk.call('ttk::combobox::PopdownWindow', self)
            return bool(int(self.tk.call('winfo', 'ismapped', popdown)))
        except Exception:
            return False

    def _verificar_texto_livre(self):
        """Verifica se o texto digitado é um item novo e, se for, pergunta
        se o usuário quer adicioná-lo - mas só uma vez por texto (ver
        self._texto_recusado), e nunca se já houver uma pergunta aberta
        (ver self._dialogo_aberto).
        """
        if self._dialogo_aberto:
            return

        texto_digitado = self.get().strip().upper()

        if not texto_digitado:
            return

        if texto_digitado in [v.upper() for v in self.valores_originais]:
            return

        if texto_digitado == self._texto_recusado:
            return

        self.confirmar_adicao_item(texto_digitado)

    def confirmar_adicao_item(self, novo_item):
        """Confirma se o usuário quer adicionar um novo item.

        Desabilita o próprio campo enquanto a pergunta está na tela. Isso
        não é cosmético: em testes, a caixa de diálogo não bloqueou de
        forma confiável a interação com o combobox por baixo dela (dava
        para clicar no dropdown e disparar uma nova seleção com a
        pergunta ainda aberta) - o que produzia exatamente o tipo de
        estado inconsistente relatado (campo com um valor, pergunta
        referindo-se a outro). Desabilitar o campo neste intervalo garante
        que isso não aconteça, independente do grab nativo do Tk estar ou
        não funcionando como esperado no ambiente do usuário.
        """
        if self._dialogo_aberto:
            return
        self._dialogo_aberto = True

        estado_anterior = str(self.cget('state'))
        self.configure(state='disabled')

        try:
            tipo_item = "Etapa da Obra" if self.config_key == "etapas_obra" else "Insumo"

            resposta = messagebox.askyesno(
                "Adicionar Novo Item",
                f"Essa {tipo_item} '{novo_item}' não existe na lista.\n\n"
                f"Deseja incluir?",
                parent=self.winfo_toplevel()
            )
        finally:
            self.configure(state=estado_anterior)
            self._dialogo_aberto = False

        if resposta:
            self.adicionar_novo_item(novo_item)
        else:
            # Limpa o campo e segue o processo, em vez de deixar o texto
            # inválido parado ali - pedido explícito do usuário: "Não"
            # deve desbloquear o fluxo, não travar em cima do mesmo texto.
            self.delete(0, tk.END)
            self._ultimo_texto_processado = ''
            self._texto_recusado = None

    def adicionar_novo_item(self, novo_item):
        """Adiciona um novo item ao arquivo de configurações"""
        try:
            # Import local mantido de propósito (evita import circular na
            # inicialização do módulo — GerenciadorConfiguracoes só é
            # necessário quando o usuário efetivamente adiciona um item novo).
            from src.configuracoes_sistema import GerenciadorConfiguracoes
            config = GerenciadorConfiguracoes.carregar_configuracoes()

            if not config:
                messagebox.showerror("Erro", "Não foi possível carregar as configurações.")
                return

            if self.config_key not in config:
                config[self.config_key] = {'lista': [], 'historico_alteracoes': []}

            if novo_item not in config[self.config_key]['lista']:
                config[self.config_key]['lista'].append(novo_item)
                config[self.config_key]['lista'].sort()

                config_path = GerenciadorConfiguracoes.CONFIG_PATH
                with open(config_path, 'w', encoding='utf-8') as f:
                    json.dump(config, f, indent=4, ensure_ascii=False)

                GerenciadorConfiguracoes._atualizar_cache(config)

                self.valores_originais = config[self.config_key]['lista']
                self['values'] = self.valores_originais

                self.set(novo_item)
                self._texto_recusado = None
                self._ultimo_texto_processado = novo_item.upper()

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
