# src/comprovantes_beneficios/interface.py
"""
Interface (Tkinter) do fluxo descrito na especificação, seção 5.

Reescrita nesta versão para:
  - usar `controle_registros` (aba na planilha do cliente) no lugar do
    SQLite descontinuado — sem caminho de banco separado, o controle
    vive dentro do próprio arquivo do cliente já selecionado.
  - usar `gerador_recibo` baseado em reportlab — sem caminhos de script
    Node/LibreOffice a configurar.
  - registrar as emissões em **lote** (uma única abertura/gravação da
    planilha do cliente por clique em "Emitir selecionados"), reduzindo
    tempo e risco de conflito de escrita.
"""

import os
import platform
import re
import socket
import sys
import subprocess
import unicodedata
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import date
from pathlib import Path

from .dados_candidatos import (
    obter_competencias_disponiveis, obter_candidatos, obter_dados_pagador,
    BENEFICIO_TRANSPORTE, BENEFICIO_CAFE, BENEFICIO_CESTA_BASICA, BENEFICIO_CESTA_NATAL,
)
from .gerador_recibo import gerar_recibo_pdf, ErroGeracaoRecibo
from .controle_registros import (
    ja_emitido, listar_emitidos, registrar_lote, NovoRegistro,
)
from .normalizacao import formatar_cpf, formatar_valor_monetario


BENEFICIOS_LABELS = {
    BENEFICIO_CESTA_BASICA: 'Cesta Básica',
    BENEFICIO_CESTA_NATAL: 'Cesta de Natal',
    BENEFICIO_TRANSPORTE: 'Transporte',
    BENEFICIO_CAFE: 'Café',
}
_LABEL_TO_BENEFICIO = {v: k for k, v in BENEFICIOS_LABELS.items()}


class InterfaceComprovantesBeneficios:

    def __init__(
        self,
        parent,
        pasta_clientes: Path,
        arquivo_clientes: Path,
        arquivo_fornecedores: Path,
        pasta_saida_base: Path,
        usuario: str = None,
    ):
        self.root = parent
        self.pasta_clientes = Path(pasta_clientes)
        self.arquivo_clientes = Path(arquivo_clientes)
        self.arquivo_fornecedores = Path(arquivo_fornecedores)
        self.pasta_saida_base = Path(pasta_saida_base)
        self.usuario = usuario or os.environ.get('USERNAME') or os.environ.get('USER') or 'desconhecido'
        self.maquina = platform.node() or socket.gethostname()

        hoje = date.today()
        self.competencia_atual = f"{hoje.year:04d}-{hoje.month:02d}"

        self.candidatos_atuais = []
        self.avisos_atuais = []
        self.arquivo_cliente_selecionado = None
        self.pagador_atual = None
        self.beneficio_atual = None
        self.competencia_selecionada = None
        self._mapa_clientes = {}

        self._montar_interface()

    # ------------------------------------------------------------------
    # Montagem
    # ------------------------------------------------------------------

    def _montar_interface(self):
        try:
            from src.config.window_config import configurar_janela
            configurar_janela(self.root, "Comprovantes de Benefícios")
        except ImportError:
            self.root.title("Comprovantes de Benefícios")
            self.root.geometry("900x650")

        frame_topo = ttk.Frame(self.root, padding=10)
        frame_topo.pack(fill='x')

        ttk.Label(frame_topo, text="Cliente:").grid(row=0, column=0, sticky='w')
        self.combo_cliente = ttk.Combobox(frame_topo, state='readonly', width=45)
        self.combo_cliente.grid(row=0, column=1, columnspan=3, padx=5, pady=(0, 8), sticky='w')
        self.combo_cliente.bind('<<ComboboxSelected>>', self._ao_selecionar_cliente)

        ttk.Label(frame_topo, text="Competência:").grid(row=1, column=0, sticky='w')
        self.combo_competencia = ttk.Combobox(frame_topo, state='readonly', width=10)
        self.combo_competencia.grid(row=1, column=1, padx=5, sticky='w')

        ttk.Label(frame_topo, text="Benefício:").grid(row=1, column=2, sticky='w', padx=(15, 0))
        self.combo_beneficio = ttk.Combobox(
            frame_topo, state='readonly', width=16, values=list(BENEFICIOS_LABELS.values()),
        )
        self.combo_beneficio.grid(row=1, column=3, padx=5, sticky='w')

        ttk.Button(
            frame_topo, text="Buscar candidatos", command=self._buscar_candidatos,
        ).grid(row=1, column=4, padx=(15, 0), sticky='w')

        ttk.Button(
            frame_topo, text="Fechar", command=self.root.destroy,
        ).grid(row=1, column=5, padx=(15, 0), sticky='w')

        frame_avisos = ttk.LabelFrame(self.root, text="Avisos", padding=5)
        frame_avisos.pack(fill='x', padx=10, pady=(8, 0))
        self.texto_avisos = tk.Text(frame_avisos, height=4, wrap='word', state='disabled')
        self.texto_avisos.pack(fill='x')

        frame_tabela = ttk.Frame(self.root, padding=10)
        frame_tabela.pack(fill='both', expand=True)

        colunas = ('nome', 'cpf', 'valor', 'vencimento', 'status')
        self.tabela = ttk.Treeview(
            frame_tabela, columns=colunas, show='headings', selectmode='extended',
        )
        self.tabela.heading('nome', text='Colaborador')
        self.tabela.heading('cpf', text='CPF')
        self.tabela.heading('valor', text='Valor')
        self.tabela.heading('vencimento', text='Vencimento')
        self.tabela.heading('status', text='Status')
        self.tabela.column('nome', width=300)
        self.tabela.column('cpf', width=130)
        self.tabela.column('valor', width=100, anchor='e')
        self.tabela.column('vencimento', width=100, anchor='center')
        self.tabela.column('status', width=160)
        self.tabela.pack(side='left', fill='both', expand=True)
        self.tabela.tag_configure('ja_emitido', background='#fff3cd')

        scroll = ttk.Scrollbar(frame_tabela, orient='vertical', command=self.tabela.yview)
        self.tabela.configure(yscrollcommand=scroll.set)
        scroll.pack(side='left', fill='y')

        frame_acoes = ttk.Frame(self.root, padding=10)
        frame_acoes.pack(fill='x')

        self.var_reemitir = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            frame_acoes,
            text="Permitir reemitir (2ª via) os já marcados como 'Já emitido' na seleção",
            variable=self.var_reemitir,
        ).pack(side='left')

        ttk.Button(
            frame_acoes, text="Selecionar todos", command=self._selecionar_todos,
        ).pack(side='left', padx=(15, 5))
        ttk.Button(
            frame_acoes, text="Emitir selecionados", command=self._emitir_selecionados,
        ).pack(side='right')

        self._carregar_clientes()

    # ------------------------------------------------------------------
    # Carregamento
    # ------------------------------------------------------------------

    def _normalizar_para_comparacao(self, texto: str) -> str:
        """
        Normaliza um nome pra comparação tolerante: maiúsculas, sem
        acento, underscore/traço tratados como espaço, espaços
        colapsados. Usado só pra CASAR o nome do cliente (vindo de
        Clientes.xlsx) com o arquivo .xlsx correspondente em
        PASTA_CLIENTES — os nomes de arquivo real usam underscore
        (ex.: CLEVER_LUIZ_SALVADOR.xlsx) enquanto Clientes.xlsx guarda
        o nome com espaços (ex.: "CLEVER LUIZ SALVADOR").
        """
        sem_acento = unicodedata.normalize('NFKD', texto).encode('ascii', 'ignore').decode()
        sem_acento = sem_acento.upper()
        sem_acento = re.sub(r'[_\-]+', ' ', sem_acento)
        return re.sub(r'\s+', ' ', sem_acento).strip()

    def _carregar_clientes(self):
        """
        Carrega clientes ATIVOS a partir de Clientes.xlsx (fonte de
        verdade — coluna E "Data Final" vazia OU futura, conforme
        obter_clientes_ativos), e resolve o arquivo .xlsx real de cada
        um em PASTA_CLIENTES por correspondência tolerante de nome.
        """
        try:
            try:
                from src.config.utils import obter_clientes_ativos
            except ImportError:
                from config.utils import obter_clientes_ativos

            nomes_ativos, _info_clientes = obter_clientes_ativos(mostrar_inativos=False)
        except ImportError:
            messagebox.showerror(
                "Erro",
                "Não foi possível importar obter_clientes_ativos de src/config/utils.py. "
                "Verifique se o módulo está acessível.",
            )
            nomes_ativos = []
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler Clientes.xlsx:\n{e}")
            nomes_ativos = []

        if not nomes_ativos:
            self.combo_cliente['values'] = []
            return

        # Mapa arquivo real (normalizado) -> Path, a partir do que existe de fato em PASTA_CLIENTES
        arquivos_disponiveis = {
            self._normalizar_para_comparacao(p.stem): p
            for p in self.pasta_clientes.glob('*.xlsx')
        }

        self._mapa_clientes = {}
        sem_arquivo = []
        for nome in nomes_ativos:
            chave = self._normalizar_para_comparacao(nome)
            arquivo = arquivos_disponiveis.get(chave)
            if arquivo:
                self._mapa_clientes[nome] = arquivo
            else:
                sem_arquivo.append(nome)

        self.combo_cliente['values'] = list(self._mapa_clientes.keys())

        if sem_arquivo:
            # Não bloqueia a tela — só avisa. Cliente pode estar ativo em
            # Clientes.xlsx mas ainda sem planilha própria criada.
            self.texto_avisos.configure(state='normal')
            self.texto_avisos.delete('1.0', 'end')
            self.texto_avisos.insert(
                'end',
                f"• {len(sem_arquivo)} cliente(s) ativo(s) em Clientes.xlsx sem planilha "
                f"correspondente encontrada em PASTA_CLIENTES: {', '.join(sem_arquivo)}\n",
            )
            self.texto_avisos.configure(state='disabled')

    def _ao_selecionar_cliente(self, event=None):
        nome_exibicao = self.combo_cliente.get()
        arquivo = self._mapa_clientes.get(nome_exibicao)
        if not arquivo:
            return
        self.arquivo_cliente_selecionado = arquivo

        try:
            self.pagador_atual = obter_dados_pagador(nome_exibicao, str(self.arquivo_clientes))
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler dados do pagador:\n{e}")
            self.pagador_atual = None
            return

        try:
            competencias = obter_competencias_disponiveis(str(arquivo), self.competencia_atual)
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível ler competências:\n{e}")
            competencias = []

        self.combo_competencia['values'] = competencias
        self.combo_competencia.set(competencias[0] if competencias else '')

    # ------------------------------------------------------------------
    # Busca de candidatos
    # ------------------------------------------------------------------

    def _buscar_candidatos(self):
        if not self.arquivo_cliente_selecionado:
            messagebox.showwarning("Atenção", "Selecione um cliente.")
            return
        competencia = self.combo_competencia.get()
        label_beneficio = self.combo_beneficio.get()
        if not competencia or not label_beneficio:
            messagebox.showwarning("Atenção", "Selecione a competência e o benefício.")
            return

        beneficio = _LABEL_TO_BENEFICIO[label_beneficio]

        try:
            candidatos, avisos = obter_candidatos(
                str(self.arquivo_cliente_selecionado), str(self.arquivo_fornecedores),
                beneficio, competencia, self.competencia_atual,
            )
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao buscar candidatos:\n{e}")
            return

        self.candidatos_atuais = candidatos
        self.avisos_atuais = avisos
        self.beneficio_atual = beneficio
        self.competencia_selecionada = competencia

        try:
            emitidos = {
                r.cpf for r in listar_emitidos(
                    str(self.arquivo_cliente_selecionado), competencia, beneficio,
                )
                if r.status == 'EMITIDO'
            }
        except Exception as e:
            messagebox.showerror("Erro", f"Falha ao ler controle de emissões:\n{e}")
            emitidos = set()

        self.tabela.delete(*self.tabela.get_children())
        for c in candidatos:
            ja = c.cpf in emitidos
            status = 'Já emitido' if ja else 'Disponível'
            valor_fmt = formatar_valor_monetario(c.valor) if c.valor is not None else '—'
            venc_fmt = c.data_vencimento.strftime('%d/%m/%Y') if c.data_vencimento else '—'
            self.tabela.insert(
                '', 'end', iid=c.cpf,
                values=(c.nome, formatar_cpf(c.cpf), valor_fmt, venc_fmt, status),
                tags=('ja_emitido',) if ja else (),
            )

        self.texto_avisos.configure(state='normal')
        self.texto_avisos.delete('1.0', 'end')
        if avisos:
            for a in avisos:
                linha = f"• {a.cpf}: {a.mensagem}\n" if a.cpf else f"• {a.mensagem}\n"
                self.texto_avisos.insert('end', linha)
        else:
            self.texto_avisos.insert('end', 'Nenhum aviso.')
        self.texto_avisos.configure(state='disabled')

    def _selecionar_todos(self):
        self.tabela.selection_set(self.tabela.get_children())

    # ------------------------------------------------------------------
    # Emissão
    # ------------------------------------------------------------------

    def _emitir_selecionados(self):
        selecionados = self.tabela.selection()
        if not selecionados:
            messagebox.showwarning("Atenção", "Selecione ao menos um colaborador.")
            return
        if not self.pagador_atual:
            messagebox.showerror("Erro", "Dados do pagador não carregados.")
            return

        permitir_reemissao = self.var_reemitir.get()
        cpf_para_candidato = {c.cpf: c for c in self.candidatos_atuais}

        pasta_saida = (
            self.pasta_saida_base / self.arquivo_cliente_selecionado.stem
            / 'Comprovantes' / self.competencia_selecionada
        )

        # Fase 1: gera todos os PDFs (cada geração é independente — um
        # erro em um colaborador não impede os demais).
        novos_registros: list[NovoRegistro] = []
        erros_pdf: list[str] = []

        for cpf in selecionados:
            candidato = cpf_para_candidato.get(cpf)
            if candidato is None:
                continue
            try:
                caminho_pdf = gerar_recibo_pdf(
                    self.beneficio_atual, candidato, self.pagador_atual,
                    date.today(), pasta_saida,
                )
            except ErroGeracaoRecibo as e:
                erros_pdf.append(f"{candidato.nome}: {e}")
                continue

            novos_registros.append(NovoRegistro(
                beneficio=self.beneficio_atual,
                competencia=self.competencia_selecionada,
                cpf=candidato.cpf,
                nome=candidato.nome,
                caminho_pdf=str(caminho_pdf),
                valor=candidato.valor,
                dias=candidato.dias,
                data_vencimento=(
                    candidato.data_vencimento.isoformat()
                    if candidato.data_vencimento else None
                ),
                usuario=self.usuario,
                maquina=self.maquina,
            ))

        # Fase 2: grava o lote inteiro numa única abertura/gravação da
        # planilha do cliente.
        try:
            gravados, pulados = registrar_lote(
                str(self.arquivo_cliente_selecionado), novos_registros,
                permitir_reemissao=permitir_reemissao,
            )
        except Exception as e:
            messagebox.showerror(
                "Erro",
                f"PDFs foram gerados em:\n{pasta_saida}\n\n"
                f"mas houve falha ao gravar o controle na planilha do cliente:\n{e}",
            )
            return

        resumo = f"{len(gravados)} comprovante(s) emitido(s) em:\n{pasta_saida}"
        if pulados:
            nomes_pulados = ", ".join(p.nome for p in pulados)
            resumo += f"\n\n{len(pulados)} pulado(s) — já emitido(s): {nomes_pulados}"
        if erros_pdf:
            resumo += "\n\nErros ao gerar PDF:\n" + "\n".join(erros_pdf)

        messagebox.showinfo("Emissão concluída", resumo)

        if gravados and pasta_saida.exists():
            self._abrir_pasta(pasta_saida)

        self._buscar_candidatos()

    @staticmethod
    def _abrir_pasta(caminho: Path):
        try:
            if sys.platform.startswith('win'):
                os.startfile(str(caminho))
            elif sys.platform == 'darwin':
                subprocess.run(['open', str(caminho)])
            else:
                subprocess.run(['xdg-open', str(caminho)])
        except Exception:
            pass  # não crítico — o caminho já consta no resumo exibido
