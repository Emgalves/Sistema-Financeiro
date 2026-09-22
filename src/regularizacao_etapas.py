# -*- coding: utf-8 -*-
"""
regularizacao_etapas.py
=======================
Ferramenta para informar a etapa da obra nos lançamentos que ficaram sem ETAPA_OBRA.

Janela dedicada (mesmo padrão de RelatorioConsistenciaDados):
    RegularizacaoEtapas(parent=self.root, cliente_inicial="NOME DO CLIENTE")

Abas:
    1. Mão de obra por quinzena: sugere a etapa em andamento na quinzena.
    2. Material, locação, serviço e diversos: agrupado por fornecedor, com sugestão (contrato de
       medição, fornecedor já classificado, palavra-chave). A data não é critério.
    3. Auditoria (somente lista): etapas já gravadas que conflitam com as regras. Nada é alterado.

Nada é gravado até o usuário confirmar a "Prévia da gravação". A gravação é feita por
etapas_obra_core.gravar_etapas (backup, conferência linha a linha, troca atômica).

Também expõe abrir_ordem_e_marco(...): diálogo para ajustar a ordem das etapas do relatório e o
marco de data (quinzena a partir da qual a ausência de etapa é considerada pendência).
"""

import logging
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk
from typing import Callable, Dict, List, Optional

import pandas as pd

from src import etapas_obra_core as core

try:
    from src.config.config import ARQUIVO_CLIENTES, PASTA_CLIENTES
except ImportError:                                            # pragma: no cover
    from config.config import ARQUIVO_CLIENTES, PASTA_CLIENTES

try:
    from src.config.window_config import configurar_janela
except ImportError:                                            # pragma: no cover
    def configurar_janela(janela, titulo, largura=1150, altura=820):
        janela.title(titulo)
        sw, sh = janela.winfo_screenwidth(), janela.winfo_screenheight()
        janela.geometry(f"{min(largura, sw)}x{min(altura, sh)}+0+0")
        janela.resizable(True, True)
        janela.lift()
        janela.focus_force()

logger = logging.getLogger(__name__)

COR_SEM_SUGESTAO = "#FDEBD0"
COR_APLICADO = "#DCEFDC"


def listar_clientes_ativos() -> List[str]:
    """Clientes ativos (Data Final vazia), em ordem alfabética."""
    df = pd.read_excel(ARQUIVO_CLIENTES)
    col_nome = next((c for c in ("Nome", "nome", "NOME", "Cliente") if c in df.columns), None)
    if not col_nome:
        return []
    col_final = next((c for c in ("Data Final", "data_final") if c in df.columns), None)
    if col_final:
        df = df[df[col_final].isna()]
    return sorted(str(n) for n in df[col_nome].dropna().tolist())


def arquivo_do_cliente(nome: str) -> Path:
    return Path(PASTA_CLIENTES) / f"{nome}.xlsx"


def _moeda(v) -> str:
    return "R$ " + core.fmt_moeda(v)


# ---------------------------------------------------------------------------
# Árvore de fila (grupo -> lançamentos) reutilizada nas abas 1 e 2
# ---------------------------------------------------------------------------
class FilaArvore:
    COLUNAS = (("cat", "Cat.", 70, "w"), ("data", "Data", 80, "w"), ("valor", "Valor (R$)", 95, "e"),
               ("sug", "Sugestão", 190, "w"), ("conf", "Conf.", 55, "center"), ("aplicar", "Aplicar a", 190, "w"),
               ("motivo", "Motivo", 320, "w"))

    def __init__(self, master, app: "RegularizacaoEtapas", texto_raiz: str):
        self.app = app
        self.frame = ttk.Frame(master)
        self.tree = ttk.Treeview(self.frame, columns=[c[0] for c in self.COLUNAS], show="tree headings", selectmode="extended")
        self.tree.heading("#0", text=texto_raiz)
        self.tree.column("#0", width=320, anchor="w")
        for cid, tit, larg, anc in self.COLUNAS:
            self.tree.heading(cid, text=tit)
            self.tree.column(cid, width=larg, anchor=anc)
        sy = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        sx = ttk.Scrollbar(self.frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=sy.set, xscrollcommand=sx.set)
        sy.pack(side="right", fill="y")
        sx.pack(side="bottom", fill="x")
        self.tree.pack(side="left", fill="both", expand=True)
        self.tree.tag_configure("sem_sug", background=COR_SEM_SUGESTAO)
        self.tree.tag_configure("aplicado", background=COR_APLICADO)
        self.grupos: List[core.Grupo] = []
        self._por_linha: Dict[int, core.Item] = {}
        self._por_chave: Dict[str, core.Grupo] = {}

    # ------------------------------------------------------------------ carga
    def carregar(self, grupos: List[core.Grupo]):
        self.grupos = grupos
        self._por_linha = {it.linha: it for g in grupos for it in g.itens}
        self._por_chave = {g.chave: g for g in grupos}
        self.tree.delete(*self.tree.get_children())
        for g in grupos:
            self.tree.insert("", "end", iid=g.chave, open=False)
            for it in g.itens:
                self.tree.insert(g.chave, "end", iid=f"L|{it.linha}")
        self.refrescar()

    def refrescar(self):
        pend = self.app.pendentes
        for g in self.grupos:
            sugs = {i.sug for i in g.itens}
            apls = {pend.get(i.linha) for i in g.itens}
            confs = {i.conf for i in g.itens if i.sug}
            sug = next(iter(sugs)) if len(sugs) == 1 and None not in sugs else ("varia" if sugs - {None} else "-")
            conf = "Alta" if confs == {"Alta"} else ("Média" if confs else "-")
            if len(apls) == 1 and None not in apls:
                apl = next(iter(apls))
            elif apls - {None}:
                apl = "várias"
            else:
                apl = ""
            motivos = {i.motivo for i in g.itens}
            motivo = next(iter(motivos)) if len(motivos) == 1 else "diversos (expanda para ver)"
            cats = "/".join(sorted({i.categoria for i in g.itens}))
            tags = ("aplicado",) if apl and apl != "várias" else (("sem_sug",) if sug == "-" and not apl else ())
            self.tree.item(g.chave, text=f"{g.rotulo}  ({len(g.itens)})", tags=tags,
                           values=(cats, "", core.fmt_moeda(g.valor), sug, conf, apl, motivo))
            for it in g.itens:
                apl_i = pend.get(it.linha, "")
                ttags = ("aplicado",) if apl_i else (("sem_sug",) if not it.sug else ())
                texto = f"{it.nome} — {it.referencia}" if g.chave.startswith("Q|") else it.referencia
                self.tree.item(f"L|{it.linha}", text=texto, tags=ttags,
                               values=(it.categoria, core.fmt_data(it.data), core.fmt_moeda(it.valor), it.sug or "-",
                                       it.conf, apl_i, it.motivo))

    # ------------------------------------------------------------------ seleção / ações
    def itens_selecionados(self) -> List[core.Item]:
        vistos, out = set(), []
        for iid in self.tree.selection():
            if iid.startswith("L|"):
                cand = [self._por_linha[int(iid[2:])]]
            else:
                cand = self._por_chave[iid].itens
            for it in cand:
                if it.linha not in vistos:
                    vistos.add(it.linha)
                    out.append(it)
        return out

    def itens_todos(self) -> List[core.Item]:
        return [it for g in self.grupos for it in g.itens]


# ---------------------------------------------------------------------------
# Janela principal
# ---------------------------------------------------------------------------
class RegularizacaoEtapas:
    def __init__(self, parent=None, cliente_inicial: Optional[str] = None):
        self.parent = parent
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        configurar_janela(self.root, "Regularizar Etapas da Obra", 1200, 820)
        self.root.protocol("WM_DELETE_WINDOW", self._fechar)

        self.cliente_atual: Optional[str] = None
        self.arquivo: Optional[Path] = None
        self.ctx: Optional[core.Contexto] = None
        self.bruto = None
        self.d = None
        self.contratos = None
        self.pendentes: Dict[int, str] = {}         # linha -> etapa escolhida (ainda não gravada)
        self._cliente_inicial = cliente_inicial
        self._setup_gui()

    # ------------------------------------------------------------------ interface
    def _setup_gui(self):
        fp = ttk.Frame(self.root, padding=10)
        fp.pack(fill="both", expand=True)
        fp.rowconfigure(2, weight=1)
        fp.columnconfigure(0, weight=1)

        sel = ttk.LabelFrame(fp, text="Cliente")
        sel.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        inner = ttk.Frame(sel)
        inner.pack(fill="x", padx=10, pady=8)
        ttk.Label(inner, text="Cliente:", font=("Arial", 11)).pack(side="left")
        self.cmb_cliente = ttk.Combobox(inner, width=48, font=("Arial", 11), state="readonly")
        self.cmb_cliente.pack(side="left", padx=8)
        self.cmb_cliente.bind("<<ComboboxSelected>>", self._on_cliente)
        ttk.Button(inner, text="Recarregar", command=self._recarregar).pack(side="left", padx=4)
        ttk.Button(inner, text="Ordem e marco...", command=self._abrir_ordem_marco).pack(side="left", padx=4)
        self.lbl_marco = ttk.Label(inner, text="", foreground="#555")
        self.lbl_marco.pack(side="left", padx=12)

        self.lbl_resumo = ttk.Label(fp, text="Selecione um cliente.", font=("Arial", 10, "bold"), foreground="#8A4B00")
        self.lbl_resumo.grid(row=1, column=0, sticky="w", pady=(0, 4))

        self.nb = ttk.Notebook(fp)
        self.nb.grid(row=2, column=0, sticky="nsew")
        self._montar_aba_mo()
        self._montar_aba_fornecedores()
        self._montar_aba_auditoria()

        bot = ttk.Frame(fp)
        bot.grid(row=3, column=0, sticky="ew", pady=(6, 0))
        self.lbl_pend = ttk.Label(bot, text="Nada atribuído ainda.")
        self.lbl_pend.pack(side="left")
        ttk.Button(bot, text="Fechar", command=self._fechar).pack(side="right", padx=4)
        ttk.Button(bot, text="Prévia da gravação...", command=self._previa, style="Accentuated.TButton").pack(side="right", padx=4)

        self._carregar_lista_clientes()

    def _barra(self, parent, com_filtro=False):
        barra = ttk.Frame(parent)
        barra.pack(fill="x", pady=(6, 4))
        return barra

    def _montar_aba_mo(self):
        aba = ttk.Frame(self.nb, padding=6)
        self.nb.add(aba, text="1. Mão de obra por quinzena")
        ttk.Label(aba, text="Mão de obra sem etapa, por quinzena. Escolha uma linha (quinzena ou lançamento) e a etapa em andamento.",
                  foreground="#555").pack(anchor="w")
        barra = self._barra(aba)
        self.cmb_etapa_mo = ttk.Combobox(barra, width=34, state="readonly")
        self.cmb_etapa_mo.pack(side="left", padx=(0, 4))
        ttk.Button(barra, text="Atribuir às selecionadas", command=lambda: self._atribuir(self.fila_mo, self.cmb_etapa_mo.get())).pack(side="left", padx=2)
        ttk.Button(barra, text="Marcar como CUSTOS GERAIS", command=lambda: self._atribuir(self.fila_mo, self.ctx.gerais if self.ctx else "")).pack(side="left", padx=2)
        ttk.Button(barra, text="Aceitar sugestões (>= 80%)", command=lambda: self._aceitar(self.fila_mo, ("Alta", "Média"))).pack(side="left", padx=12)
        ttk.Button(barra, text="Limpar atribuição", command=lambda: self._limpar(self.fila_mo)).pack(side="left", padx=2)
        self.fila_mo = FilaArvore(aba, self, "Quinzena / lançamento")
        self.fila_mo.frame.pack(fill="both", expand=True)

    def _montar_aba_fornecedores(self):
        aba = ttk.Frame(self.nb, padding=6)
        self.nb.add(aba, text="2. Material, locação, serviço e diversos")
        ttk.Label(aba, text="Agrupado por fornecedor: escolher a etapa no fornecedor aplica a todos os lançamentos dele; expanda para dividir.",
                  foreground="#555").pack(anchor="w")
        barra = self._barra(aba)
        self.cmb_etapa_fo = ttk.Combobox(barra, width=34, state="readonly")
        self.cmb_etapa_fo.pack(side="left", padx=(0, 4))
        ttk.Button(barra, text="Atribuir às selecionadas", command=lambda: self._atribuir(self.fila_fo, self.cmb_etapa_fo.get())).pack(side="left", padx=2)
        ttk.Button(barra, text="Marcar como CUSTOS GERAIS", command=lambda: self._atribuir(self.fila_fo, self.ctx.gerais if self.ctx else "")).pack(side="left", padx=2)
        ttk.Button(barra, text="Aceitar sugestões Alta", command=lambda: self._aceitar(self.fila_fo, ("Alta",))).pack(side="left", padx=(12, 2))
        ttk.Button(barra, text="Aceitar Alta e Média", command=lambda: self._aceitar(self.fila_fo, ("Alta", "Média"))).pack(side="left", padx=2)
        ttk.Button(barra, text="Limpar atribuição", command=lambda: self._limpar(self.fila_fo)).pack(side="left", padx=2)
        barra2 = ttk.Frame(aba)
        barra2.pack(fill="x", pady=(0, 4))
        ttk.Label(barra2, text="Categoria:").pack(side="left")
        self.cmb_filtro_cat = ttk.Combobox(barra2, width=8, state="readonly", values=["Todas", "MAT", "LOC", "SERV", "DIV"])
        self.cmb_filtro_cat.current(0)
        self.cmb_filtro_cat.pack(side="left", padx=(4, 12))
        self.cmb_filtro_cat.bind("<<ComboboxSelected>>", lambda e: self._aplicar_filtro())
        ttk.Label(barra2, text="Buscar:").pack(side="left")
        self.var_busca = tk.StringVar()
        ent = ttk.Entry(barra2, textvariable=self.var_busca, width=30)
        ent.pack(side="left", padx=4)
        ent.bind("<KeyRelease>", lambda e: self._aplicar_filtro())
        self.fila_fo = FilaArvore(aba, self, "Fornecedor / lançamento")
        self.fila_fo.frame.pack(fill="both", expand=True)
        self._grupos_fo_todos: List[core.Grupo] = []

    def _montar_aba_auditoria(self):
        aba = ttk.Frame(self.nb, padding=6)
        self.nb.add(aba, text="3. Auditoria (só lista)")
        ttk.Label(aba, text="Lançamentos com etapa gravada que conflitam com as regras. Esta ferramenta não altera etapas já preenchidas "
                            "(para corrigir, use o Editor em Massa do Gerenciador de Lançamentos).", foreground="#555",
                  wraplength=1100, justify="left").pack(anchor="w", pady=(0, 4))
        cols = ("linha", "data", "cat", "nome", "referencia", "valor", "etapa", "situacao")
        frame = ttk.Frame(aba)
        frame.pack(fill="both", expand=True)
        self.tree_aud = ttk.Treeview(frame, columns=cols, show="headings")
        for c, t, w, a in (("linha", "Linha", 50, "center"), ("data", "Data", 80, "w"), ("cat", "Cat.", 50, "w"),
                           ("nome", "Nome", 200, "w"), ("referencia", "Referência", 260, "w"), ("valor", "Valor (R$)", 90, "e"),
                           ("etapa", "Etapa gravada", 170, "w"), ("situacao", "Situação", 420, "w")):
            self.tree_aud.heading(c, text=t)
            self.tree_aud.column(c, width=w, anchor=a)
        sy = ttk.Scrollbar(frame, orient="vertical", command=self.tree_aud.yview)
        self.tree_aud.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        self.tree_aud.pack(side="left", fill="both", expand=True)

    # ------------------------------------------------------------------ clientes / carga
    def _carregar_lista_clientes(self):
        try:
            nomes = listar_clientes_ativos()
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro", f"Erro ao carregar clientes:\n{e}", parent=self.root)
            return
        self.cmb_cliente["values"] = nomes
        if self._cliente_inicial and self._cliente_inicial in nomes:
            self.cmb_cliente.set(self._cliente_inicial)
            self._on_cliente()

    def _on_cliente(self, event=None):
        nome = self.cmb_cliente.get()
        if not nome:
            return
        if self.pendentes and not messagebox.askyesno(
                "Atribuições não gravadas", "Há atribuições que ainda não foram gravadas. Trocar de cliente descarta essas atribuições. Continuar?",
                parent=self.root):
            if self.cliente_atual:
                self.cmb_cliente.set(self.cliente_atual)
            return
        self.cliente_atual = nome
        self.arquivo = arquivo_do_cliente(nome)
        self.pendentes = {}
        self._recarregar()

    def _recarregar(self):
        if not self.arquivo:
            return
        if not self.arquivo.exists():
            messagebox.showerror("Erro", f"Arquivo não encontrado:\n{self.arquivo}", parent=self.root)
            return
        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            self.ctx = core.criar_contexto(self.arquivo)
            self.bruto = core.carregar_lancamentos(self.arquivo)
            self.d = core.preparar(self.bruto, self.ctx)
            self.contratos = core.carregar_contratos(self.arquivo)
            self._grupos_mo = core.fila_mao_de_obra(self.d)
            self._grupos_fo_todos = core.fila_fornecedores(self.d, self.contratos, self.ctx)
            vivos = {it.linha for g in self._grupos_mo + self._grupos_fo_todos for it in g.itens}
            self.pendentes = {l: e for l, e in self.pendentes.items() if l in vivos}
            opcoes = self.ctx.etapas_para_combo()
            for cmb in (self.cmb_etapa_mo, self.cmb_etapa_fo):
                atual = cmb.get()
                cmb["values"] = opcoes
                cmb.set(atual if atual in opcoes else "")
            self.fila_mo.carregar(self._grupos_mo)
            self._aplicar_filtro()
            self._carregar_auditoria()
            self._atualizar_resumo()
            if not self.ctx.gerais_cadastrada:
                messagebox.showwarning(
                    "Etapa não cadastrada",
                    f"'{core.ETAPA_GERAIS}' não está na lista de etapas de Configurações.\n"
                    "A marcação funciona, mas cadastre a etapa para que ela também apareça na entrada de dados.",
                    parent=self.root)
        except Exception as e:  # noqa: BLE001
            logger.error(f"Erro ao carregar cliente: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao carregar a planilha:\n{e}", parent=self.root)
        finally:
            self.root.config(cursor="")

    def _aplicar_filtro(self):
        cat = self.cmb_filtro_cat.get() if hasattr(self, "cmb_filtro_cat") else "Todas"
        busca = core.normalizar(self.var_busca.get()) if hasattr(self, "var_busca") else ""
        grupos = []
        for g in self._grupos_fo_todos:
            itens = [i for i in g.itens if (cat in ("Todas", "") or i.categoria == cat)
                     and (not busca or busca in core.normalizar(g.rotulo) or busca in core.normalizar(i.referencia))]
            if itens:
                grupos.append(core.Grupo(chave=g.chave, rotulo=g.rotulo, itens=itens))
        self.fila_fo.carregar(grupos)

    def _carregar_auditoria(self):
        self.tree_aud.delete(*self.tree_aud.get_children())
        for a in core.auditoria(self.d, self.contratos, self.ctx):
            self.tree_aud.insert("", "end", values=(a["linha"], core.fmt_data(a["data"]), a["categoria"], a["nome"], a["referencia"],
                                                    core.fmt_moeda(a["valor"]), a["etapa"], a["situacao"]))

    def _atualizar_resumo(self):
        d = self.d
        total = float(d["VALOR"].sum()) if len(d) else 0.0
        sem = d[d["DESTINO"] == core.D_SEM_ETAPA]
        ant = d[d["DESTINO"] == core.D_ANTERIOR]
        pct = (sem["VALOR"].sum() / total) if total else 0
        self.lbl_resumo.config(
            text=f"Sem etapa informada: {len(sem)} lançamentos · {_moeda(sem['VALOR'].sum())} ({core.fmt_pct(pct)} da obra)     |     "
                 f"Anteriores ao marco: {len(ant)} lançamentos · {_moeda(ant['VALOR'].sum())} (fora da fila)")
        origem = "salvo neste cliente" if core.ler_config_cliente(self.arquivo).get("marco") else "padrão"
        self.lbl_marco.config(text=f"Marco: {core.fmt_data(self.ctx.marco)} ({origem})")
        self._atualizar_pendentes()

    def _atualizar_pendentes(self):
        itens = {it.linha: it for g in self.fila_mo.grupos + self._grupos_fo_todos for it in g.itens}
        valor = sum(itens[l].valor for l in self.pendentes if l in itens)
        self.lbl_pend.config(text=f"Atribuições aguardando gravação: {len(self.pendentes)} lançamentos · {_moeda(valor)}"
                             if self.pendentes else "Nada atribuído ainda.")
        for fila in (self.fila_mo, self.fila_fo):
            fila.refrescar()

    # ------------------------------------------------------------------ ações de atribuição
    def _atribuir(self, fila: FilaArvore, etapa: str):
        if not self.ctx:
            return
        if not etapa:
            messagebox.showinfo("Etapa", "Escolha a etapa na lista ao lado do botão.", parent=self.root)
            return
        itens = fila.itens_selecionados()
        if not itens:
            messagebox.showinfo("Seleção", "Selecione uma ou mais linhas (grupo ou lançamento).", parent=self.root)
            return
        canon = self.ctx.canonica(etapa)
        if not canon:
            messagebox.showerror("Etapa", f"Etapa '{etapa}' não está na lista de Configurações.", parent=self.root)
            return
        for it in itens:
            self.pendentes[it.linha] = canon
        self._atualizar_pendentes()

    def _aceitar(self, fila: FilaArvore, confs):
        n = 0
        for it in fila.itens_todos():
            if it.sug and it.conf in confs:
                self.pendentes[it.linha] = it.sug
                n += 1
        if not n:
            messagebox.showinfo("Sugestões", "Não há sugestões nesse nível de confiança.", parent=self.root)
        self._atualizar_pendentes()

    def _limpar(self, fila: FilaArvore):
        itens = fila.itens_selecionados() if fila.tree.selection() else fila.itens_todos()
        for it in itens:
            self.pendentes.pop(it.linha, None)
        self._atualizar_pendentes()

    # ------------------------------------------------------------------ prévia e gravação
    def montar_alteracoes(self) -> List[dict]:
        itens = {it.linha: it for g in self.fila_mo.grupos + self._grupos_fo_todos for it in g.itens}
        return [{"linha": l, "id": itens[l].id, "valor": itens[l].valor, "etapa": e, "data": itens[l].data,
                 "nome": itens[l].nome, "referencia": itens[l].referencia}
                for l, e in sorted(self.pendentes.items()) if l in itens]

    def _previa(self):
        alts = self.montar_alteracoes()
        if not alts:
            messagebox.showinfo("Prévia", "Não há atribuições para gravar.", parent=self.root)
            return
        DialogoPrevia(self, alts)

    def _abrir_ordem_marco(self):
        if not self.arquivo:
            messagebox.showinfo("Cliente", "Selecione um cliente.", parent=self.root)
            return
        DialogoOrdemMarco(self.root, self.arquivo, on_saved=self._recarregar)

    def _fechar(self):
        if self.pendentes and not messagebox.askyesno(
                "Atribuições não gravadas", "Há atribuições que ainda não foram gravadas e serão descartadas. Fechar mesmo assim?",
                parent=self.root):
            return
        self.root.destroy()
        if self.parent:
            try:
                self.parent.deiconify()
                self.parent.lift()
                self.parent.focus_force()
            except Exception:  # noqa: BLE001
                pass


# ---------------------------------------------------------------------------
# Diálogo: prévia da gravação
# ---------------------------------------------------------------------------
class DialogoPrevia:
    def __init__(self, app: RegularizacaoEtapas, alteracoes: List[dict]):
        self.app = app
        self.alts = alteracoes
        self.win = tk.Toplevel(app.root)
        self.win.title("Prévia da gravação")
        self.win.geometry("1000x620")
        self.win.transient(app.root)
        fr = ttk.Frame(self.win, padding=10)
        fr.pack(fill="both", expand=True)
        total = sum(a["valor"] for a in alteracoes)
        ttk.Label(fr, text=f"{len(alteracoes)} lançamentos · {_moeda(total)} · ETAPA_OBRA (coluna Q) vazia em todos · nada foi gravado ainda",
                  font=("Arial", 10, "bold")).pack(anchor="w")
        cols = ("linha", "id", "data", "nome", "referencia", "valor", "etapa")
        box = ttk.Frame(fr)
        box.pack(fill="both", expand=True, pady=6)
        tv = ttk.Treeview(box, columns=cols, show="headings")
        for c, t, w, a in (("linha", "Linha", 50, "center"), ("id", "ID", 55, "center"), ("data", "Data rel.", 90, "w"),
                           ("nome", "Nome", 200, "w"), ("referencia", "Referência", 250, "w"), ("valor", "Valor (R$)", 90, "e"),
                           ("etapa", "ETAPA_OBRA: antes → depois", 250, "w")):
            tv.heading(c, text=t)
            tv.column(c, width=w, anchor=a)
        sy = ttk.Scrollbar(box, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=sy.set)
        sy.pack(side="right", fill="y")
        tv.pack(side="left", fill="both", expand=True)
        for a in alteracoes:
            tv.insert("", "end", values=(a["linha"], a["id"], core.fmt_data(a["data"]), a["nome"], a["referencia"],
                                         core.fmt_moeda(a["valor"]), f"(vazia)  →  {a['etapa']}"))
        modelo_backup = core._pasta_backup(app.arquivo) / f"{app.arquivo.stem}_ANTES_ETAPAS_aaaammdd_hhmmss{app.arquivo.suffix}"
        info = (f"• Backup antes de gravar: {modelo_backup}\n"
                "• Histórico (coluna P), formato do sistema: EDIÇÃO: ETAPA_OBRA:  → ETAPA - dd/mm/aaaa hh:mm:ss\n"
                "• Antes de gravar a planilha é relida: ID, valor e coluna Q vazia de cada linha são conferidos; se alguém alterou, a linha é pulada\n"
                "• Grava somente Dados!P e Dados!Q, numa única gravação. Não ordena, insere nem apaga linhas\n"
                "• Se a planilha estiver aberta no Excel ou bloqueada, nada é gravado")
        ttk.Label(fr, text=info, justify="left", foreground="#333", wraplength=960).pack(anchor="w", pady=(0, 8))
        bt = ttk.Frame(fr)
        bt.pack(fill="x")
        ttk.Button(bt, text="Cancelar", command=self.win.destroy).pack(side="right", padx=4)
        self.btn = ttk.Button(bt, text=f"Gravar {len(alteracoes)} lançamentos", command=self._gravar, style="Accentuated.TButton")
        self.btn.pack(side="right", padx=4)
        self.win.grab_set()

    def _gravar(self):
        self.btn.state(["disabled"])
        self.win.config(cursor="watch")
        self.win.update_idletasks()
        try:
            res = core.gravar_etapas(self.app.arquivo, self.alts, self.app.ctx)
        except core.ArquivoEmUsoError as e:
            messagebox.showerror("Planilha em uso", f"{e}\n\nFeche a planilha no Excel e tente novamente. Nada foi gravado.", parent=self.win)
            self.btn.state(["!disabled"]); self.win.config(cursor="")
            return
        except core.PlanilhaAlteradaError as e:
            messagebox.showerror("Planilha alterada", f"{e}\n\nRecarregue a tela e refaça a gravação.", parent=self.win)
            self.btn.state(["!disabled"]); self.win.config(cursor="")
            return
        except Exception as e:  # noqa: BLE001
            logger.error(f"Erro ao gravar etapas: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao gravar. A planilha original não foi alterada.\n\n{e}", parent=self.win)
            self.btn.state(["!disabled"]); self.win.config(cursor="")
            return
        for g in res.gravados:
            self.app.pendentes.pop(g["linha"], None)
        msg = f"{len(res.gravados)} lançamento(s) gravado(s)."
        if res.backup:
            msg += f"\n\nBackup: {res.backup}"
        if res.pulados:
            msg += f"\n\n{len(res.pulados)} lançamento(s) NÃO gravado(s):\n" + "\n".join(
                f"  linha {l}: {m}" for l, m in res.pulados[:15]) + ("\n  ..." if len(res.pulados) > 15 else "")
        messagebox.showinfo("Gravação concluída", msg, parent=self.win)
        self.win.destroy()
        self.app._recarregar()


# ---------------------------------------------------------------------------
# Diálogo: ordem das etapas e marco de data
# ---------------------------------------------------------------------------
class DialogoOrdemMarco:
    def __init__(self, parent_widget, arquivo, on_saved: Optional[Callable] = None):
        self.arquivo = Path(arquivo)
        self.on_saved = on_saved
        self.win = tk.Toplevel(parent_widget)
        self.win.title(f"Ordem das etapas e marco de data — {self.arquivo.stem}")
        self.win.geometry("900x600")
        self.win.transient(parent_widget)
        try:
            self.ctx = core.criar_contexto(self.arquivo)
            self.bruto = core.carregar_lancamentos(self.arquivo)
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro", f"Erro ao ler a planilha:\n{e}", parent=parent_widget)
            self.win.destroy()
            return
        self.d = core.preparar(self.bruto, self.ctx)
        cfg = core.ler_config_cliente(self.arquivo)
        self.ordem, _ = core.ordem_efetiva(self.d, cfg.get("ordem"), self.ctx)
        self.quinzenas = core.quinzenas_do_cliente(self.bruto)

        fr = ttk.Frame(self.win, padding=10)
        fr.pack(fill="both", expand=True)
        topo = ttk.LabelFrame(fr, text="Marco de data")
        topo.pack(fill="x", pady=(0, 8))
        linha = ttk.Frame(topo)
        linha.pack(fill="x", padx=10, pady=8)
        ttk.Label(linha, text="Considerar pendência a partir da quinzena:").pack(side="left")
        datas = [core.fmt_data(q) for q in self.quinzenas]
        if core.fmt_data(self.ctx.marco) not in datas:
            datas = sorted(datas + [core.fmt_data(self.ctx.marco)], key=lambda s: pd.Timestamp(pd.to_datetime(s, dayfirst=True)))
        self.cmb_marco = ttk.Combobox(linha, width=12, state="readonly", values=datas)
        self.cmb_marco.set(core.fmt_data(self.ctx.marco))
        self.cmb_marco.pack(side="left", padx=8)
        self.cmb_marco.bind("<<ComboboxSelected>>", lambda e: self._efeito())
        self.lbl_efeito = ttk.Label(topo, text="", foreground="#333", wraplength=840, justify="left")
        self.lbl_efeito.pack(anchor="w", padx=10, pady=(0, 8))
        ttk.Label(topo, text="Lançamentos sem etapa com data de relatório ANTERIOR ao marco não são pendência: permanecem como estão e saem no bloco "
                             "'Anteriores ao controle por etapa'. Para regularizar períodos antigos aos poucos, recue o marco.",
                  foreground="#555", wraplength=840, justify="left").pack(anchor="w", padx=10, pady=(0, 8))

        meio = ttk.LabelFrame(fr, text="Ordem das etapas no relatório")
        meio.pack(fill="both", expand=True)
        cols = ("ordem", "etapa", "mediana", "prazo", "n", "valor")
        box = ttk.Frame(meio)
        box.pack(side="left", fill="both", expand=True, padx=8, pady=8)
        self.tv = ttk.Treeview(box, columns=cols, show="headings", selectmode="browse")
        for c, t, w, a in (("ordem", "Ordem", 60, "center"), ("etapa", "Etapa", 220, "w"), ("mediana", "Data mediana", 105, "center"),
                           ("prazo", "Prazo (1ª a última quinzena)", 215, "w"), ("n", "Lanç.", 55, "center"), ("valor", "Valor (R$)", 105, "e")):
            self.tv.heading(c, text=t)
            self.tv.column(c, width=w, anchor=a)
        self.tv.pack(side="left", fill="both", expand=True)
        bts = ttk.Frame(meio)
        bts.pack(side="right", fill="y", padx=8, pady=8)
        ttk.Button(bts, text="Subir", width=18, command=lambda: self._mover(-1)).pack(pady=2)
        ttk.Button(bts, text="Descer", width=18, command=lambda: self._mover(1)).pack(pady=2)
        ttk.Button(bts, text="Restaurar ordem\ncronológica", width=18, command=self._restaurar).pack(pady=(14, 2))
        rod = ttk.Frame(fr)
        rod.pack(fill="x", pady=(8, 0))
        ttk.Button(rod, text="Fechar", command=self.win.destroy).pack(side="right", padx=4)
        ttk.Button(rod, text="Salvar", command=self._salvar, style="Accentuated.TButton").pack(side="right", padx=4)
        self._popular()
        self._efeito()
        self.win.grab_set()

    def _popular(self, selecionar=None):
        self.tv.delete(*self.tv.get_children())
        e = self.d[self.d["DESTINO"] == core.D_ETAPA]
        for i, nome in enumerate(self.ordem, 1):
            g = e[e["ETAPA_FINAL"] == nome]
            s = g["DATA_REL"].dropna().sort_values()
            mediana = core.fmt_data(s.iloc[len(s) // 2]) if len(s) else ""
            prazo = core.fmt_data(s.iloc[0]) if len(s) and s.iloc[0] == s.iloc[-1] else (
                f"{core.fmt_data(s.iloc[0])} a {core.fmt_data(s.iloc[-1])}" if len(s) else "")
            self.tv.insert("", "end", iid=nome, values=(i, nome, mediana, prazo, len(g), core.fmt_moeda(g["VALOR"].sum())))
        if selecionar and selecionar in self.ordem:
            self.tv.selection_set(selecionar)
            self.tv.see(selecionar)

    def _mover(self, delta):
        sel = self.tv.selection()
        if not sel:
            return
        nome = sel[0]
        i = self.ordem.index(nome)
        j = i + delta
        if 0 <= j < len(self.ordem):
            self.ordem[i], self.ordem[j] = self.ordem[j], self.ordem[i]
            self._popular(nome)

    def _restaurar(self):
        self.ordem = core.ordem_cronologica(self.d)
        self._popular()

    def _efeito(self):
        m = core.parse_data(self.cmb_marco.get())
        ef = core.efeito_do_marco(self.bruto, self.ctx, m)
        self.lbl_efeito.config(text=f"Com este marco: {ef['n_ant']} lançamentos anteriores ({_moeda(ef['v_ant'])}) ficam fora da fila; "
                                    f"{ef['n_fila']} lançamentos ({_moeda(ef['v_fila'])}) ficam na fila de regularização.")

    def _salvar(self):
        try:
            backup = core.salvar_config_cliente(self.arquivo, self.ordem, core.parse_data(self.cmb_marco.get()))
        except core.ArquivoEmUsoError as e:
            messagebox.showerror("Planilha em uso", f"{e}\n\nNada foi gravado.", parent=self.win)
            return
        except core.PlanilhaAlteradaError as e:
            messagebox.showerror("Planilha alterada", f"{e}", parent=self.win)
            return
        except Exception as e:  # noqa: BLE001
            logger.error(f"Erro ao salvar ordem/marco: {e}", exc_info=True)
            messagebox.showerror("Erro", f"Erro ao salvar. A planilha original não foi alterada.\n\n{e}", parent=self.win)
            return
        messagebox.showinfo("Salvo", f"Ordem e marco salvos na aba {core.ABA_ETAPAS}.\n\nBackup: {backup}", parent=self.win)
        self.win.destroy()
        if self.on_saved:
            self.on_saved()


def abrir_ordem_e_marco(parent_widget, nome_cliente: str, on_saved: Optional[Callable] = None):
    """Abre o diálogo de ordem e marco para um cliente (usado pelo painel em Relatórios)."""
    arq = arquivo_do_cliente(nome_cliente)
    if not arq.exists():
        messagebox.showerror("Erro", f"Arquivo não encontrado:\n{arq}", parent=parent_widget)
        return None
    return DialogoOrdemMarco(parent_widget, arq, on_saved)


def main():                                                    # pragma: no cover
    app = RegularizacaoEtapas()
    app.root.mainloop()


if __name__ == "__main__":                                     # pragma: no cover
    main()
