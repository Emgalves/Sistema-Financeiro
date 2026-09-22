# -*- coding: utf-8 -*-
"""
relatorio_por_etapa.py
======================
Gera o PDF "Relatório por Etapa da Obra" (A4 paisagem, cabeçalho com logo3.png, tabelas cinza,
numeração n/N no rodapé), seguindo o padrão do relatorio_despesas_aprimorado.

Todas as regras vêm de src/etapas_obra_core.py. Este módulo só monta o documento.
O gráfico é desenhado com reportlab (sem matplotlib), para não criar dependência nova no executável.

Uso pela interface:
    from src.relatorio_por_etapa import gerar_relatorio_por_etapa
    caminho, nome, dados = gerar_relatorio_por_etapa(arquivo_cliente, posicao)
"""

import logging
import os
import re
import sys
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
from dateutil.relativedelta import relativedelta
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.pdfgen import canvas
from reportlab.platypus import Image, KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle
from reportlab.lib.utils import ImageReader

from src import etapas_obra_core as core

logger = logging.getLogger(__name__)

COR_LARANJA = colors.HexColor("#F08300")
COR_CINZA = colors.HexColor("#575756")
COR_ALERTA_BG = colors.HexColor("#FDEBD0")
COR_ALERTA_TX = colors.HexColor("#8A4B00")
COR_TEXTO = colors.HexColor("#333333")

TEXTO_EMPRESA = [
    "Rua Zodiaco, 87 Sala 07 – Santa Lúcia - Belo Horizonte - MG",
    "(31) 3654-6616 / (31) 99974-1241 / (31) 98711-1139",
    "rvr.engenharia@gmail.com",
]
CORES_CATEGORIA = {
    "MO": "#575756", "MAT": "#F08300", "LOC": "#F7B267", "SERV": "#9A9A98",
    "DIV": "#D4D4D2", "ADM": "#8A4B00", "TAX": "#5B7C99", "TP": "#4A90A4",
}
ESTILOS = getSampleStyleSheet()


# ---------------------------------------------------------------------------
# Localização do logo e dados do cliente
# ---------------------------------------------------------------------------
def localizar_logo(logo_path: Optional[str] = None) -> Optional[str]:
    """Procura logo3.png nos mesmos lugares que o sistema usa (_MEIPASS, pasta atual, pasta do módulo)."""
    aqui = Path(__file__).resolve().parent
    candidatos = [logo_path]
    if getattr(sys, "_MEIPASS", None):
        candidatos.append(os.path.join(sys._MEIPASS, "logo3.png"))
    candidatos += [
        os.path.join(os.getcwd(), "logo3.png"), str(aqui / "logo3.png"), str(aqui.parent / "logo3.png"),
        str(aqui.parent / "imagens" / "logo3.png"), str(aqui.parent / "assets" / "logo3.png"),
    ]
    for c in candidatos:
        if c and os.path.exists(c):
            return c
    logger.warning("logo3.png não encontrado; o relatório será gerado sem logo.")
    return None


def obter_dados_cliente(arquivo_cliente) -> Tuple[str, str, str]:
    """(nome, endereço, CNO) do cliente, lidos de clientes.xlsx; fallback: nome do arquivo."""
    nome_arquivo = Path(arquivo_cliente).stem
    try:
        try:
            from src.config.config import ARQUIVO_CLIENTES
        except ImportError:
            from config.config import ARQUIVO_CLIENTES
        df = pd.read_excel(ARQUIVO_CLIENTES, sheet_name="Clientes")
        linha = df[df["Nome"] == nome_arquivo]
        if not linha.empty:
            r = linha.iloc[0]
            col_end = next((c for c in df.columns if "nder" in str(c).lower()), None)
            end = str(r[col_end]).strip() if col_end and pd.notna(r[col_end]) else ""
            cno = str(r["CNO"]).strip() if "CNO" in df.columns and pd.notna(r["CNO"]) else ""
            return str(r["Nome"]), end, cno
    except Exception as e:  # noqa: BLE001
        logger.warning(f"Não foi possível ler clientes.xlsx: {e}")
    return nome_arquivo, "", ""


# ---------------------------------------------------------------------------
# Estilos e helpers
# ---------------------------------------------------------------------------
def _p(txt, size=7.5, bold=False, align=0, color=colors.black, leading=None):
    return Paragraph(str(txt), ParagraphStyle(
        "x", parent=ESTILOS["Normal"], fontName="Helvetica-Bold" if bold else "Helvetica",
        fontSize=size, leading=leading or size + 2, alignment=align, textColor=color))


def _prazo(a, b) -> str:
    return core.fmt_data(a) if pd.Timestamp(a) == pd.Timestamp(b) else f"{core.fmt_data(a)} a {core.fmt_data(b)}"


def _meses(inicio, posicao) -> str:
    rd = relativedelta(pd.Timestamp(posicao), pd.Timestamp(inicio))
    m = rd.years * 12 + rd.months
    return "menos de 1 mês" if m <= 0 else ("1 mês" if m == 1 else f"{m} meses")


def _plural(n, singular, plural) -> str:
    return f"{n} {singular if n == 1 else plural}"


class _NumberedCanvas(canvas.Canvas):
    """Numeração n/N no canto inferior direito (mesmo padrão do relatório de despesas)."""

    def __init__(self, *a, **k):
        super().__init__(*a, **k)
        self._saved = []

    def showPage(self):
        self._saved.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        n = len(self._saved)
        for s in self._saved:
            self.__dict__.update(s)
            self.setFont("Helvetica", 9)
            self.drawRightString(A4[1] - 30, 18, f"{self._pageNumber}/{n}")
            super().showPage()
        super().save()


# ---------------------------------------------------------------------------
# Blocos do documento
# ---------------------------------------------------------------------------
def _cabecalho(dados: core.DadosRelatorio, nome, endereco, cno, logo) -> list:
    ex = []
    est_emp = ParagraphStyle("e", parent=ESTILOS["Normal"], fontSize=10, leading=12, alignment=2)
    if logo:
        w, h_px = 170, None
        try:
            iw, ih = ImageReader(logo).getSize()
            h_px = w * ih / iw
        except Exception:  # noqa: BLE001
            h_px = 85
        t = Table([[Image(logo, width=w, height=h_px), [Paragraph(l, est_emp) for l in TEXTO_EMPRESA]]], colWidths=[200, 582])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                               ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
        ex += [t, Spacer(1, 10)]
    else:
        ex += [Paragraph(l, est_emp) for l in TEXTO_EMPRESA] + [Spacer(1, 40)]
    n = ParagraphStyle("n", parent=ESTILOS["Normal"], fontName="Helvetica-Bold", fontSize=14, leading=16)
    i = ParagraphStyle("i", parent=ESTILOS["Normal"], fontSize=10, leading=12)
    ri = ParagraphStyle("ri", parent=i, alignment=2)
    linhas = [
        [Paragraph(nome, n), Paragraph("Relatório por Etapa da Obra", ParagraphStyle("rt", parent=ri, fontName="Helvetica-Bold", fontSize=11))],
        [Paragraph(endereco or "ENDEREÇO NÃO INFORMADO", i), Paragraph(f"Posição em: {core.fmt_data(dados.posicao)}", ri)],
        [Paragraph(f"CNO {cno}" if cno else "", i),
         Paragraph(f"Início da obra: {core.fmt_data(dados.inicio)}  ·  Tempo decorrido: {_meses(dados.inicio, dados.posicao)}", ri)],
    ]
    t2 = Table(linhas, colWidths=[470, 312])
    t2.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                            ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    return ex + [t2, Spacer(1, 10)]


def _quadro_resumo(dados: core.DadosRelatorio) -> Table:
    total = dados.total or 1.0
    cab = ["Nº", "Etapa", "Prazo (1ª a última quinzena)", "Mão de obra", "Material", "Locação", "Serviços", "Diversos",
           "Adm./Taxas/<br/>Tarifas", "Total (R$)", "% obra"]
    rows = [[_p(c, 7, True, 1, colors.whitesmoke) for c in cab]]
    estilos = []

    def linha(num, nome, prazo, cats, adm, tot, bold=False, bg=None, cor=colors.black):
        cel = [_p(num, 7.5, bold, 1, cor), _p(nome, 7.5, bold, 0, cor), _p(prazo, 7, False, 1, cor)]
        for c in core.CATS_ETAPA:
            v = cats.get(c, 0.0) if cats else 0.0
            cel.append(_p(core.fmt_moeda(v) if v else "-", 7.5, bold, 2, cor))
        cel += [_p(core.fmt_moeda(adm) if adm else "-", 7.5, bold, 2, cor), _p(core.fmt_moeda(tot), 7.5, True, 2, cor),
                _p(core.fmt_pct(tot / total), 7.5, bold, 2, cor)]
        rows.append(cel)
        if bg is not None:
            estilos.append(("BACKGROUND", (0, len(rows) - 1), (-1, len(rows) - 1), bg))

    for i, b in enumerate(dados.etapas, 1):
        linha(i, b.nome, _prazo(b.primeira, b.ultima), b.cats, 0, b.total)

    def soma_cats(*blocos):
        return {c: sum(b.cats.get(c, 0.0) for b in blocos) for c in core.CATS_ETAPA}

    def adm(b):
        return sum(b.cats.get(c, 0.0) for c in core.CATS_REGRA)

    rg, mc, se, an = dados.regra, dados.marcados, dados.sem_etapa, dados.anteriores
    linha("", "CUSTOS GERAIS DA OBRA", "", soma_cats(rg, mc, se), adm(rg), dados.total_gerais, True, colors.lightgrey)
    linha("", "&nbsp;&nbsp;&nbsp;Taxa, administração e tarifas públicas (regra)", "", rg.cats, adm(rg), rg.total, False, colors.whitesmoke)
    linha("", "&nbsp;&nbsp;&nbsp;Marcados como CUSTOS GERAIS DA OBRA", "", mc.cats, 0, mc.total, False, colors.whitesmoke)
    linha("", "&nbsp;&nbsp;&nbsp;SEM ETAPA INFORMADA (pendente de classificação)", "", se.cats, 0, se.total, True, COR_ALERTA_BG, COR_ALERTA_TX)
    if an.n:
        linha("", f"ANTERIORES AO CONTROLE POR ETAPA (antes de {core.fmt_data(dados.marco)})", "", an.cats, 0, an.total, True, colors.lightgrey)
    tot_cats = {c: sum(b.cats.get(c, 0.0) for b in dados.etapas) + soma_cats(rg, mc, se, an)[c] for c in core.CATS_ETAPA}
    linha("", "TOTAL DA OBRA", "", tot_cats, adm(rg), dados.total, True, colors.grey, colors.white)

    t = Table(rows, colWidths=[22, 150, 100, 70, 70, 60, 70, 55, 70, 75, 40], repeatRows=1)
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.grey), ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("TOPPADDING", (0, 0), (-1, -1), 3),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 3)] + estilos))
    return t


def _grafico(dados: core.DadosRelatorio, largura=760, altura=330) -> Drawing:
    """Colunas 100% empilhadas: categorias por etapa (mais custos gerais, sem etapa e anteriores)."""
    barras = [(b.nome, b.cats) for b in dados.etapas]
    barras.append(("GERAIS (regra)", dados.regra.cats))
    if dados.marcados.n:
        barras.append(("GERAIS (marcados)", dados.marcados.cats))
    barras.append(("SEM ETAPA INFORMADA", dados.sem_etapa.cats))
    if dados.anteriores.n:
        barras.append(("ANTERIORES AO CONTROLE", dados.anteriores.cats))
    barras = [(n, c) for n, c in barras if sum(c.values()) > 0]

    d = Drawing(largura, altura)
    if not barras:
        return d
    series = [s for s in core.CATS_TODAS if sum(c.get(s, 0.0) for _, c in barras) > 0]
    dados_series = []
    for s in series:
        dados_series.append([100.0 * c.get(s, 0.0) / sum(c.values()) for _, c in barras])
    ch = VerticalBarChart()
    ch.x, ch.y, ch.width, ch.height = 40, 110, largura - 55, altura - 130
    ch.data = dados_series
    ch.categoryAxis.categoryNames = [n for n, _ in barras]
    ch.categoryAxis.style = "stacked"
    ch.categoryAxis.labels.angle = 30
    ch.categoryAxis.labels.boxAnchor = "ne"
    ch.categoryAxis.labels.dx = 4
    ch.categoryAxis.labels.dy = -3
    ch.categoryAxis.labels.fontSize = 6.5
    ch.categoryAxis.labels.fontName = "Helvetica"
    ch.valueAxis.valueMin, ch.valueAxis.valueMax, ch.valueAxis.valueStep = 0, 100, 20
    ch.valueAxis.labelTextFormat = "%d%%"
    ch.valueAxis.labels.fontSize = 7
    ch.valueAxis.labels.fontName = "Helvetica"
    ch.barWidth = 22
    ch.groupSpacing = 8
    for i, s in enumerate(series):
        ch.bars[i].fillColor = colors.HexColor(CORES_CATEGORIA[s])
        ch.bars[i].strokeColor = colors.white
        ch.bars[i].strokeWidth = 0.4
    d.add(ch)
    x = 40
    for s in series:                                  # legenda
        d.add(Rect(x, 12, 8, 8, fillColor=colors.HexColor(CORES_CATEGORIA[s]), strokeColor=colors.white))
        rotulo = core.NOMES_CATEGORIA[s]
        d.add(String(x + 12, 13, rotulo, fontName="Helvetica", fontSize=7))
        x += 22 + 4.2 * len(rotulo)
    return d


def _tabela_esquerda(b: core.BlocoEtapa) -> Table:
    linhas, est = [], []

    def add(nome, v, bold=False, indent=0):
        if not v:
            return
        linhas.append([_p("&nbsp;" * indent + nome, 7.5, bold), _p(core.fmt_moeda(v), 7.5, bold, 2),
                       _p(core.fmt_pct(v / b.total) if b.total else "", 7.5, bold, 2)])

    add("MÃO DE OBRA", b.cats["MO"], True)
    for s in core.SUBGRUPOS_MO:
        add(s.capitalize(), b.mo_sub.get(s, 0.0), False, 6)
    for k in ("MAT", "LOC", "SERV", "DIV"):
        add(core.NOMES_CATEGORIA[k].upper(), b.cats[k], True)
    linhas.append([_p("TOTAL DA ETAPA", 7.5, True), _p(core.fmt_moeda(b.total), 7.5, True, 2), _p("100,0%", 7.5, True, 2)])
    est.append(("BACKGROUND", (0, len(linhas) - 1), (-1, len(linhas) - 1), colors.lightgrey))
    t = Table(linhas, colWidths=[200, 100, 60])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey), ("TOPPADDING", (0, 0), (-1, -1), 2),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 2)] + est))
    return t


def _tabela_fornecedores(forn, titulo="Principais fornecedores") -> Table:
    linhas = [[_p(titulo, 7.5, True, 0, colors.whitesmoke), _p("Valor (R$)", 7.5, True, 2, colors.whitesmoke)]]
    for n, v in forn:
        linhas.append([_p(n, 7.5), _p(core.fmt_moeda(v), 7.5, False, 2)])
    if not forn:
        linhas.append([_p("-", 7.5), _p("", 7.5)])
    t = Table(linhas, colWidths=[270, 100], hAlign="LEFT")
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.grey), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                           ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    return t


def _titulo_bloco(titulo, sub, alerta=False) -> Table:
    cor = COR_ALERTA_TX if alerta else colors.white
    t = Table([[_p(titulo, 9, True, 0, cor), _p(sub, 8, False, 2, cor)]], colWidths=[382, 400])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), COR_ALERTA_BG if alerta else COR_CINZA),
                           ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4)]))
    return t


def _bloco_etapa(i, b: core.BlocoEtapa, total_obra) -> KeepTogether:
    sub = (f"Prazo: {_prazo(b.primeira, b.ultima)}   |   R$ {core.fmt_moeda(b.total)}   |   "
           f"{core.fmt_pct(b.total / total_obra if total_obra else 0)} da obra   |   {_plural(b.n, 'lançamento', 'lançamentos')}")
    corpo = Table([[_tabela_esquerda(b), _tabela_fornecedores(b.fornecedores)]], colWidths=[392, 390])
    corpo.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0)]))
    return KeepTogether([_titulo_bloco(f"{i}  {b.nome}", sub), Spacer(1, 3), corpo, Spacer(1, 12)])


def _composicao(bloco: core.BlocoSimples) -> str:
    itens = sorted(((k, v) for k, v in bloco.cats.items() if v), key=lambda t: -t[1])
    return "; ".join(f"{k} {core.fmt_moeda(v)}" for k, v in itens) or "-"


def _bloco_gerais(dados: core.DadosRelatorio) -> KeepTogether:
    total = dados.total or 1.0
    linhas = [[_p(h, 7.5, True, a, colors.whitesmoke) for h, a in
               (("Linha", 0), ("Composição por categoria (R$)", 0), ("Lançamentos", 1), ("Valor (R$)", 2), ("% obra", 2))]]
    partes = [("1. Taxa, administração e tarifas públicas (regra)", dados.regra, False),
              ("2. Marcados como CUSTOS GERAIS DA OBRA", dados.marcados, False),
              ("3. SEM ETAPA INFORMADA (pendente de classificação)", dados.sem_etapa, True)]
    for nome, b, alerta in partes:
        cor = COR_ALERTA_TX if alerta else colors.black
        linhas.append([_p(nome, 7.5, alerta, 0, cor), _p(_composicao(b), 7, False, 0, cor), _p(b.n, 7.5, False, 1, cor),
                       _p(core.fmt_moeda(b.total), 7.5, True, 2, cor), _p(core.fmt_pct(b.total / total), 7.5, False, 2, cor)])
    n_total = dados.regra.n + dados.marcados.n + dados.sem_etapa.n
    linhas.append([_p("TOTAL CUSTOS GERAIS DA OBRA", 7.5, True), _p("", 7), _p(n_total, 7.5, True, 1),
                   _p(core.fmt_moeda(dados.total_gerais), 7.5, True, 2), _p(core.fmt_pct(dados.total_gerais / total), 7.5, True, 2)])
    t = Table(linhas, colWidths=[250, 300, 70, 100, 62])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), colors.grey), ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                           ("BACKGROUND", (0, 3), (-1, 3), COR_ALERTA_BG), ("BACKGROUND", (0, 4), (-1, 4), colors.lightgrey),
                           ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3)]))
    itens = [_titulo_bloco("CUSTOS GERAIS DA OBRA", f"R$ {core.fmt_moeda(dados.total_gerais)}   |   {core.fmt_pct(dados.total_gerais / total)} da obra"),
             Spacer(1, 3), t, Spacer(1, 4)]
    if dados.sem_etapa.n:
        itens += [_p(f"<b>Observação:</b> {_plural(dados.sem_etapa.n, 'lançamento', 'lançamentos')} (R$ {core.fmt_moeda(dados.sem_etapa.total)}) "
                     f"ainda não {'tem' if dados.sem_etapa.n == 1 else 'têm'} etapa informada e "
                     f"{'está' if dados.sem_etapa.n == 1 else 'estão'} em classificação. Após a classificação, "
                     f"passam para a etapa correspondente e saem deste bloco.", 7.5, False, 0, COR_ALERTA_TX, 9.5),
                  Spacer(1, 6), _tabela_fornecedores(dados.sem_etapa.fornecedores, "Principais fornecedores sem etapa informada")]
    itens.append(Spacer(1, 14))
    return KeepTogether(itens)


def _bloco_anteriores(dados: core.DadosRelatorio) -> Optional[KeepTogether]:
    an = dados.anteriores
    if not an.n:
        return None
    total = dados.total or 1.0
    txt = (f"<b>{_plural(an.n, 'lançamento', 'lançamentos')} · R$ {core.fmt_moeda(an.total)} ({core.fmt_pct(an.total / total)} da obra)</b> "
           f"com data de relatório anterior a {core.fmt_data(dados.marco)}, data de início do controle por etapa definida para esta obra. "
           f"Permanecem como estão, sem etapa. Composição: {_composicao(an)}.")
    return KeepTogether([_titulo_bloco(f"ANTERIORES AO CONTROLE POR ETAPA (antes de {core.fmt_data(dados.marco)})",
                                       f"R$ {core.fmt_moeda(an.total)}   |   {core.fmt_pct(an.total / total)} da obra"),
                         Spacer(1, 3), _p(txt, 7.5, False, 0, colors.black, 9.5)])


# ---------------------------------------------------------------------------
# Função pública
# ---------------------------------------------------------------------------
def montar_story(dados: core.DadosRelatorio, arquivo_cliente, logo: Optional[str]) -> list:
    nome, endereco, cno = obter_dados_cliente(arquivo_cliente)
    h1 = ParagraphStyle("h1", parent=ESTILOS["Normal"], fontName="Helvetica-Bold", fontSize=13, leading=15, spaceBefore=4, spaceAfter=6)
    story = _cabecalho(dados, nome, endereco, cno, logo) + [Paragraph("QUADRO-RESUMO POR ETAPA", h1), _quadro_resumo(dados), PageBreak()]
    crit = ("<b>Critérios:</b> mão de obra na etapa em andamento na quinzena; material, locação, serviços e diversos na etapa a que se destinam; "
            "taxa de administração, administração e tarifas públicas em Custos Gerais. Base: lançamentos ativos, todos os tipos de despesa, "
            "até a data de posição. Linhas de etapa sem lançamentos não são impressas.")
    story += [Paragraph("DISTRIBUIÇÃO DE CATEGORIAS POR ETAPA DA OBRA", h1), _grafico(dados),
              _p("Cada coluna soma 100% do valor da respectiva etapa. As colunas de custos gerais, sem etapa informada e anteriores "
                 "aparecem para que todo o valor da obra esteja representado.", 7.5, False, 0, COR_TEXTO),
              Spacer(1, 6), _p(crit, 7.5, False, 0, COR_TEXTO, 9.5), PageBreak()]
    for i, b in enumerate(dados.etapas, 1):
        story.append(_bloco_etapa(i, b, dados.total))
    story.append(_bloco_gerais(dados))
    ant = _bloco_anteriores(dados)
    if ant is not None:
        story.append(ant)
    return story


def nome_pdf(dados: core.DadosRelatorio, arquivo_cliente) -> str:
    nome, _, _ = obter_dados_cliente(arquivo_cliente)
    nome = re.sub(r'[\\/:*?"<>|]', "-", nome).strip()
    return f"REL ETAPAS - {nome} - {pd.Timestamp(dados.posicao).strftime('%d-%m-%Y')}.pdf"


def gerar_relatorio_por_etapa(arquivo_cliente, posicao=None, logo_path: Optional[str] = None,
                              pasta_saida: Optional[str] = None):
    """
    Gera o PDF na pasta do cliente (ou em pasta_saida). Devolve (caminho, nome_do_arquivo, dados).
    Levanta ValueError se a conferência de totais falhar (nenhum PDF é gerado nesse caso).
    """
    arquivo_cliente = str(arquivo_cliente)
    dados = core.montar_relatorio(arquivo_cliente, posicao)            # já confere soma dos blocos = total
    for aviso in dados.avisos:
        logger.warning(aviso)
    logo = localizar_logo(logo_path)
    nome = nome_pdf(dados, arquivo_cliente)
    pasta = pasta_saida or os.path.dirname(os.path.abspath(arquivo_cliente))
    caminho = os.path.join(pasta, nome)
    doc = SimpleDocTemplate(caminho, pagesize=landscape(A4), rightMargin=30, leftMargin=30, topMargin=40, bottomMargin=30,
                            title=f"Relatório por Etapa da Obra - {Path(arquivo_cliente).stem}")
    doc.build(montar_story(dados, arquivo_cliente, logo), canvasmaker=_NumberedCanvas)
    logger.info(f"Relatório por etapa gerado: {caminho}")
    return caminho, nome, dados


if __name__ == "__main__":      # uso rápido: python -m src.relatorio_por_etapa "arquivo.xlsx" [dd/mm/aaaa]
    import sys as _sys
    logging.basicConfig(level=logging.INFO)
    if len(_sys.argv) < 2:
        print(__doc__)
        _sys.exit(1)
    print(gerar_relatorio_por_etapa(_sys.argv[1], _sys.argv[2] if len(_sys.argv) > 2 else None)[0])
