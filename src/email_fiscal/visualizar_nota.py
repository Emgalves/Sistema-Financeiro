# -*- coding: utf-8 -*-
"""
visualizar_nota.py - Mostra NF-e / NFS-e (XML) de forma legível.

O XML é convertido numa página HTML simples, no formato de um DANFE resumido
(emitente, destinatário, itens, totais, parcelas), salva na pasta temporária
do Windows e aberta no navegador padrão. O XML original não é alterado; a
página tem um link para ele. Qualquer outro arquivo (PDF etc.) abre no
programa padrão do Windows, como antes.
"""
VERSAO = "2026-11-12.9"   # igual em todos os arquivos do módulo

import html
import os
import tempfile
import webbrowser
from pathlib import Path

from .leitor_nfe import (carregar_raiz, identificar_xml, _achar, _achar_todos, _txt, _num,
                         _data, _so_digitos, _endereco)


# ---------------------------------------------------------------------------
# Formatação
# ---------------------------------------------------------------------------
def _e(x):
    return html.escape(str(x or ""))


def _moeda(v):
    v = _num(v)
    if v is None:
        return ""
    return f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def _qtd(v):
    try:
        x = float(str(v).replace(",", "."))
    except (TypeError, ValueError):
        return ""
    s = f"{x:,.4f}".rstrip("0").rstrip(".")
    return s.replace(",", "X").replace(".", ",").replace("X", ".")


def _doc_fmt(d):
    d = _so_digitos(d)
    if len(d) == 14:
        return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}"
    if len(d) == 11:
        return f"{d[:3]}.{d[3:6]}.{d[6:9]}-{d[9:]}"
    return d


def _dt(t):
    d = _data(t)
    return d.strftime("%d/%m/%Y") if d else ""


def _bloco_doc(b):
    for tag in ("CNPJ", "CPF", "Cnpj", "Cpf"):
        v = _txt(b, tag)
        if v:
            return _doc_fmt(v)
    return ""


def _pessoa(raiz, bloco, enderecos):
    b = _achar(raiz, bloco)
    if b is None:
        return None
    return {
        "nome": _txt(b, "xNome") or _txt(b, "RazaoSocial"),
        "fantasia": _txt(b, "xFant") or _txt(b, "NomeFantasia"),
        "doc": _bloco_doc(b),
        "ie": _txt(b, "IE"),
        "endereco": _endereco(b, *enderecos),
        "fone": _txt(b, "fone") or _txt(b, "Telefone"),
        "email": _txt(b, "email") or _txt(b, "Email"),
    }


CSS = """
body{font-family:Arial,Helvetica,sans-serif;font-size:13px;color:#111;background:#f4f4f4;margin:0;padding:20px}
.folha{max-width:1000px;margin:0 auto;background:#fff;padding:24px 28px;border:1px solid #ccc}
h1{font-size:18px;margin:0 0 4px}
.sub{color:#555;margin-bottom:14px}
.caixa{border:1px solid #999;border-radius:4px;padding:8px 10px;margin-bottom:10px}
.caixa h2{font-size:11px;text-transform:uppercase;letter-spacing:.04em;color:#555;margin:0 0 4px}
.grade{display:grid;grid-template-columns:1fr 1fr;gap:10px}
table{width:100%;border-collapse:collapse;margin-top:4px}
th,td{border:1px solid #bbb;padding:4px 6px;vertical-align:top}
th{background:#eee;font-size:11px;text-transform:uppercase}
td.n,th.n{text-align:right;white-space:nowrap}
.total{font-size:16px;font-weight:bold}
.status-ok{color:#0a6b1f;font-weight:bold}.status-x{color:#b00020;font-weight:bold}
.chave{font-family:Consolas,monospace;letter-spacing:.05em}
.rodape{margin-top:14px;color:#666;font-size:11px}
@media print{body{background:#fff;padding:0}.folha{border:none}.rodape a{display:none}}
"""


def _caixa_pessoa(titulo, p):
    if not p:
        return ""
    linhas = [f"<strong>{_e(p['nome'])}</strong>"]
    if p.get("fantasia") and p["fantasia"] != p["nome"]:
        linhas.append(_e(p["fantasia"]))
    ident = " &nbsp;|&nbsp; ".join(x for x in (
        f"CNPJ/CPF: {_e(p['doc'])}" if p["doc"] else "",
        f"IE: {_e(p['ie'])}" if p.get("ie") else "") if x)
    if ident:
        linhas.append(ident)
    if p.get("endereco"):
        linhas.append(_e(p["endereco"]))
    contato = " &nbsp;|&nbsp; ".join(x for x in (
        f"Fone: {_e(p['fone'])}" if p.get("fone") else "", _e(p.get("email"))) if x)
    if contato:
        linhas.append(contato)
    return f'<div class="caixa"><h2>{titulo}</h2>{"<br>".join(linhas)}</div>'


# ---------------------------------------------------------------------------
# NF-e
# ---------------------------------------------------------------------------
def _html_nfe(raiz):
    inf = _achar(raiz, "infNFe")
    chave = _so_digitos(inf.get("Id", "")) if inf is not None else ""
    emit = _pessoa(raiz, "emit", ("enderEmit",))
    dest = _pessoa(raiz, "dest", ("enderDest",))
    c_stat, motivo = _txt(raiz, "protNFe", "cStat"), _txt(raiz, "protNFe", "xMotivo")
    if c_stat in ("100", "150"):
        status = f'<span class="status-ok">Autorizada</span> — protocolo {_e(_txt(raiz, "protNFe", "nProt"))}'
    elif c_stat:
        status = f'<span class="status-x">{_e(c_stat)} - {_e(motivo)}</span>'
    else:
        status = '<span class="status-x">XML sem protocolo de autorização</span>'

    itens = []
    for det in _achar_todos(raiz, "det"):
        itens.append("<tr>" + "".join([
            f"<td>{_e(_txt(det, 'prod', 'cProd'))}</td>",
            f"<td>{_e(_txt(det, 'prod', 'xProd'))}</td>",
            f"<td>{_e(_txt(det, 'prod', 'uCom'))}</td>",
            f"<td class='n'>{_qtd(_txt(det, 'prod', 'qCom'))}</td>",
            f"<td class='n'>{_moeda(_txt(det, 'prod', 'vUnCom'))}</td>",
            f"<td class='n'>{_moeda(_txt(det, 'prod', 'vProd'))}</td>"]) + "</tr>")

    tot = _achar(raiz, "ICMSTot")
    def t(tag):
        return _moeda(_txt(tot, tag)) if tot is not None else ""
    totais = [("Produtos", t("vProd")), ("Frete", t("vFrete")), ("Seguro", t("vSeg")),
              ("Desconto", t("vDesc")), ("Outras despesas", t("vOutro")), ("IPI", t("vIPI")),
              ("ICMS ST", t("vST"))]
    totais_html = "".join(f"<tr><td>{n}</td><td class='n'>{v}</td></tr>"
                          for n, v in totais if v and v != "0,00")

    dups = "".join(f"<tr><td>{_e(_txt(d, 'nDup'))}</td><td>{_dt(_txt(d, 'dVenc'))}</td>"
                   f"<td class='n'>{_moeda(_txt(d, 'vDup'))}</td></tr>" for d in _achar_todos(raiz, "dup"))
    pag = {"01": "Dinheiro", "02": "Cheque", "03": "Cartão de crédito", "04": "Cartão de débito",
           "05": "Crédito loja", "15": "Boleto bancário", "16": "Depósito bancário", "17": "PIX",
           "18": "Transferência", "90": "Sem pagamento", "91": "Pagamento posterior", "99": "Outros"}
    formas = ", ".join(pag.get(_txt(p, 'tPag'), _txt(p, 'tPag'))
                       + (f" R$ {_moeda(_txt(p, 'vPag'))}" if _num(_txt(p, 'vPag')) else "")
                       for p in _achar_todos(raiz, "detPag"))

    return f"""
<h1>NF-e nº {_e(_txt(raiz, 'ide', 'nNF'))} &nbsp; série {_e(_txt(raiz, 'ide', 'serie'))}</h1>
<div class="sub">Emissão {_dt(_txt(raiz, 'ide', 'dhEmi') or _txt(raiz, 'ide', 'dEmi'))} &nbsp;|&nbsp;
{_e(_txt(raiz, 'ide', 'natOp'))} &nbsp;|&nbsp; {status}</div>
<div class="caixa"><h2>Chave de acesso</h2><span class="chave">{_e(' '.join(chave[i:i+4] for i in range(0, len(chave), 4)))}</span></div>
<div class="grade">{_caixa_pessoa('Emitente (fornecedor)', emit)}{_caixa_pessoa('Destinatário', dest)}</div>
<div class="caixa"><h2>Produtos ({len(itens)} itens)</h2>
<table><tr><th>Código</th><th>Descrição</th><th>Un</th><th class="n">Qtd</th><th class="n">V. unit.</th><th class="n">V. total</th></tr>
{''.join(itens)}</table></div>
<div class="grade">
<div class="caixa"><h2>Totais</h2><table>{totais_html}
<tr><td class="total">TOTAL DA NOTA</td><td class="n total">{t('vNF')}</td></tr></table></div>
<div class="caixa"><h2>Cobrança</h2>
{('<table><tr><th>Parcela</th><th>Vencimento</th><th class="n">Valor</th></tr>' + dups + '</table>') if dups else 'Nota sem parcelas (duplicatas) informadas.'}
{f'<p>Forma de pagamento: {_e(formas)}</p>' if formas else ''}</div>
</div>
{f'<div class="caixa"><h2>Informações complementares</h2>{_e(_txt(raiz, "infAdic", "infCpl"))}</div>' if _txt(raiz, "infAdic", "infCpl") else ''}
"""


# ---------------------------------------------------------------------------
# NFS-e (padrão nacional e ABRASF)
# ---------------------------------------------------------------------------
def _html_nfse(raiz):
    nacional = _achar(raiz, "infNFSe") is not None or _achar(raiz, "infDPS") is not None
    if nacional:
        prest = _pessoa(raiz, "emit", ("enderNac", "end")) or _pessoa(raiz, "prest", ("end",))
        toma = _pessoa(raiz, "toma", ("end", "endNac"))
        numero = _txt(raiz, "nNFSe") or _txt(raiz, "nDPS")
        emissao = _dt(_txt(raiz, "dhEmi") or _txt(raiz, "dhProc"))
        compet = _dt(_txt(raiz, "dCompet"))
        descricao = _txt(raiz, "xDescServ")
        valores = [("Valor do serviço", _txt(raiz, "vServ")), ("Desconto incondicionado", _txt(raiz, "vDescIncond")),
                   ("ISSQN", _txt(raiz, "vISSQN")), ("Retenções", _txt(raiz, "vTotalRet"))]
        liquido = _txt(raiz, "vLiq") or _txt(raiz, "vServ")
        inf = _achar(raiz, "infNFSe")
        chave = _so_digitos(inf.get("Id", "")) if inf is not None else ""
    else:
        prest = _pessoa(raiz, "PrestadorServico", ("Endereco",)) or _pessoa(raiz, "Prestador", ("Endereco",))
        toma = _pessoa(raiz, "TomadorServico", ("Endereco",)) or _pessoa(raiz, "Tomador", ("Endereco",))
        numero = _txt(raiz, "InfNfse", "Numero") or _txt(raiz, "Numero")
        emissao = _dt(_txt(raiz, "DataEmissao"))
        compet = _dt(_txt(raiz, "Competencia"))
        descricao = _txt(raiz, "Discriminacao")
        valores = [("Valor do serviço", _txt(raiz, "ValorServicos")), ("Deduções", _txt(raiz, "ValorDeducoes")),
                   ("ISS", _txt(raiz, "ValorIss")), ("ISS retido", _txt(raiz, "ValorIssRetido"))]
        liquido = _txt(raiz, "ValorLiquidoNfse") or _txt(raiz, "ValorServicos")
        chave = _txt(raiz, "CodigoVerificacao")
    linhas = "".join(f"<tr><td>{n}</td><td class='n'>{_moeda(v)}</td></tr>" for n, v in valores if _num(v))
    return f"""
<h1>NFS-e nº {_e(numero)}</h1>
<div class="sub">Emissão {emissao}{f' &nbsp;|&nbsp; Competência {compet}' if compet else ''}</div>
{f'<div class="caixa"><h2>Chave / código de verificação</h2><span class="chave">{_e(chave)}</span></div>' if chave else ''}
<div class="grade">{_caixa_pessoa('Prestador (fornecedor)', prest)}{_caixa_pessoa('Tomador', toma)}</div>
<div class="caixa"><h2>Serviço</h2>{_e(descricao).replace(chr(10), '<br>')}</div>
<div class="caixa"><h2>Valores</h2><table>{linhas}
<tr><td class="total">VALOR LÍQUIDO</td><td class="n total">{_moeda(liquido)}</td></tr></table></div>
"""


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def gerar_html(conteudo: bytes, caminho_original: str = ""):
    """HTML legível da nota, ou None se o XML não for NF-e/NFS-e."""
    raiz = carregar_raiz(conteudo)
    tipo = identificar_xml(raiz)
    if tipo == "NFE":
        corpo = _html_nfe(raiz)
    elif tipo == "NFSE":
        corpo = _html_nfse(raiz)
    else:
        return None
    link = ""
    if caminho_original:
        link = f' &nbsp;|&nbsp; <a href="{_e(Path(caminho_original).as_uri())}">ver o XML original</a>'
    titulo = "NFS-e" if tipo == "NFSE" else "NF-e"
    return (f'<!DOCTYPE html><html lang="pt-BR"><head><meta charset="utf-8"><title>{titulo}</title>'
            f'<style>{CSS}</style></head><body><div class="folha">{corpo}'
            f'<div class="rodape">Visualização gerada pelo Sistema Financeiro a partir do XML da nota '
            f'(documento auxiliar, sem valor fiscal){link}</div></div></body></html>')


def abrir_documento(caminho):
    """Abre o arquivo: XML de nota -> página legível no navegador; qualquer outro -> programa padrão."""
    caminho = str(caminho)
    if caminho.lower().endswith(".xml"):
        try:
            with open(caminho, "rb") as f:
                pagina = gerar_html(f.read(), caminho)
            if pagina:
                pasta = Path(tempfile.gettempdir()) / "SistemaFinanceiro_notas"
                pasta.mkdir(exist_ok=True)
                destino = pasta / (Path(caminho).stem[:120] + ".html")
                destino.write_text(pagina, encoding="utf-8")
                webbrowser.open(destino.as_uri())
                return
        except Exception:
            pass  # se algo der errado, abre o XML como antes
    os.startfile(caminho)
