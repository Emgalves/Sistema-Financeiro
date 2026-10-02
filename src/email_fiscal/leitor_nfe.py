# -*- coding: utf-8 -*-
"""
leitor_nfe.py - Lê XML de NF-e (modelo 55/65), NFS-e (padrão nacional e ABRASF)
e eventos de cancelamento de NF-e.

A leitura ignora namespaces (procura pelo "nome local" das tags), porque cada
emissor/prefeitura grava o XML de um jeito ligeiramente diferente.
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import re
import xml.etree.ElementTree as ET
from datetime import datetime, date


# ---------------------------------------------------------------------------
# Utilitários de XML sem namespace
# ---------------------------------------------------------------------------
def _local(tag):
    return tag.rsplit("}", 1)[-1] if isinstance(tag, str) else ""


def _achar(no, *caminho):
    """Procura, em qualquer profundidade, a sequência de tags informada.
    _achar(raiz, 'emit', 'CNPJ') -> primeiro <CNPJ> dentro do primeiro <emit>."""
    atual = [no]
    for nome in caminho:
        prox = None
        for base in atual:
            for el in base.iter():
                if el is not base and _local(el.tag) == nome:
                    prox = el
                    break
            if prox is not None:
                break
        if prox is None:
            return None
        atual = [prox]
    return atual[0]


def _achar_todos(no, nome):
    return [el for el in no.iter() if _local(el.tag) == nome]


def _txt(no, *caminho):
    el = _achar(no, *caminho) if caminho else no
    if el is None or el.text is None:
        return ""
    return el.text.strip()


def _num(texto):
    try:
        return round(float(str(texto).replace(",", ".")), 2)
    except (TypeError, ValueError):
        return None


def _data(texto):
    """Aceita '2025-08-20', '2025-08-20T10:00:00-03:00' e '20/08/2025'."""
    if not texto:
        return None
    texto = texto.strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(texto[:10], fmt).date()
        except ValueError:
            pass
    return None


def _so_digitos(s):
    return re.sub(r"\D", "", str(s or ""))


def _doc(no, bloco):
    """CPF ou CNPJ dentro de um bloco (emit, dest, toma, Prestador...)."""
    b = _achar(no, bloco)
    if b is None:
        return ""
    for tag in ("CNPJ", "CPF", "Cnpj", "Cpf"):
        v = _txt(b, tag)
        if v:
            return _so_digitos(v)
    return ""


def _endereco(no, *blocos):
    """Monta 'LOGRADOURO, NUM - COMPL - BAIRRO - CIDADE/UF - CEP: 00000-000'."""
    for bloco in blocos:
        b = _achar(no, bloco)
        if b is None:
            continue
        g = lambda *t: next((_txt(b, x) for x in t if _txt(b, x)), "")
        lgr, nro, cpl = g("xLgr", "Endereco"), g("nro", "Numero"), g("xCpl", "Complemento")
        bairro, mun, uf = g("xBairro", "Bairro"), g("xMun", "xCidade"), g("UF", "Uf")
        cep = _so_digitos(g("CEP", "Cep"))
        partes = [f"{lgr}, {nro}" if nro else lgr, cpl, bairro,
                  f"{mun}/{uf}" if mun and uf else (mun or uf),
                  f"CEP: {cep[:5]}-{cep[5:]}" if len(cep) == 8 else ""]
        texto = " - ".join(p for p in partes if p)
        if texto:
            return texto.upper()
    return ""


def dados_emitente(raiz):
    """Dados para cadastrar o fornecedor (NF-e ou NFS-e)."""
    for bloco in ("emit", "prest", "PrestadorServico", "Prestador"):
        b = _achar(raiz, bloco)
        if b is None:
            continue
        doc = ""
        for tag in ("CNPJ", "CPF", "Cnpj", "Cpf"):
            if _txt(b, tag):
                doc = _so_digitos(_txt(b, tag))
                break
        return {
            "doc": doc,
            "razao_social": (_txt(b, "xNome") or _txt(b, "RazaoSocial")).upper(),
            "fantasia": (_txt(b, "xFant") or _txt(b, "NomeFantasia")).upper(),
            "telefone": _so_digitos(_txt(b, "fone") or _txt(b, "Telefone")),
            "email": _txt(b, "email") or _txt(b, "Email"),
            "endereco": _endereco(b, "enderEmit", "enderNac", "end", "Endereco"),
        }
    return {}


def ler_emitente_arquivo(caminho):
    """Lê os dados do emitente de um XML já salvo em disco."""
    try:
        with open(caminho, "rb") as f:
            raiz = carregar_raiz(f.read())
        return dados_emitente(raiz) if raiz is not None else {}
    except OSError:
        return {}


# ---------------------------------------------------------------------------
# Identificação do tipo de XML
# ---------------------------------------------------------------------------
def carregar_raiz(conteudo: bytes):
    try:
        return ET.fromstring(conteudo)
    except ET.ParseError:
        # alguns XML chegam com BOM ou lixo antes do '<'
        inicio = conteudo.find(b"<")
        if inicio > 0:
            try:
                return ET.fromstring(conteudo[inicio:])
            except ET.ParseError:
                return None
        return None


def identificar_xml(raiz):
    """Retorna 'NFE', 'NFSE', 'EVENTO_NFE', 'CTE' ou None."""
    if raiz is None:
        return None
    nomes = {_local(el.tag) for el in raiz.iter()}
    if "infNFe" in nomes:
        return "NFE"
    if "infEvento" in nomes and ("chNFe" in nomes):
        return "EVENTO_NFE"
    if "infCte" in nomes:
        return "CTE"
    if nomes & {"infNFSe", "InfNfse", "CompNfse", "infDPS", "Nfse", "NFSe"}:
        return "NFSE"
    return None


# ---------------------------------------------------------------------------
# NF-e
# ---------------------------------------------------------------------------
def ler_nfe(raiz):
    inf = _achar(raiz, "infNFe")
    chave = _so_digitos(inf.get("Id", "")) if inf is not None else ""

    itens = []
    for det in _achar_todos(raiz, "det"):
        itens.append(_txt(det, "prod", "xProd"))

    duplicatas = []
    for dup in _achar_todos(raiz, "dup"):
        duplicatas.append({
            "numero": _txt(dup, "nDup"),
            "vencimento": _data(_txt(dup, "dVenc")),
            "valor": _num(_txt(dup, "vDup")),
        })

    # endereço de entrega / destinatário ajuda a achar a obra
    textos_endereco = []
    for bloco in ("enderDest", "entrega"):
        b = _achar(raiz, bloco)
        if b is not None:
            textos_endereco.append(" ".join(
                (el.text or "").strip() for el in b.iter() if el is not b and el.text))

    c_stat = _txt(raiz, "protNFe", "cStat")

    return {
        "tipo": "NFE",
        "chave": chave,
        "numero": _txt(raiz, "ide", "nNF").lstrip("0") or _txt(raiz, "ide", "nNF"),
        "serie": _txt(raiz, "ide", "serie"),
        "emissao": _data(_txt(raiz, "ide", "dhEmi") or _txt(raiz, "ide", "dEmi")),
        "tipo_operacao": _txt(raiz, "ide", "tpNF"),       # 1 = saída do emitente
        "natureza": _txt(raiz, "ide", "natOp"),
        "emit_doc": _doc(raiz, "emit"),
        "emit_nome": _txt(raiz, "emit", "xNome"),
        "dest_doc": _doc(raiz, "dest"),
        "dest_nome": _txt(raiz, "dest", "xNome"),
        "valor_total": _num(_txt(raiz, "ICMSTot", "vNF")),
        "duplicatas": duplicatas,
        "itens": itens,
        "inf_complementar": _txt(raiz, "infAdic", "infCpl"),
        "endereco_entrega": " | ".join(textos_endereco),
        "autorizada": c_stat in ("100", "150") if c_stat else None,  # None = XML sem protocolo
        "c_stat": c_stat,
    }


def ler_evento_nfe(raiz):
    return {
        "tipo": "EVENTO_NFE",
        "chave": _so_digitos(_txt(raiz, "chNFe")),
        "tipo_evento": _txt(raiz, "tpEvento"),   # 110111 = cancelamento
        "descricao": _txt(raiz, "descEvento"),
    }


# ---------------------------------------------------------------------------
# NFS-e (padrão nacional e ABRASF, os dois mais comuns)
# ---------------------------------------------------------------------------
def ler_nfse(raiz):
    if _achar(raiz, "infNFSe") is not None or _achar(raiz, "infDPS") is not None:
        return _ler_nfse_nacional(raiz)
    return _ler_nfse_abrasf(raiz)


def _ler_nfse_nacional(raiz):
    inf = _achar(raiz, "infNFSe")
    chave = _so_digitos(inf.get("Id", "")) if inf is not None else ""
    valor = (_num(_txt(raiz, "valores", "vLiq"))
             or _num(_txt(raiz, "vServPrest", "vServ"))
             or _num(_txt(raiz, "vServ")))
    return {
        "tipo": "NFSE",
        "chave": chave,
        "numero": _txt(raiz, "nNFSe") or _txt(raiz, "nDPS"),
        "serie": _txt(raiz, "serie"),
        "emissao": _data(_txt(raiz, "dhEmi") or _txt(raiz, "dhProc") or _txt(raiz, "dCompet")),
        "emit_doc": _doc(raiz, "emit") or _doc(raiz, "prest"),
        "emit_nome": _txt(raiz, "emit", "xNome") or _txt(raiz, "prest", "xNome"),
        "dest_doc": _doc(raiz, "toma"),
        "dest_nome": _txt(raiz, "toma", "xNome"),
        "valor_total": valor,
        "duplicatas": [],
        "itens": [_txt(raiz, "xDescServ")] if _txt(raiz, "xDescServ") else [],
        "inf_complementar": _txt(raiz, "xInfComp"),
        "endereco_entrega": "",
        "autorizada": None,
        "c_stat": "",
    }


def _ler_nfse_abrasf(raiz):
    numero = _txt(raiz, "InfNfse", "Numero") or _txt(raiz, "Numero")
    valor = (_num(_txt(raiz, "ValorLiquidoNfse"))
             or _num(_txt(raiz, "ValorServicos")))
    prest = _achar(raiz, "PrestadorServico")
    if prest is None:
        prest = _achar(raiz, "Prestador")
    toma = _achar(raiz, "TomadorServico")
    if toma is None:
        toma = _achar(raiz, "Tomador")

    def doc_de(bloco):
        if bloco is None:
            return ""
        for tag in ("Cnpj", "Cpf", "CNPJ", "CPF"):
            v = _txt(bloco, tag)
            if v:
                return _so_digitos(v)
        return ""

    chave = _txt(raiz, "CodigoVerificacao")
    emit_doc = doc_de(prest)
    return {
        "tipo": "NFSE",
        # ABRASF não tem chave nacional: usa CNPJ + número como identificador
        "chave": f"NFSE-{emit_doc}-{numero}" if numero else chave,
        "numero": numero,
        "serie": "",
        "emissao": _data(_txt(raiz, "DataEmissao")),
        "emit_doc": emit_doc,
        "emit_nome": (_txt(prest, "RazaoSocial") if prest is not None else "")
                     or (_txt(prest, "NomeFantasia") if prest is not None else ""),
        "dest_doc": doc_de(toma),
        "dest_nome": _txt(toma, "RazaoSocial") if toma is not None else "",
        "valor_total": valor,
        "duplicatas": [],
        "itens": [_txt(raiz, "Discriminacao")] if _txt(raiz, "Discriminacao") else [],
        "inf_complementar": _txt(raiz, "OutrasInformacoes"),
        "endereco_entrega": "",
        "autorizada": None,
        "c_stat": "",
    }


# ---------------------------------------------------------------------------
# Ponto de entrada
# ---------------------------------------------------------------------------
def ler_xml(conteudo: bytes):
    """Retorna um dict do documento ou None se o XML não for fiscal/reconhecido."""
    raiz = carregar_raiz(conteudo)
    tipo = identificar_xml(raiz)
    if tipo == "NFE":
        return ler_nfe(raiz)
    if tipo == "NFSE":
        return ler_nfse(raiz)
    if tipo == "EVENTO_NFE":
        return ler_evento_nfe(raiz)
    return None
