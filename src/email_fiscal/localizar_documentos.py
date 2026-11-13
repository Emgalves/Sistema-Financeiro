# -*- coding: utf-8 -*-
"""
localizar_documentos.py - Botão "🔍 Documentos" nas telas de lançamentos.

A partir de um lançamento JÁ GRAVADO na planilha do cliente, encontra o
boleto, a NF-e/NFS-e e o e-mail de origem, em três níveis:

  EXATO     - o lançamento foi importado pelo e-mail (OBSERVAÇÃO tem
              "IMPORTADO EMAIL #n") ou a triagem já o reconheceu como
              JA_LANCADO apontando para este ID_LANCAMENTO.
  PROVÁVEL  - documento lido do e-mail com a mesma empresa (raiz do CNPJ)
              e mesmo valor, com NF ou vencimento iguais.
  GMAIL     - nada no arquivo local: abre a pesquisa do Gmail no navegador
              com o número da NF e o valor do lançamento.
"""
VERSAO = "2026-11-12.9"   # igual em todos os arquivos do módulo

import os
import re
import urllib.parse
import webbrowser
from datetime import datetime, date, timedelta

from . import config_email as cfg
from .caixa_entrada import CaixaEntrada
from .visualizar_nota import abrir_documento
from .roteador import _chave_doc, _raiz, _nf_texto, _para_data, _para_float, _normalizar_cabecalho


# ---------------------------------------------------------------------------
# Busca (sem interface - pode ser testada isoladamente)
# ---------------------------------------------------------------------------
def _norm_lanc(lanc: dict):
    d = {_normalizar_cabecalho(k): v for k, v in dict(lanc).items()}
    return {
        "id": str(d.get("ID_LANCAMENTO") or "").split(".")[0],
        "doc": _chave_doc(d.get("CNPJ_CPF")),
        "nf": _nf_texto(d.get("NF")),
        "valor": _para_float(d.get("VALOR")),
        "venc": _para_data(d.get("DT_VENCTO")),
        "obs": str(d.get("OBSERVACAO") or ""),
        "nome": str(d.get("NOME") or ""),
    }


def _arquivos(texto):
    return [c.strip() for c in str(texto or "").split(" ; ") if c.strip()]


def buscar_documentos(lanc: dict, caixa: CaixaEntrada = None):
    """Retorna lista de achados: {nivel, motivo, arquivos, msg_id, remetente, assunto, data, id_sugestao}."""
    caixa = caixa or CaixaEntrada()
    L = _norm_lanc(lanc)
    achados, vistos = [], set()

    def add(nivel, motivo, arquivos, sug=None, msg_id=""):
        arquivos = [a for a in arquivos if a]
        chave = tuple(sorted(arquivos)) or (msg_id,)
        if chave in vistos:
            return
        ja_mostrados = {x for a in achados for x in a["arquivos"]}
        if arquivos and nivel != "EXATO" and set(arquivos) <= ja_mostrados:
            return  # mesmos arquivos de um achado anterior: não repete
        vistos.add(chave)
        achados.append({
            "nivel": nivel, "motivo": motivo, "arquivos": arquivos,
            "msg_id": str((sug or {}).get("MSG_ID") or msg_id or ""),
            "remetente": str((sug or {}).get("EMAIL_REMETENTE") or ""),
            "assunto": str((sug or {}).get("EMAIL_ASSUNTO") or ""),
            "data": (sug or {}).get("EMAIL_DATA"),
            "id_sugestao": (sug or {}).get("ID_SUGESTAO"),
        })

    # 1. EXATO pelo marcador da OBSERVAÇÃO
    for m in re.finditer(rf"{re.escape(cfg.PREFIXO_OBSERVACAO)}\s*#(\d+)", L["obs"]):
        for s in caixa.lancamentos:
            if str(s.get("ID_SUGESTAO")).split(".")[0] == m.group(1):
                add("EXATO", f"Importado pelo e-mail (sugestão #{m.group(1)})", _arquivos(s.get("ARQUIVOS")), s)

    # 2. EXATO pela triagem: sugestão reconhecida como já lançada NESTE ID_LANCAMENTO
    if L["id"]:
        padrao = re.compile(rf"ID_LANCAMENTO {re.escape(L['id'])}(?:\D|$)")
        for s in caixa.lancamentos:
            if padrao.search(str(s.get("ALERTAS") or "")) and str(s.get("STATUS")).upper() == "JA_LANCADO":
                add("EXATO", f"Triagem reconheceu este lançamento (sugestão #{s.get('ID_SUGESTAO')})",
                    _arquivos(s.get("ARQUIVOS")), s)

    # 3. PROVÁVEL: mesma empresa + mesmo valor + (mesma NF ou mesmo vencimento)
    if L["doc"] and L["valor"] is not None:
        for s in caixa.lancamentos:
            if _raiz(_chave_doc(s.get("CNPJ_CPF"))) != _raiz(L["doc"]):
                continue
            v = _para_float(s.get("VALOR"))
            if v is None or abs(v - L["valor"]) > 0.01:
                continue
            nf_sug = _nf_texto(s.get("NF"))
            if L["nf"] and nf_sug and nf_sug != L["nf"]:
                continue  # as duas têm número de NF e são diferentes: é outra nota
            mesma_nf = L["nf"] and nf_sug == L["nf"]
            mesmo_venc = L["venc"] and _para_data(s.get("DT_VENCTO")) == L["venc"]
            if mesma_nf or mesmo_venc:
                motivo = "mesma empresa e valor, " + ("mesma NF" if mesma_nf else "mesmo vencimento")
                add("PROVÁVEL", f"{motivo} (sugestão #{s.get('ID_SUGESTAO')}, {s.get('STATUS')})",
                    _arquivos(s.get("ARQUIVOS")), s)

    # 4. PROVÁVEL: nota lida (XML/PDF) da mesma empresa com o mesmo número
    if L["doc"] and L["nf"]:
        for d in caixa.documentos:
            if _raiz(_chave_doc(d.get("EMIT_DOC"))) == _raiz(L["doc"]) and _nf_texto(d.get("NUMERO")) == L["nf"]:
                add("PROVÁVEL", f"{d.get('TIPO')} da mesma empresa com NF {L['nf']}",
                    _arquivos(d.get("ARQUIVO")), msg_id=str(d.get("MSG_ID") or ""))

    ordem = {"EXATO": 0, "PROVÁVEL": 1}
    achados.sort(key=lambda a: ordem.get(a["nivel"], 9))
    return achados


def consulta_gmail(lanc: dict):
    """Texto de pesquisa do Gmail para o lançamento (NF e/ou valor, período do vencimento)."""
    L = _norm_lanc(lanc)
    termos = []
    if L["nf"]:
        termos.append(L["nf"])
    if L["valor"] is not None:
        termos.append('"' + f"{L['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + '"')
    q = "(" + " OR ".join(termos) + ")" if termos else f'"{L["nome"]}"'
    if L["venc"]:
        q += f" after:{(L['venc'] - timedelta(days=120)):%Y/%m/%d} before:{(L['venc'] + timedelta(days=15)):%Y/%m/%d}"
    return q


def abrir_pesquisa_gmail(lanc: dict):
    q = consulta_gmail(lanc)
    webbrowser.open("https://mail.google.com/mail/u/0/#search/" + urllib.parse.quote(q))
    return q


# ---------------------------------------------------------------------------
# Interface
# ---------------------------------------------------------------------------
def abrir_documentos_do_lancamento(tela):
    """Chamado pelas telas 'Gerenciar Lançamentos' e 'Lançamentos - <fornecedor>'.
    Usa o lançamento selecionado na lista (coluna oculta 'ID') e os dados completos
    em tela.dados_originais."""
    import tkinter as tk
    from tkinter import ttk
    from .importar_para_sistema import _msg

    sel = tela.tree_lancamentos.selection()
    if not sel:
        _msg("info", "Seleção", "Selecione um lançamento.")
        return
    try:
        id_lanc = int(float(str(tela.tree_lancamentos.set(sel[0], "ID"))))
        import pandas as pd
        df = tela.dados_originais
        linha = df[pd.to_numeric(df["ID_LANCAMENTO"], errors="coerce") == id_lanc]
        if linha.empty:
            raise ValueError
        lanc = linha.iloc[0].to_dict()
    except Exception:
        _msg("error", "Erro", "Não foi possível ler o lançamento selecionado. Clique em Atualizar e tente de novo.")
        return

    try:
        achados = buscar_documentos(lanc)
    except Exception as e:
        _msg("error", "Erro", f"Não foi possível ler a planilha de triagem:\n{e}")
        return

    L = _norm_lanc(lanc)
    jan = tk.Toplevel(tela.janela)
    jan.title(f"Documentos do lançamento {L['id']} - {L['nome']}")
    jan.geometry("1000x380")
    jan.transient(tela.janela)

    v = f"{L['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") if L["valor"] is not None else "-"
    ttk.Label(jan, font=("Helvetica", 10, "bold"),
              text=f"{L['nome']}  |  CNPJ {L['doc'] or '-'}  |  NF {L['nf'] or '-'}  |  R$ {v}  |  "
                   f"venc. {L['venc'].strftime('%d/%m/%Y') if L['venc'] else '-'}").pack(anchor="w", padx=10, pady=6)

    cols = ("NÍVEL", "MOTIVO", "E-MAIL", "ARQUIVOS")
    tree = ttk.Treeview(jan, columns=cols, show="headings", height=8)
    for c, w in zip(cols, (80, 330, 260, 300)):
        tree.heading(c, text=c)
        tree.column(c, width=w, anchor="w")
    tree.pack(fill="both", expand=True, padx=10)
    itens = {}
    for a in achados:
        email = " | ".join(x for x in (a["remetente"], a["assunto"]) if x)
        iid = tree.insert("", "end", values=(a["nivel"], a["motivo"], email,
                                             " ; ".join(os.path.basename(x) for x in a["arquivos"])))
        itens[iid] = a
    if itens:
        tree.selection_set(next(iter(itens)))

    status = ttk.Label(jan, text=(f"{len(achados)} documento(s) encontrado(s) no arquivo local."
                                  if achados else
                                  "Nada encontrado no arquivo local. Use 'Pesquisar no Gmail'."))
    status.pack(anchor="w", padx=10, pady=4)

    def atual():
        s = tree.selection()
        return itens.get(s[0]) if s else None

    def abrir_doc():
        a = atual()
        if not a:
            return
        existentes = [cfg.resolver_arquivo(c) for c in a["arquivos"] if os.path.exists(cfg.resolver_arquivo(c))]
        if not existentes:
            _msg("warning", "Documento", "Arquivo não encontrado:\n" + "\n".join(a["arquivos"] or ["(sem arquivo)"]))
        for c in sorted(existentes, key=lambda x: not x.lower().endswith(".pdf"))[:2]:
            abrir_documento(c)

    def abrir_email():
        a = atual()
        if not a or not a["msg_id"] or a["msg_id"].startswith("ARQ-"):
            _msg("info", "E-mail", "Este documento não veio de e-mail (arquivo local).")
            return
        webbrowser.open(f"https://mail.google.com/mail/u/0/#all/{a['msg_id']}")

    def pesquisar():
        q = abrir_pesquisa_gmail(lanc)
        status.config(text=f"Pesquisa aberta no Gmail: {q}")

    tree.bind("<Double-1>", lambda e: abrir_doc())
    rod = ttk.Frame(jan)
    rod.pack(fill="x", padx=10, pady=8)
    ttk.Button(rod, text="📄 Abrir documento", command=abrir_doc).pack(side="left", padx=4)
    ttk.Button(rod, text="✉️ Abrir e-mail", command=abrir_email).pack(side="left", padx=4)
    ttk.Button(rod, text="🔎 Pesquisar no Gmail", command=pesquisar).pack(side="left", padx=4)
    ttk.Button(rod, text="Fechar", command=jan.destroy).pack(side="right", padx=4)
