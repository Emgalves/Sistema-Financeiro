# -*- coding: utf-8 -*-
"""
coletor.py - Lê o Gmail (ou uma pasta local), identifica NF-e, NFS-e e boletos,
e grava SUGESTÕES de lançamento em Documentos_Recebidos.xlsx.

Uso pela linha de comando (a partir da pasta raiz do projeto, a que contém 'src'):
    python -m src.email_fiscal.coletor               # Gmail, últimos N dias do config
    python -m src.email_fiscal.coletor --dias 90     # Gmail, últimos 90 dias
    python -m src.email_fiscal.coletor --pasta "C:\\caminho\\com\\xmls_e_pdfs"
"""

VERSAO = "2026-11-12.8"   # igual em todos os arquivos do módulo
import argparse
import threading
import hashlib
import re
from datetime import datetime, date
from pathlib import Path

from . import config_email as cfg
from .caixa_entrada import CaixaEntrada
from .leitor_boleto import (
    chaves_da_danfe, cnpjs_sem_pontuacao, conferir_boleto_com_texto, numeros_documento, corrigir_vencimento_pelo_impresso, dados_da_chave_nfe, dados_da_chave_nfse, encontrar_boletos,
    encontrar_chaves_nfe, encontrar_chaves_nfse, encontrar_documentos, extrair_texto_pdf,
    ler_nota_debito, numero_ao_lado, parece_danfe, parece_nfse, valor_liquido_danfse, vencimento_no_texto,
)
from .leitor_nfe import ler_xml
from .roteador import BaseDados

PALAVRAS_FISCAIS = re.compile(r"boleto|nf|nota|fatura|danfe|cobran|duplicata", re.IGNORECASE)


# ===========================================================================
# Leitura dos anexos de UMA mensagem
# ===========================================================================
def _ler_anexos(msg, caixa, senhas=None):
    ref = msg["data"].date() if msg.get("data") else date.today()
    # conjunto VIVO: cada documento aceito entra nele na hora, de modo que o mesmo
    # boleto/nota anexado duas vezes (ou reenviado no mesmo lote) é lido uma vez só
    vistos = caixa.chaves_documentos()
    arquivos_vistos = set()
    r = {"notas": [], "eventos": [], "boletos": [], "danfes": [], "nfse_pdf": [],
         "problemas": [], "duplicados": 0, "textos_pdf": []}

    def novo(chave):
        if not chave or chave in vistos:
            r["duplicados"] += 1 if chave else 0
            return False
        vistos.add(chave)
        return True

    for nome, conteudo in msg["anexos"]:
        assinatura = hashlib.sha1(conteudo or b"").hexdigest()
        if assinatura in arquivos_vistos:
            continue  # arquivo idêntico anexado de novo
        arquivos_vistos.add(assinatura)
        baixo = nome.lower()

        if baixo.endswith(".xml"):
            doc = ler_xml(conteudo)
            if not doc:
                continue
            doc["_arquivo"], doc["_bytes"] = nome, conteudo
            if doc["tipo"] == "EVENTO_NFE":
                chave_evt = f"EVT-{doc['chave']}-{doc['tipo_evento']}"
                if novo(chave_evt):
                    doc["_chave_registro"] = chave_evt
                    r["eventos"].append(doc)
            elif novo(doc["chave"]):
                r["notas"].append(doc)

        elif baixo.endswith(".pdf"):
            texto, status = extrair_texto_pdf(conteudo, senhas if senhas is not None else cfg.SENHAS_PDF)
            if status != "OK":
                if PALAVRAS_FISCAIS.search(nome) or PALAVRAS_FISCAIS.search(msg.get("assunto", "")):
                    r["problemas"].append((nome, conteudo, status))
                continue
            r["textos_pdf"].append(texto)
            docs_texto = encontrar_documentos(texto)
            docs_texto["sem_pontuacao"] = cnpjs_sem_pontuacao(texto)
            boletos = [corrigir_vencimento_pelo_impresso(b, texto) for b in encontrar_boletos(texto, ref)]
            for b in boletos:
                if novo(b["linha_digitavel"]):
                    b.update(_arquivo=nome, _bytes=conteudo, _docs=docs_texto, _texto=texto,
                             _nf_doc=numeros_documento(texto))
                    r["boletos"].append(b)
            # nota em PDF (pode estar no mesmo arquivo do boleto)
            nota_pdf = False
            if parece_danfe(texto):
                for chave in chaves_da_danfe(texto):
                    nota_pdf = True
                    r["danfes"].append({"chave": chave, "modelo_doc": "NFE", **dados_da_chave_nfe(chave),
                                        "_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto,
                                        "_texto": texto})
            if not boletos and not nota_pdf:
                for chave in encontrar_chaves_nfse(texto + " " + nome):
                    nota_pdf = True
                    r["danfes"].append({"chave": chave, "modelo_doc": "NFSE", **dados_da_chave_nfse(chave),
                                        "_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto,
                                        "_texto": texto,
                                        "_valor": valor_liquido_danfse(conteudo, senhas or ()),
                                        "_venc": vencimento_no_texto(texto)})
            if not boletos and not nota_pdf and parece_nfse(texto):
                r["nfse_pdf"].append({"_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto,
                                      "_texto": texto,
                                      "_nf_doc": numero_ao_lado(conteudo, senhas or ()) or numeros_documento(texto)})

    # linha digitável colada no corpo do e-mail (sem PDF)
    if not r["boletos"] and msg.get("corpo"):
        docs_corpo = encontrar_documentos(msg["corpo"])
        docs_corpo["sem_pontuacao"] = cnpjs_sem_pontuacao(msg["corpo"])
        for b in encontrar_boletos(msg["corpo"], ref):
            if novo(b["linha_digitavel"]):
                b.update(_arquivo="", _bytes=None, _docs=docs_corpo, _texto="",
                         _nf_doc=numeros_documento(msg["corpo"]))
                r["boletos"].append(b)

    # nota em PDF: descarta se o XML veio junto ou se já foi lida antes
    chaves_xml = {n["chave"] for n in r["notas"]}
    unicas = []
    for d in r["danfes"]:
        if d["chave"] in chaves_xml:
            continue
        if d["chave"] in {u["chave"] for u in unicas}:
            continue
        if novo(d["chave"]):
            unicas.append(d)
    r["danfes"] = unicas
    return r


# ===========================================================================
# Arquivamento
# ===========================================================================
def _salvar_arquivo(msg, prefixo, nome, conteudo):
    if not conteudo:
        return ""
    origem = msg.get("_origem_path")
    if origem and Path(origem).name == nome:
        try:  # arquivo já guardado em Documentos_Email: usa o próprio, sem duplicar
            Path(origem).resolve().relative_to(Path(cfg.PASTA_DOCUMENTOS).resolve())
            return str(origem)
        except ValueError:
            pass
    ref = msg.get("data") or datetime.now()
    pasta = cfg.PASTA_DOCUMENTOS / ref.strftime("%Y-%m")
    pasta.mkdir(parents=True, exist_ok=True)
    limpo = re.sub(r'[\\/:*?"<>|]+', "_", nome)[:80]
    assinatura = hashlib.sha1(conteudo).hexdigest()[:8]   # nomes iguais nunca se sobrescrevem
    destino = pasta / f"{prefixo}_{assinatura}_{limpo}"
    if not destino.exists():
        destino.write_bytes(conteudo)
    return str(destino)


# ===========================================================================
# Montagem dos lançamentos
# ===========================================================================
def _referencia_nota(nota, i, n):
    itens = [x for x in nota.get("itens", []) if x]
    if nota["tipo"] == "NFSE":
        base = (itens[0] if itens else "SERVIÇOS")[:45]
    elif len(itens) == 1:
        base = itens[0][:45]
    elif itens:
        base = f"{itens[0][:28]} E OUTROS ({len(itens)} ITENS)"
    else:
        base = "MATERIAIS"
    if n > 1:
        base += f" - PARC {i}/{n}"
    return re.sub(r"\s+", " ", base).upper().strip()


def motivo_sem_cobranca(natureza, formas_pagamento, tem_duplicatas):
    """Devolve o motivo (texto) se a NF-e não gera conta a pagar; "" caso contrário."""
    import unicodedata
    nat = unicodedata.normalize("NFKD", str(natureza or "")).encode("ascii", "ignore").decode().upper()
    palavra = next((p for p in cfg.NATUREZAS_SEM_COBRANCA if re.search(rf"\b{p}\b", nat)), None)
    formas = [f for f in (formas_pagamento or []) if f]
    sem_pag = bool(formas) and all(f in cfg.PAGAMENTOS_SEM_COBRANCA for f in formas)
    if tem_duplicatas:
        return ""
    if palavra:
        return f"natureza da operação \"{natureza}\""
    if sem_pag:
        return "meio de pagamento 90 - Sem pagamento"
    return ""


def _identificar_beneficiario(bol, base):
    """Quem recebe o boleto (o fornecedor). Retorna (cnpj, alerta).

    Candidatos = CNPJs válidos impressos no boleto (com ou sem pontuação),
    EXCETO o pagador, os CPFs dos clientes e os CNPJs da própria empresa.
      1. o CNPJ rotulado como Beneficiário/Cedente/Sacador, se for candidato;
      2. se sobrar um único candidato, é ele;
      3. se vários, o único que já é fornecedor cadastrado;
      4. sem CNPJ impresso: razão social de um fornecedor cadastrado escrita no boleto."""
    docs = bol["_docs"]
    excluir = set(cfg.CNPJS_EMPRESA) | {c["doc"] for c in base.clientes if c["doc"]}
    if docs.get("pagador"):
        excluir.add(docs["pagador"])
    candidatos = []
    for d in docs.get("todos", []) + docs.get("sem_pontuacao", []):
        if len(d) == 14 and d not in excluir and d not in candidatos:
            candidatos.append(d)

    if docs.get("beneficiario") in candidatos:
        return docs["beneficiario"], ""
    if len(candidatos) == 1:
        return candidatos[0], ""
    cadastrados = [d for d in candidatos if base.buscar_fornecedor(d)]
    if len(cadastrados) == 1:
        return cadastrados[0], ""
    if candidatos:
        return "", f"Vários CNPJs no boleto ({', '.join(candidatos)}) - informe o do beneficiário"
    return _beneficiario_pelo_nome(bol.get("_texto", ""), base, excluir)


def _beneficiario_pelo_nome(texto, base, excluir):
    import unicodedata
    def norm(x):
        x = unicodedata.normalize("NFKD", str(x or "")).encode("ascii", "ignore").decode().upper()
        return re.sub(r"[^A-Z0-9]+", " ", x).strip()
    t = " " + norm(texto) + " "
    achados = {}
    for doc, f in base.fornecedores.items():
        if len(doc) != 14 or doc in excluir:
            continue
        razao = re.sub(r"\b(LTDA|ME|EPP|EIRELI|S A|SA)\b", " ", norm(f["razao"]))
        razao = re.sub(r"\s+", " ", razao).strip()
        if len(razao) >= 10 and f" {razao} " in t:
            achados[doc] = f
    raizes = {d[:8] for d in achados}
    if len(raizes) != 1:
        return "", ""
    if len(achados) == 1:
        doc = next(iter(achados))
        return doc, f"CNPJ não impresso no documento: deduzido pelo nome ({achados[doc]['razao']}) - confira"
    docs = sorted(achados)
    nome = achados[docs[0]]["razao"]
    return "", (f"CNPJ não impresso no documento. Fornecedor pelo nome: {nome}, que tem mais de um "
                f"CNPJ cadastrado ({', '.join(docs)}) - use Informar CNPJ")


def _nota_do_boleto(bol, doc_forn, danfes):
    """Qual nota em PDF do mesmo e-mail corresponde a este boleto.
    Nunca "chuta": o número da nota tem de estar impresso no boleto, ou então
    tem de haver UMA única nota daquele fornecedor no e-mail."""
    if not danfes:
        return None
    texto = bol.get("_texto") or ""
    do_fornecedor = [d for d in danfes if not doc_forn or d["emit_doc"] == doc_forn]
    for d in do_fornecedor:
        if re.search(rf"(?<![\d.,/])0*{re.escape(d['numero'])}(?:-\d)?(?![\d]|[.,/]\d)", texto):
            return d
    if len(do_fornecedor) == 1 and doc_forn:
        return do_fornecedor[0]
    if len(danfes) == 1 and not doc_forn:
        return danfes[0]
    return None


def _achar_boleto(boletos, valor, vencimento, usados):
    for idx, b in enumerate(boletos):
        if idx in usados or b.get("valor") is None or valor is None:
            continue
        if abs(b["valor"] - valor) > cfg.TOLERANCIA_VALOR:
            continue
        if vencimento and b.get("vencimento") and b["vencimento"] != vencimento:
            continue
        usados.add(idx)
        return b
    return None


def _completar(lanc, doc_fornecedor, nome_emitente, tipo, docs_cliente, textos, msg, base, alertas):
    """Preenche cliente, fornecedor, categoria, alertas e status."""
    forn = base.buscar_fornecedor(doc_fornecedor) if doc_fornecedor else None
    parente = None if forn or not doc_fornecedor else base.buscar_mesma_empresa(doc_fornecedor)
    if forn:
        lanc["NOME"] = forn["nome"]
        lanc["CATEGORIA"] = forn["categoria"] or cfg.CATEGORIA_PADRAO_POR_TIPO.get(tipo, "")
        if forn["status"] and forn["status"] != "ATIVO":
            alertas.append(f"Fornecedor com STATUS {forn['status']} na base")
        if not lanc.get("DADOS_BANCARIOS") and forn.get("dados_bancarios") not in ("", "-"):
            lanc["DADOS_BANCARIOS"] = forn["dados_bancarios"]
    elif parente:
        # outra filial de uma empresa já cadastrada (mesma raiz de 8 dígitos)
        lanc["NOME"] = parente["nome"]
        lanc["CATEGORIA"] = parente["categoria"] or cfg.CATEGORIA_PADRAO_POR_TIPO.get(tipo, "")
        alertas.append(f"FILIAL NOVA: CNPJ {doc_fornecedor} não cadastrado; mesma empresa de "
                       f"{parente['nome']} ({parente['doc']}) - cadastre antes de importar")
    else:
        lanc["NOME"] = (nome_emitente or "").upper()[:60]
        lanc["CATEGORIA"] = cfg.CATEGORIA_PADRAO_POR_TIPO.get(tipo, "")
        if doc_fornecedor:
            alertas.append(f"Fornecedor {doc_fornecedor} NÃO cadastrado na Base_Fornecedores - cadastre antes de importar")
        else:
            alertas.append("CNPJ do fornecedor não identificado - preencha CNPJ_CPF e NOME")

    lanc["CNPJ_CPF"] = doc_fornecedor or ""
    lanc["TP_DESP"] = ("" if cfg.REGRA_TP_DESP_POR_DATA and tipo in cfg.TIPOS_COM_REGRA_POR_DATA
                       else cfg.TP_DESP_POR_TIPO.get(tipo, ""))
    lanc["TIPO_DOC"] = tipo

    lanc["DOC_CLIENTE"] = next((re.sub(r"\D", "", d) for d in docs_cliente if d and re.sub(r"\D", "", d)), "")
    cliente, metodo, motivo = base.identificar_cliente(docs_cliente, textos)
    lanc["CLIENTE"], lanc["IDENTIFICACAO"], lanc["MOTIVO_IDENTIFICACAO"] = cliente, metodo, motivo
    if metodo != "CONFIRMADO":
        alertas.append("Confira/escolha o CLIENTE")

    if lanc.get("VALOR") in (None, ""):
        alertas.append("Sem VALOR - preencha")
    if not lanc.get("DT_VENCTO"):
        alertas.append("Sem vencimento - preencha DT_VENCTO")

    if cliente:
        ja, motivo_ja = base.ja_lancado(cliente, doc=doc_fornecedor, nf=lanc.get("NF"),
                                        valor=lanc.get("VALOR"), vencimento=lanc.get("DT_VENCTO"),
                                        nfs_provaveis=str(lanc.get("NF_PROVAVEL") or "").split(";"),
                                        parcelado=" - PARC " in str(lanc.get("REFERENCIA") or ""))
        if ja:
            lanc["STATUS"] = "JA_LANCADO"
            alertas.append(f"Já existe na planilha do cliente: {motivo_ja}")
        else:
            dup = base.possivel_duplicidade(cliente, doc=doc_fornecedor, nf=lanc.get("NF"),
                                            valor=lanc.get("VALOR"), vencimento=lanc.get("DT_VENCTO"))
            if dup:
                alertas.insert(0, dup)

    lanc["ALERTAS"] = " | ".join(dict.fromkeys(alertas))
    lanc["EMAIL_DATA"] = msg.get("data")
    lanc["EMAIL_REMETENTE"] = msg.get("remetente", "")[:120]
    lanc["EMAIL_ASSUNTO"] = msg.get("assunto", "")[:200]
    lanc["MSG_ID"] = msg["id"]
    return lanc


def processar_mensagem(msg, base: BaseDados, caixa: CaixaEntrada, registrar=True):
    """Processa UMA mensagem lógica (já baixada). Devolve um texto-resumo.
    registrar=False: quem chama registra a mensagem do Gmail uma vez só, no fim."""
    lido = _ler_anexos(msg, caixa, base.senhas_pdf())
    textos_base = [msg.get("assunto", ""), (msg.get("corpo") or "")[:5000]]
    criados, atualizados = 0, 0

    # ---- 1. Eventos de cancelamento -----------------------------------------
    for evt in lido["eventos"]:
        caixa.registrar_documento(evt["_chave_registro"], "EVENTO_NFE", evt["tipo_evento"], "",
                                  None, msg["id"], evt["_arquivo"])
        if evt["tipo_evento"] != "110111":
            continue
        for l in caixa.lancamentos:
            if str(l.get("CHAVE_DOC")) == evt["chave"]:
                if str(l.get("STATUS")).upper() == "PENDENTE":
                    l["STATUS"] = "CANCELADA"
                l["ALERTAS"] = ("NF-e CANCELADA NA SEFAZ | " + str(l.get("ALERTAS") or "")).strip(" |")
                atualizados += 1

    # ---- 2. Notas fiscais em XML ---------------------------------------------
    boletos, usados = lido["boletos"], set()
    for nota in lido["notas"]:
        arq = _salvar_arquivo(msg, f"{nota['emit_doc']}_{nota['tipo']}_{nota['numero']}",
                              nota["_arquivo"], nota["_bytes"])
        caixa.registrar_documento(nota["chave"], nota["tipo"], nota["numero"], nota["emit_doc"],
                                  nota["valor_total"], msg["id"], arq)

        parcelas = [d for d in nota["duplicatas"] if d.get("valor")] or \
                   [{"numero": "", "vencimento": None, "valor": nota["valor_total"]}]
        n = len(parcelas)
        for i, p in enumerate(parcelas, start=1):
            alertas = []
            if nota["tipo"] == "NFE" and nota.get("autorizada") is None:
                alertas.append("XML sem protocolo de autorização - confirme a NF na SEFAZ")
            elif nota["tipo"] == "NFE" and nota.get("autorizada") is False:
                alertas.append(f"NF-e com cStat {nota['c_stat']} (não autorizada)")
            if nota["tipo"] == "NFE" and nota.get("tipo_operacao") == "0":
                alertas.append("NF-e de ENTRADA (tpNF=0) - confira se é conta a pagar")

            bol = _achar_boleto(boletos, p["valor"], p["vencimento"], usados)
            venc = p["vencimento"] or (bol["vencimento"] if bol else None)
            if bol:
                alertas += conferir_boleto_com_texto(bol, bol.get("_texto"))
            arquivos = [arq]
            if bol:
                arquivos.append(_salvar_arquivo(msg, f"{nota['emit_doc']}_BOLETO_{nota['numero']}",
                                                bol["_arquivo"], bol["_bytes"]))
                caixa.registrar_documento(bol["linha_digitavel"], "BOLETO", nota["numero"],
                                          nota["emit_doc"], bol["valor"], msg["id"], arquivos[-1])

            lanc = {
                "REFERENCIA": _referencia_nota(nota, i, n),
                "NF": nota["numero"],
                "VALOR": p["valor"],
                "DT_VENCTO": venc,
                "DADOS_BANCARIOS": f"BOLETO: {bol['linha_formatada']}" if bol else "",
                "LINHA_DIGITAVEL": bol["linha_digitavel"] if bol else "",
                "CHAVE_DOC": nota["chave"],
                "ARQUIVOS": " ; ".join(a for a in arquivos if a),
            }
            docs_cli = [nota["dest_doc"]] + ([bol["_docs"]["pagador"]] if bol else [])
            textos = textos_base + [nota["inf_complementar"], nota["endereco_entrega"], nota["dest_nome"]]
            _completar(lanc, nota["emit_doc"], nota["emit_nome"], nota["tipo"],
                       docs_cli, textos, msg, base, alertas)
            if nota["tipo"] == "NFE":
                motivo = motivo_sem_cobranca(nota.get("natureza"), nota.get("formas_pagamento"),
                                             bool(nota["duplicatas"]) or bool(bol))
                if motivo:
                    lanc["STATUS"] = "IGNORADO"
                    lanc["ALERTAS"] = f"NF-e SEM COBRANÇA ({motivo}) - ignorada automaticamente | {lanc['ALERTAS']}".strip(" |")
            caixa.adicionar_lancamento(lanc)
            criados += 1

    # ---- 3. Boletos que sobraram (sem XML no mesmo e-mail) -------------------
    danfes_usadas = set()
    for idx, bol in enumerate(boletos):
        if idx in usados:
            continue
        doc_forn, alerta_forn = _identificar_beneficiario(bol, base)
        danfe = _nota_do_boleto(bol, doc_forn, lido["danfes"])
        if not doc_forn and danfe:
            doc_forn, alerta_forn = danfe["emit_doc"], ""

        arq = _salvar_arquivo(msg, f"{doc_forn or 'SEMCNPJ'}_BOLETO",
                              bol["_arquivo"] or "boleto_corpo_email.txt",
                              bol["_bytes"] or bol["linha_digitavel"].encode())
        caixa.registrar_documento(bol["linha_digitavel"], "BOLETO", "", doc_forn,
                                  bol["valor"], msg["id"], arq)

        # a nota pode ter chegado antes, em outro e-mail
        existente = caixa.buscar_lancamento_sem_boleto(doc_forn, bol["valor"], bol["vencimento"]) if doc_forn else None
        if existente:
            existente["LINHA_DIGITAVEL"] = bol["linha_digitavel"]
            existente["DADOS_BANCARIOS"] = f"BOLETO: {bol['linha_formatada']}"
            if not existente.get("DT_VENCTO") and bol["vencimento"]:
                existente["DT_VENCTO"] = bol["vencimento"]
                existente["ALERTAS"] = str(existente.get("ALERTAS") or "").replace(
                    "Sem vencimento - preencha DT_VENCTO", "").replace("|  |", "|").strip(" |")
            existente["ARQUIVOS"] = " ; ".join(x for x in [existente.get("ARQUIVOS"), arq] if x)
            atualizados += 1
            continue

        alertas = conferir_boleto_com_texto(bol, bol.get("_texto"))
        if alerta_forn:
            alertas.append(alerta_forn)
        if bol["subtipo"] == "ARRECADACAO":
            alertas.append("Guia de arrecadação/concessionária: vencimento não vem no código")
        nf = ""
        if danfe:
            nf = danfe["numero"]
            tipo_pdf = "DANFSE_PDF" if danfe["modelo_doc"] == "NFSE" else "DANFE_PDF"
            arq_d = _salvar_arquivo(msg, f"{danfe['emit_doc']}_{tipo_pdf}_{nf}", danfe["_arquivo"], danfe["_bytes"])
            if danfe["chave"] not in danfes_usadas:   # a mesma nota pode ter vários boletos (parcelas)
                caixa.registrar_documento(danfe["chave"], tipo_pdf, nf, danfe["emit_doc"], None, msg["id"], arq_d)
            danfes_usadas.add(danfe["chave"])
            arq = f"{arq} ; {arq_d}"
            alertas.append("NF só em PDF (sem XML)")
        lanc = {
            "REFERENCIA": f"BOLETO NF {nf}" if nf else "BOLETO - AGUARDANDO NF",
            "NF": nf,
            "VALOR": bol["valor"],
            "DT_VENCTO": bol["vencimento"],
            "DADOS_BANCARIOS": f"BOLETO: {bol['linha_formatada']}",
            "LINHA_DIGITAVEL": bol["linha_digitavel"],
            "NF_PROVAVEL": "" if nf else ";".join(bol.get("_nf_doc") or []),
            "CHAVE_DOC": danfe["chave"] if danfe else "",
            "ARQUIVOS": arq,
        }
        pag = bol["_docs"]["pagador"]
        docs_cli = [pag] if pag and pag not in cfg.CNPJS_EMPRESA else []
        _completar(lanc, doc_forn, "", "BOLETO", docs_cli, textos_base + [bol["_texto"][:3000]],
                   msg, base, alertas)
        caixa.adicionar_lancamento(lanc)
        criados += 1

    # ---- 4. DANFE em PDF sem XML e sem boleto -----------------------------------
    for d in lido["danfes"]:
        if d["chave"] in danfes_usadas:
            continue
        servico = d["modelo_doc"] == "NFSE"
        tipo_pdf = "DANFSE_PDF" if servico else "DANFE_PDF"
        arq = _salvar_arquivo(msg, f"{d['emit_doc']}_{tipo_pdf}_{d['numero']}", d["_arquivo"], d["_bytes"])
        caixa.registrar_documento(d["chave"], tipo_pdf, d["numero"], d["emit_doc"], None, msg["id"], arq)
        lanc = {"REFERENCIA": "SERVIÇOS" if servico else "MATERIAIS", "NF": d["numero"],
                "VALOR": d.get("_valor"), "DT_VENCTO": d.get("_venc"), "DADOS_BANCARIOS": "",
                "LINHA_DIGITAVEL": "", "CHAVE_DOC": d["chave"], "ARQUIVOS": arq}
        docs_cli = [x for x in d["_docs"]["todos"] if x != d["emit_doc"]]
        lidos = [x for x in (("valor" if d.get("_valor") is not None else ""),
                             ("vencimento (das informações complementares)" if d.get("_venc") else "")) if x]
        faltam = [x for x in (("valor" if d.get("_valor") is None else ""),
                              ("vencimento" if not d.get("_venc") else "")) if x]
        aviso = (f"NFS-e só em PDF: CNPJ e número lidos da chave de acesso"
                 + (f"; {' e '.join(lidos)} lido(s) do PDF - confira" if lidos else "")
                 + (f"; preencha {' e '.join(faltam)}" if faltam else "")
                 if servico else
                 "NF só em PDF (sem XML) - peça o XML ao fornecedor ou preencha valor/vencimento")
        _completar(lanc, d["emit_doc"], "", "NFSE" if servico else "NFE", docs_cli,
                   textos_base + [d.get("_texto", "")], msg, base, [aviso])
        if not servico:
            m = re.search(r"NATUREZA\s+D[AE]\s+OPERA[ÇC][ÃA]O\s*[:\n]?\s*([^\n]{3,80})", d.get("_texto", ""), re.IGNORECASE)
            motivo = motivo_sem_cobranca(m.group(1).strip() if m else "", [], False)
            if motivo:
                lanc["STATUS"] = "IGNORADO"
                lanc["ALERTAS"] = f"NF-e SEM COBRANÇA ({motivo}) - ignorada automaticamente | {lanc['ALERTAS']}".strip(" |")
        caixa.adicionar_lancamento(lanc)
        criados += 1

    # ---- 5. PDFs fiscais que não deu para ler ------------------------------------
    manuais = [(nome, cont, st, [], None) for nome, cont, st in lido["problemas"]]
    manuais += [(x["_arquivo"], x["_bytes"], "NFSE_PDF", x["_docs"]["todos"], x) for x in lido["nfse_pdf"]]
    for nome, conteudo, status, docs_pdf, extra in manuais:
        arq = _salvar_arquivo(msg, f"MANUAL_{status}", nome, conteudo)
        motivo = {"SEM_TEXTO": "PDF é imagem (escaneado) - leitura manual",
                  "PROTEGIDO": "PDF protegido por senha (nenhuma senha conhecida funcionou) - inclua a senha em SENHAS_PDF ou lance manualmente",
                  "GRANDE": "PDF muito grande para leitura automática - leitura manual",
                  "NFSE_PDF": "NFS-e só em PDF - peça o XML ou lance manualmente",
                  }.get(status, f"PDF não lido ({status})")
        lanc = {"REFERENCIA": "", "NF": "", "VALOR": None, "DT_VENCTO": None, "DADOS_BANCARIOS": "",
                "LINHA_DIGITAVEL": "", "CHAVE_DOC": "", "ARQUIVOS": arq}
        doc_forn, alerta_nome = "", ""
        if extra:
            excluir = set(cfg.CNPJS_EMPRESA) | {c["doc"] for c in base.clientes if c["doc"]}
            cnpjs = [d for d in extra["_docs"]["todos"] if len(d) == 14 and d not in excluir]
            if len(cnpjs) == 1:
                doc_forn = cnpjs[0]
            else:
                doc_forn, alerta_nome = _beneficiario_pelo_nome(extra.get("_texto", ""), base, excluir)
            lanc["NF_PROVAVEL"] = ";".join(extra.get("_nf_doc") or [])
        nd = ler_nota_debito(extra.get("_texto", "")) if extra else None
        if nd:
            lanc.update({"NF": nd["numero"], "NF_PROVAVEL": "", "VALOR": nd["valor"],
                         "DT_VENCTO": nd["vencimento"], "REFERENCIA": nd["referencia"]})
            motivo = "Nota de débito: número, valor e vencimento lidos do documento - confira"
        if not doc_forn:
            doc_forn = caixa.cnpj_por_dominio(msg.get("remetente"))
            if doc_forn:
                alerta_nome = "CNPJ deduzido pelo e-mail do remetente (mesmo domínio de outras notas) - confira"
        avisos = [motivo] + ([alerta_nome] if alerta_nome else [])
        _completar(lanc, doc_forn, "", "NFSE" if status == "NFSE_PDF" else "BOLETO", docs_pdf, textos_base,
                   msg, base, avisos)
        lanc["TIPO_DOC"] = "MANUAL"
        caixa.adicionar_lancamento(lanc)
        criados += 1

    partes = []
    if criados:
        partes.append(f"{criados} lançamento(s) sugerido(s)")
    if atualizados:
        partes.append(f"{atualizados} atualizado(s)")
    if lido["duplicados"]:
        partes.append(f"{lido['duplicados']} documento(s) já lido(s) antes")
    resumo = ", ".join(partes) or "sem documento fiscal"
    if registrar:
        caixa.registrar_mensagem(msg["id"], msg.get("data"), msg.get("remetente", ""),
                                 msg.get("assunto", ""), resumo)
    return resumo


# ===========================================================================
# Releitura de sugestões antigas (criadas por versões anteriores do código)
# ===========================================================================
def _reler_manual(l, base, caixa):
    """Linha MANUAL de PDF não lido (erro/senha): tenta abrir de novo, agora com as
    senhas conhecidas. Se houver boleto, a linha vira uma sugestão normal de boleto."""
    alertas = str(l.get("ALERTAS") or "")
    senhas = base.senhas_pdf()
    texto, status = "", "SEM_ARQUIVO"
    for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
        caminho = cfg.resolver_arquivo(caminho.strip())
        if caminho.lower().endswith(".pdf") and Path(caminho).exists():
            texto, status = extrair_texto_pdf(Path(caminho).read_bytes(), senhas)
            if status == "OK":
                break
    if status != "OK":
        motivo = {"PROTEGIDO": "senha não encontrada entre as conhecidas",
                  "SEM_ARQUIVO": "PDF não encontrado"}.get(status, f"PDF continua sem leitura ({status})")
        l["ALERTAS"] = f"{alertas} | Relido: {motivo}".strip(" |")
        return False
    ref = l.get("EMAIL_DATA")
    ref = ref.date() if isinstance(ref, datetime) else date.today()
    boletos = [corrigir_vencimento_pelo_impresso(b, texto) for b in encontrar_boletos(texto, ref)]
    novos = [b for b in boletos if b["linha_digitavel"] not in caixa.chaves_documentos()]
    if not novos:
        l["ALERTAS"] = f"{alertas} | Relido: PDF aberto, " + \
            ("boleto já registrado em outra sugestão" if boletos else "sem boleto - leitura manual")
        if boletos:
            l["STATUS"] = "IGNORADO"
        return bool(boletos)
    bol = novos[0]
    docs = encontrar_documentos(texto)
    docs["sem_pontuacao"] = cnpjs_sem_pontuacao(texto)
    bol.update(_docs=docs, _texto=texto)
    doc_forn, alerta_forn = _identificar_beneficiario(bol, base)
    novas_alertas = conferir_boleto_com_texto(bol, texto)
    if alerta_forn:
        novas_alertas.append(alerta_forn)
    if len(novos) > 1:
        novas_alertas.append(f"PDF tem {len(novos)} boletos: só o primeiro foi lido - confira")
    l.update({"REFERENCIA": "BOLETO - AGUARDANDO NF", "NF": "", "VALOR": bol["valor"],
              "NF_PROVAVEL": ";".join(numeros_documento(texto)),
              "DT_VENCTO": bol["vencimento"], "DADOS_BANCARIOS": f"BOLETO: {bol['linha_formatada']}",
              "LINHA_DIGITAVEL": bol["linha_digitavel"], "CHAVE_DOC": ""})
    pag = docs.get("pagador")
    msg = {"id": l.get("MSG_ID") or "", "data": l.get("EMAIL_DATA"),
           "remetente": str(l.get("EMAIL_REMETENTE") or ""), "assunto": str(l.get("EMAIL_ASSUNTO") or "")}
    _completar(l, doc_forn, "", "BOLETO", [pag] if pag and pag not in cfg.CNPJS_EMPRESA else [],
               [msg["assunto"], texto[:3000]], msg, base, novas_alertas)
    caixa.registrar_documento(bol["linha_digitavel"], "BOLETO", "", doc_forn, bol["valor"],
                              l.get("MSG_ID") or "", l.get("ARQUIVOS") or "")
    l["ALERTAS"] = f"{l.get('ALERTAS') or ''} | Relido: PDF aberto com senha conhecida".strip(" |")
    return True


def reidentificar_clientes(caixa, base, ler_pdf=True):
    """A identificação do cliente é feita na coleta. Se depois disso o CPF/CNPJ
    for cadastrado (ou corrigido) na Clientes.xlsx, as sugestões PENDENTES que
    ficaram SUGERIDO / NAO_IDENTIFICADO / AMBIGUO são reavaliadas pelo documento
    do destinatário (XML) ou do pagador (boleto). Escolhas manuais não mudam."""
    from .leitor_nfe import ler_xml
    alterados = tocados = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE":
            continue
        if str(l.get("IDENTIFICACAO") or "") not in ("SUGERIDO", "NAO_IDENTIFICADO", "AMBIGUO"):
            continue
        doc = re.sub(r"\D", "", str(l.get("DOC_CLIENTE") or ""))
        if not doc and str(l.get("DOC_CLIENTE") or "") != "-":
            # linha criada por versão anterior: busca o documento no arquivo salvo
            for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
                caminho = cfg.resolver_arquivo(caminho.strip())
                if not Path(caminho).exists():
                    continue
                if caminho.lower().endswith(".xml"):
                    nota = ler_xml(Path(caminho).read_bytes()) or {}
                    doc = nota.get("dest_doc") or ""
                elif caminho.lower().endswith(".pdf"):
                    if not ler_pdf:
                        continue  # PDF fica para o "Reprocessar triagem" / próxima coleta
                    t, st = extrair_texto_pdf(Path(caminho).read_bytes(), base.senhas_pdf())
                    if st == "OK":
                        pag = encontrar_documentos(t).get("pagador") or ""
                        doc = "" if pag in cfg.CNPJS_EMPRESA else pag
                if doc:
                    break
            tem_pdf_pendente = (not ler_pdf and not doc and
                                ".pdf" in str(l.get("ARQUIVOS") or "").lower())
            if not tem_pdf_pendente:
                l["DOC_CLIENTE"] = doc or "-"
                tocados += 1
        if not doc:
            continue
        cliente, metodo, motivo = base.identificar_cliente([doc], [])
        if metodo != "CONFIRMADO":
            continue
        antes = str(l.get("CLIENTE") or "")
        l["CLIENTE"], l["IDENTIFICACAO"], l["MOTIVO_IDENTIFICACAO"] = cliente, metodo, motivo
        itens = [a.strip() for a in str(l.get("ALERTAS") or "").split("|")
                 if a.strip() and a.strip() != "Confira/escolha o CLIENTE"]
        if antes and antes != cliente:
            itens.insert(0, f"Cliente corrigido pelo CPF/CNPJ: era {antes}")
        l["ALERTAS"] = " | ".join(itens)
        alterados += 1
        tocados += 1
    return alterados, tocados


def preencher_nf_provavel(caixa, base):
    """Boletos pendentes sem NF, criados antes desta versão: lê o número do
    documento no PDF salvo (uma vez; "-" quando não há)."""
    n = 0
    for l in caixa.lancamentos:
        if (str(l.get("STATUS") or "").upper() != "PENDENTE" or str(l.get("TIPO_DOC")) != "BOLETO"
                or l.get("NF") or str(l.get("NF_PROVAVEL") or "")):
            continue
        achados = []
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = cfg.resolver_arquivo(caminho.strip())
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                t, st = extrair_texto_pdf(Path(caminho).read_bytes(), base.senhas_pdf())
                if st == "OK":
                    achados = numeros_documento(t)
                    break
        l["NF_PROVAVEL"] = ";".join(achados) or "-"
        n += 1 if achados else 0
    return n


def ignorar_sem_cobranca(caixa, so_sem_vencimento=False):
    """Sugestões PENDENTES de NF-e (XML) que são remessa/retorno ou "sem pagamento":
    IGNORADO. so_sem_vencimento=True (abertura da janela) olha só as sem vencimento."""
    from .leitor_nfe import ler_xml
    n = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE" or str(l.get("TIPO_DOC")) != "NFE":
            continue
        if so_sem_vencimento and l.get("DT_VENCTO"):
            continue
        if l.get("LINHA_DIGITAVEL"):
            continue  # tem boleto: há cobrança
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = cfg.resolver_arquivo(caminho.strip())
            if not (caminho.lower().endswith(".xml") and Path(caminho).exists()):
                continue
            nota = ler_xml(Path(caminho).read_bytes()) or {}
            if nota.get("tipo") != "NFE":
                continue
            motivo = motivo_sem_cobranca(nota.get("natureza"), nota.get("formas_pagamento"),
                                         bool(nota.get("duplicatas")))
            if motivo:
                l["STATUS"] = "IGNORADO"
                l["ALERTAS"] = f"NF-e SEM COBRANÇA ({motivo}) - ignorada automaticamente | {l.get('ALERTAS') or ''}".strip(" |")
                n += 1
            break
    return n


def completar_valor_danfse(caixa, base):
    """Sugestões PENDENTES de NFS-e em PDF (chave de 50 dígitos) sem valor: lê o
    valor líquido do PDF salvo com as regras atuais."""
    n = 0
    for l in caixa.lancamentos:
        if (str(l.get("STATUS") or "").upper() != "PENDENTE" or str(l.get("TIPO_DOC")) != "NFSE"
                or l.get("VALOR") not in (None, "") or len(str(l.get("CHAVE_DOC") or "")) != 50):
            continue
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = cfg.resolver_arquivo(caminho.strip())
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                v = valor_liquido_danfse(Path(caminho).read_bytes(), base.senhas_pdf())
                if v is not None:
                    l["VALOR"] = v
                    itens = [x.strip() for x in str(l.get("ALERTAS") or "").split("|") if x.strip()
                             and not x.strip().startswith("Sem VALOR")]
                    itens.append("Valor líquido lido do PDF - confira")
                    l["ALERTAS"] = " | ".join(dict.fromkeys(itens))
                    n += 1
                break
    return n


def completar_notas_debito(caixa, base):
    """Sugestões MANUAL pendentes sem valor que são notas de débito: preenche número,
    valor, vencimento e referência; CNPJ pelo domínio do remetente, se possível."""
    n = 0
    for l in caixa.lancamentos:
        if (str(l.get("STATUS") or "").upper() != "PENDENTE" or str(l.get("TIPO_DOC")) != "MANUAL"
                or l.get("VALOR") not in (None, "") or "Nota de débito" in str(l.get("ALERTAS") or "")):
            continue
        texto = ""
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = cfg.resolver_arquivo(caminho.strip())
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                texto, st = extrair_texto_pdf(Path(caminho).read_bytes(), base.senhas_pdf())
                break
        nd = ler_nota_debito(texto)
        if not nd:
            continue
        l.update({"NF": nd["numero"], "NF_PROVAVEL": "", "VALOR": nd["valor"],
                  "DT_VENCTO": nd["vencimento"], "REFERENCIA": nd["referencia"]})
        itens = [x.strip() for x in str(l.get("ALERTAS") or "").split("|") if x.strip()
                 and not x.strip().startswith(("Sem VALOR", "Sem vencimento", "NFS-e só em PDF", "Relido"))]
        itens.insert(0, "Nota de débito: número, valor e vencimento lidos do documento - confira")
        if not str(l.get("CNPJ_CPF") or "").strip():
            doc = caixa.cnpj_por_dominio(l.get("EMAIL_REMETENTE"))
            if doc:
                l["CNPJ_CPF"] = doc
                forn = base.buscar_fornecedor(doc)
                if forn:
                    l["NOME"], l["CATEGORIA"] = forn["nome"], forn["categoria"] or l.get("CATEGORIA")
                itens = [x for x in itens if not x.startswith("CNPJ do fornecedor não identificado")]
                itens.append("CNPJ deduzido pelo e-mail do remetente (mesmo domínio de outras notas) - confira")
        l["ALERTAS"] = " | ".join(dict.fromkeys(itens))
        n += 1
    return n


def completar_manuais_nfse(caixa, base):
    """Sugestões MANUAL de NFS-e em PDF (ex.: relatório de medição) sem CNPJ:
    fornecedor pelo CNPJ impresso ou pela razão social cadastrada, e NF ao lado de
    "NFS-e:". Uma vez por sugestão (marca "Relido")."""
    n = 0
    excluir = set(cfg.CNPJS_EMPRESA) | {c["doc"] for c in base.clientes if c["doc"]}
    for l in caixa.lancamentos:
        alertas = str(l.get("ALERTAS") or "")
        if (str(l.get("STATUS") or "").upper() != "PENDENTE" or str(l.get("TIPO_DOC")) != "MANUAL"
                or "NFS-e só em PDF" not in alertas or "Relido" in alertas or str(l.get("CNPJ_CPF") or "").strip()):
            continue
        conteudo, texto = None, ""
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = cfg.resolver_arquivo(caminho.strip())
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                conteudo = Path(caminho).read_bytes()
                texto, st = extrair_texto_pdf(conteudo, base.senhas_pdf())
                if st == "OK":
                    break
        if not texto:
            l["ALERTAS"] = f"{alertas} | Relido: PDF não encontrado ou sem texto"
            continue
        docs = encontrar_documentos(texto)
        cnpjs = [d for d in docs["todos"] if len(d) == 14 and d not in excluir]
        doc, alerta = (cnpjs[0], "") if len(cnpjs) == 1 else _beneficiario_pelo_nome(texto, base, excluir)
        itens = [a.strip() for a in alertas.split("|") if a.strip()]
        if doc:
            l["CNPJ_CPF"] = doc
            forn = base.buscar_fornecedor(doc)
            if forn:
                l["NOME"], l["CATEGORIA"] = forn["nome"], forn["categoria"] or l.get("CATEGORIA")
            itens = [a for a in itens if not a.startswith("CNPJ do fornecedor não identificado")]
        if alerta:
            itens.append(alerta)
        nfs = numero_ao_lado(conteudo, base.senhas_pdf()) or numeros_documento(texto)
        if nfs:
            l["NF_PROVAVEL"] = ";".join(nfs)
        partes = [x for x in (f"fornecedor {doc}" if doc else "", f"NF {nfs[0]}" if nfs else "") if x]
        itens.append("Relido: " + (", ".join(partes) if partes else "sem fornecedor/NF identificáveis"))
        l["ALERTAS"] = " | ".join(dict.fromkeys(itens))
        n += 1 if partes else 0
    return n


def reler_pendentes(caixa, base):
    """Sugestões PENDENTES de boleto sem CNPJ, ou com vencimento divergente do
    impresso, são relidas do PDF já salvo, com as regras atuais. Assim,
    correções no código valem para o que já está na triagem, sem coletar de novo.
    Cada linha é relida uma única vez (marca "Relido" nos ALERTAS)."""
    from .leitor_boleto import decodificar_bancario
    alterados = tocados = 0
    for l in caixa.lancamentos:
        if str(l.get("STATUS") or "").upper() != "PENDENTE":
            continue
        linha = str(l.get("LINHA_DIGITAVEL") or "").strip()
        alertas = str(l.get("ALERTAS") or "")
        if (str(l.get("TIPO_DOC")) == "MANUAL" and "Relido" not in alertas
                and ("PDF não lido" in alertas or "protegido" in alertas)):
            tocados += 1
            if _reler_manual(l, base, caixa):
                alterados += 1
            continue
        sem_cnpj = not str(l.get("CNPJ_CPF") or "").strip()
        venc_div = "Vencimento do código" in alertas and "não aparece impresso" in alertas
        if not linha or "Relido" in alertas or not (sem_cnpj or venc_div):
            continue
        texto = ""
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = caminho.strip()
            caminho = cfg.resolver_arquivo(caminho)
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                t, st = extrair_texto_pdf(Path(caminho).read_bytes(), base.senhas_pdf())
                if st == "OK" and linha in re.sub(r"\D", "", t):
                    texto = t
                    break
        tocados += 1
        if not texto:
            l["ALERTAS"] = f"{alertas} | Relido: PDF do boleto não encontrado".strip(" |")
            continue
        docs = encontrar_documentos(texto)
        docs["sem_pontuacao"] = cnpjs_sem_pontuacao(texto)
        itens = [a.strip() for a in alertas.split("|") if a.strip()]
        mudou = []

        if sem_cnpj:
            doc, alerta = _identificar_beneficiario({"_docs": docs, "_texto": texto}, base)
            if doc:
                l["CNPJ_CPF"] = doc
                forn = base.buscar_fornecedor(doc) or base.buscar_mesma_empresa(doc)
                if forn:
                    l["NOME"] = forn["nome"]
                    l["CATEGORIA"] = forn["categoria"] or l.get("CATEGORIA")
                if not base.buscar_fornecedor(doc):
                    itens.append(f"Fornecedor {doc} NÃO cadastrado na Base_Fornecedores - cadastre antes de importar")
                itens = [a for a in itens if not a.startswith("CNPJ do fornecedor não identificado")]
                mudou.append(f"CNPJ {doc}")
            elif alerta:
                itens.append(alerta)
        pag = docs.get("pagador")
        if pag and pag not in cfg.CNPJS_EMPRESA and l.get("IDENTIFICACAO") != "CONFIRMADO":
            cliente, metodo, motivo = base.identificar_cliente([pag], [])
            if metodo == "CONFIRMADO":
                l["CLIENTE"], l["IDENTIFICACAO"], l["MOTIVO_IDENTIFICACAO"] = cliente, metodo, motivo
                itens = [a for a in itens if a != "Confira/escolha o CLIENTE"]
                mudou.append("cliente confirmado pelo CPF do pagador")
        if venc_div and len(linha) == 47:
            b = corrigir_vencimento_pelo_impresso(decodificar_bancario(linha), texto)
            if b.get("_alerta_venc"):
                l["DT_VENCTO"] = b["vencimento"]
                itens = [a for a in itens if not a.startswith("Vencimento do código")]
                itens.append(b["_alerta_venc"])
                mudou.append("vencimento corrigido")
        itens.append("Relido: " + (", ".join(mudou) if mudou else "sem alteração"))
        l["ALERTAS"] = " | ".join(dict.fromkeys(itens))
        if mudou:
            alterados += 1
    return alterados, tocados


# ===========================================================================
# Execuções completas
# ===========================================================================
def _processar_lote(logicas, base, caixa):
    """Processa as mensagens lógicas de UMA mensagem do Gmail e a registra uma vez."""
    resumos = []
    for m in logicas:
        if not m["anexos"] and not m.get("corpo"):
            continue
        r = processar_mensagem(m, base, caixa, registrar=False)
        if r != "sem documento fiscal":
            resumos.append(r)
    topo = logicas[0]
    if len(logicas) > 1:
        resumo = f"{len(logicas) - 1} e-mail(s) encaminhado(s) dentro desta mensagem; " + \
                 ("; ".join(resumos) if resumos else "sem documento fiscal")
    else:
        resumo = resumos[0] if resumos else "sem documento fiscal"
    caixa.registrar_mensagem(topo["id"], topo.get("data"), topo.get("remetente", ""),
                             topo.get("assunto", ""), resumo[:500])
    return resumo


_COLETA = threading.Lock()


def executar_gmail(dias=None, progresso=print):
    if not _COLETA.acquire(blocking=False):
        raise RuntimeError("Já existe uma coleta em andamento neste computador. Aguarde ela terminar.")
    try:
        return _executar_gmail(dias, progresso)
    finally:
        _COLETA.release()


def _executar_gmail(dias, progresso):
    from .gmail_cliente import obter_servico, listar_ids, obter_mensagens
    from .caixa_entrada import ArquivoEmUso

    base, caixa = BaseDados(), CaixaEntrada()
    servico = obter_servico()
    ids = listar_ids(servico, dias=dias)
    ja = caixa.mensagens_processadas()
    novos = [i for i in ids if i not in ja]
    progresso(f"{len(ids)} e-mail(s) com anexo no período; {len(novos)} ainda não processado(s).")

    for n, msg_id in enumerate(reversed(novos), start=1):   # do mais antigo para o mais novo
        tamanhos = (len(caixa.lancamentos), len(caixa.documentos), len(caixa.mensagens))
        progresso(f"[{n}/{len(novos)}] lendo mensagem {msg_id}...")
        try:
            logicas = obter_mensagens(servico, msg_id)
            resumo = _processar_lote(logicas, base, caixa)
            progresso(f"[{n}/{len(novos)}] {logicas[0]['assunto'][:60]!r}: {resumo}")
        except Exception as e:
            # desfaz o que esta mensagem acrescentou, para não duplicar na próxima execução
            del caixa.lancamentos[tamanhos[0]:], caixa.documentos[tamanhos[1]:], caixa.mensagens[tamanhos[2]:]
            progresso(f"[{n}/{len(novos)}] ERRO na mensagem {msg_id}: {e} (será tentada de novo na próxima execução)")
        if n % cfg.SALVAR_A_CADA == 0:
            try:
                caixa.salvar(base.nomes_clientes())
                progresso(f"Triagem gravada parcialmente ({n} de {len(novos)} mensagens).")
            except ArquivoEmUso:
                progresso("Gravação parcial adiada: a planilha de triagem está aberta.")

    reident, _ = reidentificar_clientes(caixa, base)
    if reident:
        progresso(f"{reident} sugestão(ões) com cliente confirmado pelo cadastro atual.")
    relidos, _ = reler_pendentes(caixa, base)
    completar_valor_danfse(caixa, base)
    completar_notas_debito(caixa, base)
    completar_manuais_nfse(caixa, base)
    preencher_nf_provavel(caixa, base)
    if relidos:
        progresso(f"{relidos} sugestão(ões) antiga(s) relida(s) e corrigida(s) com as regras atuais.")
    caixa.ignorar_chaves_referenciadas()
    caixa.anexar_complementares()
    unidos = caixa.unir_boletos_e_notas()
    if unidos:
        progresso(f"{unidos} boleto(s) unido(s) à nota fiscal correspondente recebida em outro e-mail.")
    caixa.salvar(base.nomes_clientes())
    return _resumo_final(caixa, progresso)


def executar_pasta_protegido(pasta, progresso=print):
    """executar_pasta com a mesma trava da coleta do Gmail (uma operação por vez)."""
    if not _COLETA.acquire(blocking=False):
        raise RuntimeError("Já existe uma coleta ou reprocessamento em andamento. Aguarde terminar.")
    try:
        return executar_pasta(pasta, progresso)
    finally:
        _COLETA.release()


def executar_pasta(pasta, progresso=print):
    """Mesmo processamento, para XML/PDF/ZIP/EML salvos numa pasta (ex.: recebidos por WhatsApp)."""
    from .gmail_cliente import expandir_zips, mensagens_de_eml
    base, caixa = BaseDados(), CaixaEntrada()
    ja = caixa.mensagens_processadas()
    arquivos = sorted((a for a in Path(pasta).rglob("*")
                       if a.is_file() and a.suffix.lower() in (".xml", ".pdf", ".zip", ".eml")),
                      key=lambda a: (a.parent.name, a.name))
    progresso(f"{len(arquivos)} arquivo(s) encontrados em {pasta} (inclui subpastas).")
    for arq in arquivos:
        conteudo = arq.read_bytes()
        msg_id = "ARQ-" + hashlib.sha1(conteudo).hexdigest()[:16]
        if msg_id in ja:
            continue
        ja.add(msg_id)
        if arq.suffix.lower() == ".eml":
            logicas = mensagens_de_eml(conteudo, msg_id)
        else:
            logicas = [{"id": msg_id, "origem": 0, "data": datetime.fromtimestamp(arq.stat().st_mtime),
                        "remetente": "(arquivo local)", "assunto": arq.name, "corpo": "",
                        "anexos": expandir_zips([(arq.name, conteudo)]), "_origem_path": str(arq)}]
        progresso(f"{arq.name}: {_processar_lote(logicas, base, caixa)}")
    reident, _ = reidentificar_clientes(caixa, base)
    if reident:
        progresso(f"{reident} sugestão(ões) com cliente confirmado pelo cadastro atual.")
    relidos, _ = reler_pendentes(caixa, base)
    completar_valor_danfse(caixa, base)
    completar_notas_debito(caixa, base)
    completar_manuais_nfse(caixa, base)
    preencher_nf_provavel(caixa, base)
    if relidos:
        progresso(f"{relidos} sugestão(ões) antiga(s) relida(s) e corrigida(s) com as regras atuais.")
    caixa.ignorar_chaves_referenciadas()
    caixa.anexar_complementares()
    unidos = caixa.unir_boletos_e_notas()
    if unidos:
        progresso(f"{unidos} boleto(s) unido(s) à nota fiscal correspondente recebida em outro e-mail.")
    caixa.salvar(base.nomes_clientes())
    return _resumo_final(caixa, progresso)


def reprocessar_triagem(progresso=print):
    """Releitura completa (com PDFs) da triagem, sem acessar o Gmail."""
    if not _COLETA.acquire(blocking=False):
        raise RuntimeError("Já existe uma coleta ou reprocessamento em andamento. Aguarde terminar.")
    try:
        base, caixa = BaseDados(), CaixaEntrada()
        progresso("Reavaliando clientes pelo cadastro atual...")
        rid, _ = reidentificar_clientes(caixa, base, ler_pdf=True)
        progresso(f"{rid} sugestão(ões) com cliente confirmado. Relendo PDFs pendentes...")
        rel, _ = reler_pendentes(caixa, base)
        vdf = completar_valor_danfse(caixa, base)
        if vdf:
            progresso(f"{vdf} NFS-e em PDF com valor líquido lido do documento.")
        nds = completar_notas_debito(caixa, base)
        if nds:
            progresso(f"{nds} nota(s) de débito com número, valor e vencimento lidos do documento.")
        man = completar_manuais_nfse(caixa, base)
        if man:
            progresso(f"{man} documento(s) de NFS-e em PDF com fornecedor/NF identificados.")
        nfp = preencher_nf_provavel(caixa, base)
        if nfp:
            progresso(f"{nfp} boleto(s) com número de NF identificado no próprio boleto.")
        caixa.ignorar_chaves_referenciadas()
        caixa.anexar_complementares()
        semcob = ignorar_sem_cobranca(caixa)
        if semcob:
            progresso(f"{semcob} NF-e sem cobrança (remessa/retorno/sem pagamento) marcada(s) como IGNORADO.")
        uni = caixa.unir_boletos_e_notas()
        from .importar_para_sistema import sincronizar_status
        progresso("Conferindo com as planilhas de todos os clientes...")
        sincronizar_status(caixa, base)
        caixa.salvar(base.nomes_clientes())
        progresso(f"Reprocessamento concluído: {rel} relida(s) e corrigida(s), {uni} boleto(s) unido(s) à nota.")
        return {"reidentificados": rid, "relidos": rel, "unidos": uni}
    finally:
        _COLETA.release()


def _resumo_final(caixa, progresso):
    pend = caixa.pendentes()
    sem_cliente = [l for l in pend if not l.get("CLIENTE")]
    progresso(f"Pendentes na triagem: {len(pend)} (sem cliente definido: {len(sem_cliente)}).")
    progresso(f"Planilha: {caixa.arquivo}")
    return {"pendentes": len(pend), "sem_cliente": len(sem_cliente)}


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Coleta NF-e/NFS-e/boletos do Gmail para a triagem.")
    ap.add_argument("--dias", type=int, default=None, help="quantos dias para trás buscar")
    ap.add_argument("--pasta", default=None, help="processar arquivos de uma pasta em vez do Gmail")
    a = ap.parse_args()
    if a.pasta:
        executar_pasta(a.pasta)
    else:
        executar_gmail(a.dias)
