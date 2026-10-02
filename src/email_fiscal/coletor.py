# -*- coding: utf-8 -*-
"""
coletor.py - Lê o Gmail (ou uma pasta local), identifica NF-e, NFS-e e boletos,
e grava SUGESTÕES de lançamento em Documentos_Recebidos.xlsx.

Uso pela linha de comando (a partir da pasta raiz do projeto, a que contém 'src'):
    python -m src.email_fiscal.coletor               # Gmail, últimos N dias do config
    python -m src.email_fiscal.coletor --dias 90     # Gmail, últimos 90 dias
    python -m src.email_fiscal.coletor --pasta "C:\\caminho\\com\\xmls_e_pdfs"
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import argparse
import hashlib
import re
from datetime import datetime, date
from pathlib import Path

from . import config_email as cfg
from .caixa_entrada import CaixaEntrada
from .leitor_boleto import (
    cnpjs_sem_pontuacao, conferir_boleto_com_texto, corrigir_vencimento_pelo_impresso, dados_da_chave_nfe, dados_da_chave_nfse, encontrar_boletos,
    encontrar_chaves_nfe, encontrar_chaves_nfse, encontrar_documentos, extrair_texto_pdf,
    parece_danfe, parece_nfse,
)
from .leitor_nfe import ler_xml
from .roteador import BaseDados

PALAVRAS_FISCAIS = re.compile(r"boleto|nf|nota|fatura|danfe|cobran|duplicata", re.IGNORECASE)


# ===========================================================================
# Leitura dos anexos de UMA mensagem
# ===========================================================================
def _ler_anexos(msg, caixa):
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
            texto, status = extrair_texto_pdf(conteudo, cfg.SENHAS_PDF)
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
                    b.update(_arquivo=nome, _bytes=conteudo, _docs=docs_texto, _texto=texto)
                    r["boletos"].append(b)
            # nota em PDF (pode estar no mesmo arquivo do boleto)
            nota_pdf = False
            if parece_danfe(texto):
                for chave in encontrar_chaves_nfe(texto):
                    nota_pdf = True
                    r["danfes"].append({"chave": chave, "modelo_doc": "NFE", **dados_da_chave_nfe(chave),
                                        "_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto,
                                        "_texto": texto})
            if not boletos and not nota_pdf:
                for chave in encontrar_chaves_nfse(texto + " " + nome):
                    nota_pdf = True
                    r["danfes"].append({"chave": chave, "modelo_doc": "NFSE", **dados_da_chave_nfse(chave),
                                        "_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto,
                                        "_texto": texto})
            if not boletos and not nota_pdf and parece_nfse(texto):
                r["nfse_pdf"].append({"_arquivo": nome, "_bytes": conteudo, "_docs": docs_texto})

    # linha digitável colada no corpo do e-mail (sem PDF)
    if not r["boletos"] and msg.get("corpo"):
        docs_corpo = encontrar_documentos(msg["corpo"])
        docs_corpo["sem_pontuacao"] = cnpjs_sem_pontuacao(msg["corpo"])
        for b in encontrar_boletos(msg["corpo"], ref):
            if novo(b["linha_digitavel"]):
                b.update(_arquivo="", _bytes=None, _docs=docs_corpo, _texto="")
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
        return doc, f"CNPJ não impresso no boleto: deduzido pelo nome ({achados[doc]['razao']}) - confira"
    docs = sorted(achados)
    nome = achados[docs[0]]["razao"]
    return "", (f"CNPJ não impresso no boleto. Beneficiário pelo nome: {nome}, que tem mais de um "
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
                                        valor=lanc.get("VALOR"), vencimento=lanc.get("DT_VENCTO"))
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
    lido = _ler_anexos(msg, caixa)
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
        lanc = {"REFERENCIA": "SERVIÇOS" if servico else "MATERIAIS", "NF": d["numero"], "VALOR": None,
                "DT_VENCTO": None, "DADOS_BANCARIOS": "", "LINHA_DIGITAVEL": "", "CHAVE_DOC": d["chave"],
                "ARQUIVOS": arq}
        docs_cli = [x for x in d["_docs"]["todos"] if x != d["emit_doc"]]
        aviso = ("NFS-e só em PDF: CNPJ e número lidos da chave de acesso - preencha valor e vencimento"
                 if servico else
                 "NF só em PDF (sem XML) - peça o XML ao fornecedor ou preencha valor/vencimento")
        _completar(lanc, d["emit_doc"], "", "NFSE" if servico else "NFE", docs_cli,
                   textos_base + [d.get("_texto", "")], msg, base, [aviso])
        caixa.adicionar_lancamento(lanc)
        criados += 1

    # ---- 5. PDFs fiscais que não deu para ler ------------------------------------
    manuais = [(nome, cont, st, []) for nome, cont, st in lido["problemas"]]
    manuais += [(x["_arquivo"], x["_bytes"], "NFSE_PDF", x["_docs"]["todos"]) for x in lido["nfse_pdf"]]
    for nome, conteudo, status, docs_pdf in manuais:
        arq = _salvar_arquivo(msg, f"MANUAL_{status}", nome, conteudo)
        motivo = {"SEM_TEXTO": "PDF é imagem (escaneado) - leitura manual",
                  "PROTEGIDO": "PDF protegido por senha - inclua a senha em SENHAS_PDF ou lance manualmente",
                  "NFSE_PDF": "NFS-e só em PDF - peça o XML ou lance manualmente",
                  }.get(status, f"PDF não lido ({status})")
        lanc = {"REFERENCIA": "", "NF": "", "VALOR": None, "DT_VENCTO": None, "DADOS_BANCARIOS": "",
                "LINHA_DIGITAVEL": "", "CHAVE_DOC": "", "ARQUIVOS": arq}
        _completar(lanc, "", "", "NFSE" if status == "NFSE_PDF" else "BOLETO", docs_pdf, textos_base,
                   msg, base, [motivo])
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
        sem_cnpj = not str(l.get("CNPJ_CPF") or "").strip()
        venc_div = "Vencimento do código" in alertas and "não aparece impresso" in alertas
        if not linha or "Relido" in alertas or not (sem_cnpj or venc_div):
            continue
        texto = ""
        for caminho in str(l.get("ARQUIVOS") or "").split(" ; "):
            caminho = caminho.strip()
            if caminho.lower().endswith(".pdf") and Path(caminho).exists():
                t, st = extrair_texto_pdf(Path(caminho).read_bytes(), cfg.SENHAS_PDF)
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


def executar_gmail(dias=None, progresso=print):
    from .gmail_cliente import obter_servico, listar_ids, obter_mensagens

    base, caixa = BaseDados(), CaixaEntrada()
    servico = obter_servico()
    ids = listar_ids(servico, dias=dias)
    ja = caixa.mensagens_processadas()
    novos = [i for i in ids if i not in ja]
    progresso(f"{len(ids)} e-mail(s) com anexo no período; {len(novos)} ainda não processado(s).")

    for n, msg_id in enumerate(reversed(novos), start=1):   # do mais antigo para o mais novo
        tamanhos = (len(caixa.lancamentos), len(caixa.documentos), len(caixa.mensagens))
        try:
            logicas = obter_mensagens(servico, msg_id)
            resumo = _processar_lote(logicas, base, caixa)
            progresso(f"[{n}/{len(novos)}] {logicas[0]['assunto'][:60]!r}: {resumo}")
        except Exception as e:
            # desfaz o que esta mensagem acrescentou, para não duplicar na próxima execução
            del caixa.lancamentos[tamanhos[0]:], caixa.documentos[tamanhos[1]:], caixa.mensagens[tamanhos[2]:]
            progresso(f"[{n}/{len(novos)}] ERRO na mensagem {msg_id}: {e} (será tentada de novo na próxima execução)")

    relidos, _ = reler_pendentes(caixa, base)
    if relidos:
        progresso(f"{relidos} sugestão(ões) antiga(s) relida(s) e corrigida(s) com as regras atuais.")
    unidos = caixa.unir_boletos_e_notas()
    if unidos:
        progresso(f"{unidos} boleto(s) unido(s) à nota fiscal correspondente recebida em outro e-mail.")
    caixa.salvar(base.nomes_clientes())
    return _resumo_final(caixa, progresso)


def executar_pasta(pasta, progresso=print):
    """Mesmo processamento, para XML/PDF/ZIP/EML salvos numa pasta (ex.: recebidos por WhatsApp)."""
    from .gmail_cliente import expandir_zips, mensagens_de_eml
    base, caixa = BaseDados(), CaixaEntrada()
    ja = caixa.mensagens_processadas()
    for arq in sorted(Path(pasta).glob("*")):
        if arq.suffix.lower() not in (".xml", ".pdf", ".zip", ".eml"):
            continue
        conteudo = arq.read_bytes()
        msg_id = "ARQ-" + hashlib.sha1(conteudo).hexdigest()[:16]
        if msg_id in ja:
            continue
        if arq.suffix.lower() == ".eml":
            logicas = mensagens_de_eml(conteudo, msg_id)
        else:
            logicas = [{"id": msg_id, "origem": 0, "data": datetime.fromtimestamp(arq.stat().st_mtime),
                        "remetente": "(arquivo local)", "assunto": arq.name, "corpo": "",
                        "anexos": expandir_zips([(arq.name, conteudo)])}]
        progresso(f"{arq.name}: {_processar_lote(logicas, base, caixa)}")
    relidos, _ = reler_pendentes(caixa, base)
    if relidos:
        progresso(f"{relidos} sugestão(ões) antiga(s) relida(s) e corrigida(s) com as regras atuais.")
    unidos = caixa.unir_boletos_e_notas()
    if unidos:
        progresso(f"{unidos} boleto(s) unido(s) à nota fiscal correspondente recebida em outro e-mail.")
    caixa.salvar(base.nomes_clientes())
    return _resumo_final(caixa, progresso)


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
