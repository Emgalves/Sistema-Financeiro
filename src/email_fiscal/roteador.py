# -*- coding: utf-8 -*-
"""
roteador.py - Decide A QUAL CLIENTE (obra) e A QUAL FORNECEDOR cada documento
pertence, usando as planilhas que você já mantém:

  Clientes.xlsx        -> coluna CPF (e CNO, Nome, Logradouro/Numero)
  Base_Fornecedores    -> coluna CNPJ/CPF (NOME, CATEGORIA, STATUS)
  <CLIENTE>.xlsx       -> para saber se o documento JÁ foi lançado

Regras de identificação do cliente (em ordem):
  1. CPF/CNPJ do destinatário da NF-e / tomador da NFS-e / pagador do boleto
     bate com o CPF de UM cliente                          -> CONFIRMADO
     (se bater com vários, prefere quem não tem Data Final; persistindo -> AMBIGUO)
  2. Nome completo do cliente aparece no assunto, corpo do e-mail,
     informações complementares ou endereço de entrega      -> SUGERIDO
  3. Número do CNO aparece nesses textos                    -> SUGERIDO
  4. Nada disso                                             -> NAO_IDENTIFICADO
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import re
import unicodedata
from datetime import datetime, date

from openpyxl import load_workbook

from . import config_email as cfg


def so_digitos(s):
    return re.sub(r"\D", "", str(s or ""))


def normalizar_texto(s):
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s.upper()).strip()


def _chave_doc(doc):
    """Normaliza CPF/CNPJ para comparação (Excel às vezes perde zeros à esquerda)."""
    d = so_digitos(doc)
    if not d:
        return ""
    if len(d) <= 11:
        return d.zfill(11)
    return d.zfill(14)


def _normalizar_cabecalho(h):
    return normalizar_texto(h).replace(":", "").strip()


def _ler_tabela(arquivo, aba=None):
    wb = load_workbook(arquivo, read_only=True, data_only=True)
    ws = wb[aba] if aba and aba in wb.sheetnames else wb.worksheets[0]
    linhas = ws.iter_rows(values_only=True)
    cab = [_normalizar_cabecalho(c) for c in next(linhas)]
    dados = [dict(zip(cab, lin)) for lin in linhas if any(v not in (None, "") for v in lin)]
    wb.close()
    return dados


class BaseDados:
    def __init__(self):
        self.clientes = self._carregar_clientes()
        self.fornecedores = self._carregar_fornecedores()
        self._cache_lancados = {}

    # ------------------------------------------------------------------ clientes
    def _carregar_clientes(self):
        clientes = []
        for lin in _ler_tabela(cfg.ARQUIVO_CLIENTES, "Clientes"):
            nome = str(lin.get("NOME") or "").strip()
            if not nome:
                continue
            clientes.append({
                "nome": nome,
                "nome_norm": normalizar_texto(nome),
                "doc": _chave_doc(lin.get("CPF")),
                "cno": so_digitos(lin.get("CNO")),
                "encerrado": lin.get("DATA FINAL") not in (None, "", " "),
            })
        return clientes

    def nomes_clientes(self):
        return sorted(c["nome"] for c in self.clientes)

    def identificar_cliente(self, documentos, textos):
        """documentos: CPF/CNPJ candidatos (destinatário/tomador/pagador).
        textos: assunto, corpo, infCpl, endereço...
        Retorna (nome_cliente, metodo, observacao)."""
        # Regra 1 - documento
        for doc in documentos:
            k = _chave_doc(doc)
            if not k:
                continue
            achados = [c for c in self.clientes if c["doc"] and c["doc"] == k]
            if len(achados) > 1:
                ativos = [c for c in achados if not c["encerrado"]]
                achados = ativos or achados
            if len(achados) == 1:
                return achados[0]["nome"], "CONFIRMADO", f"documento {doc} = CPF do cliente"
            if len(achados) > 1:
                nomes = "; ".join(c["nome"] for c in achados)
                return "", "AMBIGUO", f"documento {doc} pertence a vários clientes: {nomes}"

        texto = normalizar_texto(" ".join(t for t in textos if t))
        texto_dig = so_digitos(" ".join(t for t in textos if t))

        # Regra 2 - nome completo no texto
        por_nome = [c for c in self.clientes
                    if len(c["nome_norm"]) >= 8 and c["nome_norm"] in texto]
        if len(por_nome) > 1:
            ativos = [c for c in por_nome if not c["encerrado"]]
            por_nome = ativos or por_nome
        if len(por_nome) == 1:
            return por_nome[0]["nome"], "SUGERIDO", "nome do cliente encontrado no e-mail/nota"

        # Regra 3 - CNO no texto
        por_cno = [c for c in self.clientes if len(c["cno"]) >= 10 and c["cno"] in texto_dig]
        if len(por_cno) == 1:
            return por_cno[0]["nome"], "SUGERIDO", "CNO da obra encontrado no e-mail/nota"

        return "", "NAO_IDENTIFICADO", "nenhum CPF/nome/CNO de cliente encontrado"

    # --------------------------------------------------------------- fornecedores
    def _carregar_fornecedores(self):
        base = {}
        for lin in _ler_tabela(cfg.ARQUIVO_FORNECEDORES, "Fornecedores"):
            k = _chave_doc(lin.get("CNPJ/CPF"))
            if not k:
                continue
            reg = {
                "doc": k,
                "nome": str(lin.get("NOME") or lin.get("RAZAO SOCIAL") or "").strip(),
                "razao": str(lin.get("RAZAO SOCIAL") or "").strip(),
                "categoria": str(lin.get("CATEGORIA") or "").strip(),
                "status": str(lin.get("STATUS") or "").strip().upper(),
                "dados_bancarios": str(lin.get("DADOS BANCARIOS") or "").strip(),
            }
            # se houver duplicado, fica o ATIVO
            if k not in base or (base[k]["status"] != "ATIVO" and reg["status"] == "ATIVO"):
                base[k] = reg
        return base

    def buscar_fornecedor(self, doc):
        return self.fornecedores.get(_chave_doc(doc))

    def buscar_mesma_empresa(self, doc):
        """Outra filial/matriz da mesma empresa (8 primeiros dígitos do CNPJ)."""
        k = _chave_doc(doc)
        if len(k) != 14:
            return None
        achados = [f for d, f in self.fornecedores.items() if len(d) == 14 and d[:8] == k[:8]]
        achados.sort(key=lambda f: f["status"] != "ATIVO")
        return achados[0] if achados else None

    # --------------------------------------------------- já lançado no cliente?
    def _lancamentos_cliente(self, cliente):
        if cliente in self._cache_lancados:
            return self._cache_lancados[cliente]
        arquivo = cfg.PASTA_CLIENTES / f"{cliente}.xlsx"
        linhas = []
        if arquivo.exists():
            try:
                linhas = _ler_tabela(arquivo, cfg.ABA_DADOS_CLIENTE)
            except Exception:
                linhas = []
        self._cache_lancados[cliente] = linhas
        return linhas

    def ja_lancado(self, cliente, id_sugestao=None, doc=None, nf=None, valor=None, vencimento=None):
        """Verifica na planilha do cliente. Retorna (True/False, motivo).

        1. Marcador "IMPORTADO EMAIL #id" na OBSERVAÇÃO -> já importado por esta rotina.
        2. MESMA EMPRESA (mesmo CNPJ ou outra filial: 8 primeiros dígitos iguais)
           + mesmo valor + mesma NF (se houver) + mesmo vencimento (se houver)."""
        if not cliente:
            return False, ""
        marcador = f"{cfg.PREFIXO_OBSERVACAO} #{int(float(id_sugestao))}" if id_sugestao else None
        padrao_marcador = re.compile(re.escape(marcador) + r"(?!\d)") if marcador else None
        k_doc = _chave_doc(doc)
        nf_txt = _nf_texto(nf)

        for lin in self._lancamentos_cliente(cliente):
            obs = str(lin.get("OBSERVACAO") or "")
            if padrao_marcador and padrao_marcador.search(obs):
                return True, f"marcador '{marcador}' encontrado (ID_LANCAMENTO {lin.get('ID_LANCAMENTO')})"
            if not k_doc or str(lin.get("STATUS") or "ATIVO").upper() != "ATIVO":
                continue
            k_lin = _chave_doc(lin.get("CNPJ_CPF"))
            if _raiz(k_lin) != _raiz(k_doc):
                continue
            # 3. Nota SEM vencimento (NF só em PDF, NFS-e, XML sem duplicatas): o mesmo
            #    emitente não repete número de NF, então mesma NF + mesmo CNPJ = mesmo documento,
            #    desde que o valor seja igual ou desconhecido.
            if (nf_txt and not vencimento and k_lin == k_doc and _nf_texto(lin.get("NF")) == nf_txt
                    and (valor is None or abs((_para_float(lin.get("VALOR")) or 0) - valor) <= 0.01)):
                return True, (f"mesma NF {nf_txt} do mesmo emitente já lançada "
                              f"(ID_LANCAMENTO {lin.get('ID_LANCAMENTO')}, R$ {lin.get('VALOR')})")
            if not _mesma_conta(lin, nf_txt, valor, vencimento):
                continue
            quem = "mesmo fornecedor" if k_lin == k_doc else f"outra filial da mesma empresa (CNPJ {k_lin})"
            return True, (f"{quem}, valor{', NF' if nf_txt else ''}"
                          f"{' e vencimento' if vencimento else ''} iguais (ID_LANCAMENTO {lin.get('ID_LANCAMENTO')})")
        return False, ""

    def possivel_duplicidade(self, cliente, doc=None, nf=None, valor=None, vencimento=None):
        """Mesma NF, mesmo valor e mesmo vencimento, mas lançada com o CNPJ de
        OUTRA empresa. Não descarta nada: só devolve o texto do alerta ("" se não houver)."""
        nf_txt = _nf_texto(nf)
        k_doc = _chave_doc(doc)
        if not cliente or not nf_txt:
            return ""
        # Nota sem vencimento cuja NF já foi lançada pela mesma empresa com OUTRO valor
        if not vencimento and k_doc:
            for lin in self._lancamentos_cliente(cliente):
                if str(lin.get("STATUS") or "ATIVO").upper() != "ATIVO":
                    continue
                if _raiz(_chave_doc(lin.get("CNPJ_CPF"))) == _raiz(k_doc) and _nf_texto(lin.get("NF")) == nf_txt:
                    return (f"POSSÍVEL DUPLICIDADE: NF {nf_txt} já lançada com outro valor (ID_LANCAMENTO {lin.get('ID_LANCAMENTO')}, "
                            f"R$ {_para_float(lin.get('VALOR'))}, venc. "
                            f"{(_para_data(lin.get('DT_VENCTO')) or '-') if not _para_data(lin.get('DT_VENCTO')) else _para_data(lin.get('DT_VENCTO')).strftime('%d/%m/%Y')})"
                            f" - confira parcelas/valor")
        if valor is None or not vencimento:
            return ""
        for lin in self._lancamentos_cliente(cliente):
            if str(lin.get("STATUS") or "ATIVO").upper() != "ATIVO":
                continue
            k_lin = _chave_doc(lin.get("CNPJ_CPF"))
            if k_doc and _raiz(k_lin) == _raiz(k_doc):
                continue  # mesma empresa: é caso de ja_lancado, não deste alerta
            if _mesma_conta(lin, nf_txt, valor, vencimento, exigir_vencimento=True):
                return (f"POSSÍVEL DUPLICIDADE com ID_LANCAMENTO {lin.get('ID_LANCAMENTO')} "
                        f"({lin.get('NOME')}, CNPJ {k_lin}): mesma NF, valor e vencimento - confira o CNPJ")
        return ""


def _raiz(k):
    """CNPJ: 8 primeiros dígitos (a empresa, qualquer filial). CPF: o próprio CPF."""
    return k[:8] if len(k) == 14 else k


def _mesma_conta(lin, nf_txt, valor, vencimento, exigir_vencimento=False):
    v_lin = _para_float(lin.get("VALOR"))
    if valor is None or v_lin is None or abs(v_lin - valor) > 0.01:
        return False
    if nf_txt:
        nf_lin = _nf_texto(lin.get("NF"))
        ref = normalizar_texto(lin.get("REFERENCIA"))
        if nf_lin != nf_txt and not re.search(rf"\bNF\s*{re.escape(nf_txt)}\b", ref):
            return False
    if vencimento:  # parcelas da mesma NF têm o mesmo número: o vencimento diferencia
        d_lin = _para_data(lin.get("DT_VENCTO"))
        if d_lin and d_lin != vencimento:
            return False
        if exigir_vencimento and not d_lin:
            return False
    return True


def _nf_texto(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if re.fullmatch(r"\d+\.0", s):   # resíduo do pandas gravado como texto
        s = s[:-2]
    return s.lstrip("0")


def _para_float(v):
    try:
        return round(float(str(v).replace(",", ".")), 2)
    except (TypeError, ValueError):
        return None


def _para_data(v):
    """Suas planilhas têm DT_VENCTO ora como texto '06/09/2024', ora como data/serial."""
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    if isinstance(v, (int, float)):
        from datetime import timedelta
        return date(1899, 12, 30) + timedelta(days=int(v))
    try:
        return datetime.strptime(str(v).strip()[:10], "%d/%m/%Y").date()
    except ValueError:
        return None
