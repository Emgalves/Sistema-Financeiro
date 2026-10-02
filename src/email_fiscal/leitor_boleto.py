# -*- coding: utf-8 -*-
"""
leitor_boleto.py - Extrai de PDFs (e de textos de e-mail):
  * linhas digitáveis de boleto bancário (47 dígitos) e de arrecadação/
    concessionária (48 dígitos), VALIDADAS pelos dígitos verificadores;
  * vencimento e valor, calculados a partir da própria linha digitável
    (não depende do layout do PDF de cada banco);
  * chave de acesso de DANFE (44 dígitos) quando a NF chega só em PDF;
  * CPF/CNPJ de beneficiário e pagador.

Limitação: PDF que é só imagem (escaneado) não tem texto -> vai para triagem manual.
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import io
import re
from datetime import date, timedelta

try:
    import pdfplumber
except ImportError:  # o módulo continua funcionando para XML
    pdfplumber = None


# ---------------------------------------------------------------------------
# Texto do PDF
# ---------------------------------------------------------------------------
def extrair_texto_pdf(conteudo: bytes, senhas=()):
    """Retorna (texto, status). status: 'OK', 'SEM_TEXTO', 'PROTEGIDO', 'ERRO', 'SEM_BIBLIOTECA'."""
    if pdfplumber is None:
        return "", "SEM_BIBLIOTECA"
    for senha in [None, *senhas]:
        try:
            with pdfplumber.open(io.BytesIO(conteudo), password=senha) as pdf:
                partes = [(p.extract_text() or "") for p in pdf.pages]
            texto = "\n".join(partes)
            return (texto, "OK") if texto.strip() else ("", "SEM_TEXTO")
        except Exception as e:
            nome = type(e).__name__.lower()
            if "password" in nome or "encrypt" in str(e).lower() or "password" in str(e).lower():
                continue
            return "", "ERRO"
    return "", "PROTEGIDO"


# ---------------------------------------------------------------------------
# Dígitos verificadores
# ---------------------------------------------------------------------------
def _mod10(numero: str) -> int:
    soma, peso = 0, 2
    for d in reversed(numero):
        p = int(d) * peso
        soma += p // 10 + p % 10
        peso = 1 if peso == 2 else 2
    return (10 - soma % 10) % 10


def _mod11_bancario(numero: str) -> int:
    soma, peso = 0, 2
    for d in reversed(numero):
        soma += int(d) * peso
        peso = 2 if peso == 9 else peso + 1
    dv = 11 - soma % 11
    return 1 if dv in (0, 10, 11) else dv


def _mod11_arrecadacao(numero: str) -> int:
    soma, peso = 0, 2
    for d in reversed(numero):
        soma += int(d) * peso
        peso = 2 if peso == 9 else peso + 1
    dv = 11 - soma % 11
    return 0 if dv in (10, 11) else dv


def _mod11_chave_nfe(chave43: str) -> int:
    soma, peso = 0, 2
    for d in reversed(chave43):
        soma += int(d) * peso
        peso = 2 if peso == 9 else peso + 1
    resto = soma % 11
    return 0 if resto in (0, 1) else 11 - resto


# ---------------------------------------------------------------------------
# Boleto bancário (47 dígitos)
# ---------------------------------------------------------------------------
def _linha47_para_barras(l):
    return l[0:4] + l[32] + l[33:47] + l[4:9] + l[10:20] + l[21:31]


def validar_linha_bancaria(l: str) -> bool:
    if len(l) != 47 or not l.isdigit():
        return False
    if _mod10(l[0:9]) != int(l[9]):
        return False
    if _mod10(l[10:20]) != int(l[20]):
        return False
    if _mod10(l[21:31]) != int(l[31]):
        return False
    barras = _linha47_para_barras(l)
    return _mod11_bancario(barras[:4] + barras[5:]) == int(barras[4])


def _vencimento_por_fator(fator: int, referencia: date):
    """O fator de vencimento 'reiniciou' em 22/02/2025 (9999 -> 1000).
    Calcula as duas possibilidades e fica com a mais próxima da data de referência."""
    if fator == 0:
        return None
    base = date(1997, 10, 7)
    candidatos = [base + timedelta(days=fator), base + timedelta(days=fator + 9000)]
    return min(candidatos, key=lambda d: abs((d - referencia).days))


def decodificar_bancario(l: str, referencia: date = None):
    referencia = referencia or date.today()
    fator = int(l[33:37])
    valor = int(l[37:47]) / 100
    return {
        "tipo": "BOLETO",
        "subtipo": "BANCARIO",
        "linha_digitavel": l,
        "linha_formatada": f"{l[0:5]}.{l[5:10]} {l[10:15]}.{l[15:21]} {l[21:26]}.{l[26:32]} {l[32]} {l[33:47]}",
        "codigo_barras": _linha47_para_barras(l),
        "banco": l[0:3],
        "vencimento": _vencimento_por_fator(fator, referencia),
        "valor": round(valor, 2) if valor > 0 else None,
    }


# ---------------------------------------------------------------------------
# Arrecadação / concessionária (48 dígitos, começa com 8)
# ---------------------------------------------------------------------------
def validar_linha_arrecadacao(l: str) -> bool:
    if len(l) != 48 or not l.isdigit() or l[0] != "8":
        return False
    ident = l[2]
    if ident in "67":
        dv_func = _mod10
    elif ident in "89":
        dv_func = _mod11_arrecadacao
    else:
        return False
    for i in range(4):
        bloco = l[i * 12:i * 12 + 11]
        if dv_func(bloco) != int(l[i * 12 + 11]):
            return False
    barras = "".join(l[i * 12:i * 12 + 11] for i in range(4))
    return dv_func(barras[:3] + barras[4:]) == int(barras[3])


def decodificar_arrecadacao(l: str, referencia: date = None):
    barras = "".join(l[i * 12:i * 12 + 11] for i in range(4))
    valor = None
    if l[2] in "68":  # valor efetivo em reais
        v = int(barras[4:15]) / 100
        valor = round(v, 2) if v > 0 else None
    return {
        "tipo": "BOLETO",
        "subtipo": "ARRECADACAO",
        "linha_digitavel": l,
        "linha_formatada": " ".join(l[i * 12:i * 12 + 12] for i in range(4)),
        "codigo_barras": barras,
        "banco": "",
        "vencimento": None,   # não existe vencimento padronizado nesse tipo
        "valor": valor,
    }


# ---------------------------------------------------------------------------
# Busca no texto
# ---------------------------------------------------------------------------
_RE_BLOCO_NUMERICO = re.compile(r"\d[\d.\s\-]{40,80}\d")


def encontrar_boletos(texto: str, referencia: date = None):
    """Todas as linhas digitáveis VÁLIDAS encontradas no texto (sem repetição)."""
    achados, vistos = [], set()
    for trecho in _RE_BLOCO_NUMERICO.findall(texto or ""):
        dig = re.sub(r"\D", "", trecho)
        for tam, valida, decod in ((47, validar_linha_bancaria, decodificar_bancario),
                                   (48, validar_linha_arrecadacao, decodificar_arrecadacao)):
            for i in range(0, len(dig) - tam + 1):
                janela = dig[i:i + tam]
                if janela not in vistos and valida(janela):
                    vistos.add(janela)
                    achados.append(decod(janela, referencia))
    return achados


def encontrar_chaves_nfe(texto: str):
    """Chaves de acesso (44 dígitos) com DV válido - aparecem no DANFE em PDF."""
    chaves, vistas = [], set()
    for trecho in re.findall(r"\d[\d\s.]{43,70}\d", texto or ""):
        dig = re.sub(r"\D", "", trecho)
        for i in range(0, len(dig) - 44 + 1):
            c = dig[i:i + 44]
            if c in vistas:
                continue
            if c[20:22] in ("55", "65") and _mod11_chave_nfe(c[:43]) == int(c[43]):
                vistas.add(c)
                chaves.append(c)
    return chaves


def dados_da_chave_nfe(chave: str):
    """A própria chave informa emitente, modelo, série e número da NF."""
    return {
        "emit_doc": chave[6:20],
        "modelo": chave[20:22],
        "serie": str(int(chave[22:25])),
        "numero": str(int(chave[25:34])),
    }


# ---------------------------------------------------------------------------
# CPF / CNPJ
# ---------------------------------------------------------------------------
def cpf_valido(c):
    c = re.sub(r"\D", "", c)
    if len(c) != 11 or c == c[0] * 11:
        return False
    for n in (9, 10):
        s = sum(int(c[i]) * (n + 1 - i) for i in range(n))
        if (s * 10 % 11) % 10 != int(c[n]):
            return False
    return True


def cnpj_valido(c):
    c = re.sub(r"\D", "", c)
    if len(c) != 14 or c == c[0] * 14:
        return False
    for n, pesos in ((12, [5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2]),
                     (13, [6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2])):
        s = sum(int(c[i]) * pesos[i] for i in range(n))
        dv = 0 if s % 11 < 2 else 11 - s % 11
        if dv != int(c[n]):
            return False
    return True


_RE_CNPJ = re.compile(r"\d{2}\.\d{3}\.\d{3}\s?/\s?\d{4}\s?-\s?\d{2}")
_RE_CPF = re.compile(r"\d{3}\.\d{3}\.\d{3}\s?-\s?\d{2}")
# "Bene...ciário" tolera a perda da ligadura "fi" na extração de texto (ex.: "Beneciário").
# "Sacador/Avalista" é o credor original -> conta como beneficiário.
_RE_PAPEL = re.compile(r"(bene\w{0,4}ci[aá]rio|cedente|sacador|pagador|sacado(?!r))", re.IGNORECASE)


# CPF/CNPJ escrito só com números logo depois do rótulo (ex.: "CPF / CNPJ: 01292830611")
_RE_DOC_ROTULADO = re.compile(r"(?:CPF|CNPJ)(?:\s*/\s*(?:CPF|CNPJ))?\s*:?\s*(\d{14}|\d{11})(?!\d)", re.IGNORECASE)


def encontrar_documentos(texto: str):
    """Retorna {'beneficiario': doc, 'pagador': doc, 'todos': [...]}.

    Cada ocorrência de CPF/CNPJ recebe um "voto" do rótulo mais próximo antes
    dela (Beneficiário/Cedente/Sacador x Pagador/Sacado). Como o texto extraído
    de alguns PDFs embaralha rótulos e valores, um mesmo número pode receber
    votos dos dois papéis: vence o papel com mais votos, e um número NUNCA fica
    com os dois papéis. Empate: CNPJ -> beneficiário, CPF -> pagador."""
    texto = texto or ""
    achados = []
    for rx, validador in ((_RE_CNPJ, cnpj_valido), (_RE_CPF, cpf_valido)):
        for m in rx.finditer(texto):
            dig = re.sub(r"\D", "", m.group())
            if validador(dig):
                achados.append((m.start(), dig))
    for m in _RE_DOC_ROTULADO.finditer(texto):
        dig = m.group(1)
        if (len(dig) == 14 and cnpj_valido(dig)) or (len(dig) == 11 and cpf_valido(dig)):
            achados.append((m.start(1), dig))
    achados.sort()

    papeis = [(m.start(), m.group().lower()) for m in _RE_PAPEL.finditer(texto)]
    votos, ordem = {}, []
    for pos, dig in achados:
        if dig not in votos:
            votos[dig] = {"beneficiario": 0, "pagador": 0}
            ordem.append(dig)
        anteriores = [nome for p, nome in papeis if p < pos and pos - p < 400]
        if anteriores:
            nome = anteriores[-1]
            papel = "pagador" if (nome.startswith("pagador") or nome == "sacado") else "beneficiario"
            votos[dig][papel] += 1

    def papel_de(dig):
        b, p = votos[dig]["beneficiario"], votos[dig]["pagador"]
        if b == p == 0:
            return None
        if b != p:
            return "beneficiario" if b > p else "pagador"
        return "beneficiario" if len(dig) == 14 else "pagador"

    resultado = {"beneficiario": "", "pagador": "", "todos": ordem}
    for dig in ordem:
        papel = papel_de(dig)
        if papel and not resultado[papel]:
            resultado[papel] = dig
    return resultado


def parece_danfe(texto: str):
    t = (texto or "").upper()
    return "DANFE" in t or "DOCUMENTO AUXILIAR DA NOTA FISCAL" in t


def parece_nfse(texto: str):
    t = (texto or "").upper()
    return "NFS-E" in t or "NOTA FISCAL DE SERVI" in t or "NOTA FISCAL ELETRÔNICA DE SERVI" in t


# ---------------------------------------------------------------------------
# NFS-e padrão nacional em PDF (DANFSe): chave de acesso de 50 dígitos
# ---------------------------------------------------------------------------
# Estrutura: município(7) ambiente(1) tipo_inscrição(1) CNPJ/CPF(14)
#            número(13) AAMM(4) código(9) DV(1)
def encontrar_chaves_nfse(texto: str):
    chaves, vistas = [], set()
    for trecho in re.findall(r"\d[\d\s.]{48,80}\d|\d{50}", texto or ""):
        dig = re.sub(r"\D", "", trecho)
        for i in range(0, len(dig) - 50 + 1):
            c = dig[i:i + 50]
            if c in vistas or c[7] not in "12" or c[8] not in "12":
                continue
            insc = c[9:23]
            ok = cnpj_valido(insc) if c[8] == "2" else (insc[:3] == "000" and cpf_valido(insc[3:]))
            mes = c[38:40]
            if ok and "01" <= mes <= "12" and int(c[23:36]) > 0:
                vistas.add(c)
                chaves.append(c)
    return chaves


def dados_da_chave_nfse(chave: str):
    insc = chave[9:23] if chave[8] == "2" else chave[12:23]
    return {"emit_doc": insc, "numero": str(int(chave[23:36])), "municipio": chave[0:7],
            "competencia": f"{chave[38:40]}/20{chave[36:38]}"}


def cnpjs_sem_pontuacao(texto: str):
    """CNPJs escritos só com números (14 dígitos isolados) e válidos."""
    return [c for c in re.findall(r"(?<!\d)\d{14}(?!\d)", texto or "") if cnpj_valido(c)]


_RE_DINHEIRO = re.compile(r"(?<![\d.,])\d{1,3}(?:\.\d{3})*,\d{2}(?![\d,])")


_RE_DATA = re.compile(r"(?<!\d)(\d{2})/(\d{2})/(\d{4})(?!\d)")


def corrigir_vencimento_pelo_impresso(boleto, texto):
    """Alguns emissores geram o fator de vencimento errado depois do reinício do
    fator (22/02/2025): o código aponta para uma data exatamente 1000 (ou 2000...)
    dias antes/depois da data impressa. Se o vencimento do código NÃO está impresso
    e existe UMA data impressa que difere dele em múltiplo exato de 1000 dias,
    usa a data impressa e registra o motivo em boleto['_alerta_venc']."""
    venc = boleto.get("vencimento")
    if not venc or not texto or venc.strftime("%d/%m/%Y") in texto:
        return boleto
    candidatas = set()
    for d, m, a in _RE_DATA.findall(texto):
        try:
            dt = date(int(a), int(m), int(d))
        except ValueError:
            continue
        dif = (dt - venc).days
        if dif and dif % 1000 == 0 and abs(dif) <= 20000:
            candidatas.add(dt)
    if len(candidatas) == 1:
        impressa = candidatas.pop()
        boleto["_vencimento_codigo"] = venc
        boleto["vencimento"] = impresso = impressa
        boleto["_alerta_venc"] = (f"Vencimento do código de barras ({venc:%d/%m/%Y}) difere do impresso "
                                  f"({impresso:%d/%m/%Y}) em {abs((impresso - venc).days)} dias - erro do emissor "
                                  f"no fator de vencimento; usado o impresso")
    return boleto


def conferir_boleto_com_texto(boleto, texto):
    """A linha digitável é a fonte do valor e do vencimento. Aqui só se confere se
    o que ela diz também está impresso no boleto; se não estiver, devolve alertas."""
    alertas = [boleto["_alerta_venc"]] if boleto.get("_alerta_venc") else []
    if not texto:
        return alertas
    if boleto.get("valor") is not None:
        impressos = {round(float(v.replace(".", "").replace(",", ".")), 2) for v in _RE_DINHEIRO.findall(texto)}
        if impressos and boleto["valor"] not in impressos:
            v = f"{boleto['valor']:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            alertas.append(f"ATENÇÃO: o código de barras vale R$ {v}, mas esse valor não aparece "
                           f"impresso no boleto - confira antes de pagar")
    if boleto.get("vencimento") and boleto["vencimento"].strftime("%d/%m/%Y") not in texto:
        alertas.append(f"Vencimento do código ({boleto['vencimento']:%d/%m/%Y}) não aparece impresso no boleto - confira")
    return alertas
