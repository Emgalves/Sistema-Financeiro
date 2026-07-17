# src/leitura_guias.py
"""
Extração de dados de guias fiscais (GFD/FGTS Digital e DARF) a partir de PDFs nativos.
Sem dependência de UI — testável isoladamente.
"""
import re
import pdfplumber


class GuiaNaoReconhecida(Exception):
    pass


def extrair_dados_guia(caminho_pdf):
    with pdfplumber.open(caminho_pdf) as pdf:
        texto = "\n".join(p.extract_text() or "" for p in pdf.pages)

    texto_normalizado = re.sub(r'\s+', ' ', texto)

    if "GFD - Guia do FGTS Digital" in texto_normalizado:
        return _extrair_gfd(texto)
    elif "Documento de Arrecadação de Receitas Federais" in texto_normalizado:
        return _extrair_darf(texto)
    else:
        raise GuiaNaoReconhecida(
            "PDF não reconhecido como GFD ou DARF. Verifique se é o documento correto."
        )

CNPJ_FGTS = '00.360.305/0001-04'
CNPJ_INSS_IRRF = '00.394.460/0001-41'

NOME_FORNECEDOR_POR_TIPO = {
    'FGTS': 'FGTS',        # <-- CONFIRMAR: nome exato cadastrado em Configurações
    'DARF': 'INSS/IRRF',   # <-- CONFIRMAR: nome exato cadastrado em Configurações
}

MESES_PT = {
    'JANEIRO': 1, 'FEVEREIRO': 2, 'MARÇO': 3, 'ABRIL': 4, 'MAIO': 5, 'JUNHO': 6,
    'JULHO': 7, 'AGOSTO': 8, 'SETEMBRO': 9, 'OUTUBRO': 10, 'NOVEMBRO': 11, 'DEZEMBRO': 12,
}

def competencia_para_mm_aaaa(texto_competencia, tipo):
    """Normaliza a competência extraída para o formato MM/AAAA, independente do
    formato de origem (a GFD já vem como '05/2026'; o DARF vem como 'Maio/2026')."""
    if tipo == 'FGTS':
        return texto_competencia  # já vem como MM/AAAA

    # DARF: "Maio/2026" -> "05/2026"
    partes = texto_competencia.split('/')
    if len(partes) != 2:
        return None
    mes_nome, ano = partes
    mes_num = MESES_PT.get(mes_nome.strip().upper())
    if not mes_num:
        return None
    return f"{mes_num:02d}/{ano.strip()}"

def cnpj_esperado_para_tipo(tipo):
    return {'FGTS': CNPJ_FGTS, 'DARF': CNPJ_INSS_IRRF}.get(tipo)

def _extrair_gfd(texto):
    # Linha do topo: CPF/CNPJ + Nome + Data de vencimento
    # (ex: "298.493.306-00 EDIGAR BATISTA FERREIRA 19/06/2026")
    m_topo = re.search(
        r'(\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})\s+.+?\s+(\d{2}/\d{2}/\d{4})',
        texto
    )

    # Linha da guia: Identificador (16 dígitos + dígito verificador) + Tag + Competência + Valor
    # (ex: "0126061343835794-6 FGTS MENSAL 05/2026 340,08")
    m_guia = re.search(
        r'(\d{16}-\d)\s+FGTS[^\d\n]*?(\d{2}/\d{4})\s+([\d\.]+,\d{2})',
        texto
    )

    cnpj_empregador = m_topo.group(1) if m_topo else None
    vencimento      = m_topo.group(2) if m_topo else None
    identificador   = m_guia.group(1) if m_guia else None
    competencia     = m_guia.group(2) if m_guia else None
    valor           = _para_float(m_guia.group(3)) if m_guia else None

    campos_faltando = [nome for nome, v in [('valor', valor), ('vencimento', vencimento)] if v is None]

    return {
        'tipo': 'FGTS',
        'cnpj_agencia_arrecadadora': '00.360.305/0001-04',
        'cnpj_empregador': cnpj_empregador,
        'valor': valor,
        'vencimento': vencimento,
        'competencia': competencia,
        'identificador': identificador,
        'campos_faltando': campos_faltando,
    }

def _extrair_darf(texto):
    cpf_cnpj_match = re.search(
        r'(\d{3}\.\d{3}\.\d{3}-\d{2}|\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto
    )

    # Captura a linha de dados: "Maio/2026  19/06/2026  07.16.26164.5952504-8  19/06/2026"
    # na mesma ordem dos rótulos "Período de Apuração / Data de Vencimento /
    # Número do Documento / Pagar esse documento até"
    m_dados = re.search(
        r'([A-ZÇÁÉÍÓÚÂÊÔ][a-zçáéíóúâêô]+/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+([\d\.]+-\d)\s+(\d{2}/\d{2}/\d{4})',
        texto
    )

    # O valor total é sempre o último número em formato monetário do
    # documento — testado tanto neste layout simplificado quanto no
    # layout real com múltiplos códigos (a linha "Totais" também vem por
    # último no texto extraído)
    valores_monetarios = re.findall(r'[\d\.]+,\d{2}', texto)

    periodo_apuracao = m_dados.group(1) if m_dados else None
    vencimento       = m_dados.group(2) if m_dados else None
    numero_documento = m_dados.group(3) if m_dados else None
    valor            = _para_float(valores_monetarios[-1]) if valores_monetarios else None

    campos_faltando = [nome for nome, v in [('valor', valor), ('vencimento', vencimento)] if v is None]

    return {
        'tipo': 'DARF',
        'cnpj_agencia_arrecadadora': CNPJ_INSS_IRRF,
        'cpf_cnpj_contribuinte': cpf_cnpj_match.group(1) if cpf_cnpj_match else None,
        'valor': valor,
        'vencimento': vencimento,
        'numero_documento': numero_documento,
        'periodo_apuracao': periodo_apuracao,
        'campos_faltando': campos_faltando,
    }


def _para_float(valor_str):
    return float(valor_str.replace('.', '').replace(',', '.'))