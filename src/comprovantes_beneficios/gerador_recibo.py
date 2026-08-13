# src/comprovantes_beneficios/gerador_recibo.py
"""
Gera o PDF do comprovante diretamente em Python, com reportlab.

Texto e layout revisados a partir de RECIBOS.docx (modelo definitivo
fornecido pelo usuário) e da lista de ajustes visuais pedida depois de
ver os primeiros PDFs reais:
  - sem caixa/borda
  - endereço sem negrito, com vírgula ao final
  - linha de local/data sem negrito; dia em branco (cestas, preenchimento
    manual) ou vencimento real do lançamento (transporte/café)
  - assinatura (linha + nome + CPF) alinhada à esquerda, sem negrito,
    sem o rótulo "Nome:"
  - "VALE TRANSPORTE" / "VALE CAFÉ" em maiúsculo e negrito no corpo
  - valor por extenso em minúsculo, sem negrito
"""

from datetime import date
from pathlib import Path

from reportlab.lib.pagesizes import letter
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import ParagraphStyle

from .normalizacao import valor_por_extenso, formatar_valor_monetario, formatar_cpf
from .dados_candidatos import (
    Candidato, DadosPagador,
    BENEFICIO_TRANSPORTE, BENEFICIO_CAFE, BENEFICIO_CESTA_BASICA, BENEFICIO_CESTA_NATAL,
)

try:
    from src.feriados import calcular_enesimo_dia_util
except ImportError:
    from feriados import calcular_enesimo_dia_util

_MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro',
}


class ErroGeracaoRecibo(RuntimeError):
    pass


def _mes_extenso(mes: int) -> str:
    return _MESES_PT[mes]


def _titulo(tipo: str) -> str:
    return {
        BENEFICIO_CESTA_BASICA: 'RECIBO DE ENTREGA DE CESTA BASICA',
        BENEFICIO_CESTA_NATAL: 'RECIBO DE ENTREGA DE CESTA DE NATAL',
        BENEFICIO_TRANSPORTE: 'RECIBO DE VALE TRANSPORTE',
        BENEFICIO_CAFE: 'RECIBO DE VALE CAFÉ',
    }[tipo]


def _corpo_html(tipo: str, pagador: DadosPagador, candidato: Candidato) -> str:
    ano_comp, mes_comp = candidato.competencia.split('-')
    mes_extenso = _mes_extenso(int(mes_comp))
    abertura = (
        f"Pelo presente declaro para os devidos fins, que recebi de "
        f"<b>{pagador.nome}</b>, {pagador.endereco},"
    )

    if tipo == BENEFICIO_CESTA_BASICA:
        return f"{abertura} a <b>CESTA BASICA</b> referente ao mês de {mes_extenso} de {ano_comp}."

    if tipo == BENEFICIO_CESTA_NATAL:
        return (
            f"{abertura} a <b>CESTA DE NATAL</b>, como gratificação pela "
            f"prestação de serviços durante o ano de {ano_comp}."
        )

    if tipo in (BENEFICIO_TRANSPORTE, BENEFICIO_CAFE):
        if candidato.valor is None:
            raise ErroGeracaoRecibo(
                f"Candidato {candidato.nome} não tem valor definido para {tipo}."
            )
        valor_fmt = formatar_valor_monetario(candidato.valor)
        valor_ext = valor_por_extenso(candidato.valor).lower()
        if tipo == BENEFICIO_TRANSPORTE:
            return (
                f"{abertura} o valor de R$ {valor_fmt} ({valor_ext}), referente ao "
                f"<b>VALE TRANSPORTE</b> para uso em {mes_extenso} de {ano_comp}, "
                f"juntamente com minha remuneração mensal."
            )
        else:
            return (
                f"{abertura} o valor de R$ {valor_fmt} ({valor_ext}), referente ao "
                f"<b>VALE CAFÉ</b> a ser consumido em {mes_extenso} de {ano_comp}, "
                f"juntamente com minha remuneração mensal."
            )

    raise ErroGeracaoRecibo(f"Tipo de benefício desconhecido: {tipo}")


def _dia_assinatura(tipo: str, candidato: Candidato) -> str:
    """
    Cestas: espaço em branco para preenchimento manual do dia na
    assinatura. Transporte/Café: dia do vencimento real do lançamento
    (aba Dados) e, só na ausência dele, o 5º dia útil calculado.
    """
    if tipo in (BENEFICIO_CESTA_BASICA, BENEFICIO_CESTA_NATAL):
        return '&nbsp;' * 12  # espaço visível para preenchimento manual do dia

    if candidato.data_vencimento:
        return str(candidato.data_vencimento.day)

    ano_comp, mes_comp = candidato.competencia.split('-')
    dia_calculado = calcular_enesimo_dia_util(int(ano_comp), int(mes_comp))
    return str(dia_calculado.day)


# ============================================================================
# Estilos (sem caixa, sem negrito na assinatura, alinhamento à esquerda)
# ============================================================================

_ESTILO_TITULO = ParagraphStyle(
    'TituloRecibo', fontName='Helvetica-Bold', fontSize=13,
    alignment=TA_CENTER, spaceAfter=50,
)
_ESTILO_CORPO = ParagraphStyle(
    'CorpoRecibo', fontName='Helvetica', fontSize=11, leading=16,
    alignment=TA_LEFT, spaceAfter=24,
)
_ESTILO_DATA = ParagraphStyle(
    'DataRecibo', fontName='Helvetica', fontSize=11,
    alignment=TA_LEFT, spaceAfter=40,
)
_ESTILO_ASSINATURA = ParagraphStyle(
    'AssinaturaRecibo', fontName='Helvetica', fontSize=11,
    alignment=TA_LEFT, leading=15,
)


def _nome_arquivo_pdf(beneficio: str, candidato: Candidato) -> str:
    return f"RECIBO_{beneficio}_{candidato.cpf}_{candidato.competencia}.pdf"


def gerar_recibo_pdf(
    beneficio: str,
    candidato: Candidato,
    pagador: DadosPagador,
    data_emissao: date,
    pasta_saida: Path,
) -> Path:
    pasta_saida.mkdir(parents=True, exist_ok=True)
    caminho_pdf = pasta_saida / _nome_arquivo_pdf(beneficio, candidato)

    dia_txt = _dia_assinatura(beneficio, candidato)
    linha_data = f"{pagador.cidade_emissao}, {dia_txt} de {_mes_extenso(data_emissao.month)} de {data_emissao.year}"
    linha_assinatura = (
        f"__________________________________________________________<br/>"
        f"{candidato.nome}<br/>"
        f"CPF: {formatar_cpf(candidato.cpf)}"
    )

    conteudo = [
        Paragraph(_titulo(beneficio), _ESTILO_TITULO),
        Paragraph(_corpo_html(beneficio, pagador, candidato), _ESTILO_CORPO),
        Paragraph(linha_data, _ESTILO_DATA),
        Spacer(1, 8 * mm),
        Paragraph(linha_assinatura, _ESTILO_ASSINATURA),
    ]

    doc = SimpleDocTemplate(
        str(caminho_pdf), pagesize=letter,
        topMargin=30 * mm, bottomMargin=30 * mm,
        leftMargin=25 * mm, rightMargin=25 * mm,
    )

    try:
        doc.build(conteudo)
    except Exception as e:
        raise ErroGeracaoRecibo(f"Falha ao gerar PDF para {candidato.nome}: {e}")

    return caminho_pdf
