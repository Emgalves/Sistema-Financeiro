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

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.platypus import (
    SimpleDocTemplate, BaseDocTemplate, PageTemplate, Frame, FrameBreak,
    Paragraph, Spacer,
)
from reportlab.lib.styles import ParagraphStyle

from .normalizacao import valor_por_extenso, formatar_valor_monetario, formatar_cpf
from .dados_candidatos import (
    Candidato, DadosPagador,
    BENEFICIO_TRANSPORTE, BENEFICIO_CAFE, BENEFICIO_CESTA_BASICA, BENEFICIO_CESTA_NATAL,
)

_MESES_PT = {
    1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
    5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
    9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro',
}


class ErroGeracaoRecibo(RuntimeError):
    pass


def _mes_extenso(mes: int) -> str:
    return _MESES_PT[mes]


def _mes_ano_anterior(competencia: str) -> tuple[str, int]:
    """Mês (por extenso) e ano imediatamente anteriores à competência
    (ex.: competência 2026-01 -> ('Dezembro', 2025))."""
    ano, mes = (int(p) for p in competencia.split('-'))
    if mes == 1:
        return _mes_extenso(12), ano - 1
    return _mes_extenso(mes - 1), ano


def _titulo(tipo: str) -> str:
    return {
        BENEFICIO_CESTA_BASICA: 'RECIBO DE ENTREGA DE CESTA BASICA',
        BENEFICIO_CESTA_NATAL: 'RECIBO DE ENTREGA DE CESTA DE NATAL',
        BENEFICIO_TRANSPORTE: 'RECIBO DE TRANSPORTE',
        BENEFICIO_CAFE: 'RECIBO DE CAFÉ',
    }[tipo]


def _corpo_html(tipo: str, pagador: DadosPagador, candidato: Candidato) -> str:
    ano_comp, mes_comp = candidato.competencia.split('-')
    mes_extenso = _mes_extenso(int(mes_comp))
    abertura = (
        f"Pelo presente declaro para os devidos fins, que recebi de "
        f"<b>{pagador.nome}</b>, {pagador.endereco},"
    )

    if tipo == BENEFICIO_CESTA_BASICA:
        mes_ref, ano_ref = _mes_ano_anterior(candidato.competencia)
        return f"{abertura} a <b>CESTA BASICA</b> referente ao mês de {mes_ref} de {ano_ref}."

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
                f"{abertura} o valor de R$ {valor_fmt} ({valor_ext}), referente ao fornecimento de "
                f"<b>CAFÉ</b> a ser consumido em {mes_extenso} de {ano_comp}, "
                f"juntamente com minha remuneração mensal."
            )

    raise ErroGeracaoRecibo(f"Tipo de benefício desconhecido: {tipo}")


def _dia_assinatura(tipo: str, candidato: Candidato) -> str:
    """
    Dia em branco para preenchimento manual na assinatura física, para
    os 4 tipos de benefício (Cestas, Transporte e Café) — pedido
    explícito do usuário: as assinaturas de Transporte/Café também são
    recolhidas a posteriori, então não faz sentido imprimir um dia já
    calculado.

    ATENÇÃO: isto contraria o que está documentado na especificação,
    seção 4.1 — lá o 5º dia útil calculado (via `feriados.py`) é a
    fonte primária do dia impresso no recibo de Transporte/Café,
    decisão tomada e revisada explicitamente na v6. Ver aviso enviado
    junto com esta alteração sobre a necessidade de atualizar a
    especificação.
    """
    return '&nbsp;' * 12  # espaço visível para preenchimento manual do dia


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


def _construir_bloco(beneficio: str, candidato: Candidato, pagador: DadosPagador, data_emissao: date) -> list:
    """Monta a lista de flowables de UM recibo (título, corpo, data, assinatura) — reaproveitado tanto no PDF individual quanto no combinado Transporte+Café."""
    dia_txt = _dia_assinatura(beneficio, candidato)
    linha_data = f"{pagador.cidade_emissao}, {dia_txt} de {_mes_extenso(data_emissao.month)} de {data_emissao.year}"
    linha_assinatura = (
        f"__________________________________________________________<br/>"
        f"{candidato.nome}<br/>"
        f"CPF: {formatar_cpf(candidato.cpf)}"
    )
    return [
        Paragraph(_titulo(beneficio), _ESTILO_TITULO),
        Paragraph(_corpo_html(beneficio, pagador, candidato), _ESTILO_CORPO),
        Paragraph(linha_data, _ESTILO_DATA),
        Spacer(1, 8 * mm),
        Paragraph(linha_assinatura, _ESTILO_ASSINATURA),
    ]


def gerar_recibo_pdf(
    beneficio: str,
    candidato: Candidato,
    pagador: DadosPagador,
    data_emissao: date,
    pasta_saida: Path,
) -> Path:
    pasta_saida.mkdir(parents=True, exist_ok=True)
    caminho_pdf = pasta_saida / _nome_arquivo_pdf(beneficio, candidato)

    conteudo = _construir_bloco(beneficio, candidato, pagador, data_emissao)

    doc = SimpleDocTemplate(
        str(caminho_pdf), pagesize=A4,
        topMargin=30 * mm, bottomMargin=30 * mm,
        leftMargin=25 * mm, rightMargin=25 * mm,
    )

    try:
        doc.build(conteudo)
    except Exception as e:
        raise ErroGeracaoRecibo(f"Falha ao gerar PDF para {candidato.nome}: {e}")

    return caminho_pdf


def gerar_recibo_transporte_cafe_pdf(
    candidato_transporte: Candidato,
    candidato_cafe: Candidato,
    pagador: DadosPagador,
    data_emissao: date,
    pasta_saida: Path,
) -> Path:
    """
    Gera Transporte e Café do MESMO colaborador em UMA página A4 —
    Transporte na metade de cima, Café na metade de baixo. Cada bloco
    continua sendo um recibo individual e completo (título, texto legal
    próprio, assinatura própria) — só a folha é compartilhada, para
    economizar impressão. Não substitui a emissão separada (opção do
    usuário na interface).

    Usa duas Frames fixas (metade de cima / metade de baixo) em vez de
    um Spacer calculado, para o corte ficar sempre exatamente na metade
    da página independente do tamanho do texto de cada bloco.
    """
    pasta_saida.mkdir(parents=True, exist_ok=True)
    nome_arquivo = f"RECIBO_TRANSPORTE_CAFE_{candidato_transporte.cpf}_{candidato_transporte.competencia}.pdf"
    caminho_pdf = pasta_saida / nome_arquivo

    largura_pagina, altura_pagina = A4
    margem_lateral = 25 * mm
    margem_topo = 15 * mm
    margem_fundo = 15 * mm

    largura_frame = largura_pagina - 2 * margem_lateral
    altura_frame = (altura_pagina - margem_topo - margem_fundo) / 2

    frame_cima = Frame(
        margem_lateral, margem_fundo + altura_frame, largura_frame, altura_frame,
        id='cima', topPadding=6 * mm,
    )
    frame_baixo = Frame(
        margem_lateral, margem_fundo, largura_frame, altura_frame,
        id='baixo', topPadding=6 * mm,
    )

    doc = BaseDocTemplate(str(caminho_pdf), pagesize=A4)
    doc.addPageTemplates([PageTemplate(id='DuasMetades', frames=[frame_cima, frame_baixo])])

    conteudo = _construir_bloco(BENEFICIO_TRANSPORTE, candidato_transporte, pagador, data_emissao)
    conteudo.append(FrameBreak())
    conteudo += _construir_bloco(BENEFICIO_CAFE, candidato_cafe, pagador, data_emissao)

    try:
        doc.build(conteudo)
    except Exception as e:
        raise ErroGeracaoRecibo(
            f"Falha ao gerar PDF combinado Transporte+Café para {candidato_transporte.nome}: {e}"
        )

    return caminho_pdf
