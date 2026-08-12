# src/comprovantes_beneficios/normalizacao.py
"""
Normalização de dados vindos das planilhas de origem (Clientes.xlsx e
planilhas por cliente), que apresentam inconsistências conhecidas:

  - CPF gravado com/sem máscara, e às vezes com zeros à esquerda extras
    (ex.: '07573876670', '003.544.326-05', '00007573876670').
  - Nome do mesmo colaborador grafado de forma diferente entre lançamentos
    (ex.: 'ROBEVAL PORTACIO DOS SANTOS' vs 'ROBERVAL PORTACIO DOS SANTOS').

Este módulo não tem dependência de UI nem de planilha — só funções puras,
testáveis isoladamente, conforme a ordem de construção combinada.
"""

import re
from num2words import num2words


# ============================================================================
# CPF
# ============================================================================

def normalizar_cpf(cpf_bruto) -> str | None:
    """
    Remove toda formatação e zeros à esquerda em excesso, retornando
    sempre uma string de 11 dígitos (ou None se não for possível obter 11
    dígitos válidos a partir da entrada).

    Trata o caso observado nos dados reais em que o CPF vem com dígitos
    extras à esquerda (ex.: '00007573876670' -> só os últimos 11 dígitos
    são o CPF real: '07573876670').
    """
    if cpf_bruto is None:
        return None

    apenas_digitos = re.sub(r'\D', '', str(cpf_bruto))

    if len(apenas_digitos) == 0:
        return None

    # CPF sempre tem 11 dígitos. Se vier mais longo (zeros à esquerda
    # indevidos, como visto na planilha real), usamos os 11 últimos.
    if len(apenas_digitos) > 11:
        apenas_digitos = apenas_digitos[-11:]

    # Se vier mais curto, completa à esquerda com zero (CPFs que começam
    # com zero às vezes perdem o zero em campos numéricos do Excel).
    apenas_digitos = apenas_digitos.zfill(11)

    if len(apenas_digitos) != 11:
        return None

    return apenas_digitos


def cpf_valido(cpf_normalizado: str) -> bool:
    """
    Valida dígitos verificadores. Usado apenas para sinalizar CPFs
    suspeitos na interface (aviso), nunca para bloquear a seleção —
    a planilha de origem pode ter CPFs digitados errado há anos e o
    módulo não deve travar a operação por causa disso.
    """
    if cpf_normalizado is None or len(cpf_normalizado) != 11:
        return False

    if cpf_normalizado == cpf_normalizado[0] * 11:
        return False

    def _digito(cpf_parcial: str) -> str:
        soma = sum(
            int(d) * peso
            for d, peso in zip(cpf_parcial, range(len(cpf_parcial) + 1, 1, -1))
        )
        resto = (soma * 10) % 11
        return '0' if resto == 10 else str(resto)

    d1 = _digito(cpf_normalizado[:9])
    d2 = _digito(cpf_normalizado[:9] + d1)

    return cpf_normalizado[-2:] == d1 + d2


def formatar_cpf(cpf_normalizado: str) -> str:
    """Formata um CPF de 11 dígitos como 000.000.000-00 para exibição/impressão."""
    if cpf_normalizado is None or len(cpf_normalizado) != 11:
        return cpf_normalizado or ''
    c = cpf_normalizado
    return f"{c[0:3]}.{c[3:6]}.{c[6:9]}-{c[9:11]}"


# ============================================================================
# NOME
# ============================================================================

def normalizar_nome(nome_bruto: str) -> str:
    """
    Padroniza espaçamento e caixa alta para comparação/exibição.
    Não corrige erros de grafia — isso é tratado em dados_candidatos.py,
    que decide qual grafia usar quando há divergência para o mesmo CPF.
    """
    if not nome_bruto:
        return ''
    nome = re.sub(r'\s+', ' ', str(nome_bruto)).strip()
    return nome.upper()


# ============================================================================
# VALOR POR EXTENSO
# ============================================================================

def valor_por_extenso(valor: float) -> str:
    """
    Converte um valor em reais (float) para o formato usado nos modelos
    originais, ex.:
        745.60  -> "SETECENTOS E QUARENTA E CINCO REAIS E SESSENTA CENTAVOS"
        1.00    -> "UM REAL"
        0.01    -> "UM CENTAVO"
        2000.00 -> "DOIS MIL REAIS"

    Levanta ValueError para valor negativo (não deve ocorrer neste
    módulo, mas é melhor falhar alto do que gerar um recibo errado).
    """
    if valor is None:
        raise ValueError("valor_por_extenso: valor não pode ser None")
    if valor < 0:
        raise ValueError(f"valor_por_extenso: valor negativo não suportado ({valor})")

    # Evita erro de ponto flutuante (ex.: 775.1999999999999 vindo da planilha)
    centavos_totais = round(valor * 100)
    reais, centavos = divmod(centavos_totais, 100)

    partes = []

    if reais > 0:
        texto_reais = num2words(reais, lang='pt_BR')
        sufixo_reais = 'REAL' if reais == 1 else 'REAIS'
        partes.append(f"{texto_reais.upper()} {sufixo_reais}")

    if centavos > 0:
        texto_centavos = num2words(centavos, lang='pt_BR')
        sufixo_centavos = 'CENTAVO' if centavos == 1 else 'CENTAVOS'
        partes.append(f"{texto_centavos.upper()} {sufixo_centavos}")

    if not partes:
        return "ZERO REAIS"

    return " E ".join(partes)


def formatar_valor_monetario(valor: float) -> str:
    """Formata um float como R$ 0.000,00 (padrão brasileiro)."""
    if valor is None:
        return ''
    texto = f"{valor:,.2f}"
    # troca separadores: 1,234.56 (en-US) -> 1.234,56 (pt-BR)
    texto = texto.replace(',', '_').replace('.', ',').replace('_', '.')
    return texto
