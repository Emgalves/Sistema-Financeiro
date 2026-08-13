"""
Módulo central de feriados — nacional, estadual e municipal.
Usado por qualquer cálculo de 'dia útil' no sistema.

O endereço de referência é o da CONSTRUTORA (sede fixa, mesma para todos
os clientes/obras) — não o endereço de cada obra. Por isso UF e cidade
têm padrão embutido aqui, e normalmente não precisam ser passados pelo
chamador.
"""
import holidays
from datetime import date

# Sede da construtora — usada como padrão em todo cálculo de dia útil.
UF_PADRAO = 'MG'
CIDADE_PADRAO = 'BELO HORIZONTE'

# ✅ Cadastro manual de feriados municipais por cidade.
# Chave = nome da cidade em MAIÚSCULO sem acento.
FERIADOS_MUNICIPAIS_FIXOS = {
    "BELO HORIZONTE": [
        (8, 15),   # Nossa Senhora da Assunção (padroeira de BH)
        (12, 12),  # Aniversário de Belo Horizonte
    ],
}


def obter_feriados_ano(ano: int, uf: str = UF_PADRAO, cidade: str = CIDADE_PADRAO) -> set:
    """
    Retorna um set de objetos date com todos os feriados (nacional +
    estadual + municipal) de um ano específico.
    """
    feriados_set = set()

    try:
        br_feriados = holidays.Brazil(years=ano, subdiv=uf) if uf else holidays.Brazil(years=ano)
        feriados_set.update(br_feriados.keys())
    except Exception:
        br_feriados = holidays.Brazil(years=ano)
        feriados_set.update(br_feriados.keys())

    if cidade:
        cidade_normalizada = cidade.strip().upper()
        for mes, dia in FERIADOS_MUNICIPAIS_FIXOS.get(cidade_normalizada, []):
            try:
                feriados_set.add(date(ano, mes, dia))
            except ValueError:
                continue

    return feriados_set


def eh_dia_util(dia: date, feriados: set) -> bool:
    """
    Dia útil segundo a regra de negócio: segunda a SÁBADO,
    excluindo domingo e feriados.
    """
    DOMINGO = 6  # weekday(): 0=segunda ... 6=domingo
    return dia.weekday() != DOMINGO and dia not in feriados

def calcular_enesimo_dia_util(ano: int, mes: int, n: int = 5, uf: str = UF_PADRAO, cidade: str = CIDADE_PADRAO):
    """
    Retorna a data do N-ésimo dia útil do mês, pela regra:
    segunda a sábado, excluindo domingo e feriados. UF/cidade usam o
    padrão da sede da construtora (Belo Horizonte/MG) se não informados.
    """
    from datetime import date, timedelta

    feriados = obter_feriados_ano(ano, uf=uf, cidade=cidade)
    if mes == 12:
        feriados |= obter_feriados_ano(ano + 1, uf=uf, cidade=cidade)

    dia = date(ano, mes, 1)
    contador = 0
    while dia.month == mes:
        if eh_dia_util(dia, feriados):
            contador += 1
            if contador == n:
                return dia
        dia += timedelta(days=1)

    return dia - timedelta(days=1)