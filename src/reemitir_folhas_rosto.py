# reemitir_folhas_rosto.py
# ============================================================
# Reemite a FOLHA DE ROSTO (cabecalho + resumo + saldo, 1 pagina - sem
# os detalhes) de TODOS os relatorios ja emitidos de um cliente, usando
# os dados ATUAIS do cadastro (nome, endereco, etc.). Uso tipico:
# corrigiu um erro de cadastro (ex.: numero do endereco) e precisa
# reemitir as folhas de rosto de todo o historico, sem refazer os
# relatorios completos (que teriam paginas de detalhes identicas,
# ja corretas, que nao precisam ser regeneradas).
#
# COMO FUNCIONA:
#   1) Descobre as datas de relatorio ja emitidas olhando os valores
#      distintos de DATA_REL na aba Dados do cliente (mesmo criterio
#      que "Relatorio no: N" ja usa - testado: bate exatamente com o
#      numero do relatorio mais recente).
#   2) Para cada data, roda o MESMO pipeline que a tela usa
#      (RelatoriosDespesasService.processar_para_preview), so trocando
#      gerar_relatorio_pdf por gerar_folha_rosto_pdf no final.
#   3) Salva numa pasta SEPARADA (nunca sobrescreve os PDFs originais
#      automaticamente) - depois de conferir, voce copia/substitui.
#
# IMPORTANTE: rode isso DEPOIS de corrigir o cadastro do cliente (nome/
# endereco) em Clientes.xlsx - a correcao e lida automaticamente em
# cada folha de rosto gerada, nao precisa fazer nada especial.
#
# Uso:
#   python reemitir_folhas_rosto.py --cliente "RONALDO ROLIM DE OLIVEIRA"
#
# Opcional, para reemitir so um intervalo (ex.: so as ultimas 47
# quinzenas, sem mexer nas mais antigas):
#   python reemitir_folhas_rosto.py --cliente "NOME" --desde 2024-01-05 --ate 2026-08-20
#
# Opcional, pasta de saida (padrao: subpasta REEMISSAO_FOLHA_ROSTO
# dentro de PASTA_CLIENTES):
#   python reemitir_folhas_rosto.py --cliente "NOME" --saida "C:/caminho/pasta"
# ============================================================

import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


def resolver_caminho_por_nome_cliente(nome_cliente):
    try:
        from src.config.config import PASTA_CLIENTES
    except ImportError:
        from config.config import PASTA_CLIENTES

    caminho = Path(PASTA_CLIENTES) / f"{nome_cliente}.xlsx"
    if not caminho.exists():
        raise FileNotFoundError(
            f"Não encontrei '{caminho.name}' em PASTA_CLIENTES ({PASTA_CLIENTES})."
        )
    return caminho, Path(PASTA_CLIENTES)


def descobrir_datas_relatorio(arquivo_excel, desde=None, ate=None):
    """
    Le a aba Dados e devolve a lista ordenada de datas distintas de
    DATA_REL - cada uma corresponde a um relatorio ja emitido.
    """
    df = pd.read_excel(arquivo_excel, sheet_name='Dados')
    datas = pd.to_datetime(df['DATA_REL'], errors='coerce').dropna().unique()
    datas = sorted(pd.Timestamp(d) for d in datas)

    if desde is not None:
        datas = [d for d in datas if d >= pd.Timestamp(desde)]
    if ate is not None:
        datas = [d for d in datas if d <= pd.Timestamp(ate)]

    return datas


def reemitir(nome_cliente, arquivo_excel, pasta_saida, desde=None, ate=None):
    # Import tardio: RelatoriosDespesasService/RelatorioHandler dependem
    # de tkinter/xlwings/tkcalendar, so precisam existir quando o script
    # realmente roda (nao no ambiente de quem so revisa o codigo).
    from relatorio_despesas_service import RelatoriosDespesasService

    service = RelatoriosDespesasService()

    datas = descobrir_datas_relatorio(arquivo_excel, desde, ate)
    if not datas:
        print("Nenhuma data de relatório encontrada no intervalo informado.")
        return

    pasta_saida = Path(pasta_saida)
    pasta_saida.mkdir(parents=True, exist_ok=True)

    print(f"Cliente: {nome_cliente}")
    print(f"Arquivo: {arquivo_excel}")
    print(f"Pasta de saída: {pasta_saida}")
    print(f"Datas encontradas: {len(datas)} (de {datas[0].strftime('%d/%m/%Y')} "
          f"a {datas[-1].strftime('%d/%m/%Y')})")
    print("-" * 70)

    sucesso, falha = 0, 0

    for data in datas:
        nome_arquivo = f"REL - {nome_cliente} - {data.strftime('%d-%m-%Y')}.pdf"
        caminho_saida = pasta_saida / nome_arquivo

        try:
            config = {
                'arquivo': str(arquivo_excel),
                'data': data.date(),
                'incluir_excluidos': False,
                'incluir_futuros': False,
                'incluir_notas': False,
                'texto_notas': '',
            }
            dados_completos = service.processar_para_preview(config)
            service.handler.gerar_folha_rosto_pdf(dados_completos, str(caminho_saida), str(arquivo_excel))

            print(f"  ✅ {data.strftime('%d/%m/%Y')} -> {nome_arquivo}")
            sucesso += 1

        except Exception as e:
            print(f"  ❌ {data.strftime('%d/%m/%Y')} -> ERRO: {e}")
            falha += 1

    print("-" * 70)
    print(f"Concluído: {sucesso} folha(s) de rosto gerada(s), {falha} erro(s).")
    print(f"Revise o conteúdo em: {pasta_saida}")
    print("Nada foi sobrescrito nos arquivos originais - após conferir, "
          "substitua manualmente os PDFs antigos pelos novos.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Reemite as folhas de rosto (1 página cada) de todos os "
                    "relatórios já emitidos de um cliente."
    )
    grupo = parser.add_mutually_exclusive_group(required=True)
    grupo.add_argument('--cliente', metavar='NOME',
                        help='Nome do cliente exatamente como em PASTA_CLIENTES (sem .xlsx).')
    grupo.add_argument('--arquivo', metavar='CAMINHO',
                        help='Caminho completo do arquivo .xlsx (uso manual).')

    parser.add_argument('--desde', metavar='AAAA-MM-DD', default=None,
                         help='Só reemite relatórios a partir desta data (inclusive).')
    parser.add_argument('--ate', metavar='AAAA-MM-DD', default=None,
                         help='Só reemite relatórios até esta data (inclusive).')
    parser.add_argument('--saida', metavar='PASTA', default=None,
                         help='Pasta de saída (padrão: REEMISSAO_FOLHA_ROSTO dentro de PASTA_CLIENTES).')

    args = parser.parse_args()

    if args.cliente:
        caminho_arquivo, pasta_clientes = resolver_caminho_por_nome_cliente(args.cliente)
        nome_cliente = args.cliente
        pasta_saida_padrao = pasta_clientes / "REEMISSAO_FOLHA_ROSTO" / args.cliente
    else:
        caminho_arquivo = Path(args.arquivo)
        nome_cliente = caminho_arquivo.stem
        pasta_saida_padrao = caminho_arquivo.parent / "REEMISSAO_FOLHA_ROSTO" / nome_cliente

    pasta_saida = Path(args.saida) if args.saida else pasta_saida_padrao

    reemitir(
        nome_cliente,
        caminho_arquivo,
        pasta_saida,
        desde=args.desde,
        ate=args.ate,
    )
