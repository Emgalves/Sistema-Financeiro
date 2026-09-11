# reemitir_relatorios.py
# ============================================================
# Reemite relatórios já emitidos de um cliente, por intervalo de datas,
# usando os dados ATUAIS do cadastro/planilha. Substitui e generaliza o
# antigo reemitir_folhas_rosto.py: agora dá pra escolher entre reemitir
# só a folha de rosto (1 página) ou o RELATÓRIO COMPLETO (com as
# páginas de detalhe), usando a flag --modo.
#
# python reemitir_relatorios.py --cliente "NOME" --modo rosto      (padrão,
#     equivalente ao antigo reemitir_folhas_rosto.py)
# python reemitir_relatorios.py --cliente "NOME" --modo completo   (novo)
#
# COMO FUNCIONA (igual ao script antigo):
#   1) Descobre as datas de relatório já emitidas olhando os valores
#      distintos de DATA_REL na aba Dados do cliente.
#   2) Para cada data, roda o MESMO pipeline que a tela usa
#      (RelatoriosDespesasService.processar_para_preview) e gera o PDF
#      com gerar_folha_rosto_pdf (--modo rosto) ou gerar_relatorio_pdf
#      (--modo completo).
#   3) Salva numa pasta SEPARADA (nunca sobrescreve os PDFs originais
#      automaticamente) - depois de conferir, você copia/substitui.
#
# ------------------------------------------------------------
# LEIA ANTES DE USAR --modo completo (diferença importante em relação
# à folha de rosto):
#
# A folha de rosto é segura por natureza: as páginas de detalhe do PDF
# original não são tocadas, só o cabeçalho/resumo é refeito. Já o
# --modo completo REGENERA O RELATÓRIO INTEIRO a partir do estado ATUAL
# da planilha - inclusive os valores/lançamentos daquela data. Isso é
# geralmente o que você quer quando o motivo da reemissão é corrigir um
# lançamento errado (não só o endereço). Mas duas consequências
# precisam estar claras:
#
#   a) Se você corrigiu um lançamento de uma data X, o "acumulado" (o
#      total corrido) de TODOS os relatórios emitidos DEPOIS de X
#      também mudou - mesmo que os lançamentos deles não tenham sido
#      alterados. Reemitir só o relatório da data X deixa os relatórios
#      seguintes com um acumulado desatualizado em relação ao que a
#      planilha calcularia hoje. Se o objetivo é manter os PDFs
#      coerentes com a planilha, use --desde a partir da data X (ou
#      antes) até a data mais recente, não só a data pontual corrigida.
#
#   b) "Lançamentos futuros" (incluir_futuros) NÃO faz diferença aqui:
#      no relatório completo atual, essa seção foi removida do PDF
#      principal e virou um PDF separado, sob demanda
#      (gerar_relatorio_lancamentos_futuros_pdf). Este script não mexe
#      nisso - se precisar reemitir também os relatórios de lançamentos
#      futuros de datas antigas, é outro script (posso montar se for o
#      caso).
#
# NOTAS DO RELATÓRIO (texto livre que aparece na 1ª página):
#   O script tenta recuperar automaticamente a nota que foi salva na
#   época, lendo a aba "HISTÓRICO_NOTAS" do próprio arquivo do cliente
#   e casando pela data do relatório (é a única aba que o sistema
#   realmente usa para salvar notas hoje - existe também uma aba/rotina
#   "Notas" no código do handler, mas é código morto: nunca é
#   preenchida na prática, porque foi substituída pela rotina que grava
#   em "HISTÓRICO_NOTAS" sem que a antiga tenha sido removida). Se
#   houver mais de uma nota salva para a mesma data, usa a mais
#   recente. Se não encontrar nada, gera sem notas.
#   Use --sem-notas para forçar a reemissão sem nenhuma nota, mesmo que
#   o script encontre uma no histórico.
# ------------------------------------------------------------

import argparse
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook


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
    Lê a aba Dados e devolve a lista ordenada de datas distintas de
    DATA_REL - cada uma corresponde a um relatório já emitido.
    """
    df = pd.read_excel(arquivo_excel, sheet_name='Dados')
    datas = pd.to_datetime(df['DATA_REL'], errors='coerce').dropna().unique()
    datas = sorted(pd.Timestamp(d) for d in datas)

    if desde is not None:
        datas = [d for d in datas if d >= pd.Timestamp(desde)]
    if ate is not None:
        datas = [d for d in datas if d <= pd.Timestamp(ate)]

    return datas


def buscar_nota_historico(arquivo_excel, data_relatorio):
    """
    Procura, na aba HISTÓRICO_NOTAS, a nota mais recente salva para a
    data de relatório informada (casamento por data formatada
    dd/mm/aaaa, que é como a rotina de salvar grava a coluna A).
    Retorna string vazia se não achar nada.
    """
    try:
        wb = load_workbook(arquivo_excel, data_only=True)
    except Exception:
        return ''

    if "HISTÓRICO_NOTAS" not in wb.sheetnames:
        return ''

    ws = wb["HISTÓRICO_NOTAS"]
    if ws.max_row <= 1:
        return ''

    alvo = data_relatorio.strftime('%d/%m/%Y')
    nota_encontrada = ''

    # Percorre todas as linhas: se houver mais de uma nota para a
    # mesma data, fica com a última (assume-se que foram gravadas em
    # ordem cronológica, como a rotina de salvar realmente faz).
    for row in range(2, ws.max_row + 1):
        data_cel = ws.cell(row, 1).value
        texto_cel = ws.cell(row, 3).value
        if data_cel is None:
            continue
        data_cel_str = data_cel.strftime('%d/%m/%Y') if hasattr(data_cel, 'strftime') else str(data_cel)
        if data_cel_str == alvo and texto_cel:
            nota_encontrada = str(texto_cel)

    return nota_encontrada


def reemitir(nome_cliente, arquivo_excel, pasta_saida, modo, desde=None, ate=None,
             incluir_excluidos=False, sem_notas=False):
    # Import tardio: RelatoriosDespesasService/RelatorioHandler dependem
    # de tkinter/xlwings/tkcalendar, só precisam existir quando o script
    # realmente roda (não no ambiente de quem só revisa o código).
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
    print(f"Modo: {'RELATÓRIO COMPLETO' if modo == 'completo' else 'FOLHA DE ROSTO'}")
    print(f"Pasta de saída: {pasta_saida}")
    print(f"Datas encontradas: {len(datas)} (de {datas[0].strftime('%d/%m/%Y')} "
          f"a {datas[-1].strftime('%d/%m/%Y')})")
    print("-" * 70)

    sucesso, falha = 0, 0

    for data in datas:
        sufixo = " (com excluídos)" if incluir_excluidos else ""
        if modo == 'completo':
            nome_arquivo = f"REL - {nome_cliente} - {data.strftime('%d-%m-%Y')}{sufixo}.pdf"
        else:
            nome_arquivo = f"REL ROSTO - {nome_cliente} - {data.strftime('%d-%m-%Y')}{sufixo}.pdf"
        caminho_saida = pasta_saida / nome_arquivo

        try:
            texto_notas = ''
            incluir_notas = False
            if modo == 'completo' and not sem_notas:
                texto_notas = buscar_nota_historico(arquivo_excel, data)
                incluir_notas = bool(texto_notas)

            config = {
                'arquivo': str(arquivo_excel),
                'data': data.date(),
                'incluir_excluidos': incluir_excluidos,
                'incluir_futuros': False,  # irrelevante: removido do PDF principal
                'incluir_notas': incluir_notas,
                'texto_notas': texto_notas,
            }
            dados_completos = service.processar_para_preview(config)

            if modo == 'completo':
                service.handler.gerar_relatorio_pdf(dados_completos, str(caminho_saida), str(arquivo_excel))
            else:
                service.handler.gerar_folha_rosto_pdf(dados_completos, str(caminho_saida), str(arquivo_excel))

            aviso_nota = " [nota recuperada]" if incluir_notas else ""
            print(f"  ✅ {data.strftime('%d/%m/%Y')} -> {nome_arquivo}{aviso_nota}")
            sucesso += 1

        except Exception as e:
            print(f"  ❌ {data.strftime('%d/%m/%Y')} -> ERRO: {e}")
            falha += 1

    print("-" * 70)
    print(f"Concluído: {sucesso} PDF(s) gerado(s), {falha} erro(s).")
    print(f"Revise o conteúdo em: {pasta_saida}")
    print("Nada foi sobrescrito nos arquivos originais - após conferir, "
          "substitua manualmente os PDFs antigos pelos novos.")


if __name__ == '__main__':
    parser = argparse.ArgumentParser(
        description="Reemite folhas de rosto ou relatórios completos, já emitidos, "
                    "de um cliente, por intervalo de datas."
    )
    grupo = parser.add_mutually_exclusive_group(required=True)
    grupo.add_argument('--cliente', metavar='NOME',
                        help='Nome do cliente exatamente como em PASTA_CLIENTES (sem .xlsx).')
    grupo.add_argument('--arquivo', metavar='CAMINHO',
                        help='Caminho completo do arquivo .xlsx (uso manual).')

    parser.add_argument('--modo', choices=['rosto', 'completo'], default='rosto',
                         help="'rosto' (padrão) reemite só a folha de rosto (1 página). "
                              "'completo' reemite o relatório inteiro, com as páginas de detalhe.")
    parser.add_argument('--desde', metavar='AAAA-MM-DD', default=None,
                         help='Só reemite relatórios a partir desta data (inclusive).')
    parser.add_argument('--ate', metavar='AAAA-MM-DD', default=None,
                         help='Só reemite relatórios até esta data (inclusive).')
    parser.add_argument('--saida', metavar='PASTA', default=None,
                         help='Pasta de saída (padrão: subpasta REEMISSAO_* dentro de PASTA_CLIENTES).')
    parser.add_argument('--incluir-excluidos', action='store_true',
                         help='Inclui lançamentos marcados como EXCLUÍDO (padrão: não inclui).')
    parser.add_argument('--sem-notas', action='store_true',
                         help='(só --modo completo) Não tenta recuperar nem incluir notas, mesmo que existam no histórico.')

    args = parser.parse_args()

    if args.cliente:
        caminho_arquivo, pasta_clientes = resolver_caminho_por_nome_cliente(args.cliente)
        nome_cliente = args.cliente
        subpasta = "REEMISSAO_RELATORIO_COMPLETO" if args.modo == 'completo' else "REEMISSAO_FOLHA_ROSTO"
        pasta_saida_padrao = pasta_clientes / subpasta / args.cliente
    else:
        caminho_arquivo = Path(args.arquivo)
        nome_cliente = caminho_arquivo.stem
        subpasta = "REEMISSAO_RELATORIO_COMPLETO" if args.modo == 'completo' else "REEMISSAO_FOLHA_ROSTO"
        pasta_saida_padrao = caminho_arquivo.parent / subpasta / nome_cliente

    pasta_saida = Path(args.saida) if args.saida else pasta_saida_padrao

    reemitir(
        nome_cliente,
        caminho_arquivo,
        pasta_saida,
        modo=args.modo,
        desde=args.desde,
        ate=args.ate,
        incluir_excluidos=args.incluir_excluidos,
        sem_notas=args.sem_notas,
    )
