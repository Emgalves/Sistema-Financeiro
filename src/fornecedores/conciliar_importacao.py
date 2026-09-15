# -*- coding: utf-8 -*-
r"""
conciliar_importacao.py
========================
Camada de conciliação entre colaboradores vindos de planilhas de RH
(Transporte, Folha RH, Diárias) e o cadastro mestre de fornecedores
(Base_Fornecedores.xlsx).

Por que isto é separado de regularizar_fornecedor.py
------------------------------------------------------
`regularizar_fornecedor.py` resolve CORREÇÃO de um cadastro que já existe e
está errado (CPF ou nome digitados errado), com propagação para todos os
clientes onde o fornecedor aparece — operação manual, disparada por um botão
próprio, um fornecedor de cada vez.

Este módulo resolve um problema anterior na linha do tempo: TRIAGEM em massa,
no momento em que uma planilha de RH está sendo importada, para os três
casos que podem acontecer com cada colaborador:

  1. já está corretamente cadastrado                -> nada a fazer;
  2. está cadastrado, mas o nome/dados bancários     -> atualização simples
     divergem do que veio na planilha                  (campo editável, não
                                                         precisa de
                                                         regularização);
  3. não existe nenhum cadastro para aquele CPF      -> precisa ser criado.

Em nenhum dos três casos este módulo decide ou altera CNPJ/CPF: o CPF é
sempre a chave de busca, nunca o alvo de correção automática. Se o CPF que
vier na planilha de origem estiver errado, isso é um problema de dado de
origem e deve ser corrigido manualmente (ou via `regularizar_fornecedor`,
depois de identificado), nunca inferido por nome aqui.

Uso típico (a partir de ImportadorRH, ANTES de finalizar os lançamentos):

    from src.fornecedores.conciliar_importacao import classificar_colaboradores

    ok, pendencias = classificar_colaboradores(candidatos, self.sistema, origem="Diárias")
    if pendencias:
        from src.fornecedores.triagem_importacao_rh import TriagemImportacaoRH
        resultado = TriagemImportacaoRH(self.sistema.root, self.sistema, pendencias).executar()
        if resultado['cancelado']:
            return  # aborta a importação inteira, nada foi gravado
        cpfs_fora = set(resultado['cpfs_excluidos'])
        registros_lote = [r for r in registros_lote
                           if ''.join(filter(str.isdigit, r['cnpj_cpf'])) not in cpfs_fora]
"""

from dataclasses import dataclass
from difflib import SequenceMatcher


# ---------------------------------------------------------------------------
# Utilitários de comparação de nomes (independentes dos internos de
# regularizar_fornecedor.py de propósito — este módulo não deve depender de
# funções privadas de outro módulo).
# ---------------------------------------------------------------------------

def _normalizar(nome):
    return ' '.join(str(nome or '').strip().upper().split())


def _similaridade(nome_a, nome_b):
    return SequenceMatcher(None, _normalizar(nome_a), _normalizar(nome_b)).ratio()


def extrair_chave_pix(dados_bancarios_texto):
    """
    Extrai a chave PIX de um texto livre vindo de planilha de RH.
    Ex.: 'PIX: CELULAR (31)9.7511-1509' -> '(31)9.7511-1509'
         'PIX: 31987228285'             -> '31987228285'
    Não inventa dado: se não achar o padrão 'PIX', devolve string vazia.
    """
    if not dados_bancarios_texto:
        return ''
    texto = str(dados_bancarios_texto).strip()
    if not texto.upper().startswith('PIX'):
        return ''
    resto = texto.split(':', 1)[1].strip() if ':' in texto else ''
    for rotulo in ('CELULAR', 'TELEFONE', 'EMAIL', 'CPF', 'CNPJ'):
        if resto.upper().startswith(rotulo):
            resto = resto[len(rotulo):].strip()
    return resto


# ---------------------------------------------------------------------------
# Classificação
# ---------------------------------------------------------------------------

@dataclass
class Pendencia:
    cpf: str                       # só dígitos (11)
    nome_planilha: str
    status: str                    # 'ausente' | 'divergente'
    nome_atual: str = ''           # nome hoje na base ('' se ausente)
    razao_atual: str = ''
    similaridade: float = 0.0
    dados_bancarios_planilha: str = ''
    categoria_sugerida: str = 'MO'
    origem: str = ''               # nome do arquivo/aba, só para exibição na tela


def classificar_colaboradores(candidatos, sistema, categoria_padrao='MO', origem=''):
    """
    Confere cada colaborador de uma planilha de RH contra o cadastro mestre
    de fornecedores, usando CPF como chave (nunca o nome).

    candidatos: lista de dicts com pelo menos 'cpf' (com ou sem pontuação) e
        'nome'. Opcionalmente 'dados_bancarios' (texto livre da planilha,
        ex. "PIX: 31987228285").
    sistema: precisa expor `buscar_fornecedor_completo(cpf)` (já existe em
        SistemaEntradaDados).

    Retorna (ok, pendencias):
        ok: lista de CPFs (só dígitos) já corretamente cadastrados.
        pendencias: lista de `Pendencia` — precisam de decisão do usuário
            antes de seguir com a importação.

    Colaboradores repetidos (mesmo CPF em várias linhas/abas, comum em
    planilhas de diária com uma aba por período) são consolidados uma única
    vez antes da checagem.
    """
    vistos = {}
    for c in candidatos:
        cpf_d = ''.join(filter(str.isdigit, str(c.get('cpf', ''))))
        if len(cpf_d) != 11:
            continue  # CPF malformado é outro problema, não é tratado aqui
        if cpf_d not in vistos:
            vistos[cpf_d] = {
                'nome': _normalizar(c.get('nome', '')),
                'dados_bancarios': str(c.get('dados_bancarios', '') or ''),
            }
        elif not vistos[cpf_d]['dados_bancarios'] and c.get('dados_bancarios'):
            vistos[cpf_d]['dados_bancarios'] = str(c['dados_bancarios'])

    ok, pendencias = [], []
    for cpf_d, info in vistos.items():
        fornecedor = sistema.buscar_fornecedor_completo(cpf_d)

        if not fornecedor:
            pendencias.append(Pendencia(
                cpf=cpf_d, nome_planilha=info['nome'], status='ausente',
                dados_bancarios_planilha=info['dados_bancarios'],
                categoria_sugerida=categoria_padrao, origem=origem,
            ))
            continue

        nome_atual = _normalizar(fornecedor.get('nome', ''))
        razao_atual = _normalizar(fornecedor.get('razao_social', ''))

        if info['nome'] == nome_atual or info['nome'] == razao_atual:
            ok.append(cpf_d)
            continue

        pendencias.append(Pendencia(
            cpf=cpf_d, nome_planilha=info['nome'], status='divergente',
            nome_atual=nome_atual, razao_atual=razao_atual,
            similaridade=_similaridade(info['nome'], nome_atual),
            dados_bancarios_planilha=info['dados_bancarios'],
            categoria_sugerida=fornecedor.get('categoria') or categoria_padrao,
            origem=origem,
        ))

    return ok, pendencias


# ---------------------------------------------------------------------------
# Gravação em lote (mesma regra de negócio de SistemaEntradaDados.salvar_fornecedor,
# fatorada para não duplicar lógica e para gravar várias pessoas de uma vez).
# ---------------------------------------------------------------------------

COLUNAS_FORNECEDOR = [
    'cnpj_cpf', 'tipo_pessoa', 'razao_social', 'nome',
    'telefone', 'email', 'banco', 'op', 'agencia', 'conta',
    'chave_pix', 'categoria', 'especificacao', 'vinculo',
    'dados_bancarios', 'endereco', 'status', 'responsavel',
]


def upsert_fornecedores_lote(lista_dados, arquivo_fornecedores):
    """
    Insere ou atualiza, de uma só vez, um lote de fornecedores na aba
    'Fornecedores' de Base_Fornecedores.xlsx.

    Reproduz as mesmas regras de SistemaEntradaDados.salvar_fornecedor:
    upsert por CNPJ/CPF (só dígitos), preserva o STATUS de quem já existia,
    marca ATIVO quem é novo, e reordena a planilha inteira por nome ao
    final. Abre e salva o arquivo uma única vez, mesmo para várias pessoas
    — evita N aberturas/gravações redundantes num único lote de importação.

    lista_dados: lista de dicts com as mesmas chaves de COLUNAS_FORNECEDOR.
        Chaves ausentes são tratadas como string vazia. Linhas sem
        'cnpj_cpf' ou sem 'nome' são silenciosamente ignoradas (não é
        função desta rotina inventar dado obrigatório).

    Levanta PermissionError se o arquivo estiver aberto no Excel — quem
    chamar deve tratar isso (mesmo padrão do resto do sistema: nada é
    gravado parcialmente se falhar).
    """
    from openpyxl import load_workbook

    wb = load_workbook(arquivo_fornecedores)
    ws = wb['Fornecedores']

    if ws.cell(1, 17).value != 'STATUS':
        ws.cell(1, 17, 'STATUS')

    fornecedores = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        row_pad = list(row) + [None] * (len(COLUNAS_FORNECEDOR) - len(row))
        reg = {col: (str(row_pad[i]).strip() if row_pad[i] else '')
               for i, col in enumerate(COLUNAS_FORNECEDOR)}
        reg['status'] = reg['status'].upper() if reg['status'] else 'ATIVO'
        if reg['cnpj_cpf'] and reg['nome']:
            fornecedores.append(reg)

    index_por_cpf = {
        ''.join(filter(str.isdigit, reg['cnpj_cpf'])): i
        for i, reg in enumerate(fornecedores)
    }

    gravados = 0
    for dados_novo in lista_dados:
        dados = {col: str(dados_novo.get(col, '') or '').strip() for col in COLUNAS_FORNECEDOR}
        busca = ''.join(filter(str.isdigit, dados['cnpj_cpf']))
        if not busca or not dados['nome']:
            continue
        if busca in index_por_cpf:
            i = index_por_cpf[busca]
            dados['status'] = fornecedores[i].get('status', 'ATIVO')  # preserva status
            fornecedores[i] = dados
        else:
            dados['status'] = 'ATIVO'
            fornecedores.append(dados)
            index_por_cpf[busca] = len(fornecedores) - 1
        gravados += 1

    fornecedores.sort(key=lambda x: x.get('nome', '').upper())

    max_row = ws.max_row
    if max_row > 1:
        ws.delete_rows(2, max_row - 1)
    for i, reg in enumerate(fornecedores, start=2):
        for j, col in enumerate(COLUNAS_FORNECEDOR, start=1):
            ws.cell(row=i, column=j, value=reg.get(col, ''))

    wb.save(arquivo_fornecedores)
    wb.close()
    return gravados
