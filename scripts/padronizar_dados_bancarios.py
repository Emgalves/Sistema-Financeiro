"""
Padroniza a coluna DADOS BANCÁRIOS (coluna O) da base_fornecedores.xlsx
para o mesmo formato do "Dados Consolidados" do Editar Fornecedor:

    BANCO | Op: 013 | Ag: 0001 | Conta: 12345-6 | CNPJ/CPF: 00.000.000/0000-00

Mexe SOMENTE em fornecedores de TED: sem chave PIX e com banco, agência
ou conta preenchidos. Fornecedores com PIX e sem dados bancários não mudam.

Uso (na pasta raiz do projeto, com o Excel FECHADO):
    python scripts/padronizar_dados_bancarios.py            -> só mostra o que mudaria
    python scripts/padronizar_dados_bancarios.py --gravar   -> faz backup e grava
"""
import os
import sys
import shutil
from datetime import datetime
from pathlib import Path

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import load_workbook
from src.config.utils import compor_dados_bancarios


def _txt(v):
    if v is None:
        return ''
    t = str(v).strip()
    return '' if t.lower() in ('nan', 'none') else t


def main():
    gravar = '--gravar' in sys.argv
    args = [a for a in sys.argv[1:] if a != '--gravar']
    if args:
        arquivo = args[0]
    else:
        from src.config.config import ARQUIVO_FORNECEDORES
        arquivo = ARQUIVO_FORNECEDORES

    print(f"Arquivo: {arquivo}")
    wb = load_workbook(arquivo)
    ws = wb['Fornecedores']

    # Colunas fixas da aba Fornecedores (mesma ordem usada pelo Salvar)
    COL_DOC, COL_TIPO, COL_NOME = 1, 2, 4
    COL_BANCO, COL_OP, COL_AG, COL_CONTA, COL_PIX, COL_DADOS = 7, 8, 9, 10, 11, 15

    alteracoes, atencao = [], []
    for r in range(2, ws.max_row + 1):
        doc = _txt(ws.cell(r, COL_DOC).value)
        if not doc:
            continue
        pix = _txt(ws.cell(r, COL_PIX).value)
        banco = _txt(ws.cell(r, COL_BANCO).value)
        op = _txt(ws.cell(r, COL_OP).value)
        ag = _txt(ws.cell(r, COL_AG).value)
        conta = _txt(ws.cell(r, COL_CONTA).value)
        atual = _txt(ws.cell(r, COL_DADOS).value)
        nome = _txt(ws.cell(r, COL_NOME).value)

        if pix:
            # Tem PIX mas a coluna O mostra outra coisa: só avisa, não altera
            if atual and atual != f"PIX: {pix}":
                atencao.append((r, nome, atual, pix))
            continue
        if not (banco or op or ag or conta):
            continue

        novo = compor_dados_bancarios(banco=banco, op=op, agencia=ag, conta=conta,
                                      cnpj_cpf=doc, tipo_pessoa=ws.cell(r, COL_TIPO).value)
        if novo != atual:
            alteracoes.append((r, nome, atual, novo))
            if gravar:
                ws.cell(r, COL_DADOS).value = novo

    print(f"\n{len(alteracoes)} fornecedor(es) de TED com texto diferente do padrão:\n")
    for r, nome, atual, novo in alteracoes:
        print(f"Linha {r} - {nome}\n   antes : {atual or '(vazio)'}\n   depois: {novo}")

    if atencao:
        print(f"\nATENÇÃO - {len(atencao)} fornecedor(es) com PIX cadastrado, mas com outro "
              f"texto em DADOS BANCÁRIOS (NÃO alterados, confira manualmente):")
        for r, nome, atual, pix in atencao:
            print(f"Linha {r} - {nome} | PIX: {pix} | coluna O: {atual}")

    if not gravar:
        print("\nNada foi gravado. Para gravar, rode de novo com --gravar")
        return
    if not alteracoes:
        print("\nNada a gravar.")
        return

    # ARQUIVO_FORNECEDORES vem do config como Path, não como texto
    caminho = Path(arquivo)
    backup = caminho.with_name(f"{caminho.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{caminho.suffix}")
    shutil.copy2(arquivo, backup)
    wb.save(arquivo)
    print(f"\nGravado. Backup do original em: {backup}")


if __name__ == '__main__':
    main()
