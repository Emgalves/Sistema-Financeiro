# -*- coding: utf-8 -*-
"""
LIMPEZA ÚNICA: remove o ".0" das NFs já gravadas nas planilhas de clientes.

Origem do problema: a tela "Editar" do gerenciador de lançamentos gravava
de volta a NF já convertida pelo pandas ("552233" virava "552233.0").
A correção da leitura impede novos casos; este script corrige os antigos.

Regra (a única alteração feita):
  célula de TEXTO na coluna NF da aba "Dados", formada só por dígitos
  seguidos de ".0"  ->  remove o ".0"     ("552233.0" -> "552233")
Nada mais é tocado: NFs sem ".0", alfanuméricas, vazias ou numéricas
ficam como estão.

Segurança:
  - Simulação por padrão: lista cada alteração (planilha, linha, antes/depois).
  - Ao aplicar, copia cada planilha alterada para BackUp/nf_limpeza_<data_hora>/.
  - Planilha aberta no Excel é pulada e listada (feche-a e rode de novo).

USO (na pasta raiz do projeto, com o sistema FECHADO):
  python limpar_nf_planilhas.py            -> simulação
  python limpar_nf_planilhas.py --aplicar  -> grava (com backup)
  python limpar_nf_planilhas.py --desfazer -> restaura o último backup
Opcional: informar outra pasta de clientes como último argumento.
"""
import json
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

PADRAO = re.compile(r"^\d+\.0$")
PASTA_BACKUP = Path("BackUp")
MANIFESTO = "manifesto_nf.json"


def pasta_clientes():
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args:
        return Path(args[0])
    from src.config.config import PASTA_CLIENTES
    return Path(PASTA_CLIENTES)


def analisar(arquivo):
    """Devolve (workbook, [(linha, antes, depois), ...]) ou (None, motivo)."""
    try:
        wb = load_workbook(arquivo)
    except Exception as e:
        return None, f"não foi possível abrir ({e})"
    if "Dados" not in wb.sheetnames:
        wb.close()
        return None, None
    ws = wb["Dados"]
    cab = [str(c.value or "").strip().upper() for c in ws[1]]
    if "NF" not in cab:
        wb.close()
        return None, None
    col = cab.index("NF") + 1
    mudancas = []
    for linha in range(2, ws.max_row + 1):
        cel = ws.cell(row=linha, column=col)
        if isinstance(cel.value, str) and PADRAO.match(cel.value.strip()):
            mudancas.append((linha, cel.value, cel.value.strip()[:-2]))
    return (wb, mudancas, col)


def executar(aplicar):
    pasta = pasta_clientes()
    print(f"Pasta de clientes: {pasta}")
    print("MODO: APLICAR\n" if aplicar else "MODO: SIMULAÇÃO (nada será gravado)\n")
    arquivos = sorted(p for p in pasta.glob("*.xlsx") if not p.name.startswith("~$"))
    destino = PASTA_BACKUP / f"nf_limpeza_{datetime.now():%Y%m%d_%H%M%S}"
    alterados, pulados, total = [], [], 0
    for arq in arquivos:
        r = analisar(arq)
        if r[0] is None:
            if r[1]:
                pulados.append((arq.name, r[1]))
            continue
        wb, mudancas, col = r
        if not mudancas:
            wb.close()
            continue
        print(f"== {arq.name}: {len(mudancas)} NF(s)")
        for linha, antes, depois in mudancas:
            print(f"     linha {linha}: '{antes}' -> '{depois}'")
        total += len(mudancas)
        if aplicar:
            ws = wb["Dados"]
            for linha, _, depois in mudancas:
                ws.cell(row=linha, column=col).value = depois
            try:
                (destino / arq.name).parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(arq, destino / arq.name)
                wb.save(arq)
                alterados.append(str(arq))
            except PermissionError:
                pulados.append((arq.name, "planilha aberta no Excel - feche e rode de novo"))
                (destino / arq.name).unlink(missing_ok=True)
        wb.close()

    print(f"\nTotal: {total} NF(s) com '.0' em {len(arquivos)} planilha(s) verificada(s).")
    if pulados:
        print("\n== PULADAS:")
        for nome, motivo in pulados:
            print(f"   {nome}: {motivo}")
    if aplicar and alterados:
        (destino / MANIFESTO).write_text(json.dumps(alterados, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\n{len(alterados)} planilha(s) corrigida(s). Originais em: {destino}")
        print("Para desfazer: python limpar_nf_planilhas.py --desfazer")
    elif not aplicar and total:
        print("\nPara gravar: python limpar_nf_planilhas.py --aplicar")


def desfazer():
    lotes = sorted(PASTA_BACKUP.glob(f"nf_limpeza_*/{MANIFESTO}"))
    if not lotes:
        print("Nenhuma limpeza de NF aplicada foi encontrada em BackUp/.")
        return
    manifesto = lotes[-1]
    for caminho in json.loads(manifesto.read_text(encoding="utf-8")):
        shutil.copy2(manifesto.parent / Path(caminho).name, caminho)
        print(f"   restaurado: {caminho}")
    manifesto.rename(manifesto.with_suffix(".restaurado"))


if __name__ == "__main__":
    if "--desfazer" in sys.argv:
        desfazer()
    else:
        executar("--aplicar" in sys.argv)
