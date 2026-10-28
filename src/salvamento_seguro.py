"""
Salvamento seguro das planilhas dos clientes.

Problema que resolve: workbook.save(caminho) grava POR CIMA do arquivo original.
Se a gravação for interrompida (queda de rede, erro, fechamento do programa),
o original já foi destruído e o novo fica incompleto -> planilha corrompida.

Como funciona:
  1. Grava a versão nova num arquivo temporário, na mesma pasta.
  2. Reabre o temporário e confere se está íntegro (mesmas abas).
  3. Guarda uma cópia da versão atual em Backups_Planilhas (mantém as últimas N).
  4. Só então substitui o original pelo temporário (troca atômica).
Se algo falhar em 1 ou 2, o original continua intacto e o temporário é apagado.
"""

import os
import csv
import time
import inspect
import shutil
import tempfile
import logging
from datetime import datetime
from pathlib import Path

from openpyxl import load_workbook

logger = logging.getLogger(__name__)

NOME_PASTA_BACKUP = "Backups_Planilhas"
VERSOES_MANTIDAS = 10          # versões anteriores guardadas por cliente
TENTATIVAS_SUBSTITUICAO = 3    # o Drive/antivírus pode travar o arquivo por instantes
NOME_LOG_GRAVACOES = "log_gravacoes"  # um CSV por computador, dentro de Backups_Planilhas


def salvar_planilha_segura(workbook, caminho, versoes_mantidas=VERSOES_MANTIDAS):
    """Substitui workbook.save(caminho). Lança exceção se não conseguir salvar,
    mas nesse caso o arquivo original permanece intacto."""
    caminho = Path(caminho)
    pasta = caminho.parent

    # O temporário precisa estar na MESMA pasta: só assim a troca é atômica.
    # Prefixo "~$": as listagens de clientes do sistema já ignoram arquivos
    # que começam com "~", então o temporário nunca aparece como cliente.
    fd, nome_tmp = tempfile.mkstemp(prefix=f"~$tmp_{caminho.stem}_", suffix=caminho.suffix or ".xlsx", dir=pasta)
    os.close(fd)
    tmp = Path(nome_tmp)

    try:
        # 1. Grava a versão nova no temporário
        workbook.save(tmp)

        # 2. Confere se o temporário abre e tem as mesmas abas
        wb_teste = load_workbook(tmp, read_only=True)
        abas_gravadas = wb_teste.sheetnames
        wb_teste.close()
        if abas_gravadas != workbook.sheetnames:
            raise IOError(
                f"Arquivo gravado incompleto: abas esperadas {workbook.sheetnames}, "
                f"encontradas {abas_gravadas}"
            )

        # 3. Guarda a versão atual antes de substituir
        if caminho.exists():
            _guardar_versao_anterior(caminho, versoes_mantidas)

        # 4. Troca o original pelo novo
        _substituir_com_tentativas(tmp, caminho)
        logger.info(f"✅ Planilha salva com segurança: {caminho.name}")

        # 5. Registra quem gravou, de onde e por qual caminho do código
        _registrar_gravacao(caminho)

    except Exception:
        # Falhou antes da troca: o original não foi tocado. Remove o temporário.
        try:
            tmp.unlink(missing_ok=True)
        except Exception as e:
            logger.warning(f"⚠️ Não foi possível remover o temporário {tmp.name}: {e}")
        raise


def _guardar_versao_anterior(caminho, versoes_mantidas):
    """Copia a versão atual para Backups_Planilhas e apaga as mais antigas."""
    try:
        pasta_backup = caminho.parent / NOME_PASTA_BACKUP
        pasta_backup.mkdir(exist_ok=True)

        carimbo = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        shutil.copy2(caminho, pasta_backup / f"{caminho.stem}_{carimbo}.xlsx")

        # O carimbo no nome faz a ordem alfabética ser a ordem cronológica.
        # Padrão exato evita misturar clientes de nome parecido (ex.: JOAO e JOAO_SILVA).
        versoes = sorted(pasta_backup.glob(f"{caminho.stem}_????????_??????_??????.xlsx"))
        for antiga in versoes[:-versoes_mantidas]:
            antiga.unlink()
    except Exception as e:
        # Falha no backup não impede o salvamento: a troca atômica já protege o original
        logger.warning(f"⚠️ Não foi possível guardar a versão anterior de {caminho.name}: {e}")


def _substituir_com_tentativas(origem, destino):
    """os.replace com novas tentativas. Se a planilha estiver aberta no Excel,
    após as tentativas lança PermissionError (tratado em enviar_dados)."""
    for tentativa in range(1, TENTATIVAS_SUBSTITUICAO + 1):
        try:
            os.replace(origem, destino)
            return
        except PermissionError:
            if tentativa == TENTATIVAS_SUBSTITUICAO:
                raise
            logger.debug(f"Arquivo travado, nova tentativa {tentativa + 1}/{TENTATIVAS_SUBSTITUICAO}")
            time.sleep(1)


def usuario_atual():
    """Usuário do Windows e nome do computador (mesma fonte usada nos backups JSON)."""
    usuario = os.environ.get('USERNAME') or os.environ.get('USER') or 'Desconhecido'
    computador = os.environ.get('COMPUTERNAME') or 'Desconhecido'
    return usuario, computador


def _registrar_gravacao(caminho):
    """Acrescenta uma linha em Backups_Planilhas/log_gravacoes_<COMPUTADOR>.csv.

    Um arquivo por computador evita conflito de sincronização do Google Drive
    quando duas máquinas gravam ao mesmo tempo. A coluna ORIGEM traz a cadeia
    de funções que chamou o salvamento (ex.: salvar_na_planilha > enviar_dados),
    o que mostra QUAL botão ou fluxo disparou a gravação.
    Qualquer falha aqui é só registrada no log: nunca impede o salvamento."""
    try:
        usuario, computador = usuario_atual()
        # context=0: não lê código-fonte (rápido e funciona no executável)
        funcoes = [f.function for f in inspect.stack(0)[2:10]
                   if not f.function.startswith('__')
                   and f.function not in ('<module>', 'wrapper', 'mainloop')][:4]
        origem = " < ".join(funcoes)

        pasta_backup = caminho.parent / NOME_PASTA_BACKUP
        pasta_backup.mkdir(exist_ok=True)
        arquivo_log = pasta_backup / f"{NOME_LOG_GRAVACOES}_{computador}.csv"
        novo = not arquivo_log.exists()

        with open(arquivo_log, 'a', newline='', encoding='utf-8-sig') as f:
            w = csv.writer(f, delimiter=';')
            if novo:
                w.writerow(['DATA_HORA', 'USUARIO', 'COMPUTADOR', 'PLANILHA', 'ORIGEM'])
            w.writerow([datetime.now().strftime('%d/%m/%Y %H:%M:%S'),
                        usuario, computador, caminho.name, origem])
    except Exception as e:
        logger.warning(f"⚠️ Não foi possível registrar a gravação no log: {e}")
