# -*- coding: utf-8 -*-
"""
gmail_cliente.py - Acesso ao Gmail pela API oficial do Google, SOMENTE LEITURA.

Permissão pedida: gmail.readonly. O programa não apaga, não move, não marca
como lido e não envia nada. O controle do que já foi lido fica na planilha
Documentos_Recebidos.xlsx (aba Mensagens).

Dependências:
    pip install google-api-python-client google-auth-httplib2 google-auth-oauthlib
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
import base64
import io
import re
import zipfile
from email import message_from_bytes, policy
from datetime import date, timedelta
from email.utils import parsedate_to_datetime

from . import config_email as cfg

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]


def obter_servico():
    """Autentica e devolve o objeto de acesso ao Gmail.
    1ª vez: abre o navegador para você autorizar. Depois usa o token.json salvo."""
    from google.auth.transport.requests import Request
    from google.oauth2.credentials import Credentials
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build

    cred = None
    if cfg.ARQUIVO_TOKEN.exists():
        cred = Credentials.from_authorized_user_file(str(cfg.ARQUIVO_TOKEN), SCOPES)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            try:
                cred.refresh(Request())
            except Exception:
                cred = None  # token revogado/expirado -> novo login
        if not cred or not cred.valid:
            if not cfg.ARQUIVO_CREDENTIALS.exists():
                raise FileNotFoundError(
                    f"Arquivo de credenciais não encontrado: {cfg.ARQUIVO_CREDENTIALS}\n"
                    f"Siga o passo a passo do LEIA-ME para baixá-lo do Google Cloud.")
            flow = InstalledAppFlow.from_client_secrets_file(str(cfg.ARQUIVO_CREDENTIALS), SCOPES)
            cred = flow.run_local_server(port=0)
        cfg.PASTA_CREDENCIAIS.mkdir(parents=True, exist_ok=True)
        cfg.ARQUIVO_TOKEN.write_text(cred.to_json(), encoding="utf-8")

    return build("gmail", "v1", credentials=cred, cache_discovery=False)


def listar_ids(servico, consulta=None, dias=None, limite=None):
    """IDs das mensagens que atendem à busca, das mais novas para as mais antigas."""
    consulta = consulta or cfg.CONSULTA_GMAIL
    dias = cfg.DIAS_RETROATIVOS if dias is None else dias
    limite = limite or cfg.MAX_MENSAGENS_POR_EXECUCAO
    desde = (date.today() - timedelta(days=dias)).strftime("%Y/%m/%d")
    q = f"{consulta} after:{desde}"

    ids, token = [], None
    while len(ids) < limite:
        resp = servico.users().messages().list(
            userId="me", q=q, pageToken=token, maxResults=min(100, limite - len(ids))
        ).execute()
        ids += [m["id"] for m in resp.get("messages", [])]
        token = resp.get("nextPageToken")
        if not token:
            break
    return ids


def _b64(dados):
    return base64.urlsafe_b64decode(dados.encode("ascii") + b"==")


def obter_mensagens(servico, msg_id):
    """Baixa a mensagem completa (formato bruto) e a DESMONTA em mensagens lógicas.

    Uma mensagem do Gmail pode conter outras mensagens dentro dela - é o que
    acontece quando alguém seleciona vários e-mails e usa "Encaminhar como
    anexo" (ou anexa arquivos .eml). Cada e-mail original vira uma mensagem
    lógica separada, com o seu próprio remetente, assunto, data, corpo e
    anexos. Assim, a nota e o boleto de um fornecedor nunca se misturam com
    os de outro e-mail do mesmo lote.

    Retorna uma lista de dicts: [{id, data, remetente, assunto, corpo, anexos, origem}]
    (o primeiro item é a própria mensagem recebida)."""
    bruto = servico.users().messages().get(userId="me", id=msg_id, format="raw").execute()
    raiz = message_from_bytes(_b64(bruto["raw"]), policy=policy.default)
    saida = []
    _desmontar(raiz, saida, msg_id)
    return saida


def mensagens_de_eml(conteudo: bytes, msg_id: str):
    """Mesma desmontagem, para um arquivo .eml salvo em disco."""
    saida = []
    _desmontar(message_from_bytes(conteudo, policy=policy.default), saida, msg_id)
    return saida


def _desmontar(m, saida, msg_id):
    try:
        data = parsedate_to_datetime(str(m["date"])).replace(tzinfo=None)
    except Exception:
        data = None
    info = {"id": msg_id, "origem": len(saida), "data": data,
            "remetente": str(m["from"] or ""), "assunto": str(m["subject"] or ""),
            "corpo": "", "anexos": []}
    saida.append(info)
    textos, anexos = [], []

    def percorrer(parte):
        tipo = parte.get_content_type()
        if tipo == "message/rfc822":
            for sub in parte.get_payload() or []:
                _desmontar(sub, saida, msg_id)
            return
        if parte.is_multipart():
            for sub in parte.get_payload():
                percorrer(sub)
            return
        nome = parte.get_filename()
        if nome:
            try:
                conteudo = parte.get_payload(decode=True) or b""
            except Exception:
                conteudo = b""
            if nome.lower().endswith(".eml"):
                try:
                    _desmontar(message_from_bytes(conteudo, policy=policy.default), saida, msg_id)
                except Exception:
                    pass
            else:
                anexos.append((nome, conteudo))
        elif tipo in ("text/plain", "text/html"):
            try:
                txt = parte.get_content()
            except Exception:
                txt = (parte.get_payload(decode=True) or b"").decode("utf-8", errors="ignore")
            if tipo == "text/html":
                txt = re.sub(r"<[^>]+>", " ", txt)
            textos.append(txt)

    percorrer(m)
    info["corpo"] = "\n".join(textos)
    info["anexos"] = expandir_zips(anexos)


def expandir_zips(anexos):
    saida = []
    for nome, conteudo in anexos:
        if nome.lower().endswith(".zip"):
            try:
                with zipfile.ZipFile(io.BytesIO(conteudo)) as z:
                    for info in z.infolist():
                        n = info.filename.lower()
                        if n.endswith((".xml", ".pdf")) and info.file_size < 20_000_000:
                            saida.append((info.filename.split("/")[-1], z.read(info)))
            except zipfile.BadZipFile:
                pass
        else:
            saida.append((nome, conteudo))
    return saida
