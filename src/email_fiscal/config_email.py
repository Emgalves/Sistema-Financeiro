# -*- coding: utf-8 -*-
"""
config_email.py - Configurações do módulo de leitura de documentos fiscais do Gmail.

Este é o ÚNICO arquivo que você precisa ajustar. Tudo o mais lê daqui.
"""

VERSAO = "2026-09-26.4"   # igual em todos os arquivos do módulo
from pathlib import Path

# ---------------------------------------------------------------------------
# 1) Caminhos do seu sistema (reaproveita o src/config/config.py que já existe)
# ---------------------------------------------------------------------------
try:
    from src.config.config import PASTA_CLIENTES, ARQUIVO_CLIENTES, ARQUIVO_FORNECEDORES, ENV
except ImportError:
    try:
        from config.config import PASTA_CLIENTES, ARQUIVO_CLIENTES, ARQUIVO_FORNECEDORES, ENV
    except ImportError:
        # Fallback só para testes fora do sistema
        PASTA_CLIENTES = Path("Clientes")
        ARQUIVO_CLIENTES = Path("Clientes.xlsx")
        ARQUIVO_FORNECEDORES = Path("Base_Fornecedores.xlsx")
        ENV = "teste"

PASTA_CLIENTES = Path(PASTA_CLIENTES)
ARQUIVO_CLIENTES = Path(ARQUIVO_CLIENTES)
ARQUIVO_FORNECEDORES = Path(ARQUIVO_FORNECEDORES)

# ---------------------------------------------------------------------------
# Pastas do módulo, SEPARADAS POR AMBIENTE (seguem o src/config/config.py)
# ---------------------------------------------------------------------------
# Triagem e anexos ficam em <Financeiro>/Email_Fiscal, ao lado de
# 'Planilhas_Base' e 'Clientes' - fora da pasta Clientes, para nenhum
# relatório que percorra essa pasta confundir a triagem com um cliente.
#   TESTE   : C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Email_Fiscal
#   PRODUÇÃO: <BASE_DADOS>/Email_Fiscal
PASTA_EMAIL_FISCAL = PASTA_CLIENTES.parent / "Email_Fiscal"
PASTA_DOCUMENTOS = PASTA_EMAIL_FISCAL / "Documentos_Email"        # anexos: AAAA-MM/
ARQUIVO_CAIXA = PASTA_EMAIL_FISCAL / "Documentos_Recebidos.xlsx"   # planilha de triagem

# Credenciais do Google FORA do projeto e SEPARADAS POR AMBIENTE:
# o ambiente de teste lê uma caixa de e-mail e o de produção, outra.
#   C:/Users/<você>/AppData/Roaming/SistemaFinanceiro/gmail/teste/
#   C:/Users/<você>/AppData/Roaming/SistemaFinanceiro/gmail/producao/
import os
PASTA_CREDENCIAIS = (Path(os.environ.get("APPDATA", str(Path.home())))
                     / "SistemaFinanceiro" / "gmail" / str(ENV))
ARQUIVO_CREDENTIALS = PASTA_CREDENCIAIS / "credentials.json"   # baixado do Google Cloud
ARQUIVO_TOKEN = PASTA_CREDENCIAIS / "token.json"               # criado no 1º login

# Nome da aba de lançamentos nas planilhas de cada cliente
ABA_DADOS_CLIENTE = "Dados"

# ---------------------------------------------------------------------------
# 2) Busca no Gmail
# ---------------------------------------------------------------------------
# Mesma sintaxe da caixa de pesquisa do Gmail.
CONSULTA_GMAIL = "has:attachment (filename:xml OR filename:pdf OR filename:zip) -in:sent -in:chats"
DIAS_RETROATIVOS = 30             # só olha e-mails dos últimos N dias
MAX_MENSAGENS_POR_EXECUCAO = 300  # limite de segurança por execução

# Senhas para tentar abrir PDFs protegidos (alguns bancos protegem boletos).
# Ex.: ["12345", "0123"]  -> deixe vazio se não usar.
SENHAS_PDF = []

# ---------------------------------------------------------------------------
# 3) Como o documento vira lançamento
# ---------------------------------------------------------------------------
# Regra de TP_DESP pela data (definida pelo usuário), aplicada NA IMPORTAÇÃO,
# quando a DATA_REL já é conhecida:
#   data do e-mail ANTERIOR à data do relatório e vencimento POSTERIOR -> "3"
#   data do e-mail ANTERIOR à data do relatório e vencimento ATÉ ela     -> "5"
#   (vencimento no próprio dia do relatório conta como "até ela" -> "5")
# Fora disso (e-mail na data do relatório ou depois, sem data ou sem
# vencimento) vale o TP_DESP digitado na planilha de triagem; se estiver
# vazio, o item não é importado até ser preenchido.
REGRA_TP_DESP_POR_DATA = True
TIPOS_COM_REGRA_POR_DATA = ("NFE", "NFSE", "BOLETO")

# Usada só quando REGRA_TP_DESP_POR_DATA = False.
# !!! CONFIRA ESTA TABELA !!!
# TP_DESP foi inferido do seu histórico (materiais com NF e vencimento futuro
# aparecem como TP_DESP 3; serviços de terceiros com NF como 5).
TP_DESP_POR_TIPO = {
    "NFE": "3",      # NF-e de produto (materiais)
    "NFSE": "5",     # NFS-e de serviço
    "BOLETO": "3",   # boleto sem nota fiscal no mesmo e-mail
}

# Categoria usada quando o fornecedor NÃO está na Base_Fornecedores.
# Se estiver, vale a CATEGORIA cadastrada lá.
CATEGORIA_PADRAO_POR_TIPO = {
    "NFE": "MAT",
    "NFSE": "SERV",
    "BOLETO": "",
}

# Texto gravado em OBSERVAÇÃO. O "#<id>" é o que permite ao sistema saber,
# depois, que a sugestão já foi salva na planilha do cliente.
PREFIXO_OBSERVACAO = "IMPORTADO EMAIL"

# CNPJs da própria empresa (quem paga em nome das obras). Nunca são tratados
# como fornecedor ao ler um boleto: aparecem como PAGADOR em compras feitas
# pela empresa. CONFIRA e acrescente outros CNPJs do grupo, se houver.
CNPJS_EMPRESA = [
    "30104762000107",   # RVR SERVICOS DE ENGENHARIA E CONSTRUCAO LTDA (Vasconcelos & Rinaldi)
]

# Tolerância em reais para casar boleto com parcela da NF (arredondamentos).
TOLERANCIA_VALOR = 0.05
