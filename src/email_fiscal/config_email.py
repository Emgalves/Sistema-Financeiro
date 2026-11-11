# -*- coding: utf-8 -*-
"""
config_email.py - Configurações do módulo de leitura de documentos fiscais do Gmail.

Este é o ÚNICO arquivo que você precisa ajustar. Tudo o mais lê daqui.
"""

VERSAO = "2026-11-12.8"   # igual em todos os arquivos do módulo
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

# Proteções contra travamento numa coleta longa
TIMEOUT_GMAIL = 90          # segundos de espera por resposta do Gmail antes de desistir da mensagem
MAX_TAMANHO_PDF_MB = 15     # PDFs maiores não são lidos (vão para triagem manual)
MAX_PAGINAS_PDF = 8         # só as primeiras páginas são lidas (boleto/DANFE ficam no início)
SALVAR_A_CADA = 20          # grava a triagem a cada N mensagens: um travamento não perde tudo

# Muitos fornecedores protegem o boleto com os primeiros dígitos do CPF/CNPJ do
# pagador. Com True, o sistema tenta automaticamente, para cada cliente:
# 3, 4, 5 e 6 primeiros dígitos e o documento completo (além de SENHAS_PDF).
SENHAS_PELO_CPF_DOS_CLIENTES = True

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

# DATA_REL pelo vencimento: cada item vai para o relatório (dias 5 e 20) mais
# recente que seja ATÉ o vencimento, nunca anterior ao relatório em uso.
#   ex.: relatório em uso 05/11 -> venc. 17/11: 05/11 | venc. 23/11 e 26/11: 20/11
#        venc. 11/12 e 15/12: 05/12 | venc. no próprio dia 20/11: 20/11
#        venc. já passado (ou sem vencimento): relatório em uso
# O TP_DESP (regra acima) é calculado com a DATA_REL do próprio item.
DATA_REL_PELO_VENCIMENTO = True
DIAS_RELATORIO = (5, 20)
TIPOS_COM_REGRA_POR_DATA = ("NFE", "NFSE", "BOLETO", "MANUAL")  # MANUAL = PDF completado na edição

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

# Boleto importado sem NF, mas com o número da NF lido no próprio boleto:
# a REFERÊNCIA vai como "BOLETO - NF 35599 (A CONFIRMAR)" em vez de
# "BOLETO - AGUARDANDO NF". False = mantém "BOLETO - AGUARDANDO NF".
NF_PROVAVEL_NA_REFERENCIA = True

# Fecha a janela "📧 E-mail" sozinha quando, depois de uma ação (importar,
# ignorar, editar, comparar), não sobra nenhuma sugestão pendente do cliente.
FECHAR_JANELA_QUANDO_VAZIA = True

# NF-e SEM COBRANÇA não são contas a pagar e ficam IGNORADO automaticamente:
#  - natureza da operação com uma destas palavras E sem duplicatas na nota
#    (ex.: "REMESSA LOCACAO", "Remessa de bens", "Retorno de mercadorias");
#  - ou meio de pagamento = 90 (Sem pagamento) em todos os pagamentos da nota.
# Se a nota tiver duplicatas, ela NÃO é ignorada (recebe só um alerta).
NATUREZAS_SEM_COBRANCA = ("REMESSA", "RETORNO")
PAGAMENTOS_SEM_COBRANCA = ("90",)

# Tolerância em reais para casar boleto com parcela da NF (arredondamentos).
TOLERANCIA_VALOR = 0.05


# ---------------------------------------------------------------------------
# Caminhos de arquivos salvos
# ---------------------------------------------------------------------------
def resolver_arquivo(caminho):
    """Os caminhos gravados na triagem são absolutos e dependem do computador
    (Z:\\Servidor\\... numa máquina, H:\\.shortcut-targets...\\ em outra) e do
    ambiente. Se o caminho gravado não existe aqui, procura o MESMO arquivo
    dentro da pasta Documentos_Email deste computador/ambiente.
    Retorna o caminho (str) que existe, ou o original se não achar."""
    import re as _re
    original = str(caminho or "").strip()
    if not original or Path(original).exists():
        return original
    partes = [p for p in _re.split(r"[\\/]+", original) if p]
    if "Documentos_Email" in partes:
        i = len(partes) - 1 - partes[::-1].index("Documentos_Email")
        candidato = PASTA_DOCUMENTOS.joinpath(*partes[i + 1:])
        if candidato.exists():
            return str(candidato)
    return original
