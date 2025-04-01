# config.py
from pathlib import Path
import platform
import os

# Verifica o ambiente
ENV = os.getenv('SISTEMA_AMBIENTE', 'teste')
print(f"Ambiente atual: {ENV}")

# Detecta o sistema operacional
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'

if ENV == 'teste':
    if IS_WINDOWS:
        GOOGLE_DRIVE_PATH = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios")
    elif IS_MAC:
        GOOGLE_DRIVE_PATH = Path("/Users/emiliamargareth/Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive")
    
    # Define os caminhos base para diferentes pastas
    BASE_PATH = GOOGLE_DRIVE_PATH / "Financeiro/Planilhas_Base"
    PASTA_CLIENTES = GOOGLE_DRIVE_PATH / "Financeiro/Clientes"

    print(f"BASE_PATH definido como: {BASE_PATH}")
    print(f"PASTA_CLIENTES definida como: {PASTA_CLIENTES}")

else:
    # Ambiente de teste - usar caminho fixo
    BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
    PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')

# Define caminhos específicos
ARQUIVO_CLIENTES = BASE_PATH / "clientes.xlsx"
ARQUIVO_FORNECEDORES = BASE_PATH / "base_fornecedores.xlsx"
ARQUIVO_MODELO = BASE_PATH / "MODELO.xlsx"
ARQUIVO_CONTROLE = BASE_PATH / "controle_taxa_adm.xlsx"

print(f"Verificando se o diretório existe:")
print(f"GOOGLE_DRIVE_PATH existe? {GOOGLE_DRIVE_PATH.exists()}")
print(f"BASE_PATH existe? {BASE_PATH.exists()}")
print(f"PASTA_CLIENTES existe? {PASTA_CLIENTES.exists()}")

print(f"ARQUIVO_CLIENTES: {ARQUIVO_CLIENTES}")
print(f"ARQUIVO_MODELO: {ARQUIVO_MODELO}")

# Criar as pastas se não existirem no ambiente de teste
if ENV == 'producao':
    try:
        BASE_PATH.mkdir(parents=True, exist_ok=True)
        PASTA_CLIENTES.mkdir(parents=True, exist_ok=True)
        print(f"Pastas criadas/verificadas com sucesso")
    except Exception as e:
        print(f"Erro ao criar pastas: {e}")

def verificar_arquivos():
    """Verifica se todos os arquivos necessários estão acessíveis"""
    arquivos = [ARQUIVO_CLIENTES, ARQUIVO_FORNECEDORES, ARQUIVO_MODELO, ARQUIVO_CONTROLE]
    for arquivo in arquivos:
        existe = arquivo.exists()
        print(f"Verificando arquivo {arquivo}: {'existe' if existe else 'não existe'}")
        if not existe:
            raise FileNotFoundError(f"Arquivo não encontrado: {arquivo}")
