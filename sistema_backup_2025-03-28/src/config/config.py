# config.py
from pathlib import Path
import platform
import os

# Verifica o ambiente de forma mais robusta
ENV = os.getenv('SISTEMA_AMBIENTE', '')

# Se não estiver explicitamente em 'teste', assume produção
if ENV.lower() != 'teste':
    ENV = 'producao'

print(f"Ambiente atual: {ENV}")

# Detecta o sistema operacional
IS_WINDOWS = platform.system() == 'Windows'
IS_MAC = platform.system() == 'Darwin'

# Inicializa a variável GOOGLE_DRIVE_PATH como None
GOOGLE_DRIVE_PATH = None

if ENV == 'producao':
    if IS_WINDOWS:
        # Lista de possíveis caminhos do Google Drive em Windows
        possiveis_caminhos = [
            Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios"),
            # Caminho alternativo comum para Google Drive no Windows
            Path(os.path.expanduser("~")) / "Google Drive",
            # Outro formato possível 
            Path(os.path.expanduser("~")) / "AppData/Local/Google/Drive/shared_drives"
        ]
        
        # Tenta encontrar um caminho válido
        for caminho in possiveis_caminhos:
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"Google Drive encontrado em: {GOOGLE_DRIVE_PATH}")
                break
                
    elif IS_MAC:
        # Lista de possíveis caminhos do Google Drive em Mac
        possiveis_caminhos = [
            Path("/Users/emiliamargareth/Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive"),
            Path(os.path.expanduser("~")) / "Library/CloudStorage/GoogleDrive-emilia.mga@gmail.com/Meu Drive",
            Path(os.path.expanduser("~")) / "Google Drive"
        ]
        
        # Tenta encontrar um caminho válido
        for caminho in possiveis_caminhos:
            if caminho.exists():
                GOOGLE_DRIVE_PATH = caminho
                print(f"Google Drive encontrado em: {GOOGLE_DRIVE_PATH}")
                break
    
    # Se não encontrou o Google Drive, mostra um aviso
    if GOOGLE_DRIVE_PATH is None:
        print("AVISO: Não foi possível encontrar o Google Drive. Usando caminho local.")
        # Usar caminho local como fallback
        BASE_PATH = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Planilhas_Base')
        PASTA_CLIENTES = Path('C:/Users/Obras/sistema_gestao_testes/testes/Financeiro/Clientes')
    else:
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

# Verificar se os diretórios existem
print(f"Verificando se o diretório existe:")
if GOOGLE_DRIVE_PATH is not None:
    print(f"GOOGLE_DRIVE_PATH existe? {GOOGLE_DRIVE_PATH.exists()}")
print(f"BASE_PATH existe? {BASE_PATH.exists()}")
print(f"PASTA_CLIENTES existe? {PASTA_CLIENTES.exists()}")

print(f"ARQUIVO_CLIENTES: {ARQUIVO_CLIENTES}")
print(f"ARQUIVO_MODELO: {ARQUIVO_MODELO}")

# Criar as pastas se não existirem
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