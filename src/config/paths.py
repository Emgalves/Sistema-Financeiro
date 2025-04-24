# src/config/paths.py
from pathlib import Path
import os
import sys

# Determinar diretório base
BASE_DIR = Path(__file__).resolve().parent.parent.parent

# Definir caminhos padrão
SRC_DIR = BASE_DIR / 'src'
CONFIG_DIR = SRC_DIR / 'config'

# Pasta de dados
DATA_DIR = BASE_DIR / 'dados'
os.makedirs(DATA_DIR, exist_ok=True)

# Caminhos específicos
ARQUIVO_CLIENTES = DATA_DIR / "clientes.xlsx"
ARQUIVO_MODELO = DATA_DIR / "MODELO.xlsx"
PASTA_CLIENTES = DATA_DIR / "clientes"
ARQUIVO_FORNECEDORES = DATA_DIR / "fornecedores.xlsx"

# Garantir que as pastas existam
os.makedirs(PASTA_CLIENTES, exist_ok=True)

# Verificar pasta compartilhada no Google Drive
def verificar_pasta_drive():
    drive_path = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios")
    if drive_path.exists():
        return drive_path / "Financeiro"
    return None

# Verificar e atualizar caminhos se estiver no ambiente com Google Drive
drive_path = verificar_pasta_drive()
if drive_path is not None and drive_path.exists():
    ARQUIVO_CLIENTES = drive_path / "Planilhas_Base" / "clientes.xlsx"
    ARQUIVO_MODELO = drive_path / "Planilhas_Base" / "MODELO.xlsx"
    PASTA_CLIENTES = drive_path / "Clientes"
    ARQUIVO_FORNECEDORES = drive_path / "Planilhas_Base" / "fornecedores.xlsx"

# Função para obter o caminho base
def obter_base_path():
    """Retorna o caminho base para os arquivos de dados"""
    drive_path = verificar_pasta_drive()
    if drive_path is not None:
        return str(drive_path / "Planilhas_Base")
    return str(DATA_DIR)

# Função para obter a pasta de clientes
def obter_pasta_clientes():
    """Retorna o caminho para a pasta de clientes"""
    drive_path = verificar_pasta_drive()
    if drive_path is not None:
        return str(drive_path / "Clientes")
    return str(PASTA_CLIENTES)

# Adicionar caminhos ao sys.path
def configurar_paths():
    paths = [str(BASE_DIR), str(SRC_DIR), str(CONFIG_DIR)]
    for path in paths:
        if path not in sys.path:
            sys.path.append(path)

# Executar configuração quando o módulo é importado
configurar_paths()