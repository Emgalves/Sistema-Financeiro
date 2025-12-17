# src/config/paths.py
"""
SEMPRE VERIFICAR ESTES ARQUIVOS PARA MANTER CONSISTÊNCIA:
 - src/ambiente_config.py
    - src/config/paths.py
    - src/config/config.py
    - src/config/__init__.py
"""
from pathlib import Path
import os
import sys

# ============================================================================
# USAR AMBIENTE_CONFIG.PY PARA DETERMINAR CAMINHOS
# ============================================================================

try:
    from src.ambiente_config import config_ambiente
    USA_AMBIENTE_CONFIG = True
    print("✅ paths.py: Usando ambiente_config")
except ImportError:
    USA_AMBIENTE_CONFIG = False
    print("⚠️ paths.py: ambiente_config não disponível")

# Determinar diretório base
BASE_DIR = Path(__file__).resolve().parent.parent.parent

# Definir caminhos padrão
SRC_DIR = BASE_DIR / 'src'
CONFIG_DIR = SRC_DIR / 'config'

# ============================================================================
# DETERMINAR CAMINHOS BASEADO NO AMBIENTE
# ============================================================================

if USA_AMBIENTE_CONFIG and config_ambiente.eh_producao():
    # PRODUÇÃO - Usar Google Drive
    print("🟢 paths.py: Configurando para PRODUÇÃO")
    
    drive_path = Path("H:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios/Financeiro")
    
    if drive_path.exists():
        BASE_PATH = drive_path / "Planilhas_Base"
        PASTA_CLIENTES = drive_path / "Clientes"
        DATA_DIR = drive_path
        print(f"✅ Usando Google Drive: {drive_path}")
    else:
        # Fallback
        print(f"⚠️ Google Drive não encontrado em {drive_path}")
        print(f"   Tentando caminhos alternativos...")
        
        # Tentar outros caminhos
        caminhos_alternativos = [
            Path("G:/.shortcut-targets-by-id/195uuohIL_ZKum7lhwu-OzJCH_CGAb97G/Relatórios/Financeiro"),
            Path("H:/Drives compartilhados/Relatórios/Financeiro"),
            Path("G:/Drives compartilhados/Relatórios/Financeiro"),
        ]
        
        drive_encontrado = False
        for caminho in caminhos_alternativos:
            if caminho.exists():
                drive_path = caminho
                BASE_PATH = drive_path / "Planilhas_Base"
                PASTA_CLIENTES = drive_path / "Clientes"
                DATA_DIR = drive_path
                print(f"✅ Usando caminho alternativo: {drive_path}")
                drive_encontrado = True
                break
        
        if not drive_encontrado:
            print(f"⚠️ Nenhum caminho do Google Drive encontrado, usando local")
            DATA_DIR = BASE_DIR / 'dados'
            BASE_PATH = DATA_DIR
            PASTA_CLIENTES = DATA_DIR / "clientes"
else:
    # TESTE - Usar caminhos locais
    print("🟨 paths.py: Configurando para TESTE")
    
    # ✅ CAMINHO CORRETO (sem pasta "testes" extra)
    DATA_DIR = Path('C:/Users/Obras/sistema_gestao_testes/Financeiro')
    BASE_PATH = DATA_DIR / "Planilhas_Base"
    PASTA_CLIENTES = DATA_DIR / "Clientes"
    print(f"✅ Usando caminho local de teste: {DATA_DIR}")

# Garantir que as pastas existam (apenas em modo teste)
if not (USA_AMBIENTE_CONFIG and config_ambiente.eh_producao()):
    try:
        os.makedirs(DATA_DIR, exist_ok=True)
        os.makedirs(PASTA_CLIENTES, exist_ok=True)
        print(f"✅ Pastas criadas/verificadas")
    except Exception as e:
        print(f"⚠️ Erro ao criar pastas: {e}")

# Caminhos específicos
ARQUIVO_CLIENTES = BASE_PATH / "Clientes.xlsx"
ARQUIVO_MODELO = BASE_PATH / "MODELO.xlsx"
ARQUIVO_FORNECEDORES = BASE_PATH / "base_fornecedores.xlsx"

print(f"📁 BASE_PATH: {BASE_PATH}")
print(f"📁 PASTA_CLIENTES: {PASTA_CLIENTES}")
print(f"📄 ARQUIVO_CLIENTES: {ARQUIVO_CLIENTES}")

# Função para obter o caminho base
def obter_base_path():
    """Retorna o caminho base para os arquivos de dados"""
    return str(BASE_PATH)

# Função para obter a pasta de clientes
def obter_pasta_clientes():
    """Retorna o caminho para a pasta de clientes"""
    return str(PASTA_CLIENTES)

# Adicionar caminhos ao sys.path
def configurar_paths():
    paths = [str(BASE_DIR), str(SRC_DIR), str(CONFIG_DIR)]
    for path in paths:
        if path not in sys.path:
            sys.path.append(path)

# Executar configuração quando o módulo é importado
configurar_paths()