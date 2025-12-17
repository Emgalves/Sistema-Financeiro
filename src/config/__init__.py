# src/config/__init__.py
# Este arquivo torna a pasta config um pacote Python
"""
SEMPRE VERIFICAR ESTES ARQUIVOS PARA MANTER CONSISTÊNCIA:
 - src/ambiente_config.py
    - src/config/paths.py
    - src/config/config.py
    - src/config/__init__.py
"""

# Importar configurações do config.py
from .config import (
    ENV,
    GOOGLE_DRIVE_PATH,
    IS_WINDOWS,
    IS_MAC,
    BASE_PATH,
    PASTA_CLIENTES,
    ARQUIVO_CLIENTES,
    ARQUIVO_FORNECEDORES,
    ARQUIVO_MODELO,
    ARQUIVO_CONTROLE,
    PASTA_RH,
    ARQUIVO_PARAMETROS_MATERIAIS,
    verificar_arquivos,
)

# Importar funções de window_config (que já existiam)
from .window_config import configurar_janela

# Exportar tudo
__all__ = [
    # Configurações de ambiente e caminhos
    'ENV',
    'GOOGLE_DRIVE_PATH',
    'IS_WINDOWS',
    'IS_MAC',
    'BASE_PATH',
    'PASTA_CLIENTES',
    'ARQUIVO_CLIENTES',
    'ARQUIVO_FORNECEDORES',
    'ARQUIVO_MODELO',
    'ARQUIVO_CONTROLE',
    'PASTA_RH',
    'ARQUIVO_PARAMETROS_MATERIAIS',
    'verificar_arquivos',
    # Configurações de janela
    'window_config',
    'configurar_janela',
]