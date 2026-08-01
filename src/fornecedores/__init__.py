# Arquivo src/fornecedores/__init__.py
# Este arquivo torna a pasta fornecedores um pacote Python

from .cache_fornecedores import CacheFornecedores
from .gerenciador_cpfs_criados import GerenciadorCPFsCriados
from .regularizar_fornecedor import (
    regularizar_fornecedor, fundir_fornecedores, detectar_possiveis_duplicatas
)