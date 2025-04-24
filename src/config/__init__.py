# Arquivo src/config/__init__.py
# Este arquivo torna a pasta config um pacote Python

# Importar funções comuns para facilitar o acesso
from .window_config import configurar_janela

# Exporta módulos
__all__ = ['window_config', 'configurar_janela']