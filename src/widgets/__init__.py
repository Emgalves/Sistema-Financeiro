# Arquivo src/__init__.py
# Este arquivo torna a pasta src um pacote Python

# Importações básicas que serão usadas por outros módulos
from pathlib import Path
import os
import sys

# Define BASE_PATH como uma constante global
BASE_DIR = Path(__file__).resolve().parent.parent

# Exporta subpacotes
__all__ = ['config']