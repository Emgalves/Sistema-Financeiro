# Arquivo src/__init__.py
# Este arquivo torna a pasta src um pacote Python

# Importações básicas que serão usadas por outros módulos
from pathlib import Path
import os
import sys

# Define BASE_PATH como uma constante global
BASE_DIR = Path(__file__).resolve().parent.parent

# ====================================================================
# PRÉ-CARREGAR MÓDULOS ESSENCIAIS PARA PYINSTALLER
# ====================================================================
# Isso garante que o PyInstaller inclua e consiga importar estes módulos

print("🔄 Inicializando pacote src...")

# Tentar carregar módulos essenciais
_modulos_carregados = []
_modulos_falhados = []

# Módulo de ambiente
try:
    from . import ambiente_config
    _modulos_carregados.append('ambiente_config')
except ImportError as e:
    _modulos_falhados.append(('ambiente_config', str(e)))

# Módulo de controle de versão
try:
    from . import version_control
    _modulos_carregados.append('version_control')
except ImportError as e:
    _modulos_falhados.append(('version_control', str(e)))

# Módulo de relatórios (IMPORTANTE!)
try:
    from . import relatorios_interface
    _modulos_carregados.append('relatorios_interface')
except ImportError as e:
    _modulos_falhados.append(('relatorios_interface', str(e)))

# Sistema de entrada de dados
try:
    from . import Sistema_Entrada_Dados
    _modulos_carregados.append('Sistema_Entrada_Dados')
except ImportError as e:
    _modulos_falhados.append(('Sistema_Entrada_Dados', str(e)))

# Exibir resultados do carregamento
if _modulos_carregados:
    print(f"✅ Módulos carregados: {', '.join(_modulos_carregados)}")

if _modulos_falhados:
    print(f"⚠️ Módulos não carregados:")
    for modulo, erro in _modulos_falhados:
        print(f"   - {modulo}: {erro}")

# Exporta subpacotes e módulos
__all__ = [
    'config',
    'ambiente_config',
    'version_control',
    'relatorios_interface',
    'Sistema_Entrada_Dados',
]

__version__ = "1.4.5"