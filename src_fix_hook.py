"""
Hook para corrigir problemas de importação no PyInstaller
Este script é executado automaticamente durante a inicialização do aplicativo
"""
import os
import sys
import importlib.util

# Adicionar diretório atual ao caminho Python
base_dir = os._getenv('_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, base_dir)

# Criar redirecionamento para config
if 'config' not in sys.modules:
    # Adicionar alias para src.config
    sys.modules['config'] = importlib.import_module('src.config')
    
    # Adicionar aliases para submódulos de config
    for submodule in ['config', 'utils', 'window_config', 'logger_config']:
        full_name = f'src.config.{submodule}'
        alias = f'config.{submodule}'
        if importlib.util.find_spec(full_name) and alias not in sys.modules:
            try:
                module = importlib.import_module(full_name)
                sys.modules[alias] = module
            except ImportError:
                pass

print("Hook de correção de importação executado com sucesso.")