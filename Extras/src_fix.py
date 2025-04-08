
# Fix para importações relativas vs. absolutas
import sys
import importlib

# Criar alias para config -> src.config
sys.modules['config'] = importlib.import_module('src.config')
for submodule in ['config', 'utils', 'window_config', 'logger_config']:
    full_name = f'src.config.{submodule}'
    alias = f'config.{submodule}'
    if alias not in sys.modules:
        try:
            sys.modules[alias] = importlib.import_module(full_name)
        except ImportError:
            pass
