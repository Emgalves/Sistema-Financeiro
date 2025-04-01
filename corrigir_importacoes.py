"""
Script para corrigir problemas de importação na pasta dist existente
"""
import os
import shutil
import sys

def criar_arquivo_com_conteudo(caminho, conteudo):
    """Cria um arquivo com o conteúdo especificado"""
    try:
        with open(caminho, 'w', encoding='utf-8') as f:
            f.write(conteudo)
        print(f"✓ Arquivo criado: {caminho}")
        return True
    except Exception as e:
        print(f"✗ Erro ao criar arquivo {caminho}: {e}")
        return False

def main():
    # Pasta do executável existente
    dist_dir = os.path.join('dist', 'Sistema_Gestao_Financeira')
    
    if not os.path.exists(dist_dir):
        print(f"Pasta {dist_dir} não encontrada!")
        return
        
    print(f"Trabalhando na pasta: {dist_dir}")
    
    # 1. Criar arquivo config.py na raiz (redirecionamento)
    config_path = os.path.join(dist_dir, 'config.py')
    config_content = """# Redirecionamento config -> src.config
import os
import sys
import importlib.util

# Carregar variáveis de ambiente
os.environ['SISTEMA_AMBIENTE'] = 'producao'
os.environ['DEV_MODE'] = 'False'

# Importar src.config
try:
    from src.config import *
except ImportError as e:
    print(f"Erro ao importar src.config: {e}")
    
# Módulos de config que precisamos redirecionar
modules = ["config", "utils", "window_config", "logger_config"]

# Criar redirecionamentos para config.X -> src.config.X
for module_name in modules:
    src_module = f"src.config.{module_name}"
    tgt_module = f"config.{module_name}"
    
    try:
        if src_module not in sys.modules:
            continue
            
        # Adicionar ao sys.modules para redirecionamento
        if tgt_module not in sys.modules:
            sys.modules[tgt_module] = sys.modules[src_module]
    except Exception as e:
        print(f"Erro ao redirecionar {src_module} -> {tgt_module}: {e}")
"""
    criar_arquivo_com_conteudo(config_path, config_content)
    
    # 2. Criar pasta config na raiz (necessário para importações)
    config_dir = os.path.join(dist_dir, 'config')
    if not os.path.exists(config_dir):
        os.makedirs(config_dir, exist_ok=True)
        print(f"✓ Pasta criada: {config_dir}")
    
    # 3. Criar __init__.py dentro da pasta config
    init_path = os.path.join(config_dir, '__init__.py')
    init_content = """# Redirecionamento para src.config
from src.config import *
"""
    criar_arquivo_com_conteudo(init_path, init_content)
    
    # 4. Redirecionar cada módulo principal
    for module in ["config", "utils", "window_config", "logger_config"]:
        module_path = os.path.join(config_dir, f"{module}.py")
        module_content = f"""# Redirecionamento para src.config.{module}
from src.config.{module} import *
"""
        criar_arquivo_com_conteudo(module_path, module_content)
    
    # 5. Criar arquivo launcher.py na raiz
    launcher_path = os.path.join(dist_dir, 'launcher.py')
    launcher_content = """# Launcher para Sistema de Gestão Financeira
import os
import sys
import traceback

# Configurar ambiente
os.environ['SISTEMA_AMBIENTE'] = 'producao'
os.environ['DEV_MODE'] = 'False'

try:
    # Importar config para garantir redirecionamentos
    import config
    
    # Importar sistema principal
    from src import sistema_principal
    
    # Iniciar aplicação
    print("Iniciando Sistema de Gestão Financeira...")
    app = sistema_principal.SistemaPrincipal()
    app.run()
    
except Exception as e:
    # Registrar erro
    with open("erro_sistema.log", "w") as f:
        f.write(f"ERRO AO INICIAR: {str(e)}\\n")
        f.write(traceback.format_exc())
    
    # Mostrar mensagem na tela
    import tkinter as tk
    from tkinter import messagebox
    
    root = tk.Tk()
    root.withdraw()
    messagebox.showerror("Erro de Inicialização", 
                      f"Erro ao iniciar Sistema de Gestão:\\n{str(e)}\\n\\n"
                      "Detalhes salvos em 'erro_sistema.log'")
"""
    criar_arquivo_com_conteudo(launcher_path, launcher_content)
    
    # 6. Criar arquivo de lote para executar launcher
    bat_path = os.path.join(dist_dir, 'Iniciar_Sistema.bat')
    bat_content = """@echo off
echo Iniciando Sistema de Gestao Financeira...
python launcher.py
pause
"""
    criar_arquivo_com_conteudo(bat_path, bat_content)
    
    print("\nPROCESSO CONCLUÍDO!")
    print("====================")
    print("Para iniciar o sistema:")
    print(f"1. Navegue até a pasta: {dist_dir}")
    print("2. Execute o arquivo 'Iniciar_Sistema.bat'")
    print("====================")

if __name__ == "__main__":
    main()