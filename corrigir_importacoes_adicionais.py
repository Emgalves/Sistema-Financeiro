"""
Script para corrigir problemas de importação adicionais
"""
import os
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
    
    # Criar redirecionamento para configuracoes_sistema.py
    config_sistema_path = os.path.join(dist_dir, 'configuracoes_sistema.py')
    config_sistema_content = """# Redirecionamento para src.configuracoes_sistema
from src.configuracoes_sistema import *
"""
    criar_arquivo_com_conteudo(config_sistema_path, config_sistema_content)
    
    # Modificar o launcher para incluir mais redirecionamentos
    launcher_path = os.path.join(dist_dir, 'launcher.py')
    launcher_content = """# Launcher para Sistema de Gestão Financeira
import os
import sys
import traceback
import importlib

# Configurar ambiente
os.environ['SISTEMA_AMBIENTE'] = 'producao'
os.environ['DEV_MODE'] = 'False'

# Adicionar diretório atual ao path
current_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, current_dir)

# Lista de módulos a serem redirecionados (origem -> destino)
redirecionamentos = {
    'config': 'src.config',
    'config.config': 'src.config.config',
    'config.utils': 'src.config.utils',
    'config.window_config': 'src.config.window_config',
    'config.logger_config': 'src.config.logger_config',
    'configuracoes_sistema': 'src.configuracoes_sistema',
    'Sistema_Entrada_Dados': 'src.Sistema_Entrada_Dados',
    'gestao_taxas': 'src.gestao_taxas',
    'finalizacao_quinzena': 'src.finalizacao_quinzena',
    'relatorio_despesas_aprimorado': 'src.relatorio_despesas_aprimorado',
    'version_control': 'src.version_control'
}

# Aplicar todos os redirecionamentos
for destino, origem in redirecionamentos.items():
    try:
        # Verificar se o módulo de origem existe
        if importlib.util.find_spec(origem):
            # Carregar o módulo de origem
            modulo = importlib.import_module(origem)
            # Adicionar redirecionamento
            sys.modules[destino] = modulo
            print(f"Redirecionamento aplicado: {destino} -> {origem}")
    except Exception as e:
        print(f"Erro ao redirecionar {destino} -> {origem}: {e}")

try:
    # Importar sistema principal
    print("Importando sistema principal...")
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
    try:
        import tkinter as tk
        from tkinter import messagebox
        
        root = tk.Tk()
        root.withdraw()
        messagebox.showerror("Erro de Inicialização", 
                          f"Erro ao iniciar Sistema de Gestão:\\n{str(e)}\\n\\n"
                          "Detalhes salvos em 'erro_sistema.log'")
    except:
        print(f"ERRO CRÍTICO: {str(e)}")
        print("Detalhes salvos em erro_sistema.log")
"""
    criar_arquivo_com_conteudo(launcher_path, launcher_content)
    
    print("\nPROCESSO CONCLUÍDO!")
    print("====================")
    print("Para iniciar o sistema:")
    print(f"1. Navegue até a pasta: {dist_dir}")
    print("2. Execute o arquivo 'Iniciar_Sistema.bat'")
    print("====================")

if __name__ == "__main__":
    main()