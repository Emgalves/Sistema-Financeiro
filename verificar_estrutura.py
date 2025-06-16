# Criar este arquivo como: verificar_estrutura.py
# Execute-o uma vez para garantir que todas as pastas necessárias existam

import os
import sys
from pathlib import Path

def criar_estrutura_pastas():
    """
    Cria a estrutura de pastas necessária para o sistema
    """
    print("=== Verificando e Criando Estrutura de Pastas ===")
    
    # Determinar diretório base
    if getattr(sys, 'frozen', False):
        # Executável PyInstaller
        base_dir = os.path.dirname(sys.executable)
        print(f"Modo executável detectado. Base: {base_dir}")
    else:
        # Desenvolvimento
        base_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"Modo desenvolvimento detectado. Base: {base_dir}")
    
    # Pastas necessárias
    pastas_necessarias = [
        'logs',
        'dados',
        'clientes',
        'relatorios',
        'temp',
        'config',
        'backup'
    ]
    
    created_count = 0
    for pasta in pastas_necessarias:
        pasta_path = os.path.join(base_dir, pasta)
        try:
            if not os.path.exists(pasta_path):
                os.makedirs(pasta_path, exist_ok=True)
                print(f"✓ Pasta criada: {pasta}")
                created_count += 1
            else:
                print(f"✓ Pasta já existe: {pasta}")
        except Exception as e:
            print(f"✗ Erro ao criar pasta {pasta}: {str(e)}")
    
    print(f"\nResumo: {created_count} pastas criadas")
    
    # Verificar permissões de escrita
    print("\n=== Verificando Permissões de Escrita ===")
    for pasta in pastas_necessarias:
        pasta_path = os.path.join(base_dir, pasta)
        try:
            test_file = os.path.join(pasta_path, 'test_write.tmp')
            with open(test_file, 'w') as f:
                f.write('teste')
            os.remove(test_file)
            print(f"✓ Escrita OK: {pasta}")
        except Exception as e:
            print(f"✗ Erro de escrita em {pasta}: {str(e)}")
    
    # Criar arquivo de configuração básico se não existir
    config_file = os.path.join(base_dir, 'config', 'sistema_config.json')
    if not os.path.exists(config_file):
        try:
            import json
            config_basica = {
                "sistema": {
                    "versao": "1.0.0",
                    "logging_habilitado": True,
                    "pasta_logs": "logs",
                    "pasta_dados": "dados",
                    "pasta_clientes": "clientes"
                },
                "interface": {
                    "tema": "default",
                    "largura_janela": 800,
                    "altura_janela": 600
                }
            }
            
            with open(config_file, 'w', encoding='utf-8') as f:
                json.dump(config_basica, f, indent=4, ensure_ascii=False)
            print(f"✓ Configuração básica criada: {config_file}")
        except Exception as e:
            print(f"✗ Erro ao criar configuração: {str(e)}")
    
    print("\n=== Verificação Concluída ===")
    return True

if __name__ == "__main__":
    try:
        sucesso = criar_estrutura_pastas()
        if sucesso:
            print("Estrutura de pastas verificada com sucesso!")
            input("Pressione Enter para continuar...")
        else:
            print("Houve problemas na verificação da estrutura.")
            input("Pressione Enter para continuar...")
    except Exception as e:
        print(f"Erro crítico: {str(e)}")
        input("Pressione Enter para continuar...")