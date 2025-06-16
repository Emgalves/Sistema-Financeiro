# Criar este arquivo como: teste_logging.py
# Execute para verificar se o sistema de logging está funcionando

import sys
import os
from pathlib import Path

# Adicionar diretório raiz ao path
def add_project_root():
    current_dir = Path(__file__).resolve().parent
    src_dir = current_dir / 'src'
    if src_dir.exists():
        project_root = current_dir
    else:
        project_root = current_dir.parent
    
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    
    if str(project_root / 'src') not in sys.path:
        sys.path.insert(0, str(project_root / 'src'))

add_project_root()

def teste_completo_logging():
    """Testa todo o sistema de logging"""
    print("=== TESTE DO SISTEMA DE LOGGING ===\n")
    
    # 1. Testar estrutura de pastas
    print("1. Verificando estrutura de pastas...")
    
    # Determinar diretório base
    if getattr(sys, 'frozen', False):
        base_dir = os.path.dirname(sys.executable)
        print(f"   Modo: Executável PyInstaller")
    else:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        print(f"   Modo: Desenvolvimento")
    
    print(f"   Diretório base: {base_dir}")
    
    # Verificar/criar pasta logs
    logs_dir = os.path.join(base_dir, 'logs')
    if not os.path.exists(logs_dir):
        try:
            os.makedirs(logs_dir, exist_ok=True)
            print(f"   ✓ Pasta logs criada: {logs_dir}")
        except Exception as e:
            print(f"   ✗ Erro ao criar pasta logs: {str(e)}")
            return False
    else:
        print(f"   ✓ Pasta logs já existe: {logs_dir}")
    
    # 2. Testar importação do logger existente
    print("\n2. Testando importação do sistema de logging...")
    
    try:
        from src.config.logger_config import system_logger, log_action, get_logging_status, test_logging
        print("   ✓ Importação do logger_config bem-sucedida")
    except ImportError as e:
        print(f"   ✗ Erro ao importar logger_config: {str(e)}")
        try:
            from config.logger_config import system_logger, log_action, get_logging_status, test_logging
            print("   ✓ Importação alternativa do logger_config bem-sucedida")
        except ImportError as e2:
            print(f"   ✗ Erro na importação alternativa: {str(e2)}")
            return False
    
    # 3. Testar configuração do logger
    print("\n3. Testando configuração do logger...")
    
    try:
        status = get_logging_status()
        print(f"   Log directory: {status['log_dir']}")
        print(f"   Log file: {status['log_file']}")
        print(f"   Handlers count: {status['handlers_count']}")
        print(f"   Current user: {status['current_user']}")
        
        if status['handlers_count'] > 0:
            print("   ✓ Logger configurado com handlers")
        else:
            print("   ⚠ Logger sem handlers configurados")
    
    except Exception as e:
        print(f"   ✗ Erro ao obter status do logger: {str(e)}")
        return False
    
    # 4. Testar funcionalidade do logger
    print("\n4. Testando funcionalidade do logger...")
    
    try:
        success = test_logging()
        if success:
            print("   ✓ Teste de logging bem-sucedido")
        else:
            print("   ✗ Teste de logging falhou")
            return False
    except Exception as e:
        print(f"   ✗ Erro no teste de logging: {str(e)}")
        return False
    
    # 5. Testar decorator log_action
    print("\n5. Testando decorator log_action...")
    
    try:
        @log_action("teste do decorator")
        def funcao_teste():
            return "sucesso"
        
        resultado = funcao_teste()
        if resultado == "sucesso":
            print("   ✓ Decorator log_action funcionando")
        else:
            print("   ✗ Decorator retornou resultado inesperado")
    except Exception as e:
        print(f"   ✗ Erro no teste do decorator: {str(e)}")
        return False
    
    # 6. Testar mudança de usuário
    print("\n6. Testando mudança de usuário...")
    
    try:
        system_logger.set_user("usuario_teste")
        logger = system_logger.get_logger()
        logger.info("Teste com usuário alterado")
        print("   ✓ Mudança de usuário bem-sucedida")
    except Exception as e:
        print(f"   ✗ Erro na mudança de usuário: {str(e)}")
        return False
    
    # 7. Verificar se arquivo de log foi criado
    print("\n7. Verificando criação de arquivo de log...")
    
    try:
        status = get_logging_status()
        if status['log_file'] and os.path.exists(status['log_file']):
            print(f"   ✓ Arquivo de log criado: {status['log_file']}")
            
            # Verificar se tem conteúdo
            try:
                with open(status['log_file'], 'r', encoding='utf-8') as f:
                    content = f.read()
                    if content:
                        print(f"   ✓ Arquivo tem conteúdo ({len(content)} caracteres)")
                    else:
                        print("   ⚠ Arquivo de log está vazio")
            except Exception as e:
                print(f"   ⚠ Não foi possível ler arquivo de log: {str(e)}")
        else:
            print("   ⚠ Arquivo de log não foi criado (usando apenas console)")
    except Exception as e:
        print(f"   ✗ Erro ao verificar arquivo de log: {str(e)}")
    
    print("\n=== TESTE CONCLUÍDO COM SUCESSO ===")
    return True

def teste_relatorios_interface():
    """Testa especificamente a importação do módulo de relatórios"""
    print("\n=== TESTE DO MÓDULO DE RELATÓRIOS ===\n")
    
    try:
        # Simular a importação que acontece no sistema
        print("1. Testando importação do relatorios_interface...")
        
        # Primeiro limpar cache se existir
        module_names = [
            'src.relatorios_interface',
            'relatorios_interface'
        ]
        
        for module_name in module_names:
            if module_name in sys.modules:
                del sys.modules[module_name]
                print(f"   Cache limpo: {module_name}")
        
        # Tentar importar
        modulo = None
        for module_name in module_names:
            try:
                import importlib
                modulo = importlib.import_module(module_name)
                print(f"   ✓ Importação bem-sucedida: {module_name}")
                break
            except ImportError as e:
                print(f"   ⚠ Falha na importação {module_name}: {str(e)}")
                continue
        
        if modulo:
            if hasattr(modulo, 'SistemaRelatorios'):
                print("   ✓ Classe SistemaRelatorios encontrada")
                return True
            else:
                print("   ✗ Classe SistemaRelatorios não encontrada")
                return False
        else:
            print("   ✗ Nenhum módulo pôde ser importado")
            return False
            
    except Exception as e:
        print(f"   ✗ Erro geral no teste: {str(e)}")
        return False

if __name__ == "__main__":
    print("Iniciando testes do sistema de logging...\n")
    
    try:
        # Teste principal
        sucesso_logging = teste_completo_logging()
        
        # Teste específico do módulo de relatórios
        sucesso_relatorios = teste_relatorios_interface()
        
        print(f"\nResumo dos testes:")
        print(f"  Sistema de logging: {'✓ OK' if sucesso_logging else '✗ FALHOU'}")
        print(f"  Módulo de relatórios: {'✓ OK' if sucesso_relatorios else '✗ FALHOU'}")
        
        if sucesso_logging and sucesso_relatorios:
            print("\n🎉 Todos os testes passaram! O sistema deve funcionar corretamente.")
        else:
            print("\n⚠️  Alguns testes falharam. Verifique os erros acima.")
        
    except Exception as e:
        print(f"\nErro crítico durante os testes: {str(e)}")
        import traceback
        traceback.print_exc()
    
    input("\nPressione Enter para continuar...")