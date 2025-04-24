"""
Script principal simplificado para compilação inicial
Foca apenas no módulo Sistema_Entrada_Dados
"""
import os
import sys
from pathlib import Path
import logging

# Configurar logging básico
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("sistema")

# Adicionar diretório raiz ao path para facilitar importações
current_dir = Path(__file__).resolve().parent
if str(current_dir) not in sys.path:
    sys.path.append(str(current_dir))
logger.info(f"Diretório atual adicionado ao PYTHONPATH: {current_dir}")

# Adicionar diretórios src e config ao path
src_dir = current_dir / 'src'
config_dir = src_dir / 'config'

for path in [src_dir, config_dir]:
    if path.exists() and str(path) not in sys.path:
        sys.path.append(str(path))
        logger.info(f"Caminho adicionado ao PYTHONPATH: {path}")

# Verificar e criar arquivos __init__.py se necessário
for dir_path in [src_dir, config_dir]:
    init_path = dir_path / '__init__.py'
    if not init_path.exists():
        with open(init_path, 'w') as f:
            f.write("# Arquivo gerado automaticamente\n")
        logger.info(f"Criado arquivo __init__.py em {dir_path}")

# Listar todos os caminhos no PYTHONPATH para diagnóstico
logger.info("PYTHONPATH atual:")
for p in sys.path:
    logger.info(f"  - {p}")

# Pré-carregar módulos de configuração para diagnóstico
try:
    # Verifique se os módulos de configuração podem ser importados
    import config
    logger.info("Módulo config importado com sucesso")
    
    # Tentar importar submódulos
    from config import utils, logger_config, window_config
    logger.info("Submódulos config importados com sucesso")
except ImportError as e:
    logger.error(f"Erro ao importar módulos config: {e}")
    try:
        # Tentar caminho alternativo
        from src.config import utils, logger_config, window_config
        logger.info("Módulos de config importados com caminho src.config")
    except ImportError as e2:
        logger.error(f"Erro ao importar módulos via src.config: {e2}")

# Importar o módulo principal
try:
    # Primeiro tentar importação relativa
    logger.info("Tentando importar Sistema_Entrada_Dados...")
    from src.Sistema_Entrada_Dados import SistemaEntradaDados
    logger.info("Módulo Sistema_Entrada_Dados importado com sucesso")
except Exception as e:
    logger.error(f"Erro ao importar Sistema_Entrada_Dados: {str(e)}")
    import traceback
    logger.error(traceback.format_exc())
    sys.exit(1)

def main():
    try:
        # Verificar ambiente
        import os
        import sys
        from pathlib import Path
        
        # Log do diretório de execução
        current_dir = os.getcwd()
        logger.info(f"Diretório de execução: {current_dir}")
        
        # No PyInstaller, os recursos são acessados via _MEIPASS
        if hasattr(sys, '_MEIPASS'):
            base_dir = Path(sys._MEIPASS)
            logger.info(f"Rodando a partir do PyInstaller em: {base_dir}")
            
            # Listar arquivos no diretório _MEIPASS
            logger.info("Arquivos disponíveis no _MEIPASS:")
            for root, dirs, files in os.walk(base_dir):
                for file in files:
                    logger.info(f"  - {os.path.join(root, file)}")
                    
            # Verificar a existência de arquivos importantes
            for path in [
                base_dir / 'src',
                base_dir / 'src/config',
                base_dir / 'src/config/parametros_sistema.json'
            ]:
                logger.info(f"Verificando {path}: {'Existe' if path.exists() else 'Não existe'}")
        else:
            logger.info("Rodando em modo normal (não compilado)")
        
        logger.info("Iniciando Sistema de Entrada de Dados")
        app = SistemaEntradaDados()
        
        # Use mainloop em vez de run
        app.root.mainloop()
        
    except Exception as e:
        logger.error(f"Erro ao iniciar o sistema: {str(e)}")
        import traceback
        logger.error(traceback.format_exc())
        
        # Em caso de erro, mantenha a janela do console aberta
        if hasattr(sys, '_MEIPASS'):
            input("\nPressione Enter para fechar...")
    finally:
        logger.info("Sistema finalizado")

if __name__ == "__main__":
    main()