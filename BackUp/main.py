"""
Script principal simplificado com foco em caminhos corretos para PyInstaller
"""
import os
import sys
import logging
from pathlib import Path

# Configurar logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[logging.StreamHandler()]
)
logger = logging.getLogger("sistema")

# Ajustar caminhos baseado no modo de execução
if getattr(sys, 'frozen', False):
    # PyInstaller
    base_dir = Path(sys._MEIPASS)
    logger.info(f"Executando a partir do PyInstaller em: {base_dir}")
    
    # Verificar arquivos empacotados
    logger.info("Arquivos no pacote:")
    for root, dirs, files in os.walk(base_dir):
        if 'src/config' in root or 'src\\config' in root:
            for file in files:
                logger.info(f"  - {os.path.join(root, file)}")
else:
    # Execução normal
    base_dir = Path(__file__).resolve().parent
    logger.info(f"Executando em modo normal a partir de: {base_dir}")

# Garantir que caminhos importantes estão no sys.path
src_dir = base_dir / 'src'
config_dir = src_dir / 'config'

for path in [str(base_dir), str(src_dir), str(config_dir)]:
    if path not in sys.path:
        sys.path.insert(0, path)
        logger.info(f"Adicionado ao path: {path}")

# Verificar módulos primeiro
try:
    logger.info("Verificando módulos de configuração...")
    
    # Tentativa 1: Importação absoluta com src
    try:
        import src.config
        logger.info("Módulo src.config importado com sucesso")
        
        # Tentar importar submódulos
        import src.config.utils
        import src.config.logger_config
        import src.config.window_config
        import src.config.config
        logger.info("Todos os submódulos de src.config importados com sucesso")
    except ImportError as e:
        logger.error(f"Erro ao importar via src.config: {e}")
        
        # Tentativa 2: Importação direta
        try:
            import config
            logger.info("Módulo config importado com sucesso")
            
            import config.utils
            import config.logger_config
            import config.window_config
            import config.config
            logger.info("Todos os submódulos de config importados com sucesso")
        except ImportError as e:
            logger.error(f"Erro ao importar via config direta: {e}")
            
            # Se chegou aqui, há um problema sério
            logger.error("FALHA CRÍTICA: Não foi possível importar os módulos de configuração")
            if getattr(sys, 'frozen', False):
                # Mostrar mensagem e pausar no modo compilado
                print("\nERRO CRÍTICO: Módulos de configuração não encontrados!")
                print("Este problema geralmente ocorre quando o PyInstaller não empacota corretamente os arquivos.")
                input("Pressione ENTER para sair...")
                sys.exit(1)
except Exception as e:
    logger.error(f"Erro ao verificar módulos: {str(e)}")
    if getattr(sys, 'frozen', False):
        input("Erro crítico. Pressione ENTER para sair...")
        sys.exit(1)

# Importar e iniciar o sistema
try:
    logger.info("Importando Sistema_Entrada_Dados...")
    from src.Sistema_Entrada_Dados import SistemaEntradaDados
    logger.info("Sistema_Entrada_Dados importado com sucesso")
    
    app = SistemaEntradaDados()
    logger.info("Iniciando mainloop...")
    app.root.mainloop()
except Exception as e:
    logger.error(f"Erro ao iniciar sistema: {str(e)}")
    import traceback
    logger.error(traceback.format_exc())
    
    if getattr(sys, 'frozen', False):
        print(f"\nERRO: {e}")
        print(traceback.format_exc())
        input("Pressione ENTER para sair...")