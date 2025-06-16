import logging
import os
import sys
from datetime import datetime
from functools import wraps

class SystemLogger:
    def __init__(self):
        self.user_context = {'user': 'sistema'}
        self.logger = None
        self.log_dir = None
        self.log_file = None
        self._setup_logging()

    def _determine_log_directory(self):
        """Determina o diretório correto para logs baseado no ambiente"""
        try:
            if getattr(sys, 'frozen', False):
                # Executável PyInstaller - usar diretório do executável
                base_dir = os.path.dirname(sys.executable)
            else:
                # Desenvolvimento - usar diretório do projeto
                current_dir = os.path.dirname(os.path.abspath(__file__))
                # Subir dois níveis: config -> src -> projeto
                base_dir = os.path.dirname(os.path.dirname(current_dir))
            
            self.log_dir = os.path.join(base_dir, 'logs')
            return True
            
        except Exception as e:
            print(f"Erro ao determinar diretório de logs: {str(e)}")
            # Fallback para diretório temporário
            import tempfile
            self.log_dir = os.path.join(tempfile.gettempdir(), 'sistema_gestao_logs')
            return False

    def _create_log_directory(self):
        """Cria o diretório de logs com tratamento robusto de erros"""
        try:
            os.makedirs(self.log_dir, exist_ok=True)
            
            # Testar se pode escrever no diretório
            test_file = os.path.join(self.log_dir, 'test_write.tmp')
            with open(test_file, 'w') as f:
                f.write('teste')
            os.remove(test_file)
            
            return True
            
        except Exception as e:
            print(f"Não foi possível criar/usar diretório de logs: {str(e)}")
            return False

    def _setup_logging(self):
        """Configura o sistema de logging com fallbacks robustos"""
        # Determinar diretório de logs
        dir_success = self._determine_log_directory()
        
        # Formato do log
        self.log_format = '%(asctime)s - %(user)s - %(module)s - %(levelname)s - %(message)s'
        
        # Configurar logger principal
        self.logger = logging.getLogger('sistema_gestao')
        self.logger.setLevel(logging.INFO)
        
        # Evitar handlers duplicados
        if self.logger.handlers:
            return
        
        # 1. Handler para console (sempre funciona)
        console_handler = logging.StreamHandler(sys.stdout)
        console_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)
        
        # 2. Handler para arquivo (com fallback)
        file_configured = False
        
        if dir_success and self._create_log_directory():
            try:
                # Nome do arquivo com data
                self.log_file = os.path.join(
                    self.log_dir, 
                    f'sistema_{datetime.now().strftime("%Y%m%d")}.log'
                )
                
                # Configurar handler de arquivo
                file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
                file_handler.setFormatter(logging.Formatter(self.log_format))
                self.logger.addHandler(file_handler)
                
                file_configured = True
                print(f"Log configurado com sucesso: {self.log_file}")
                
            except Exception as e:
                print(f"Erro ao configurar log em arquivo: {str(e)}")
        
        if not file_configured:
            print("Sistema funcionando apenas com log de console")
        
        # Log inicial
        adapter = logging.LoggerAdapter(self.logger, self.user_context)
        adapter.info("Sistema de logging inicializado")

    def set_user(self, username):
        """Define o usuário atual"""
        self.user_context['user'] = username
        if self.logger:
            adapter = logging.LoggerAdapter(self.logger, self.user_context)
            adapter.info(f"Usuário alterado para: {username}")

    def get_logger(self):
        """Retorna o logger configurado"""
        if self.logger is None:
            self._setup_logging()
        return logging.LoggerAdapter(self.logger, self.user_context)

    def get_log_info(self):
        """Retorna informações sobre o estado do logging"""
        return {
            'log_dir': self.log_dir,
            'log_file': self.log_file,
            'handlers_count': len(self.logger.handlers) if self.logger else 0,
            'current_user': self.user_context.get('user', 'não definido')
        }

# Criar instância global
system_logger = SystemLogger()

# Decorator para logging de ações (mantido igual para compatibilidade)
def log_action(action_description):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            logger = system_logger.get_logger()
            try:
                logger.info(f"Iniciando: {action_description}")
                result = func(*args, **kwargs)
                logger.info(f"Concluído: {action_description}")
                return result
            except Exception as e:
                logger.error(f"Erro em {action_description}: {str(e)}", exc_info=True)
                raise
        return wrapper
    return decorator

# Funções auxiliares para debugging
def get_logging_status():
    """Retorna o status atual do sistema de logging"""
    return system_logger.get_log_info()

def test_logging():
    """Testa se o sistema de logging está funcionando"""
    try:
        logger = system_logger.get_logger()
        logger.info("Teste de logging executado com sucesso")
        return True
    except Exception as e:
        print(f"Erro no teste de logging: {str(e)}")
        return False