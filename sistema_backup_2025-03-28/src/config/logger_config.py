import logging
import os
from datetime import datetime
from functools import wraps

class SystemLogger:
    def __init__(self):
        # Criar diretório para logs
        self.log_dir = 'logs'
        os.makedirs(self.log_dir, exist_ok=True)

        # Formato do log com usuário e módulo
        self.log_format = '%(asctime)s - %(user)s - %(module)s - %(levelname)s - %(message)s'
        
        # Nome do arquivo com data
        self.log_file = os.path.join(
            self.log_dir, 
            f'sistema_{datetime.now().strftime("%Y%m%d")}.log'
        )

        # Configurar handler de arquivo
        file_handler = logging.FileHandler(self.log_file, encoding='utf-8')
        file_handler.setFormatter(logging.Formatter(self.log_format))

        # Configurar logger
        self.logger = logging.getLogger('sistema_gestao')
        self.logger.setLevel(logging.INFO)
        self.logger.addHandler(file_handler)

        # Contexto do usuário
        self.user_context = {'user': 'sistema'}

    def set_user(self, username):
        """Define o usuário atual"""
        self.user_context['user'] = username

    def get_logger(self):
        """Retorna o logger configurado"""
        return logging.LoggerAdapter(self.logger, self.user_context)

# Criar instância global
system_logger = SystemLogger()

# Decorator para logging de ações
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