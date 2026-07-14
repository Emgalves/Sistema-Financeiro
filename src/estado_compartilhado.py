"""
Estado compartilhado em memória entre os módulos do sistema.
Vive apenas enquanto o processo está aberto (reinicia ao fechar o sistema) —
usado para lembrar o último cliente selecionado ao trocar de módulo.
"""

class EstadoSessao:
    def __init__(self):
        self.ultimo_cliente = None

# Instância única, compartilhada por todo o processo
estado_sessao = EstadoSessao()