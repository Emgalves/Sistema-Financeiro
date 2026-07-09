"""
Cache em memória de fornecedores, usado para acelerar buscas repetidas
sem reabrir a planilha de fornecedores a cada consulta.

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].
Nenhuma alteração de lógica foi feita nesta extração — apenas mudança
de localização e ajuste de imports.
"""
import logging
import os
from datetime import datetime

from openpyxl import load_workbook

# Mesmo logger usado no restante do sistema (Sistema_Entrada_Dados.py).
# Não definimos nível aqui: ele herda o nível efetivo do logger raiz,
# configurado uma única vez no arquivo principal (logging.basicConfig).
logger = logging.getLogger("sistema")


class CacheFornecedores:
    """Cache para otimizar buscas de fornecedores"""

    def __init__(self):
        self.cache_fornecedores = None
        self.cache_timestamp = None
        self.cache_duracao = 300  # 5 minutos

    def carregar_cache_se_necessario(self, arquivo_fornecedores):
        """Carrega cache se necessário ou se arquivo foi modificado"""
        try:
            agora = datetime.now()
            arquivo_modificado = os.path.getmtime(arquivo_fornecedores)

            # Verificar se precisa recarregar
            precisa_recarregar = (
                self.cache_fornecedores is None or
                self.cache_timestamp is None or
                (agora - self.cache_timestamp).seconds > self.cache_duracao or
                arquivo_modificado > self.cache_timestamp.timestamp()
            )

            if precisa_recarregar:
                logger.debug("DEBUG: Recarregando cache de fornecedores...")
                self.cache_fornecedores = self._carregar_fornecedores(arquivo_fornecedores)
                self.cache_timestamp = agora
                logger.debug(f"DEBUG: Cache carregado com {len(self.cache_fornecedores)} fornecedores")

            return self.cache_fornecedores

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao carregar cache: {str(e)}")
            return []

    def _carregar_fornecedores(self, arquivo_fornecedores):
        """Carrega todos os fornecedores em memória"""
        fornecedores = []

        try:
            wb = load_workbook(arquivo_fornecedores, data_only=True)
            ws = wb['Fornecedores']

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0] or not row[3]:  # Pular se não tem CNPJ ou nome
                    continue

                fornecedor = {
                    'cnpj_cpf': str(row[0]).strip(),
                    'nome': str(row[3]).strip().upper(),
                    'categoria': str(row[11] or '').strip(),
                    'banco': str(row[4] or '').strip(),
                    'op': str(row[5] or '').strip(),
                    'agencia': str(row[6] or '').strip(),
                    'conta': str(row[7] or '').strip(),
                    'chave_pix': str(row[8] or '').strip()
                }

                fornecedores.append(fornecedor)

            wb.close()

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao carregar fornecedores: {str(e)}")

        return fornecedores
