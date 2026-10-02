"""
Motor de cálculo das taxas de administração: base de cálculo, percentual
contratado, recálculo quando a base muda, e verificação de consistência
entre o valor esperado e o valor efetivamente lançado na planilha.

Extraído de Sistema_Entrada_Dados.py em [DATA_DA_EXTRACAO].
Nenhuma alteração de lógica foi feita nesta extração — apenas mudança
de localização e ajuste de imports.

Dependência cruzada não resolvida por esta extração:
    recalcular_taxas_afetadas() chama
    self.sistema.obter_administradores_cliente_CORRIGIDO(cliente),
    método que permanece em SistemaEntradaDados. Isso continua
    funcionando normalmente porque a instância de GestorTaxasAdministracao
    recebe `sistema_principal` (self.sistema) no __init__, exatamente
    como antes da extração.
"""
import logging
import os
from datetime import datetime

import pandas as pd
from src.config.utils import CONVERSORES_DADOS  # NF lida como texto (ver normalizar_nf)
from dateutil.relativedelta import relativedelta
from openpyxl import load_workbook

from src.config.config import PASTA_CLIENTES

# Mesmo logger usado no restante do sistema (Sistema_Entrada_Dados.py).
# Não definimos nível aqui: ele herda o nível efetivo do logger raiz,
# configurado uma única vez no arquivo principal (logging.basicConfig).
logger = logging.getLogger("sistema")


class GestorTaxasAdministracao:
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal

    def identificar_lancamentos_taxa_admin(self, df):
        """
        Identifica lançamentos de taxa de administração com padrões mais amplos
        """
        if df.empty:
            return pd.DataFrame()

        mask_taxa = df['TP_DESP'] == 7

        taxas = df[mask_taxa].copy()
        logger.debug(f"DEBUG: Taxas encontradas (tp_desp=7): {len(taxas)} registros")

        if not taxas.empty:
            logger.debug(f"DEBUG: Referências das taxas: {taxas['REFERÊNCIA'].tolist()}")

        return taxas

    def calcular_base_calculo_taxa(self, data_referencia, df=None):
        """
        VERSÃO UNIFICADA - Calcula valor base seguindo a lógica corrigida do finalizacao_quinzena.py

        Parâmetros:
        - data_referencia: Data para cálculo (str ou datetime)
        - df: DataFrame opcional (para compatibilidade com código existente)
            Se não fornecido, lê diretamente da planilha
        """
        try:
            logger.debug(f"DEBUG: Calculando valor base para {data_referencia}")

            # Se DataFrame foi fornecido, usar lógica compatível
            if df is not None:
                return self._calcular_base_por_dataframe(df, data_referencia)

            # Caso contrário, usar lógica corrigida da planilha
            return self._calcular_base_por_planilha(data_referencia)

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao calcular valor base: {str(e)}")
            return 0

    def _calcular_base_por_planilha(self, data_referencia):
        """
        Método interno - Calcula base lendo diretamente da planilha
        (Lógica corrigida do finalizacao_quinzena.py)
        """
        try:
            cliente_atual = self.sistema.cliente_atual

            arquivo_cliente = PASTA_CLIENTES / f"{cliente_atual}.xlsx"
            wb = load_workbook(arquivo_cliente)
            ws_dados = wb["Dados"]

            # Converter data de referência se necessário
            if isinstance(data_referencia, str):
                data_ref = datetime.strptime(data_referencia, '%d/%m/%Y')
            else:
                data_ref = data_referencia

            logger.debug(f"DEBUG: Data de referência (planilha): {data_ref.strftime('%d/%m/%Y')}")

            valor_base = 0
            lancamentos_encontrados = 0

            # Usar a mesma lógica do finalizacao_quinzena.py
            for row in ws_dados.iter_rows(min_row=2, values_only=True):
                data_lancamento = row[0]  # Coluna A

                if isinstance(data_lancamento, datetime):
                    # Verificar se é da mesma data (dia, mês, ano)
                    if (data_lancamento.day == data_ref.day and
                        data_lancamento.month == data_ref.month and
                        data_lancamento.year == data_ref.year):

                        tipo_desp = row[1]  # Coluna B (TP_DESP)
                        status = row[13] if len(row) > 13 else "ATIVO"  # Coluna N (STATUS)

                        # Incluir apenas tipos 1 a 6 e status ATIVO
                        if (isinstance(tipo_desp, (int, float)) and 1 <= tipo_desp <= 6 and
                            status == "ATIVO"):
                            valor = row[8]  # Coluna I (VALOR)

                            if valor:
                                try:
                                    valor_numeric = float(str(valor).replace(',', '.'))
                                    valor_base += valor_numeric
                                    lancamentos_encontrados += 1

                                    logger.debug(f"DEBUG: Lançamento incluído - Tipo: {tipo_desp}, Valor: R$ {valor_numeric:.2f}")

                                except (ValueError, TypeError) as e:
                                    logger.debug(f"DEBUG: Erro ao processar valor '{valor}': {e}")
                                    continue

            logger.debug(f"DEBUG: Valor base total (planilha): R$ {valor_base:.2f}")
            logger.debug(f"DEBUG: Total de lançamentos incluídos: {lancamentos_encontrados}")

            wb.close()
            return valor_base

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao calcular valor base por planilha: {str(e)}")
            if 'wb' in locals():
                wb.close()
            return 0

    def _calcular_base_por_dataframe(self, df, data_referencia):
        """
        Método interno - Calcula base usando DataFrame fornecido
        (Para compatibilidade com verificações existentes)
        """
        try:
            # Converter data de referência
            if isinstance(data_referencia, str):
                data_ref = pd.to_datetime(data_referencia, format='%d/%m/%Y')
            else:
                data_ref = pd.to_datetime(data_referencia)

            logger.debug(f"DEBUG: Data de referência (DataFrame): {data_ref.strftime('%d/%m/%Y')}")

            # Garantir que DATA_REL existe e está em formato datetime
            if 'DATA_REL' not in df.columns:
                logger.debug("DEBUG: Coluna DATA_REL não encontrada no DataFrame")
                return 0

            df = df.copy()
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')

            # Filtrar dados para a data específica
            df_data = df[df['DATA_REL'].dt.date == data_ref.date()].copy()

            if df_data.empty:
                logger.debug(f"DEBUG: Nenhum lançamento encontrado para {data_ref.strftime('%d/%m/%Y')}")
                return 0

            # Garantir que STATUS existe
            if 'STATUS' not in df_data.columns:
                df_data['STATUS'] = 'ATIVO'

            # Filtrar apenas lançamentos ativos e tipos 1-6
            df_base = df_data[
                (df_data['STATUS'] == 'ATIVO') &
                (df_data['TP_DESP'].isin([1, 2, 3, 4, 5, 6]))
            ].copy()

            if df_base.empty:
                logger.debug("DEBUG: Nenhum lançamento ativo dos tipos 1-6 encontrado")
                return 0

            # Converter valores para numérico
            df_base['VALOR_NUM'] = pd.to_numeric(
                df_base['VALOR'].astype(str).str.replace('R$', '').str.replace(',', '.'),
                errors='coerce'
            ).fillna(0)

            valor_base = df_base['VALOR_NUM'].sum()

            logger.debug(f"DEBUG: Valor base total (DataFrame): R$ {valor_base:.2f}")
            logger.debug(f"DEBUG: Lançamentos incluídos: {len(df_base)}")

            return valor_base

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao calcular valor base por DataFrame: {str(e)}")
            return 0

    def obter_percentual_taxa_cliente(self, cliente):
        """
        VERSÃO CORRIGIDA - Busca percentual de taxa seguindo a mesma lógica do finalizacao_quinzena.py
        """
        try:
            logger.debug(f"DEBUG: Buscando percentual da taxa para cliente {cliente}")

            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"
            wb = load_workbook(arquivo_cliente)

            if 'Contratos_ADM' not in wb.sheetnames:
                logger.debug("DEBUG: Aba 'Contratos_ADM' não encontrada")
                wb.close()
                return 0

            ws_contratos = wb['Contratos_ADM']
            logger.debug(f"DEBUG: Aba 'Contratos_ADM' carregada")

            # CORREÇÃO 1: Usar a mesma lógica do finalizacao_quinzena.py
            # 1º Passo: Encontrar contratos ativos
            contratos_ativos = set()
            for row in ws_contratos.iter_rows(min_row=3, values_only=True):  # Começar da linha 3
                if row[0] and row[3] == 'ATIVO':  # Coluna A (Nº Contrato) e Coluna D (Status)
                    contratos_ativos.add(str(row[0]))
                    logger.debug(f"DEBUG: Contrato ativo encontrado: {row[0]}")

            logger.debug(f"DEBUG: Contratos ativos: {contratos_ativos}")

            if not contratos_ativos:
                logger.debug("DEBUG: Nenhum contrato ativo encontrado")
                wb.close()
                return 0

            # CORREÇÃO 2: Para cada contrato ativo, buscar administradores com taxa percentual
            taxa_total = 0
            administradores_encontrados = []

            for num_contrato in contratos_ativos:
                logger.debug(f"DEBUG: Verificando administradores do contrato {num_contrato}")

                for row in ws_contratos.iter_rows(min_row=3, values_only=True):
                    # CORREÇÃO 3: Verificar se pertence ao contrato (coluna G) e é do tipo Percentual (coluna J)
                    if (str(row[6]) == num_contrato and      # Coluna G (Nº Contrato)
                        row[9] == 'Percentual'):            # Coluna J (Tipo)

                        # CORREÇÃO 4: Extrair percentual da coluna K
                        percentual_raw = row[10]  # Coluna K (Valor/Percentual)

                        logger.debug(f"DEBUG: Administrador encontrado:")
                        logger.debug(f"  - CNPJ/CPF: {row[7]}")     # Coluna H
                        logger.debug(f"  - Nome: {row[8]}")         # Coluna I
                        logger.debug(f"  - Tipo: {row[9]}")         # Coluna J
                        logger.debug(f"  - Percentual bruto: '{percentual_raw}'")  # Coluna K

                        try:
                            # CORREÇÃO 5: Processar o percentual corretamente
                            if percentual_raw:
                                # Converter para string e limpar
                                percentual_str = str(percentual_raw).strip()

                                # Remover % se existir e converter vírgula para ponto
                                percentual_limpo = percentual_str.replace('%', '').replace(',', '.')

                                percentual_float = float(percentual_limpo)
                                taxa_total += percentual_float

                                administradores_encontrados.append({
                                    'cnpj_cpf': row[7],
                                    'nome': row[8],
                                    'percentual': percentual_float
                                })

                                logger.debug(f"DEBUG: Percentual processado: {percentual_float}%")

                        except (ValueError, TypeError) as e:
                            logger.debug(f"DEBUG: Erro ao processar percentual '{percentual_raw}': {e}")
                            continue

            logger.debug(f"DEBUG: Taxa total encontrada: {taxa_total}%")
            logger.debug(f"DEBUG: Administradores encontrados: {len(administradores_encontrados)}")

            wb.close()
            return taxa_total

        except Exception as e:
            logger.debug(f"DEBUG: Erro ao obter percentual: {str(e)}")
            if 'wb' in locals():
                wb.close()
            return 0

    def verificar_necessidade_recalculo(self, data_referencia, cliente=None):
        """
        VERSÃO CORRIGIDA - Verifica se há necessidade de recálculo usando os métodos unificados
        """
        try:
            if not cliente:
                cliente = self.sistema.cliente_atual

            arquivo_cliente = PASTA_CLIENTES / f"{cliente}.xlsx"

            if not os.path.exists(arquivo_cliente):
                return False, "Arquivo do cliente não encontrado"

            logger.debug(f"DEBUG: Verificando necessidade de recálculo para {cliente} em {data_referencia}")

            # Ler dados da planilha
            df = pd.read_excel(arquivo_cliente, sheet_name='Dados', converters=CONVERSORES_DADOS)
            df = df.fillna("")

            # Converter data
            if isinstance(data_referencia, str):
                data_ref_dt = pd.to_datetime(data_referencia, format='%d/%m/%Y')
            else:
                data_ref_dt = pd.to_datetime(data_referencia)

            # Filtrar para a data específica
            df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')
            df_data = df[df['DATA_REL'].dt.date == data_ref_dt.date()].copy()

            # Verificar se há taxas existentes
            taxas_existentes = self.identificar_lancamentos_taxa_admin(df_data)

            if taxas_existentes.empty:
                return False, "Nenhuma taxa encontrada para esta data"

            logger.debug(f"DEBUG: {len(taxas_existentes)} taxa(s) encontrada(s)")

            # CORREÇÃO: Usar o método unificado para calcular base
            # Primeiro tentar com DataFrame (mais rápido)
            base_atual = self.calcular_base_calculo_taxa(data_referencia, df=df)

            logger.debug(f"DEBUG: Base atual calculada: R$ {base_atual:.2f}")

            # Obter percentual da taxa
            percentual = self.obter_percentual_taxa_cliente(cliente)

            if percentual == 0:
                return False, "Percentual de taxa não configurado"

            logger.debug(f"DEBUG: Percentual de taxa: {percentual}%")

            # Calcular valor esperado da taxa
            valor_esperado = base_atual * (percentual / 100)
            logger.debug(f"DEBUG: Valor esperado da taxa: R$ {valor_esperado:.2f}")

            # Somar valor atual das taxas ATIVAS
            valor_atual_taxas = 0
            taxas_ativas = 0

            for _, taxa in taxas_existentes.iterrows():
                status = taxa.get('STATUS', 'ATIVO')
                if status != 'EXCLUIDO':
                    try:
                        valor_taxa = float(str(taxa.get('VALOR', 0)).replace(',', '.'))
                        valor_atual_taxas += valor_taxa
                        taxas_ativas += 1
                        logger.debug(f"DEBUG: Taxa ativa ID {taxa.get('ID_LANCAMENTO', 'N/A')}: R$ {valor_taxa:.2f}")
                    except (ValueError, TypeError):
                        logger.debug(f"DEBUG: Erro ao processar valor da taxa: {taxa.get('VALOR', 'N/A')}")
                        pass

            logger.debug(f"DEBUG: Valor atual total das taxas ativas: R$ {valor_atual_taxas:.2f}")
            logger.debug(f"DEBUG: Taxas ativas encontradas: {taxas_ativas}")

            # Calcular diferença
            diferenca = abs(valor_esperado - valor_atual_taxas)
            tolerancia = 0.01  # R$ 0,01

            logger.debug(f"DEBUG: Diferença: R$ {diferenca:.2f} (tolerância: R$ {tolerancia:.2f})")

            if diferenca > tolerancia:
                mensagem = f"Recálculo necessário - Base: R$ {base_atual:.2f} ({percentual}%) = R$ {valor_esperado:.2f}, Atual: R$ {valor_atual_taxas:.2f}, Diferença: R$ {diferenca:.2f}"
                return True, mensagem

            mensagem = f"Taxas consistentes - Base: R$ {base_atual:.2f} ({percentual}%) = R$ {valor_esperado:.2f}"
            return False, mensagem

        except Exception as e:
            import traceback
            logger.debug(f"DEBUG: Erro na verificação: {traceback.format_exc()}")
            return False, f"Erro na verificação: {str(e)}"
