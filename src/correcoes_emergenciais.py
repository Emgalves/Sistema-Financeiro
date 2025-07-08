# ARQUIVO: correcoes_emergenciais.py
# Este arquivo contém todas as correções necessárias para resolver os erros

import pandas as pd
import logging
from datetime import datetime
from openpyxl import load_workbook

logger = logging.getLogger(__name__)

def aplicar_todas_correcoes():
    """
    Função principal que aplica todas as correções necessárias
    """
    print("=== INICIANDO APLICAÇÃO DE CORREÇÕES ===")
    
    try:
        # Importar a classe RelatorioHandler
        from relatorio_despesas_aprimorado import RelatorioHandler
        
        # Aplicar correções na classe RelatorioHandler
        aplicar_correcoes_relatorio_handler(RelatorioHandler)
        
        print("✅ Todas as correções foram aplicadas com sucesso!")
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar correções: {str(e)}")
        return False

def aplicar_correcoes_relatorio_handler(RelatorioHandler):
    """
    Aplica todas as correções necessárias na classe RelatorioHandler
    """
    print("Aplicando correções na classe RelatorioHandler...")
    
    # 1. Corrigir método processar_dados
    RelatorioHandler.processar_dados_original = RelatorioHandler.processar_dados
    RelatorioHandler.processar_dados = processar_dados_corrigido
    
    # 2. Corrigir método criar_resumo_despesas  
    RelatorioHandler.criar_resumo_despesas_original = RelatorioHandler.criar_resumo_despesas
    RelatorioHandler.criar_resumo_despesas = criar_resumo_despesas_corrigido
    
    # 3. Adicionar/corrigir método processar_lancamentos_futuros
    RelatorioHandler.processar_lancamentos_futuros = processar_lancamentos_futuros_corrigido
    
    # 4. Adicionar método de validação
    RelatorioHandler.validar_integridade_dados = validar_integridade_dados
    
    # 5. Corrigir método carregar_dados_excel
    RelatorioHandler.carregar_dados_excel_original = RelatorioHandler.carregar_dados_excel
    RelatorioHandler.carregar_dados_excel = carregar_dados_excel_corrigido
    
    print("✅ Correções aplicadas na classe RelatorioHandler")

def processar_dados_corrigido(self, df, data_relatorio, incluir_excluidos=False):
    """Versão corrigida do método processar_dados"""
    try:
        # Converter data para datetime
        try:
            data_rel = pd.to_datetime(data_relatorio)
        except:
            data_rel = pd.to_datetime(data_relatorio, format='%d/%m/%Y')
        
        # Criar cópia do DataFrame
        df = df.copy()
        
        # Log para debug
        logger.debug(f"Colunas do DataFrame original: {df.columns.tolist()}")
        
        # Filtrar excluídos se necessário
        if not incluir_excluidos and 'STATUS' in df.columns:
            df = df[df['STATUS'] != 'EXCLUIDO'].copy()
            print(f"Processando dados - registros após filtrar excluídos: {len(df)}")
        else:
            print(f"Processando dados - incluindo todos os registros: {len(df)}")
        
        # Verificação crítica da coluna TP_DESP
        if 'TP_DESP' not in df.columns:
            logger.error("ERRO CRÍTICO: Coluna TP_DESP não encontrada!")
            raise ValueError("Coluna TP_DESP não encontrada no DataFrame")
        
        # Adicionar coluna de índice original
        df = df.reset_index(drop=True)
        df['ordem_original'] = df.index
        
        # Processar DT_VENCTO
        if 'DT_VENCTO' in df.columns:
            try:
                df['DT_VENCTO_SORT'] = pd.to_datetime(df['DT_VENCTO'], 
                                                    format='mixed', 
                                                    errors='coerce', 
                                                    dayfirst=True)
                df['DT_VENCTO_DISPLAY'] = df['DT_VENCTO_SORT'].dt.strftime('%d/%m/%Y')
            except Exception as e:
                print(f"Erro ao converter DT_VENCTO: {str(e)}")
                df['DT_VENCTO_SORT'] = pd.to_datetime('2000-01-01')
                df['DT_VENCTO_DISPLAY'] = df['DT_VENCTO']
        
        # Aplicar restrição de dados bancários
        if 'DADOS_BANCARIOS' in df.columns:
            df['DADOS_BANCARIOS_ORIGINAL'] = df['DADOS_BANCARIOS']
            df.loc[df['TP_DESP'].isin([3, 5]), 'DADOS_BANCARIOS'] = ''
        
        # Filtrar dados principais
        df_filtrado = df[
            (df['DATA_REL'] == data_rel) & 
            (df['TP_DESP'] != 1)
        ].copy()
        
        # Ordenação especial para TP_DESP == 5
        df_tp5 = df_filtrado[df_filtrado['TP_DESP'] == 5].copy()
        df_outros = df_filtrado[df_filtrado['TP_DESP'] != 5].copy()
        
        if not df_outros.empty:
            df_outros = df_outros.sort_values(
                by=['TP_DESP', 'DT_VENCTO_SORT', 'VALOR'], 
                ascending=[True, True, False]
            )
        
        if not df_tp5.empty:
            df_tp5 = df_tp5.sort_values('ordem_original')
        
        # Combinar DataFrames
        if not df_outros.empty and not df_tp5.empty:
            df_filtrado = pd.concat([df_outros, df_tp5], ignore_index=True)
        elif not df_outros.empty:
            df_filtrado = df_outros
        elif not df_tp5.empty:
            df_filtrado = df_tp5
        else:
            df_filtrado = pd.DataFrame()
        
        # Processar outros DataFrames
        df_diaria = df[
            (df['DATA_REL'] == data_rel) & 
            (df['TP_DESP'] == 1) & 
            (df['REFERÊNCIA'] == 'DIÁRIA')
        ].copy()
        
        df_tp_desp_1 = df[
            (df['DATA_REL'] == data_rel) & 
            (df['TP_DESP'] == 1) & 
            (df['REFERÊNCIA'].isin(['SALÁRIO', 'TRANSPORTE', 'CAFÉ']))
        ].copy()

        df_tp_desp_2 = df[
            (df['DATA_REL'] == data_rel) & 
            (df['TP_DESP'] == 1) & 
            (df['REFERÊNCIA'].isin(['FÉRIAS', 'RESCISÃO', '13º SALÁRIO']))
        ].copy()
        
        # Substituir DT_VENCTO pela versão formatada
        if 'DT_VENCTO_DISPLAY' in df_filtrado.columns:
            df_filtrado['DT_VENCTO'] = df_filtrado['DT_VENCTO_DISPLAY']
        
        # CORREÇÃO CRÍTICA: Preservar colunas essenciais
        colunas_essenciais = [
            'TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR', 'DATA_REL', 'DT_VENCTO',
            'DADOS_BANCARIOS', 'DIAS', 'VR_UNIT', 'NF', 'STATUS'
        ]
        
        # Remover apenas colunas temporárias
        colunas_temporarias = ['DT_VENCTO_SORT', 'DT_VENCTO_DISPLAY', 'ordem_original', 'DADOS_BANCARIOS_ORIGINAL']
        
        for df_temp in [df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2]:
            if df_temp.empty:
                continue
                
            # Log das colunas antes da limpeza
            logger.debug(f"DataFrame com {len(df_temp)} registros - Colunas antes: {df_temp.columns.tolist()}")
            
            # Remover apenas colunas temporárias que existem
            colunas_para_remover = [col for col in colunas_temporarias if col in df_temp.columns]
            
            if colunas_para_remover:
                df_temp.drop(columns=colunas_para_remover, inplace=True)
                logger.debug(f"Colunas removidas: {colunas_para_remover}")
            
            # Verificação crítica
            if 'TP_DESP' not in df_temp.columns and not df_temp.empty:
                logger.error(f"ERRO: TP_DESP removida inadvertidamente de DataFrame com {len(df_temp)} registros!")
                raise ValueError("TP_DESP foi removida inadvertidamente!")
        
        # Log final
        logger.info(f"df_filtrado final: {len(df_filtrado)} registros")
        if not df_filtrado.empty and 'TP_DESP' in df_filtrado.columns:
            logger.info(f"Tipos de despesa únicos: {df_filtrado['TP_DESP'].unique()}")
        
        return df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2
        
    except Exception as e:
        logger.error(f"Erro em processar_dados_corrigido: {str(e)}", exc_info=True)
        raise

def criar_resumo_despesas_corrigido(self, dados):
    """Versão corrigida do método criar_resumo_despesas"""
    try:
        logger.debug("Iniciando criar_resumo_despesas_corrigido")
        
        # Obter DataFrames
        df_filtrado = dados.get('df_filtrado', pd.DataFrame())
        df_tp_desp_1 = dados.get('df_tp_desp_1', pd.DataFrame())
        df_tp_desp_2 = dados.get('df_tp_desp_2', pd.DataFrame())
        df_diaria = dados.get('df_diaria', pd.DataFrame())
        
        # Verificação crítica do df_filtrado
        if not df_filtrado.empty and 'TP_DESP' not in df_filtrado.columns:
            logger.error("df_filtrado não contém TP_DESP! Tentando recuperar...")
            
            # Tentar recriar dos dados originais
            df_original = dados.get('df_original', pd.DataFrame())
            if not df_original.empty and 'TP_DESP' in df_original.columns:
                data_relatorio = dados.get('data_relatorio')
                if data_relatorio:
                    data_rel = pd.to_datetime(data_relatorio)
                    df_filtrado = df_original[
                        (df_original['DATA_REL'] == data_rel) & 
                        (df_original['TP_DESP'] != 1)
                    ].copy()
                    logger.info(f"df_filtrado recriado com {len(df_filtrado)} registros")
                else:
                    df_filtrado = pd.DataFrame(columns=['TP_DESP', 'VALOR'])
            else:
                df_filtrado = pd.DataFrame(columns=['TP_DESP', 'VALOR'])
        
        subtotais = {}
        
        # Calcular subtotais por tipo
        for tipo, descricao in self.tipos_despesas.items():
            valor = 0
            
            try:
                if tipo == 1:
                    # Somar despesas de colaboradores
                    valor1 = valor2 = valor3 = 0
                    
                    if not df_tp_desp_1.empty and 'VALOR' in df_tp_desp_1.columns:
                        try:
                            valores_numericos = pd.to_numeric(df_tp_desp_1['VALOR'], errors='coerce').fillna(0)
                            valor1 = valores_numericos.sum()
                        except:
                            valor1 = 0
                    
                    if not df_tp_desp_2.empty and 'VALOR' in df_tp_desp_2.columns:
                        try:
                            valores_numericos = pd.to_numeric(df_tp_desp_2['VALOR'], errors='coerce').fillna(0)
                            valor2 = valores_numericos.sum()
                        except:
                            valor2 = 0
                    
                    if not df_diaria.empty and 'VALOR' in df_diaria.columns:
                        try:
                            valores_numericos = pd.to_numeric(df_diaria['VALOR'], errors='coerce').fillna(0)
                            valor3 = valores_numericos.sum()
                        except:
                            valor3 = 0
                    
                    valor = valor1 + valor2 + valor3
                    
                else:
                    # Outras despesas
                    if not df_filtrado.empty and 'TP_DESP' in df_filtrado.columns and 'VALOR' in df_filtrado.columns:
                        try:
                            df_tipo = df_filtrado[df_filtrado['TP_DESP'] == tipo]
                            if not df_tipo.empty:
                                valores_numericos = pd.to_numeric(df_tipo['VALOR'], errors='coerce').fillna(0)
                                valor = valores_numericos.sum()
                        except Exception as e:
                            logger.error(f"Erro ao processar tipo {tipo}: {str(e)}")
                            valor = 0
                    else:
                        valor = 0
                        
            except Exception as e:
                logger.error(f"Erro geral para tipo {tipo}: {str(e)}")
                valor = 0
                
            subtotais[tipo] = valor
        
        # Calcular totais agrupados
        despesas_a_pagar = sum(subtotais.get(tp, 0) for tp in [1, 2, 3, 4, 7])
        despesas_pagas_cliente = sum(subtotais.get(tp, 0) for tp in [5])
        despesas_pagas_caixa = sum(subtotais.get(tp, 0) for tp in [6])
        total_quinzena = sum(subtotais.values())
        
        acumulado = dados.get('acumulado', 0)
        numero_relatorio = dados.get('numero_relatorio', 1)
        total_obra = total_quinzena + acumulado
        
        # Criar tabelas
        tabela_subtotais = []
        for tipo, descricao in self.tipos_despesas.items():
            if tipo in subtotais:
                valor_formatado = self.formatar_numero(subtotais[tipo])
                tabela_subtotais.append([descricao, valor_formatado])
        
        tabela_totais = [
            ['DESPESAS A PAGAR', self.formatar_numero(despesas_a_pagar)],
            ['DESPESAS PAGAS PELO CLIENTE', self.formatar_numero(despesas_pagas_cliente)],
            ['COMPLEMENTO DE CAIXA', self.formatar_numero(despesas_pagas_caixa)],
            [''],
            ['TOTAL DA QUINZENA', self.formatar_numero(total_quinzena)],
            [f'TOTAL ACUMULADO RELATÓRIO Nº {numero_relatorio - 1}', self.formatar_numero(acumulado)],
            ['TOTAL DA OBRA', self.formatar_numero(total_obra)]
        ]
        
        return tabela_subtotais, tabela_totais
        
    except Exception as e:
        logger.error(f"Erro em criar_resumo_despesas_corrigido: {str(e)}", exc_info=True)
        raise

def processar_lancamentos_futuros_corrigido(self, df, data_relatorio, incluir_excluidos=False):
    """Versão corrigida do método processar_lancamentos_futuros"""
    try:
        # Converter data do relatório
        try:
            self.data_ref = pd.to_datetime(data_relatorio)
        except:
            self.data_ref = pd.to_datetime(data_relatorio, format='%d/%m/%Y')
        
        df = df.copy()
        
        # Filtrar excluídos se necessário
        if not incluir_excluidos and 'STATUS' in df.columns:
            df = df[df['STATUS'] != 'EXCLUIDO'].copy()
            print(f"Lançamentos futuros - registros após filtrar excluídos: {len(df)}")
        else:
            print(f"Lançamentos futuros - incluindo todos os registros: {len(df)}")
        
        # Verificar colunas necessárias
        if 'DATA_REL' not in df.columns:
            logger.error("Coluna DATA_REL não encontrada")
            return pd.DataFrame()
        
        if 'DT_VENCTO' not in df.columns:
            logger.warning("Coluna DT_VENCTO não encontrada, usando DATA_REL")
            df['DT_VENCTO'] = df['DATA_REL']
        
        # Converter para datetime
        df['DATA_REL'] = pd.to_datetime(df['DATA_REL'], errors='coerce')
        df['DT_VENCTO'] = pd.to_datetime(df['DT_VENCTO'], format='%d/%m/%Y', errors='coerce')
        
        # Remover registros com datas inválidas
        df = df.dropna(subset=['DATA_REL'])
        
        # Formatar DT_VENCTO
        df['DT_VENCTO'] = df['DT_VENCTO'].dt.strftime('%d/%m/%Y').fillna('')
        
        # Filtrar lançamentos futuros
        df_futuro = df[(df['DATA_REL'] > self.data_ref) & (df['TP_DESP'] != 1)].copy()
        
        if df_futuro.empty:
            logger.info("Nenhum lançamento futuro encontrado")
            return df_futuro
        
        # Ordenar por data
        df_futuro = df_futuro.sort_values('DATA_REL')
        
        # Classificar por período
        def classificar_periodo(data_rel):
            try:
                diff_days = (data_rel - self.data_ref).days
                if diff_days <= 30:
                    return "Próximos 30 dias"
                elif diff_days <= 60:
                    return "31 a 60 dias"
                else:
                    return "Após 60 dias"
            except:
                return "Após 60 dias"
        
        df_futuro['periodo'] = df_futuro['DATA_REL'].apply(classificar_periodo)
        
        logger.info(f"Processados {len(df_futuro)} lançamentos futuros")
        return df_futuro
        
    except Exception as e:
        logger.error(f"Erro ao processar lançamentos futuros: {str(e)}", exc_info=True)
        return pd.DataFrame()

def carregar_dados_excel_corrigido(self, arquivo_excel, incluir_excluidos=False):
    """Versão corrigida do método carregar_dados_excel"""
    try:
        logger.info(f"Carregando dados de: {arquivo_excel}")
        
        df = pd.read_excel(arquivo_excel, sheet_name='Dados')
        df = df.fillna("")
        
        logger.info(f"Dados carregados: {len(df)} registros, {len(df.columns)} colunas")
        
        # Verificar colunas necessárias
        colunas_necessarias = {'DATA_REL', 'TP_DESP', 'REFERÊNCIA', 'DT_VENCTO', 'VALOR', 'NF'}
        colunas_faltantes = colunas_necessarias - set(df.columns)
        
        if colunas_faltantes:
            raise ValueError(f"Colunas necessárias ausentes: {colunas_faltantes}")
        
        # Adicionar STATUS se não existir
        if 'STATUS' not in df.columns:
            df['STATUS'] = 'ATIVO'
        
        # Validar integridade
        if not self.validar_integridade_dados(df, "Dados carregados"):
            raise ValueError("Falha na validação de integridade")
        
        # Filtrar excluídos
        if not incluir_excluidos:
            df_original_size = len(df)
            df = df[df['STATUS'] != 'EXCLUIDO'].copy()
            registros_excluidos = df_original_size - len(df)
            if registros_excluidos > 0:
                logger.info(f"Filtrados {registros_excluidos} registros excluídos")
            print(f"Registros após filtrar excluídos: {len(df)}")
        
        # Converter NF para string
        df['NF'] = df['NF'].astype(str)
        
        # Concatenar NF com REFERÊNCIA
        mascara = (df['TP_DESP'] != 1) & (df['NF'].notna()) & (df['NF'].str.strip() != '') & (df['NF'] != 'nan')
        df.loc[mascara, 'REFERÊNCIA'] = df[mascara].apply(
            lambda row: f"{row['REFERÊNCIA']} (NF: {row['NF'].strip()})", 
            axis=1
        )
        
        logger.info(f"Dados processados com sucesso: {len(df)} registros")
        return df
        
    except Exception as e:
        logger.error(f"Erro ao carregar arquivo: {str(e)}", exc_info=True)
        raise

def validar_integridade_dados(self, df, local="DataFrame"):
    """Valida integridade dos dados"""
    try:
        logger.debug(f"Validando integridade em: {local}")
        
        if df.empty:
            logger.warning(f"{local}: DataFrame vazio")
            return True
        
        # Verificar colunas essenciais
        colunas_essenciais = ['TP_DESP', 'NOME', 'REFERÊNCIA', 'VALOR', 'DATA_REL']
        colunas_faltantes = [col for col in colunas_essenciais if col not in df.columns]
        
        if colunas_faltantes:
            logger.error(f"{local}: Colunas ausentes: {colunas_faltantes}")
            return False
        
        # Verificar valores nulos
        for col in ['TP_DESP', 'DATA_REL']:
            nulos = df[col].isnull().sum()
            if nulos > 0:
                logger.warning(f"{local}: {nulos} valores nulos em {col}")
        
        logger.debug(f"{local}: Validação OK - {len(df)} registros")
        return True
        
    except Exception as e:
        logger.error(f"Erro na validação de {local}: {str(e)}")
        return False

# FUNÇÃO PRINCIPAL PARA APLICAR TODAS AS CORREÇÕES
def main():
    """Função principal para aplicar correções"""
    print("=== SISTEMA DE CORREÇÕES EMERGENCIAIS ===")
    
    if aplicar_todas_correcoes():
        print("✅ Sistema corrigido com sucesso!")
        print("\nPróximos passos:")
        print("1. Execute novamente o relatório de despesas")
        print("2. Verifique se os erros foram resolvidos")
        print("3. Teste com diferentes arquivos")
    else:
        print("❌ Falha ao aplicar correções")
        print("Verifique os logs para mais detalhes")

if __name__ == "__main__":
    main()