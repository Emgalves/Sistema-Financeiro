# relatorio_despesas_service.py
# =============================
import tempfile
import os
import logging
from datetime import datetime
from openpyxl import load_workbook

logger = logging.getLogger(__name__)


class RelatoriosDespesasService:
    """
    Serviço que encapsula toda lógica de negócio de despesas
    USA o RelatorioHandler existente como motor
    """
    
    def __init__(self):
        # Importa e usa o handler existente
        from relatorio_despesas_aprimorado import RelatorioHandler
        self.handler = RelatorioHandler()
    
        # 🔍 DEBUG: Verificar métodos disponíveis
        print("=" * 80)
        print("🔍 DEBUG: VERIFICANDO MÉTODOS DO HANDLER")
        print("=" * 80)
        
        # Listar todos os métodos que contém "futuro"
        metodos_futuro = [m for m in dir(self.handler) if 'futuro' in m.lower() and not m.startswith('_')]
        print(f"Métodos com 'futuro': {metodos_futuro}")
        
        # Verificar especificamente os métodos que precisamos
        tem_processar = hasattr(self.handler, 'processar_lancamentos_futuros')
        tem_adicionar = hasattr(self.handler, 'adicionar_lancamentos_futuros')
        
        print(f"✓ hasattr processar_lancamentos_futuros: {tem_processar}")
        print(f"✓ hasattr adicionar_lancamentos_futuros: {tem_adicionar}")
        
        if tem_processar:
            metodo = getattr(self.handler, 'processar_lancamentos_futuros')
            print(f"✓ Tipo do método processar: {type(metodo)}")
            print(f"✓ Assinatura: {metodo.__code__.co_varnames[:metodo.__code__.co_argcount]}")
        
        print("=" * 80)


    def processar_para_preview(self, config):
        """
        Versão SIMPLIFICADA - confia completamente no handler original
        NÃO refaz ordenações, apenas usa o que já está correto
        """
        try:
            print("🔧 PROCESSANDO PARA PREVIEW - USANDO HANDLER ORIGINAL")
            
            # 1. Carregar dados usando método original
            df_original = self.handler.carregar_dados_excel(
                config['arquivo'], 
                config['incluir_excluidos']
            )
            print(f"✅ Dados carregados: {len(df_original)} registros")
            
            # 2. Processar dados usando método original
            df_filtrado, df_diaria, df_tp_desp_1, df_tp_desp_2 = self.handler.processar_dados(
                df_original, config['data'], config['incluir_excluidos']
            )
            print("✅ Dados processados pelo handler original")
            
            # 3. Lançamentos futuros
            import pandas as pd
            df_futuro = pd.DataFrame()  # ⚠️ SEMPRE inicializar como DataFrame, NUNCA None

            if config['incluir_futuros']:
                try:
                    if hasattr(self.handler, 'processar_lancamentos_futuros'):
                        resultado = self.handler.processar_lancamentos_futuros(
                            df_original, config['data'], config['incluir_excluidos']
                        )
                        # Garantir que é DataFrame, não None
                        if resultado is not None and isinstance(resultado, pd.DataFrame):
                            df_futuro = resultado
                        else:
                            print(f"⚠️ AVISO: processar_lancamentos_futuros retornou {type(resultado)}")
                            df_futuro = pd.DataFrame()
                    else:
                        print("⚠️ Método processar_lancamentos_futuros não encontrado!")
                except Exception as e:
                    print(f"❌ ERRO ao processar lançamentos futuros: {str(e)}")
                    import traceback
                    traceback.print_exc()
                    df_futuro = pd.DataFrame()
            
            # 4. Informações do cliente
            workbook = load_workbook(config['arquivo'], data_only=True)
            ws_resumo = workbook['RESUMO']
            
            numero_relatorio = self.handler.obter_numero_relatorio(ws_resumo, config['data'])
            valor_acumulado = self.handler.calcular_acumulado_dados(
                df_original, config['data'], config['incluir_excluidos']
            )
            
            # ⭐ TRATAMENTO ROBUSTO DE NOTAS ⭐
            # Garante que texto_notas seja sempre uma string, mesmo que venha como bool ou None
            incluir_notas = bool(config.get('incluir_notas', False))
            texto_notas_raw = config.get('texto_notas', '')
            
            # Conversão segura para string
            if texto_notas_raw is None or texto_notas_raw is False:
                texto_notas = ''
            elif texto_notas_raw is True:
                texto_notas = ''
            elif isinstance(texto_notas_raw, str):
                texto_notas = texto_notas_raw
            else:
                texto_notas = str(texto_notas_raw)
            
            # 5. Montar dados completos
            dados_completos = {
                # DataFrames processados
                'df_filtrado': df_filtrado,
                'df_diaria': df_diaria,
                'df_tp_desp_1': df_tp_desp_1,
                'df_tp_desp_2': df_tp_desp_2,
                'df_futuro': df_futuro,
                'df_original': df_original,
                
                # Configurações
                'incluir_futuros': config['incluir_futuros'],
                'incluir_excluidos': config['incluir_excluidos'],
                'data_relatorio': config['data'],
                
                # Informações do cliente
                'nome_cliente': ws_resumo['A3'].value,
                'endereco_cliente': ws_resumo['A4'].value,
                'numero_relatorio': numero_relatorio,
                'acumulado': valor_acumulado,
                
                # ⭐ NOTAS TRATADAS ⭐
                'incluir_notas': incluir_notas,
                'texto_notas': texto_notas
            }
            
            workbook.close()
            
            # Debug com verificação de tipo e tamanho
            print("📊 DADOS PRONTOS:")
            print(f"   - incluir_notas: {dados_completos['incluir_notas']} (tipo: {type(dados_completos['incluir_notas']).__name__})")
            print(f"   - texto_notas tipo: {type(dados_completos['texto_notas']).__name__}")
            
            if dados_completos['incluir_notas'] and dados_completos['texto_notas']:
                comprimento = len(dados_completos['texto_notas'])
                if comprimento > 50:
                    texto_preview = dados_completos['texto_notas'][:50] + "..."
                else:
                    texto_preview = dados_completos['texto_notas']
                print(f"   - texto_notas ({comprimento} caracteres): {texto_preview}")
            else:
                print(f"   - texto_notas: (vazio)")
            
            # Debug de lançamentos futuros
            print(f"   - incluir_futuros: {dados_completos['incluir_futuros']}")
            print(f"   - df_futuro tipo: {type(dados_completos['df_futuro']).__name__}")
            if dados_completos['df_futuro'] is not None:
                import pandas as pd
                if isinstance(dados_completos['df_futuro'], pd.DataFrame):
                    print(f"   - df_futuro tamanho: {len(dados_completos['df_futuro'])} registros")
                    print(f"   - df_futuro vazio?: {dados_completos['df_futuro'].empty}")
                else:
                    print(f"   - df_futuro: não é DataFrame!")
            else:
                print(f"   - df_futuro: None")

            return dados_completos
            
        except Exception as e:
            print(f"💥 ERRO no processar_para_preview: {str(e)}")
            import traceback
            traceback.print_exc()
            raise
    
    def gerar_pdf_temporario(self, dados_completos, arquivo_original):
        """Gera PDF temporário para análise detalhada - PRESERVA ORDENAÇÕES"""
        try:
            logger.info("🔧 GERANDO PDF TEMPORÁRIO COM ORDENAÇÕES PRESERVADAS")
            
            # 1. Criar arquivo temporário com nome descritivo
            temp_dir = tempfile.gettempdir()
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_cliente = dados_completos.get('nome_cliente', 'Cliente').replace(' ', '_')
            temp_name = f"PREVIEW_{nome_cliente}_{timestamp}.pdf"
            temp_path = os.path.join(temp_dir, temp_name)
            
            logger.info(f"📄 Criando PDF temporário: {temp_name}")
            
            # 2. CRUCIAL: Usar o handler original com dados ordenados
            # Isso garante que TODAS as formatações e ordenações sejam preservadas
            self.handler.gerar_relatorio_pdf(
                dados_completos, 
                temp_path, 
                arquivo_original
            )
            
            logger.info(f"✅ PDF temporário gerado: {temp_path}")
            
            # 3. Verificar se arquivo foi criado
            if not os.path.exists(temp_path):
                raise Exception(f"PDF temporário não foi criado em: {temp_path}")
            
            # 4. Verificar tamanho do arquivo
            tamanho = os.path.getsize(temp_path)
            if tamanho < 1000:  # Muito pequeno, provável erro
                logger.warning(f"⚠️ PDF temporário muito pequeno: {tamanho} bytes")
            else:
                logger.info(f"✅ PDF temporário válido: {tamanho:,} bytes")
            
            return temp_path
            
        except Exception as e:
            logger.error(f"💥 ERRO ao gerar PDF temporário: {str(e)}", exc_info=True)
            raise
    
    def gerar_pdf_definitivo(self, dados_completos, arquivo_original):
        """Gera PDF definitivo na pasta correta"""
        # Determinar nome e caminho final
        data_formatada = dados_completos['data_relatorio'].strftime('%d-%m-%Y')
        nome_cliente = dados_completos['nome_cliente']
        nome_arquivo = f"REL - {nome_cliente} - {data_formatada}.pdf"
        
        if dados_completos['incluir_excluidos']:
            nome_arquivo = nome_arquivo.replace('.pdf', ' (com excluídos).pdf')
        
        pasta_cliente = os.path.dirname(arquivo_original)
        caminho_final = os.path.join(pasta_cliente, nome_arquivo)
        
        # USAR MÉTODO ORIGINAL (garante formatação correta)
        self.handler.gerar_relatorio_pdf(dados_completos, caminho_final, arquivo_original)
        
        return caminho_final, nome_arquivo
    
    def verificar_metodos_handler(self):
        """Método para debug - verificar que métodos existem no handler"""
        try:
            print("🔍 MÉTODOS DISPONÍVEIS NO HANDLER:")
            metodos = [method for method in dir(self.handler) if not method.startswith('_')]
            
            # Filtrar métodos relacionados a lançamentos futuros
            metodos_futuros = [m for m in metodos if 'futuro' in m.lower()]
            print(f"Métodos relacionados a futuros: {metodos_futuros}")
            
            # Filtrar métodos relacionados a processamento
            metodos_proc = [m for m in metodos if 'processo' in m.lower() or 'processar' in m.lower()]
            print(f"Métodos de processamento: {metodos_proc}")
            
            return metodos
            
        except Exception as e:
            print(f"Erro ao verificar métodos: {str(e)}")
            return []