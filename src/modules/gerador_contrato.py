"""
Módulo de Geração de Contratos de Prestação de Serviços
Integrado ao sistema de Gestão de Medições

VERSÃO 2.1 - Corrigida com tratamento robusto de colunas

LOCALIZAÇÃO: src/modules/gerador_contrato.py
"""

import os
import json
import subprocess
from datetime import datetime
from pathlib import Path
import tempfile
import shutil
import pandas as pd
import numpy as np
from openpyxl import load_workbook

# Importar configurações do sistema
from src.config.config import (
    ARQUIVO_CLIENTES,
    ARQUIVO_FORNECEDORES,
    PASTA_CLIENTES,
    BASE_PATH
)

# Importar logger
from src.config.logger_config import system_logger, log_action
logger = system_logger.get_logger()

# Importar utils
from src.config.utils import (
    formatar_cnpj_cpf,
    buscar_dados_bancarios_fornecedor
)


class GeradorContrato:
    """Classe para geração de contratos em formato DOCX"""
    
    # Caminho do arquivo JSON de serviços (junto com planilhas_base)
    SERVICOS_JSON_PATH = BASE_PATH / "servicos_construcao.json"
    
    # Pasta de contratos (única, dentro de Clientes/)
    PASTA_CONTRATOS = PASTA_CLIENTES / "Contratos"
    

    def _verificar_nodejs(self):
        """Verifica se Node.js está instalado e acessível"""
        try:
            # Tentar encontrar node
            node_path = shutil.which('node')
            
            if node_path:
                logger.info(f"Node.js encontrado em: {node_path}")
                return 'node'
            
            # Se não encontrou, tentar caminhos comuns do Windows
            caminhos_comuns = [
                r"C:\Program Files\nodejs\node.exe",
                r"C:\Program Files (x86)\nodejs\node.exe",
                os.path.expanduser(r"~\AppData\Roaming\npm\node.exe"),
                os.path.expanduser(r"~\AppData\Local\Programs\nodejs\node.exe")
            ]
            
            for caminho in caminhos_comuns:
                if os.path.exists(caminho):
                    logger.info(f"Node.js encontrado em: {caminho}")
                    return caminho
            
            # Não encontrou
            logger.error("Node.js não foi encontrado no sistema!")
            logger.error("Por favor, instale o Node.js de: https://nodejs.org/")
            return None
            
        except Exception as e:
            logger.error(f"Erro ao verificar Node.js: {e}")
            return None
    
    def __init__(self):
        """Inicializa o gerador de contratos"""
        self.servicos_json = self._carregar_servicos()
        self._garantir_pasta_contratos()
        
        # Verificar se Node.js está disponível
        self.node_path = self._verificar_nodejs()
        if not self.node_path:
            logger.warning("Node.js não encontrado - geração de contratos não funcionará!")
        
        logger.info("GeradorContrato inicializado com sucesso")
    
    def _garantir_pasta_contratos(self):
        """Garante que a pasta de contratos existe"""
        try:
            self.PASTA_CONTRATOS.mkdir(parents=True, exist_ok=True)
            logger.info(f"Pasta de contratos verificada: {self.PASTA_CONTRATOS}")
        except Exception as e:
            logger.error(f"Erro ao criar pasta de contratos: {e}")
    
    def _carregar_servicos(self):
        """Carrega a lista de serviços do arquivo JSON"""
        json_path = self.SERVICOS_JSON_PATH
        
        logger.info(f"Carregando serviços de: {json_path}")
        
        # Se não existir, criar arquivo JSON básico
        if not json_path.exists():
            logger.warning("Arquivo de serviços não encontrado, criando padrão...")
            servicos_basicos = {
                "categorias": {
                    "fundacao": {
                        "nome": "Fundação e Estrutura",
                        "servicos": [
                            "escavação e movimento de terra",
                            "fundação em sapata corrida",
                            "fundação em radier",
                            "estrutura em concreto armado",
                            "estrutura metálica",
                            "laje pré-moldada",
                            "impermeabilização de fundação"
                        ]
                    },
                    "alvenaria": {
                        "nome": "Alvenaria e Vedação",
                        "servicos": [
                            "alvenaria de tijolo cerâmico",
                            "alvenaria de bloco de concreto",
                            "parede em drywall",
                            "chapisco e emboço interno",
                            "reboco fino interno"
                        ]
                    },
                    "cobertura": {
                        "nome": "Cobertura",
                        "servicos": [
                            "estrutura de madeira para telhado",
                            "cobertura em telha cerâmica",
                            "cobertura em telha metálica",
                            "calha e rufo",
                            "forro em PVC"
                        ]
                    },
                    "instalacoes": {
                        "nome": "Instalações",
                        "servicos": [
                            "instalação hidráulica completa",
                            "instalação elétrica completa",
                            "instalação de louças e metais"
                        ]
                    },
                    "revestimentos": {
                        "nome": "Revestimentos",
                        "servicos": [
                            "contrapiso",
                            "piso cerâmico",
                            "piso porcelanato",
                            "azulejo em parede",
                            "pintura em látex"
                        ]
                    }
                }
            }
            
            try:
                json_path.parent.mkdir(parents=True, exist_ok=True)
                with open(json_path, 'w', encoding='utf-8') as f:
                    json.dump(servicos_basicos, f, ensure_ascii=False, indent=2)
                logger.info("Arquivo de serviços criado com sucesso")
            except Exception as e:
                logger.error(f"Erro ao criar arquivo de serviços: {e}")
                return servicos_basicos
        
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                servicos = json.load(f)
            logger.info(f"Serviços carregados: {len(servicos.get('categorias', {}))} categorias")
            return servicos
        except Exception as e:
            logger.error(f"Erro ao ler arquivo de serviços: {e}")
            return {"categorias": {}}
    
    def listar_categorias_servicos(self):
        """Retorna lista de categorias de serviços disponíveis"""
        categorias = []
        for key, value in self.servicos_json.get('categorias', {}).items():
            categorias.append({
                'id': key,
                'nome': value['nome'],
                'qtd_servicos': len(value['servicos'])
            })
        return categorias
    
    def listar_servicos_categoria(self, categoria_id):
        """Retorna lista de serviços de uma categoria específica"""
        categorias = self.servicos_json.get('categorias', {})
        if categoria_id in categorias:
            return categorias[categoria_id]['servicos']
        return []
    
    def _get_safe_value(self, row, col_name, default=''):
        """Obtém valor de forma segura, tratando NaN e None"""
        try:
            value = row.get(col_name, default)
            # Verificar se é NaN (pandas/numpy)
            if pd.isna(value) or value is None:
                return default
            # Converter para string e limpar
            return str(value).strip()
        except:
            return default
    
    def formatar_cno(self, cno):
        """Formata CNO no padrão XX.XXX.XXXXX/XX"""
        try:
            # Se vazio, retornar vazio
            if not cno or pd.isna(cno):
                return ''
            
            # Remover caracteres não numéricos
            cno_limpo = ''.join(filter(str.isdigit, str(cno)))
            
            # Se não tiver dígitos suficientes, retornar original
            if len(cno_limpo) < 12:
                return str(cno)
            
            # Garantir 13 dígitos
            cno_limpo = cno_limpo.zfill(13)
            
            # Formatar: XX.XXX.XXXXX/XX
            return f"{cno_limpo[:2]}.{cno_limpo[2:5]}.{cno_limpo[5:10]}/{cno_limpo[10:12]}"
        except:
            return str(cno) if cno else ''
    
    def obter_dados_cliente(self, nome_cliente):
        """Obtém dados do cliente da planilha - VERSÃO ROBUSTA"""
        try:
            logger.info(f"Obtendo dados do cliente: {nome_cliente}")
            
            # Ler planilha
            df = pd.read_excel(ARQUIVO_CLIENTES)
            
            # Verificar colunas disponíveis
            logger.info(f"Colunas disponíveis: {list(df.columns)}")
            
            # Tentar diferentes nomes para coluna de cliente
            col_cliente = None
            for nome_col in ['Nome', 'Cliente', 'cliente', 'CLIENTE', 'nome', 'NOME']:
                if nome_col in df.columns:
                    col_cliente = nome_col
                    break
            
            if not col_cliente:
                logger.error("Coluna de nome do cliente não encontrada")
                return None
            
            # Buscar cliente
            cliente_row = df[df[col_cliente] == nome_cliente]
            
            if cliente_row.empty:
                logger.error(f"Cliente '{nome_cliente}' não encontrado na planilha")
                return None
            
            cliente = cliente_row.iloc[0]
            
            # Obter dados com tratamento robusto
            dados = {
                'nome': self._get_safe_value(cliente, col_cliente, nome_cliente),
                'cnpj_cpf': self._get_safe_value(cliente, 'CPF', ''),
                'cno': self.formatar_cno(self._get_safe_value(cliente, 'CNO', '')),
                'estado_civil': self._get_safe_value(cliente, 'Estado Civil', ''),
                'cidade': self._get_safe_value(cliente, 'Cidade', ''),
                'endereco': self._get_safe_value(cliente, 'Endereço', '')
            }
            
            # Log de campos vazios
            campos_vazios = [k for k, v in dados.items() if not v]
            if campos_vazios:
                logger.warning(f"Campos vazios para cliente '{nome_cliente}': {', '.join(campos_vazios)}")
            
            logger.info(f"Dados do cliente '{nome_cliente}' carregados com sucesso")
            return dados
            
        except Exception as e:
            logger.error(f"Erro ao obter dados do cliente: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def obter_dados_fornecedor(self, cnpj_cpf):
        """Obtém dados do fornecedor da planilha"""
        try:
            df = pd.read_excel(ARQUIVO_FORNECEDORES)
            
            # Limpar CNPJ/CPF para comparação
            cnpj_limpo = ''.join(filter(str.isdigit, str(cnpj_cpf)))
            
            # Buscar fornecedor
            fornecedor = None
            for _, row in df.iterrows():
                row_cnpj = ''.join(filter(str.isdigit, str(row.get('CNPJ/CPF', ''))))
                if row_cnpj == cnpj_limpo:
                    fornecedor = row
                    break
            
            if fornecedor is None:
                logger.warning(f"Fornecedor não encontrado: {cnpj_cpf}")
                return None
            
            dados = {
                'nome': self._get_safe_value(fornecedor, 'Nome', ''),
                'cnpj_cpf': formatar_cnpj_cpf(fornecedor.get('CNPJ/CPF', '')),
                'endereco': self._get_safe_value(fornecedor, 'Endereço', ''),
                'tipo_pessoa': 'física' if len(cnpj_limpo) == 11 else 'jurídica'
            }
            
            logger.info(f"Dados do fornecedor '{dados['nome']}' carregados com sucesso")
            return dados
        except Exception as e:
            logger.error(f"Erro ao obter dados do fornecedor: {e}")
            return None
        
    def obter_dados_fornecedor_por_nome(self, nome_fornecedor):
        
        try:
            logger.info(f"Buscando fornecedor por nome: {nome_fornecedor}")
            
            from openpyxl import load_workbook
            
            wb = load_workbook(ARQUIVO_FORNECEDORES)
            ws = wb['Fornecedores']
            
            # Buscar fornecedor pelo nome
            fornecedor = None
            
            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:  # Pular sem CNPJ/CPF
                    continue
                
                razao_social = row[2] if len(row) > 2 else ''
                nome_fantasia = row[3] if len(row) > 3 else ''
                
                # Verificar se é o fornecedor
                if (nome_fantasia and str(nome_fantasia).strip() == nome_fornecedor) or \
                (razao_social and str(razao_social).strip() == nome_fornecedor):
                    fornecedor = row
                    break
            
            wb.close()
            
            if fornecedor is None:
                logger.warning(f"Fornecedor não encontrado: {nome_fornecedor}")
                return None
            
            # Extrair dados
            cnpj_cpf_raw = fornecedor[0]
            # CORREÇÃO: Endereço está na coluna P (índice 15), não na coluna O (índice 14)
            endereco_raw = fornecedor[15] if len(fornecedor) > 15 else ''
            
            # Limpar CNPJ/CPF
            cnpj_cpf_str = str(cnpj_cpf_raw).strip()
            if cnpj_cpf_str.endswith('.0'):
                cnpj_cpf_str = cnpj_cpf_str[:-2]
            
            cnpj_limpo = ''.join(filter(str.isdigit, cnpj_cpf_str))
            
            # Montar dados
            dados = {
                'nome': nome_fornecedor,
                'cnpj_cpf': formatar_cnpj_cpf(cnpj_cpf_str),
                'endereco': str(endereco_raw).strip() if not pd.isna(endereco_raw) else '',
                'tipo_pessoa': 'física' if len(cnpj_limpo) == 11 else 'jurídica'
            }
            
            logger.info(f"✅ Dados do fornecedor '{nome_fornecedor}' carregados")
            return dados
            
        except Exception as e:
            logger.error(f"❌ Erro ao obter dados do fornecedor: {e}")
            import traceback
            traceback.print_exc()
            return None
    
    def numero_por_extenso(self, numero):
        """Converte número para extenso (simplificado)"""
        unidades = ['', 'um', 'dois', 'três', 'quatro', 'cinco', 'seis', 'sete', 'oito', 'nove']
        dezenas = ['', '', 'vinte', 'trinta', 'quarenta', 'cinquenta', 'sessenta', 'setenta', 'oitenta', 'noventa']
        especiais = ['dez', 'onze', 'doze', 'treze', 'quatorze', 'quinze', 'dezesseis', 'dezessete', 'dezoito', 'dezenove']
        centenas = ['', 'cento', 'duzentos', 'trezentos', 'quatrocentos', 'quinhentos', 'seiscentos', 'setecentos', 'oitocentos', 'novecentos']
        
        try:
            numero = int(numero)
        except:
            return str(numero)
        
        if numero == 0:
            return "zero"
        
        if numero < 10:
            return unidades[numero]
        elif numero < 20:
            return especiais[numero - 10]
        elif numero < 100:
            dezena = numero // 10
            unidade = numero % 10
            if unidade == 0:
                return dezenas[dezena]
            else:
                return f"{dezenas[dezena]} e {unidades[unidade]}"
        elif numero < 1000:
            centena = numero // 100
            resto = numero % 100
            if numero == 100:
                return "cem"
            if resto == 0:
                return centenas[centena]
            else:
                return f"{centenas[centena]} e {self.numero_por_extenso(resto)}"
        elif numero < 1000000:
            milhar = numero // 1000
            resto = numero % 1000
            if milhar == 1:
                mil_text = "mil"
            else:
                mil_text = f"{self.numero_por_extenso(milhar)} mil"
            
            if resto == 0:
                return mil_text
            else:
                return f"{mil_text} e {self.numero_por_extenso(resto)}"
        elif numero < 1000000000:  # Até 999 milhões
            milhao = numero // 1000000
            resto = numero % 1000000
            
            if milhao == 1:
                milhao_text = "um milhão"
            else:
                milhao_text = f"{self.numero_por_extenso(milhao)} milhões"
            
            if resto == 0:
                return milhao_text
            else:
                return f"{milhao_text} e {self.numero_por_extenso(resto)}"
        else:
            # Para valores muito grandes, retornar em formato legível
            return f"{numero:,}".replace(',', '.')
    
    def data_por_extenso(self, data_str):
        """Converte data para extenso - apenas mês por extenso"""
        meses = [
            'janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
            'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'
        ]
        
        try:
            if isinstance(data_str, str):
                data = datetime.strptime(data_str, '%d/%m/%Y')
            else:
                data = data_str
            
            dia = data.day
            mes = meses[data.month - 1]
            ano = data.year
            
            return f"{dia} de {mes} de {ano}"
        except Exception as e:
            logger.error(f"Erro ao converter data para extenso: {e}")
            return str(data_str)
    
    def valor_por_extenso(self, valor):
        """Converte valor monetário para extenso"""
        try:
            # Limpar string de valor
            valor_str = str(valor).replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor_float = float(valor_str)
            
            reais = int(valor_float)
            centavos = int((valor_float - reais) * 100)
            
            reais_extenso = self.numero_por_extenso(reais)
            
            if centavos > 0:
                centavos_extenso = self.numero_por_extenso(centavos)
                return f"{reais_extenso} reais e {centavos_extenso} centavos"
            else:
                return f"{reais_extenso} reais"
        except Exception as e:
            logger.error(f"Erro ao converter valor para extenso: {e}")
            return str(valor)
    
    def concatenar_servicos(self, lista_servicos):
        """Concatena lista de serviços de forma gramaticalmente correta"""
        if not lista_servicos:
            return ""
        
        if len(lista_servicos) == 1:
            return lista_servicos[0]
        
        if len(lista_servicos) == 2:
            return f"{lista_servicos[0]} e {lista_servicos[1]}"
        
        # Múltiplos serviços: "x, y, z e w"
        servicos_texto = ", ".join(lista_servicos[:-1])
        return f"{servicos_texto} e {lista_servicos[-1]}"
    
    def escapar_texto_js(self, texto):
        """Escapa caracteres especiais para uso em JavaScript"""
        if texto is None:
            return ""
        texto = str(texto)
        texto = texto.replace('\\', '\\\\')
        texto = texto.replace('"', '\\"')
        texto = texto.replace('\n', '\\n')
        texto = texto.replace('\r', '\\r')
        texto = texto.replace('\t', '\\t')
        return texto
    

    def formatar_valor_monetario(self, valor_str):
        """Formata valor para padrão monetário brasileiro R$ #.##0,00"""
        try:
            # Remover caracteres não numéricos exceto vírgula e ponto
            valor_limpo = valor_str.replace('R$', '').replace('.', '').replace(',', '.').strip()
            valor_float = float(valor_limpo)
            
            # Formatar como moeda brasileira
            valor_formatado = f"R$ {valor_float:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            
            return valor_formatado
        except Exception as e:
            logger.error(f"Erro ao formatar valor monetário: {e}")
            return valor_str
    
    def gerar_nome_arquivo_contrato(self, cliente_nome, data_contrato):
        """Gera nome de arquivo com data do contrato"""
        # Converter data_contrato para string no formato YYYY-MM-DD
        if isinstance(data_contrato, str):
            # Se veio como DD/MM/YYYY, converter para YYYY-MM-DD
            try:
                data_obj = datetime.strptime(data_contrato, '%d/%m/%Y')
                data_formatada = data_obj.strftime("%Y-%m-%d")
            except:
                data_formatada = datetime.now().strftime("%Y-%m-%d")
        else:
            data_formatada = data_contrato.strftime("%Y-%m-%d")
        
        cliente_safe = "".join(c for c in cliente_nome if c.isalnum() or c in (' ', '-', '_')).strip()
        cliente_safe = cliente_safe.replace(' ', '_')
        return f"contrato_{cliente_safe}_{data_formatada}.docx"
    
    @log_action("Gerar contrato")
    def gerar_contrato(self, dados_contrato):
        """
        Gera contrato em formato DOCX usando docx-js
        
        Args:
            dados_contrato: dicionário com os dados do contrato
        
        Returns:
            caminho do arquivo gerado ou None se erro
        """
        try:
            # Validar dados obrigatórios
            campos_obrigatorios = [
                'data', 'cidade', 'cliente_nome', 'cliente_cno', 'cliente_cpf',
                'cliente_estado_civil', 'cliente_endereco', 'fornecedor_nome',
                'fornecedor_cnpj_cpf', 'fornecedor_endereco', 'descricao',
                'endereco_obra', 'dias', 'data_inicio', 'data_fim', 'valor',
                'multa', 'dados_bancarios'
            ]
            
            campos_faltantes = []
            campos_vazios = []
            for campo in campos_obrigatorios:
                if campo not in dados_contrato:
                    campos_faltantes.append(campo)
                elif not dados_contrato[campo] or str(dados_contrato[campo]).strip() == '':
                    campos_vazios.append(campo)
            
            if campos_faltantes:
                logger.error(f"Campos obrigatórios faltantes: {', '.join(campos_faltantes)}")
                raise ValueError(f"Campos obrigatórios ausentes: {', '.join(campos_faltantes)}")
            
            if campos_vazios:
                logger.warning(f"Campos vazios (serão preenchidos com valor padrão): {', '.join(campos_vazios)}")
                # Preencher campos vazios com valores padrão
                for campo in campos_vazios:
                    if campo in ['cliente_cno', 'cliente_cpf', 'cliente_estado_civil']:
                        dados_contrato[campo] = '[PREENCHER]'
                    elif campo == 'cidade':
                        dados_contrato[campo] = 'Belo Horizonte'
            
            # Preparar dados
            data_extenso = self.data_por_extenso(dados_contrato['data'])
            
            # Limpar e formatar valores monetários
            # Garantir que o valor seja tratado corretamente
            valor_original = dados_contrato['valor']
            multa_original = dados_contrato['multa']
            
            # Formatar valores monetários
            valor_global_formatado = self.formatar_valor_monetario(valor_original)
            multa_formatada = self.formatar_valor_monetario(multa_original)
            
            # Converter valores para extenso
            try:
                valor_global_extenso = self.valor_por_extenso(valor_original)
                if not valor_global_extenso or valor_global_extenso.strip() == '':
                    valor_global_extenso = 'valor não especificado'
            except Exception as e:
                logger.error(f"Erro ao converter valor para extenso: {e}")
                valor_global_extenso = 'valor não especificado'
            
            try:
                multa_extenso = self.valor_por_extenso(multa_original)
                if not multa_extenso or multa_extenso.strip() == '':
                    multa_extenso = 'valor não especificado'
            except Exception as e:
                logger.error(f"Erro ao converter multa para extenso: {e}")
                multa_extenso = 'valor não especificado'
            
            logger.info(f"Valor formatado: {valor_global_formatado} ({valor_global_extenso})")
            logger.info(f"Multa formatada: {multa_formatada} ({multa_extenso})")
            
            # Gerar nome do arquivo
            nome_arquivo = self.gerar_nome_arquivo_contrato(dados_contrato['cliente_nome'], dados_contrato['data'])
            arquivo_saida = self.PASTA_CONTRATOS / nome_arquivo
            
            logger.info(f"Gerando contrato: {arquivo_saida}")
            
            # Criar script JavaScript (mesmo código anterior)
            js_script = f"""
const fs = require('fs');
const {{ Document, Packer, Paragraph, TextRun, AlignmentType, UnderlineType, HeadingLevel }} = require('docx');

const dados = {{
    data: "{self.escapar_texto_js(dados_contrato['data'])}",
    cidade: "{self.escapar_texto_js(dados_contrato['cidade'])}",
    cliente_nome: "{self.escapar_texto_js(dados_contrato['cliente_nome'])}",
    cliente_cno: "{self.escapar_texto_js(dados_contrato['cliente_cno'])}",
    cliente_cpf: "{self.escapar_texto_js(dados_contrato['cliente_cpf'])}",
    cliente_estado_civil: "{self.escapar_texto_js(dados_contrato['cliente_estado_civil'])}",
    cliente_endereco: "{self.escapar_texto_js(dados_contrato['cliente_endereco'])}",
    fornecedor_nome: "{self.escapar_texto_js(dados_contrato['fornecedor_nome'])}",
    fornecedor_cnpj_cpf: "{self.escapar_texto_js(dados_contrato['fornecedor_cnpj_cpf'])}",
    fornecedor_endereco: "{self.escapar_texto_js(dados_contrato['fornecedor_endereco'])}",
    descricao: "{self.escapar_texto_js(dados_contrato['descricao'])}",
    endereco_obra: "{self.escapar_texto_js(dados_contrato['endereco_obra'])}",
    dias: "{self.escapar_texto_js(str(dados_contrato['dias']))}",
    data_inicio: "{self.escapar_texto_js(dados_contrato['data_inicio'])}",
    data_fim: "{self.escapar_texto_js(dados_contrato['data_fim'])}",
    valor: "{self.escapar_texto_js(dados_contrato['valor'])}",
    valor_extenso: "{self.escapar_texto_js(valor_global_extenso)}",
    multa: "{self.escapar_texto_js(dados_contrato['multa'])}",
    multa_extenso: "{self.escapar_texto_js(multa_extenso)}",
    data_extenso: "{self.escapar_texto_js(data_extenso)}",
    dados_bancarios: "{self.escapar_texto_js(dados_contrato['dados_bancarios'])}"
}};

const doc = new Document({{
    styles: {{
        default: {{ document: {{ run: {{ font: "Arial", size: 24 }} }} }},
        paragraphStyles: [
            {{ id: "Title", name: "Title", basedOn: "Normal",
                run: {{ size: 28, bold: true, color: "000000", font: "Arial" }},
                paragraph: {{ spacing: {{ before: 240, after: 240 }}, alignment: AlignmentType.CENTER }} }},
            {{ id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal",
                run: {{ size: 24, bold: true, color: "000000", font: "Arial" }},
                paragraph: {{ spacing: {{ before: 200, after: 100 }}, alignment: AlignmentType.LEFT }} }},
            {{ id: "Normal", name: "Normal",
                run: {{ size: 22, color: "000000", font: "Arial" }},
                paragraph: {{ spacing: {{ before: 0, after: 100, line: 360 }}, alignment: AlignmentType.JUSTIFIED }} }}
        ]
    }},
    sections: [{{
        properties: {{ page: {{ margin: {{ top: 1440, right: 1440, bottom: 1440, left: 1440 }} }} }},
        children: [
            new Paragraph({{ heading: HeadingLevel.TITLE,
                children: [new TextRun({{ text: "CONTRATO PARTICULAR DE PRESTAÇÃO DE SERVIÇOS POR EMPREITADA", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun(`Aos ${{dados.data}}, nesta cidade ${{dados.cidade}}, entre partes, de um lado: `)] }}),
            new Paragraph({{ children: [
                new TextRun({{ text: dados.cliente_nome, bold: true }}),
                new TextRun(`, pessoa física devidamente inscrita sob o CNO n.º ${{dados.cliente_cno}} e CPF nº ${{dados.cliente_cpf}}, ${{dados.cliente_estado_civil}}, residente na ${{dados.cliente_endereco}}, doravante denominada CONTRATANTE e, de outro `),
                new TextRun({{ text: dados.fornecedor_nome, bold: true }}),
                new TextRun(`, pessoa física devidamente inscrita sob o CPF n.º ${{dados.fornecedor_cnpj_cpf}} com residência na ${{dados.fornecedor_endereco}}, doravante denominado simplesmente de CONTRATADA, ambas representadas por seus representantes legais que ao final firmam o presente contrato, tem entre si, justo e contratado o presente, que se regerá pelas seguintes Cláusulas e Condições:`)
            ] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA PRIMEIRA - OBJETO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun(`O presente contrato tem como OBJETO a prestação de serviços especializados em ${{dados.descricao}} bem como todos os trabalhos e atividades necessárias para sua conclusão.`)] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO PRIMEIRO: ", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun(`Os serviços deverão ser prestados no imóvel situado à ${{dados.endereco_obra}}`)] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO SEGUNDO: ", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("A contratada prestará os serviços constantes em orçamento e/ou descritivo de atividades na modalidade por empreitada de forma autônoma, sem qualquer exclusividade, podendo desempenhar atividades para terceiros em geral, simultaneamente ou não.")] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA SEGUNDA - SERVIÇOS", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("Os serviços acima mencionados serão prestados pela contratada através de seus prepostos ou empregados devidamente registrados, sem qualquer vinculação com a contratante.")] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO PRIMEIRO:", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("O Contratado obrigar-se-á:")] }}),
            new Paragraph({{ children: [
                new TextRun("a) executar os serviços autônomos com toda a perfeição técnica na forma e modo ajustados, dentro das normas e especificações técnicas aplicáveis à espécie e "),
                new TextRun({{ text: "em estrito cumprimento dos detalhes, projetos e especificações, dando plena e total garantia dos mesmos;", underline: {{ type: UnderlineType.SINGLE }} }})
            ] }}),
            new Paragraph({{ children: [new TextRun("b) fornecer toda mão-de-obra necessária à execução e entrega dos serviços no prazo estabelecido, devendo registrar todos os trabalhadores em seu nome, obrigando-se pelos salários dos empregados que o mesmo utilizar na obra, comprometendo-se a respeitar as normas trabalhistas, de segurança do trabalho e previdenciárias vigentes;")] }}),
            new Paragraph({{ children: [new TextRun("c) fornecer todas as ferramentas necessárias para a execução dos serviços contratados;")] }}),
            new Paragraph({{ children: [new TextRun("d) corrigir, por sua conta e risco, qualquer defeito constatado durante a construção ou instalação ou execução e/ou oriundo de imperfeição de serviços;")] }}),
            new Paragraph({{ children: [new TextRun("e) pagamento dos encargos sociais, previdenciários e trabalhistas dos colaboradores utilizados na execução dos serviços ora contratados;")] }}),
            new Paragraph({{ children: [new TextRun("f) garantir a solidez e estabilidade do serviço prestado, assumindo, por ela, inteira responsabilidade, pelos danos oriundos de sua negligência, imprudência ou imperícia nos termos do Código Civil Brasileiro;")] }}),
            new Paragraph({{ children: [new TextRun("g) manter, por sua conta, seguro contra acidentes de trabalho em nome de todos os colaboradores que trabalharem na obra;")] }}),
            new Paragraph({{ children: [new TextRun("h) Fornecer, zelar e garantir o uso de equipamentos de proteção individuais e coletivos na execução dos serviços e ambiente da obra, como forma de atender todas as normas de segurança e higiene do trabalho vigentes e pertinentes ao ramo de sua atividade.")] }}),
            new Paragraph({{ children: [new TextRun("i) Avaliar e mitigar os riscos para iniciar a execução dos trabalhos sendo que na possibilidade de verificar o menor risco de acidente deverá comunicar o contratante sem adentrar ao ambiente de prestação de serviços, medida necessária para garantir segurança aos seus colaboradores.")] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO SEGUNDO:", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("São obrigações exclusivas do contratante:")] }}),
            new Paragraph({{ children: [new TextRun("a) Fornecer todos os detalhes, projetos e especificações para a perfeita execução dos serviços;")] }}),
            new Paragraph({{ children: [new TextRun("b) Efetuar o pagamento na forma e modo aprazados.")] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA TERCEIRA - PRAZO", bold: true }})] }}),
            new Paragraph({{ children: [
                new TextRun("Os serviços ora contratados serão executados/prestados até o limite de "),
                new TextRun({{ text: `${{dados.dias}} dias`, bold: true }}),
                new TextRun(", iniciando-se a contagem com a assinatura deste.")
            ] }}),
            new Paragraph({{ children: [new TextRun(`Iniciando-se a contagem com a entrada no campo de obras que está prevista para ${{dados.data_inicio}} e encerrando-se em ${{dados.data_fim}}.`)] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA QUARTA -- REMUNERAÇÃO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun(`Como remuneração pelos serviços a serem prestados, os contratantes pagarão ao contratado, mediante depósito/transferência bancária, o valor de ${{dados.valor}} (${{dados.valor_extenso}}), para pagamento integral dos serviços contratados por este instrumento valores fixos e irreajustáveis, valores que serão pagos mediante medição, após sua execução. Os valores convencionados deverão ser pagos na medida e prazos em que a prestação de serviços se desenvolver, podendo o contratante reter o pagamento, sem nenhum ônus, caso o serviço não seja prestado adequadamente ou integralmente nos moldes e diretrizes estabelecidas pelas partes e projetos de conhecimento.`)] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO PRIMEIRO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("A remuneração pelos serviços contratados inclui todos os encargos trabalhistas, sociais, previdenciários, securitários e outros não nominados, gastos e despesas relativos ao exercício dos serviços contratados, por mais especiais que sejam, nada mais sendo devido pelo contratante ao contratado, a qualquer título.")] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO SEGUNDO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("O presente contrato não implica em qualquer vínculo empregatício do contratado, de seus prepostos ou colaboradores pelos serviços prestados ao contratante.")] }}),
            new Paragraph({{ children: [new TextRun({{ text: "PARÁGRAFO TERCEIRO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("Os comprovantes de transferência servirão como recibo de quitação dos valores eventualmente pagos à Contratada.")] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA QUINTA - DISPOSIÇÕES GERAIS", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("a) As alterações de valores que venham a ser discutidos e aprovados pelas partes, deverão necessariamente ser objeto de Termo Aditivo.")] }}),
            new Paragraph({{ children: [new TextRun("b) A transferência ou cessão dos serviços de que trata o presente instrumento depende do consentimento expresso deste contratante, bem como a aditivo contratual, constando assinatura do contratante.")] }}),
            new Paragraph({{ children: [new TextRun("c) É expressamente vedada à Contratada a utilização de trabalhadores menores, púberes ou impúberes, para a prestação dos serviços.")] }}),
            new Paragraph({{ children: [new TextRun("d) Ao contratante fica ressalvado o direito à ação regressiva em face do contratado e ainda, a retenção da importância devida, em razão da quitação de eventuais obrigações trabalhistas dos empregados do contratado que eventualmente venha a sofrer em decorrência de acordos ou decisões judiciais.")] }}),
            new Paragraph({{ children: [new TextRun("e) Fica assegurado o direito do contratante ao ressarcimento dos danos sofridos em virtude de interpelação judicial em razão de obrigação não cumprida pelo contratado, inclusive eventuais despesas com honorários advocatícios contratuais.")] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA SEXTA -- DOS PREJUÍZOS", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("A contratada responderá por qualquer prejuízo que direta ou indiretamente cause ao contratante ou a terceiros, seja por ação ou omissão, sua ou de seus prepostos, empregados ou colaboradores.")] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA SÉTIMA -- DA RESCISÃO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("Serão casos de rescisão contratual:")] }}),
            new Paragraph({{ children: [new TextRun("a) a desistência de uma das partes antes de iniciada a prestação de serviços;")] }}),
            new Paragraph({{ children: [new TextRun("b) a falha do Contratado em executar os trabalhos ora especificados, nas condições estipuladas ou paralisação da obra por mais de 7 (sete) dias sem relevante razão;")] }}),
            new Paragraph({{ children: [new TextRun("c) qualquer outro fato ou ato que, por culpa ou dolo de uma das partes, impossibilite a execução do presente contrato.")] }}),
            new Paragraph({{ children: [
                new TextRun({{ text: "PARÁGRAFO ÚNICO -- ", bold: true }}),
                new TextRun(`Além das possibilidades elencadas no caput o inadimplemento de quaisquer das cláusulas estabelecidas neste instrumento, facultará a parte que não lhe deu causa, impor sua rescisão cumulada com ressarcimento de eventuais perdas e danos e lucros cessantes e multa pecuniária irredutível e não compensatória, no valor de ${{dados.multa}} (${{dados.multa_extenso}}).`)
            ] }}),
            new Paragraph({{ heading: HeadingLevel.HEADING_1,
                children: [new TextRun({{ text: "CLÁUSULA OITAVA - FORO", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun("Elegem as partes o foro da Comarca de Belo Horizonte, Estado de Minas Gerais, para nele serem dirimidas todas e quaisquer dúvidas ou questões oriundas do presente contrato, renunciando as partes a qualquer outro, por mais especial e privilegiado que seja.")] }}),
            new Paragraph({{ children: [new TextRun("E por estarem assim justos e contratados, assinam o presente em duas (02) vias de igual teor e forma, na presença de duas testemunhas, obrigando-se por si e seus sucessores, para que produzam todos os efeitos de direito.")] }}),
            new Paragraph({{ spacing: {{ before: 200, after: 200 }},
                children: [new TextRun(`Belo Horizonte -- MG, ${{dados.data_extenso}}.`)] }}),
            new Paragraph({{ spacing: {{ before: 400 }},
                children: [new TextRun("________________________________________________________________________")] }}),
            new Paragraph({{ alignment: AlignmentType.CENTER,
                children: [new TextRun({{ text: dados.cliente_nome, bold: true }})] }}),
            new Paragraph({{ spacing: {{ before: 200 }},
                children: [new TextRun("________________________________________________________________________")] }}),
            new Paragraph({{ alignment: AlignmentType.CENTER,
                children: [new TextRun({{ text: dados.fornecedor_nome, bold: true }})] }}),
            new Paragraph({{ spacing: {{ before: 300 }},
                children: [new TextRun({{ text: "Testemunhas:", bold: true }})] }}),
            new Paragraph({{ spacing: {{ before: 200 }},
                children: [new TextRun("________________________________________________________________________")] }}),
            new Paragraph({{ children: [new TextRun("RG n.º                                                   RG n.º")] }}),
            new Paragraph({{ spacing: {{ before: 300 }},
                children: [new TextRun({{ text: "DADOS BANCÁRIOS PARA PAGAMENTO DA PRESTAÇÃO DE SERVIÇOS:", bold: true }})] }}),
            new Paragraph({{ children: [new TextRun({{ text: dados.fornecedor_nome, bold: true }})] }}),
            new Paragraph({{ children: [new TextRun(dados.dados_bancarios)] }})
        ]
    }}]
}});

Packer.toBuffer(doc).then(buffer => {{
    fs.writeFileSync("{self.escapar_texto_js(str(arquivo_saida))}", buffer);
    console.log("Contrato gerado com sucesso!");
}}).catch(err => {{
    console.error("Erro ao gerar contrato:", err);
    process.exit(1);
}});
"""
            
            # Salvar script temporário
            # Usar diretório temporário do sistema (funciona em Windows, Linux e Mac)
            temp_dir = Path(tempfile.gettempdir())
            script_path = temp_dir / "gerar_contrato.js"
            
            # CRIAR PACKAGE.JSON E INSTALAR DOCX NO DIRETÓRIO TEMPORÁRIO
            # Isso garante que o Node.js encontre o módulo
            package_json_path = temp_dir / "package.json"
            node_modules_path = temp_dir / "node_modules"
            
            # Criar package.json se não existir
            if not package_json_path.exists():
                logger.info("Criando package.json no diretório temporário...")
                package_json = {
                    "name": "gerador-contrato-temp",
                    "version": "1.0.0",
                    "dependencies": {
                        "docx": "^9.0.0"
                    }
                }
                with open(package_json_path, 'w', encoding='utf-8') as f:
                    json.dump(package_json, f, indent=2)
            
            # Instalar docx se node_modules não existir ou não tiver docx
            docx_module_path = node_modules_path / "docx"
            if not docx_module_path.exists():
                logger.info("Instalando biblioteca docx no diretório temporário...")
                logger.info("Isso pode levar alguns segundos na primeira vez...")
                
                try:
                    # Encontrar npm
                    npm_path = shutil.which('npm')
                    if not npm_path:
                        # Tentar caminhos comuns do Windows
                        npm_paths = [
                            r"C:\Program Files\nodejs\npm.cmd",
                            r"C:\Program Files (x86)\nodejs\npm.cmd",
                        ]
                        for path in npm_paths:
                            if Path(path).exists():
                                npm_path = path
                                break
                    
                    if npm_path:
                        # Instalar docx no diretório temporário
                        install_result = subprocess.run(
                            [npm_path, 'install', 'docx'],
                            cwd=str(temp_dir),
                            capture_output=True,
                            text=True,
                            timeout=60,
                            shell=True  # Necessário no Windows para .cmd
                        )
                        
                        if install_result.returncode == 0:
                            logger.info("✅ Biblioteca docx instalada com sucesso!")
                        else:
                            logger.warning(f"Aviso ao instalar docx: {install_result.stderr}")
                    else:
                        logger.warning("NPM não encontrado - tentando usar instalação global...")
                        
                except Exception as e:
                    logger.warning(f"Erro ao instalar docx localmente: {e}")
                    logger.info("Tentando usar instalação global do docx...")
            
            # Salvar o script JavaScript
            with open(script_path, 'w', encoding='utf-8') as f:
                f.write(js_script)
            
            # Verificar se Node.js está disponível
            if not self.node_path:
                erro_msg = (
                    "Node.js não está instalado ou não foi encontrado!\n\n"
                    "Para gerar contratos, você precisa instalar o Node.js:\n"
                    "1. Baixe em: https://nodejs.org/\n"
                    "2. Instale a versão LTS (recomendada)\n"
                    "3. Reinicie o sistema\n\n"
                    "Após instalar, feche e reabra este programa."
                )
                logger.error(erro_msg)
                raise RuntimeError(erro_msg)
            
            # Executar script Node.js
            logger.info(f"Executando Node.js: {self.node_path}")
            result = subprocess.run(
                [self.node_path, str(script_path)],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode != 0:
                logger.error(f"Erro ao executar script Node.js: {result.stderr}")
                return None
            
            # Verificar se arquivo foi criado
            if not arquivo_saida.exists():
                logger.error("Arquivo de contrato não foi criado")
                return None
            
            logger.info(f"Contrato gerado com sucesso: {arquivo_saida}")
            return str(arquivo_saida)
            
        except Exception as e:
            logger.error(f"Erro ao gerar contrato: {e}")
            import traceback
            traceback.print_exc()
            return None


if __name__ == "__main__":
    # Teste básico
    print("Testando GeradorContrato...")
    gerador = GeradorContrato()
    print(f"Categorias disponíveis: {len(gerador.listar_categorias_servicos())}")