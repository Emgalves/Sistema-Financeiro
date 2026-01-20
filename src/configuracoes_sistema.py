import tkinter as tk
from tkinter import ttk, messagebox
import json
from datetime import datetime
from pathlib import Path
import os
from openpyxl import load_workbook, Workbook
from src.config.logger_config import system_logger, log_action

logger = system_logger.get_logger()

class GerenciadorConfiguracoes:
    # Importar o caminho do BASE_PATH definido no config.py
    from src.config.config import BASE_PATH
    
    # Definir o caminho do arquivo de configurações no mesmo local das planilhas base
    CONFIG_PATH = BASE_PATH / "parametros_sistema.json"
    MATERIAIS_CONFIG_PATH = BASE_PATH / "parametros_materiais.json"
    SERVICOS_JSON_PATH = BASE_PATH / "servicos_construcao.json"
    
    # Cache de configurações para acesso rápido
    _config_cache = None
    _materiais_cache = None
    
    @staticmethod
    def _atualizar_cache(config):
        """Atualiza o cache de configurações"""
        GerenciadorConfiguracoes._config_cache = config
    
    @staticmethod
    def _atualizar_cache_materiais(config):
        """Atualiza o cache de configurações de materiais"""
        GerenciadorConfiguracoes._materiais_cache = config
    
    @staticmethod
    def _garantir_estrutura_completa(config):
        """Garante que a estrutura de configurações está completa"""
        estrutura_padrao = GerenciadorConfiguracoes._obter_configuracoes_padrao_estaticas()
        
        # Verificar e adicionar seções que podem estar faltando
        for secao, valores in estrutura_padrao.items():
            if secao not in config:
                config[secao] = valores
                logger.info(f"Seção '{secao}' adicionada às configurações")
        
        # Verificar estruturas específicas
        if 'indices_correcao' in config:
            if 'indices_disponiveis' not in config['indices_correcao']:
                config['indices_correcao']['indices_disponiveis'] = estrutura_padrao['indices_correcao']['indices_disponiveis']
        
        if 'correcao_automatica' in config:
            for chave, valor in estrutura_padrao['correcao_automatica'].items():
                if chave not in config['correcao_automatica']:
                    config['correcao_automatica'][chave] = valor
        
        return config
    
    @staticmethod
    @log_action("Carregar configurações")
    def carregar_configuracoes():
        """
        Método estático para carregar configurações do sistema
        """
        # Verificar se há cache disponível
        if GerenciadorConfiguracoes._config_cache is not None:
            return GerenciadorConfiguracoes._config_cache
            
        config_path = GerenciadorConfiguracoes.CONFIG_PATH
        
        # Imprimir informação de debug
        print(f"Tentando carregar configurações de: {config_path}")
        print(f"O arquivo existe? {config_path.exists()}")
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # GARANTIR que as novas seções existam (para arquivos antigos)
                config = GerenciadorConfiguracoes._garantir_estrutura_completa(config)
                
                # Atualizar o cache
                GerenciadorConfiguracoes._atualizar_cache(config)
                return config
            except Exception as e:
                logger.error(f"Erro ao carregar configurações: {e}")
                return None
        
        # Se o arquivo não existir, criar com configurações padrão completas
        default_config = GerenciadorConfiguracoes._obter_configuracoes_padrao_estaticas()
        
        try:
            # Garantir que o diretório existe
            config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            
            GerenciadorConfiguracoes._atualizar_cache(default_config)
            return default_config
        except Exception as e:
            logger.error(f"Erro ao criar arquivo de configurações: {e}")
            return None

    @staticmethod
    def carregar_configuracoes_materiais():
        """Método estático para carregar configurações de materiais"""
        # Verificar se há cache disponível
        if GerenciadorConfiguracoes._materiais_cache is not None:
            return GerenciadorConfiguracoes._materiais_cache
            
        config_path = GerenciadorConfiguracoes.MATERIAIS_CONFIG_PATH
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                
                # Atualizar o cache
                GerenciadorConfiguracoes._atualizar_cache_materiais(config)
                return config
            except Exception as e:
                logger.error(f"Erro ao carregar configurações de materiais: {e}")
                return None
        
        # Se o arquivo não existir, criar com configurações padrão
        default_config = GerenciadorConfiguracoes._obter_configuracoes_materiais_padrao()
        
        try:
            # Garantir que o diretório existe
            config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, indent=4, ensure_ascii=False)
            
            GerenciadorConfiguracoes._atualizar_cache_materiais(default_config)
            return default_config
        except Exception as e:
            logger.error(f"Erro ao criar arquivo de configurações de materiais: {e}")
            return None

    @staticmethod
    def _obter_configuracoes_materiais_padrao():
        """Configurações padrão para materiais"""
        return {
            "categorias_materiais": {
                "REVESTIMENTO": {
                    "subcategorias": ["CERAMICA", "PORCELANATO", "PEDRA NATURAL", "MADEIRA", "VINILICO", "PAPEL PAREDE", "PASTILHA", "LAMINADO"],
                    "cor": "#8B4513"
                },
                "ACABAMENTO": {
                    "subcategorias": ["RODAPE", "MOLDURA", "SANCA", "BAGUETE", "PERFIL", "GUARNIÇÃO", "ALISAR", "ACABAMENTO"],
                    "cor": "#4682B4"
                },
                "ILUMINACAO": {
                    "subcategorias": ["LUMINARIA LED", "SPOT", "PENDENTE", "ARANDELA", "LUSTRE", "FITA LED", "LAMPADA", "DIMMER"],
                    "cor": "#FFD700"
                },
                "HIDRAULICO": {
                    "subcategorias": ["TORNEIRA", "CHUVEIRO", "VASO SANITARIO", "CUBA", "TANQUE", "TUBULACAO", "VALVULA", "REGISTRO"],
                    "cor": "#0000FF"
                },
                "ELETRICO": {
                    "subcategorias": ["TOMADA", "INTERRUPTOR", "DISJUNTOR", "QUADRO", "CABO", "ELETRODUTO", "CAIXA"],
                    "cor": "#FF4500"
                },
                "OUTROS": {
                    "subcategorias": ["DIVERSOS", "ACESSORIO", "FERRAMENTA", "CONSUMIVEL"],
                    "cor": "#808080"
                }
            },
            "ambientes": [
                "INSTALAÇÃO DA OBRA", "SALA", "COZINHA", "BANHEIRO SUITE", "BANHEIRO SOCIAL", 
                "QUARTO CASAL", "QUARTO SOLTEIRO", "QUARTO HOSPEDE",
                "VARANDA", "GARAGEM", "AREA EXTERNA", "PISCINA", 
                "SAUNA", "JARDIM", "ESCRITORIO", "LAVANDERIA",
                "DESPENSA", "ADEGA", "CHURRASQUEIRA", "TODOS AMBIENTES"
            ],
            "status_instalacao": [
                "PENDENTE", "EM INSTALACAO", "INSTALADO", "GARANTIA", "MANUTENCAO"
            ],
            "unidades": [
                "PC", "M2", "MT", "KG", "LT", "CX", "UN", "PAR", "JG", "GL", "BD", "RL"
            ]
        }
    
    @staticmethod
    def _obter_configuracoes_padrao_estaticas():
        """Configurações padrão para método estático"""
        return {
            'cafe': {
                'valor_atual': 4.00,
                'historico': [
                    {'valor': 4.00, 'data_inicio': '01/01/2024', 'data_fim': None}
                ]
            },
            'bancos': {
                'lista': ['BANCO DO BRASIL', 'BRADESCO', 'CAIXA', 'ITAU', 'SANTANDER'],
                'historico_alteracoes': []
            },
            'categorias': {
                'lista': ['ADM', 'DIV', 'LOC', 'MAT', 'MO', 'SERV', 'TP'],
                'historico_alteracoes': []
            },
            # NOVA SEÇÃO: Etapas da Obra
            'etapas_obra': {
                'lista': [
                    'DEMOLIÇÃO',
                    'FUNDAÇÃO',
                    'ESTRUTURA',
                    'ALVENARIA',
                    'INSTALAÇÕES HIDRÁULICAS',
                    'INSTALAÇÕES ELÉTRICAS',
                    'COBERTURA',
                    'ESQUADRIAS',
                    'REVESTIMENTOS',
                    'PISOS',
                    'PINTURA',
                    'ACABAMENTOS',
                    'LIMPEZA FINAL'
                ],
                'historico_alteracoes': []
            },
            'indices_correcao': {
                'indice_padrao': 'IGPM',
                'indices_disponiveis': {
                    'IGPM': {
                        'nome_completo': 'Índice Geral de Preços do Mercado',
                        'historico': [],
                        'ultimo_calculo': None
                    },
                    'IPCA': {
                        'nome_completo': 'Índice Nacional de Preços ao Consumidor Amplo',
                        'historico': [],
                        'ultimo_calculo': None
                    },
                    'INPC': {
                        'nome_completo': 'Índice Nacional de Preços ao Consumidor',
                        'historico': [],
                        'ultimo_calculo': None
                    }
                }
            },
            'correcao_automatica': {
                'ativa': True,
                'dia_calculo': 15,
                'meses_aplicacao': [1, 4, 7, 10],
                'avisar_antes_dias': 7,
                'ultimo_processamento': None
            },
            'historico_correcoes': []
        }

    @staticmethod
    def get_bancos():
        """Retorna a lista de bancos"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'bancos' in config:
            return config['bancos']['lista']
        return []

    @staticmethod
    def get_categorias_fornecedor():
        """Retorna a lista de categorias de fornecedor"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'categorias' in config:
            return config['categorias']['lista']
        return ['ADM', 'DIV', 'LOC', 'MAT', 'MO', 'SERV', 'TP']
    
    @staticmethod
    def get_etapas_obra():
        """Retorna a lista de etapas da obra"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'etapas_obra' in config:
            return config['etapas_obra']['lista']
        return ['DEMOLIÇÃO', 'FUNDAÇÃO', 'ESTRUTURA', 'ALVENARIA', 
                'INSTALAÇÕES HIDRÁULICAS', 'INSTALAÇÕES ELÉTRICAS', 
                'COBERTURA', 'ESQUADRIAS', 'REVESTIMENTOS', 'PISOS', 
                'PINTURA', 'ACABAMENTOS', 'LIMPEZA FINAL']

    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Configurações do Sistema")
        self.root.geometry("920x950")
        
        # Usar o caminho da variável de classe 
        self.config_path = GerenciadorConfiguracoes.CONFIG_PATH
        self.materiais_config_path = GerenciadorConfiguracoes.MATERIAIS_CONFIG_PATH
        
        # Carregar ou criar configurações iniciais
        self.carregar_configuracoes_locais()
        
        # IMPORTANTE: Carregar configurações de materiais
        self.carregar_configuracoes_materiais_locais()
        
        # Setup da interface
        self.setup_gui()

    def carregar_configuracoes_locais(self):
        """Carrega as configurações do sistema com suporte completo a correção monetária"""
        self.config = GerenciadorConfiguracoes.carregar_configuracoes()
        
        # Se não foi possível carregar, criar configurações padrão COMPLETAS
        if self.config is None:
            self.config = self._obter_configuracoes_padrao_completas()
            self.salvar_configuracoes()

    def carregar_configuracoes_materiais_locais(self):
        """Carrega as configurações de materiais"""
        self.materiais_config = GerenciadorConfiguracoes.carregar_configuracoes_materiais()
        
        # Se não foi possível carregar, criar configurações padrão
        if self.materiais_config is None:
            self.materiais_config = GerenciadorConfiguracoes._obter_configuracoes_materiais_padrao()
            self.salvar_configuracoes_materiais()

    def _obter_configuracoes_padrao_completas(self):
        """Retorna configurações padrão completas incluindo correção monetária"""
        return {
            # Configurações existentes
            'cafe': {
                'valor_atual': 4.00,
                'historico': [
                    {'valor': 4.00, 'data_inicio': '01/01/2024', 'data_fim': None}
                ]
            },
            'bancos': {
                'lista': ['BANCO DO BRASIL', 'BRADESCO', 'CAIXA', 'ITAU', 'SANTANDER'],
                    'historico_alteracoes': []
            },
            'categorias': {
                'lista': ['ADM', 'DIV', 'LOC', 'MAT', 'MO', 'SERV', 'TP'],
                'historico_alteracoes': []
            },
            'compromissos_recorrentes': {
                'ativos': True,
                'lista': [
                    {
                        'nome': 'FOLHA DP',
                        'dia_vencimento': 5,
                        'recorrencia': 'mensal',
                        'valor_estimado': 0.0,
                        'categoria': 'MO',
                        'tipo_despesa': 3,
                        'ativo': True,
                        'observacao': 'Gestão de folha de pagamento',
                        'mes_referencia': 'anterior',  # ✅ NOVO CAMPO
                        'mes_ref_numero': None,
                        'meses_ocorrencias': None
                    },
                    {
                        'nome': 'FGTS',
                        'dia_vencimento': 20,
                        'recorrencia': 'mensal',
                        'valor_estimado': 0.0,
                        'categoria': 'MO',
                        'tipo_despesa': 3,
                        'ativo': True,
                        'observacao': 'Recolhimento FGTS',
                        'mes_referencia': 'anterior',  # ✅ Mês ANTERIOR
                        'mes_ref_numero': None,
                        'meses_ocorrencias': None
                    },
                    {
                        'nome': 'COPASA',
                        'dia_vencimento': 20,
                        'recorrencia': 'mensal',
                        'valor_estimado': 0.0,
                        'categoria': 'SERV',
                        'tipo_despesa': 3,
                        'ativo': True,
                        'observacao': 'Conta de água',
                        'mes_referencia': 'atual',  # ✅ Mês ATUAL
                        'mes_ref_numero': None,
                        'meses_ocorrencias': None
                    },
                    {
                        'nome': 'MOTOBOY',
                        'dia_vencimento': 5,
                        'recorrencia': 'mensal',
                        'valor_estimado': 0.0,
                        'categoria': 'DIV',
                        'tipo_despesa': 2,
                        'ativo': True,
                        'observacao': 'Serviço de motoboy obra',
                        'mes_referencia': 'anterior',  # ✅ Mês ANTERIOR
                        'mes_ref_numero': None,
                        'meses_ocorrencias': None
                    }
                ],
                'historico_alteracoes': []
            },
            'etapas_obra': {
                'lista': [
                    'DEMOLIÇÃO',
                    'FUNDAÇÃO', 
                    'ESTRUTURA',
                    'ALVENARIA',
                    'INSTALAÇÕES HIDRÁULICAS',
                    'INSTALAÇÕES ELÉTRICAS',
                    'COBERTURA',
                    'ESQUADRIAS',
                    'REVESTIMENTOS',
                    'PISOS',
                    'PINTURA',
                    'ACABAMENTOS',
                    'LIMPEZA FINAL'
                ],
                'historico_alteracoes': []
            },
            'insumos': {
                'lista': [
                    'CIMENTO',
                    'TIJOLO',
                    'EMPREITEIROS',
                    'BLOCOS CONCRETO',
                    'MATERIAIS ELÉTRICO',
                    'MATERIAIS HIDRÁULICO',
                    'ESPAÇADOR',
                    'MASSA CORRIDA',
                    'BIANCO',
                    'TELAS',
                    'UNIFORMES',
                    'ARGAMASSA',
                    'LONA',
                    'MANTA ASFÁLTICA'
                ],
                'historico_alteracoes': []
            },
            'indices_correcao': {
                'indice_padrao': 'IGPM',
                'indices_disponiveis': {
                    'IGPM': {
                        'nome_completo': 'Índice Geral de Preços do Mercado',
                        'historico': [],
                        'ultimo_calculo': None
                    },
                    'IPCA': {
                        'nome_completo': 'Índice Nacional de Preços ao Consumidor Amplo',
                        'historico': [],
                        'ultimo_calculo': None
                    },
                    'INPC': {
                        'nome_completo': 'Índice Nacional de Preços ao Consumidor',
                        'historico': [],
                        'ultimo_calculo': None
                    }
                }
            },
            # NOVAS CONFIGURAÇÕES: Correção automática
            'correcao_automatica': {
                'ativa': True,
                'dia_calculo': 15,  # Dia do mês para calcular correções
                'meses_aplicacao': [1, 4, 7, 10],  # Trimestral por padrão
                'avisar_antes_dias': 7,  # Avisar 7 dias antes da correção
                'ultimo_processamento': None
            },
            # NOVA SEÇÃO: Histórico de correções aplicadas
            'historico_correcoes': []
        }
                   

    def salvar_configuracoes(self):
        """Salva as configurações no arquivo"""
        try:
            # Garantir que o diretório existe
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4, ensure_ascii=False)
            
            # Atualizar o cache ao salvar
            GerenciadorConfiguracoes._atualizar_cache(self.config)
            
            print(f"Configurações salvas com sucesso em: {self.config_path}")
        except Exception as e:
            logger.error(f"Erro ao salvar configurações: {e}")
            messagebox.showerror("Erro", f"Não foi possível salvar as configurações: {e}")

    def salvar_configuracoes_materiais(self):
        """Salva as configurações de materiais no arquivo"""
        try:
            # Garantir que o diretório existe
            self.materiais_config_path.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.materiais_config_path, 'w', encoding='utf-8') as f:
                json.dump(self.materiais_config, f, indent=2, ensure_ascii=False)
            
            # Atualizar o cache ao salvar
            GerenciadorConfiguracoes._atualizar_cache_materiais(self.materiais_config)
            
            print(f"Configurações de materiais salvas com sucesso em: {self.materiais_config_path}")
        except Exception as e:
            logger.error(f"Erro ao salvar configurações de materiais: {e}")
            messagebox.showerror("Erro", f"Não foi possível salvar as configurações de materiais: {e}")

    def setup_gui(self):
        """Configura a interface gráfica"""
        # Notebook para diferentes seções
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Abas
        self.setup_aba_cafe()
        self.setup_aba_bancos()
        self.setup_aba_categorias()
        self.setup_aba_compromissos_recorrentes()
        self.setup_aba_etapas_obra()
        self.setup_aba_insumos()
        self.setup_aba_indices_correcao()
        self.setup_aba_materiais()
        self.criar_aba_servicos_construcao()
        
        # Botões globais
        frame_botoes = ttk.Frame(self.root)
        frame_botoes.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(frame_botoes, text="Salvar Todas Alterações",
                  command=self.salvar_todas_alteracoes).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Voltar ao Menu Principal", 
                  command=self.voltar_menu_local).pack(side='right', padx=5)
        # ttk.Button(frame_botoes, text="Fechar",
        #           command=self.root.quit).pack(side='right', padx=5)

    def setup_aba_cafe(self):
        """Configura a aba de valores do café"""
        frame_cafe = ttk.Frame(self.notebook)
        self.notebook.add(frame_cafe, text='Valor do Café')
        
        # Valor atual
        frame_atual = ttk.LabelFrame(frame_cafe, text="Valor Atual")
        frame_atual.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_atual, text=f"Valor atual: R$ {self.config['cafe']['valor_atual']:.2f}").pack(padx=5, pady=5)
        
        # Novo valor
        frame_novo = ttk.LabelFrame(frame_cafe, text="Definir Novo Valor")
        frame_novo.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Novo valor:").grid(row=0, column=0, padx=5, pady=5)
        self.novo_valor_cafe = ttk.Entry(frame_novo)
        self.novo_valor_cafe.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Data início:").grid(row=1, column=0, padx=5, pady=5)
        self.data_inicio_cafe = ttk.Entry(frame_novo)
        self.data_inicio_cafe.grid(row=1, column=1, padx=5, pady=5)
        self.data_inicio_cafe.insert(0, datetime.now().strftime('%d/%m/%Y'))
        
        ttk.Button(frame_novo, text="Adicionar",
                  command=self.adicionar_valor_cafe).grid(row=2, column=0, columnspan=2, pady=10)
        
        # Histórico
        frame_historico = ttk.LabelFrame(frame_cafe, text="Histórico de Valores")
        frame_historico.pack(fill='both', expand=True, padx=5, pady=5)
        
        colunas = ('Valor', 'Data Início', 'Data Fim')
        self.tree_cafe = ttk.Treeview(frame_historico, columns=colunas, show='headings')
        for col in colunas:
            self.tree_cafe.heading(col, text=col)
        self.tree_cafe.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.atualizar_historico_cafe()

    def setup_aba_bancos(self):
        """Configura a aba de bancos"""
        frame_bancos = ttk.Frame(self.notebook)
        self.notebook.add(frame_bancos, text='Bancos')
        
        # Frame para adicionar novo banco
        frame_novo = ttk.LabelFrame(frame_bancos, text="Adicionar Novo Banco")
        frame_novo.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Nome do Banco:").grid(row=0, column=0, padx=5, pady=5)
        self.novo_banco = ttk.Entry(frame_novo)
        self.novo_banco.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo, text="Adicionar",
                  command=self.adicionar_banco).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de bancos
        frame_lista = ttk.LabelFrame(frame_bancos, text="Bancos Cadastrados")
        frame_lista.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.tree_bancos = ttk.Treeview(frame_lista, columns=('Banco',), show='headings')
        self.tree_bancos.heading('Banco', text='Banco')
        self.tree_bancos.pack(fill='both', expand=True, padx=5, pady=5)
        
        ttk.Button(frame_lista, text="Remover Selecionado",
                  command=self.remover_banco).pack(pady=5)
        
        self.atualizar_lista_bancos()

    def setup_aba_categorias(self):
        """Configura a aba de categorias"""
        frame_categorias = ttk.Frame(self.notebook)
        self.notebook.add(frame_categorias, text='Categorias')
        
        # Frame para adicionar nova categoria
        frame_novo = ttk.LabelFrame(frame_categorias, text="Adicionar Nova Categoria")
        frame_novo.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Categoria:").grid(row=0, column=0, padx=5, pady=5)
        self.nova_categoria = ttk.Entry(frame_novo)
        self.nova_categoria.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo, text="Adicionar",
                  command=self.adicionar_categoria).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de categorias
        frame_lista = ttk.LabelFrame(frame_categorias, text="Categorias Cadastradas")
        frame_lista.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.tree_categorias = ttk.Treeview(frame_lista, columns=('Categoria',), show='headings')
        self.tree_categorias.heading('Categoria', text='Categoria')
        self.tree_categorias.pack(fill='both', expand=True, padx=5, pady=5)
        
        ttk.Button(frame_lista, text="Remover Selecionada",
                  command=self.remover_categoria).pack(pady=5)
        
        self.atualizar_lista_categorias()

    def adicionar_valor_cafe(self):
        """Adiciona um novo valor para o café"""
        try:
            novo_valor = float(self.novo_valor_cafe.get().replace(',', '.'))
            data_inicio = datetime.strptime(self.data_inicio_cafe.get(), '%d/%m/%Y')
            
            # Validações
            if novo_valor <= 0:
                messagebox.showerror("Erro", "O valor deve ser maior que zero!")
                return
                
            # Atualizar valor atual
            self.config['cafe']['valor_atual'] = novo_valor
            
            # Fechar o último registro do histórico
            if self.config['cafe']['historico']:
                ultimo_registro = self.config['cafe']['historico'][-1]
                if ultimo_registro['data_fim'] is None:
                    ultimo_registro['data_fim'] = data_inicio.strftime('%d/%m/%Y')
            
            # Adicionar novo registro
            self.config['cafe']['historico'].append({
                'valor': novo_valor,
                'data_inicio': data_inicio.strftime('%d/%m/%Y'),
                'data_fim': None
            })
            
            self.salvar_configuracoes()
            self.atualizar_historico_cafe()
            
            # Limpar campos
            self.novo_valor_cafe.delete(0, tk.END)
            messagebox.showinfo("Sucesso", "Novo valor do café registrado com sucesso!")
            
        except ValueError:
            messagebox.showerror("Erro", "Valor inválido!")

    def adicionar_banco(self):
        """Adiciona um novo banco à lista"""
        banco = self.novo_banco.get().strip().upper()
        if not banco:
            messagebox.showerror("Erro", "Digite o nome do banco!")
            return
            
        if banco in self.config['bancos']['lista']:
            messagebox.showerror("Erro", "Este banco já está cadastrado!")
            return
            
        self.config['bancos']['lista'].append(banco)
        self.config['bancos']['lista'].sort()
        self.salvar_configuracoes()
        
        self.novo_banco.delete(0, tk.END)
        self.atualizar_lista_bancos()
        messagebox.showinfo("Sucesso", "Banco adicionado com sucesso!")

    def adicionar_categoria(self):
        """Adiciona uma nova categoria à lista"""
        categoria = self.nova_categoria.get().strip().upper()
        if not categoria:
            messagebox.showerror("Erro", "Digite a categoria!")
            return
            
        if categoria in self.config['categorias']['lista']:
            messagebox.showerror("Erro", "Esta categoria já está cadastrada!")
            return
            
        self.config['categorias']['lista'].append(categoria)
        self.config['categorias']['lista'].sort()
        self.salvar_configuracoes()
        
        self.nova_categoria.delete(0, tk.END)
        self.atualizar_lista_categorias()
        messagebox.showinfo("Sucesso", "Categoria adicionada com sucesso!")

    def remover_banco(self):
        """Remove o banco selecionado"""
        selecionado = self.tree_bancos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um banco para remover!")
            return
            
        banco = self.tree_bancos.item(selecionado)['values'][0]
        if messagebox.askyesno("Confirmar", f"Deseja remover o banco {banco}?"):
            self.config['bancos']['lista'].remove(banco)
            self.salvar_configuracoes()
            self.atualizar_lista_bancos()

    def remover_categoria(self):
        """Remove a categoria selecionada"""
        selecionado = self.tree_categorias.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma categoria para remover!")
            return
            
        categoria = self.tree_categorias.item(selecionado)['values'][0]
        if messagebox.askyesno("Confirmar", f"Deseja remover a categoria {categoria}?"):
            self.config['categorias']['lista'].remove(categoria)
            self.salvar_configuracoes()
            self.atualizar_lista_categorias()

    def atualizar_historico_cafe(self):
        """Atualiza a exibição do histórico de valores do café"""
        for item in self.tree_cafe.get_children():
            self.tree_cafe.delete(item)
            
        for registro in self.config['cafe']['historico']:
            self.tree_cafe.insert('', 'end', values=(
                f"R$ {registro['valor']:.2f}",
                registro['data_inicio'],
                registro['data_fim'] or 'Atual'
            ))

    def atualizar_lista_bancos(self):
        """Atualiza a exibição da lista de bancos"""
        for item in self.tree_bancos.get_children():
            self.tree_bancos.delete(item)
            
        for banco in sorted(self.config['bancos']['lista']):
            self.tree_bancos.insert('', 'end', values=(banco,))

    def atualizar_lista_categorias(self):
        """Atualiza a exibição da lista de categorias"""
        for item in self.tree_categorias.get_children():
            self.tree_categorias.delete(item)
            
        for categoria in sorted(self.config['categorias']['lista']):
            self.tree_categorias.insert('', 'end', values=(categoria,))

    def setup_aba_etapas_obra(self):
        """Configura a aba de etapas da obra"""
        frame_etapas = ttk.Frame(self.notebook)
        self.notebook.add(frame_etapas, text='Etapas da Obra')
        
        # Frame para adicionar nova etapa
        frame_novo = ttk.LabelFrame(frame_etapas, text="Adicionar Nova Etapa")
        frame_novo.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Nome da Etapa:").grid(row=0, column=0, padx=5, pady=5)
        self.nova_etapa = ttk.Entry(frame_novo, width=40)
        self.nova_etapa.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo, text="Adicionar",
                command=self.adicionar_etapa_obra).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de etapas
        frame_lista = ttk.LabelFrame(frame_etapas, text="Etapas Cadastradas")
        frame_lista.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.tree_etapas = ttk.Treeview(frame_lista, columns=('Etapa',), show='headings')
        self.tree_etapas.heading('Etapa', text='Etapa da Obra')
        self.tree_etapas.pack(fill='both', expand=True, padx=5, pady=5)
        
        frame_botoes_etapas = ttk.Frame(frame_lista)
        frame_botoes_etapas.pack(fill='x', pady=5)
        
        ttk.Button(frame_botoes_etapas, text="Mover para Cima",
                command=self.mover_etapa_cima).pack(side='left', padx=5)
        ttk.Button(frame_botoes_etapas, text="Mover para Baixo",
                command=self.mover_etapa_baixo).pack(side='left', padx=5)
        ttk.Button(frame_botoes_etapas, text="Remover Selecionada",
                command=self.remover_etapa_obra).pack(side='right', padx=5)
        
        self.atualizar_lista_etapas_obra()

    def adicionar_etapa_obra(self):
        """Adiciona uma nova etapa da obra à lista"""
        etapa = self.nova_etapa.get().strip().upper()
        if not etapa:
            messagebox.showerror("Erro", "Digite o nome da etapa!")
            return
            
        if etapa in self.config['etapas_obra']['lista']:
            messagebox.showerror("Erro", "Esta etapa já está cadastrada!")
            return
            
        self.config['etapas_obra']['lista'].append(etapa)
        self.salvar_configuracoes()
        
        self.nova_etapa.delete(0, tk.END)
        self.atualizar_lista_etapas_obra()
        messagebox.showinfo("Sucesso", "Etapa adicionada com sucesso!")

    def remover_etapa_obra(self):
        """Remove a etapa selecionada"""
        selecionado = self.tree_etapas.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma etapa para remover!")
            return
            
        etapa = self.tree_etapas.item(selecionado)['values'][0]
        if messagebox.askyesno("Confirmar", f"Deseja remover a etapa '{etapa}'?"):
            self.config['etapas_obra']['lista'].remove(etapa)
            self.salvar_configuracoes()
            self.atualizar_lista_etapas_obra()

    def mover_etapa_cima(self):
        """Move a etapa selecionada para cima na lista"""
        selecionado = self.tree_etapas.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma etapa!")
            return
        
        etapa = self.tree_etapas.item(selecionado)['values'][0]
        lista = self.config['etapas_obra']['lista']
        indice = lista.index(etapa)
        
        if indice > 0:
            lista[indice], lista[indice-1] = lista[indice-1], lista[indice]
            self.salvar_configuracoes()
            self.atualizar_lista_etapas_obra()

    def mover_etapa_baixo(self):
        """Move a etapa selecionada para baixo na lista"""
        selecionado = self.tree_etapas.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma etapa!")
            return
        
        etapa = self.tree_etapas.item(selecionado)['values'][0]
        lista = self.config['etapas_obra']['lista']
        indice = lista.index(etapa)
        
        if indice < len(lista) - 1:
            lista[indice], lista[indice+1] = lista[indice+1], lista[indice]
            self.salvar_configuracoes()
            self.atualizar_lista_etapas_obra()

    def atualizar_lista_etapas_obra(self):
        """Atualiza a exibição da lista de etapas da obra"""
        for item in self.tree_etapas.get_children():
            self.tree_etapas.delete(item)
            
        for etapa in self.config['etapas_obra']['lista']:
            self.tree_etapas.insert('', 'end', values=(etapa,))

    @staticmethod
    def get_insumos():
        """Retorna a lista de insumos"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'insumos' in config:
            return config['insumos']['lista']
        return ['CIMENTO', 'TIJOLO','EMPREITEIROS','BLOCOS CONCRETO',
                'MATERIAIS ELÉTRICO','MATERIAIS HIDRÁULICO','ESPAÇADOR',
                'MASSA CORRIDA','BIANCO','TELAS','UNIFORMES','ARGAMASSA','LONA', 'OUTROS']

    @staticmethod
    def get_compromissos_recorrentes():
        """Retorna a lista de compromissos recorrentes ativos"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'compromissos_recorrentes' in config:
            # Retornar apenas os compromissos ativos
            return [c for c in config['compromissos_recorrentes']['lista'] if c.get('ativo', True)]
        return []

    @staticmethod
    def get_compromissos_recorrentes_todos():
        """Retorna todos os compromissos recorrentes (ativos e inativos)"""
        config = GerenciadorConfiguracoes.carregar_configuracoes()
        if config and 'compromissos_recorrentes' in config:
            return config['compromissos_recorrentes']['lista']
        return []

    # ==============================================================================
    # MODIFICAÇÃO 1: Atualizar setup_aba_compromissos_recorrentes
    # ==============================================================================

    def setup_aba_compromissos_recorrentes(self):
        """
        Configura a aba de compromissos recorrentes
        VERSÃO CORRIGIDA: Campos em layout vertical
        """
        import tkinter as tk
        from tkinter import ttk
        
        # Frame principal da aba
        frame_principal = ttk.Frame(self.notebook)
        self.notebook.add(frame_principal, text='Agenda - Compromissos')
        
        # Dividir em duas seções
        paned = ttk.PanedWindow(frame_principal, orient='horizontal')
        paned.pack(fill='both', expand=True, padx=10, pady=10)
        
        # ========================================================================
        # SEÇÃO ESQUERDA: Lista de compromissos cadastrados
        # ========================================================================
        
        frame_esquerda = ttk.Frame(paned)
        paned.add(frame_esquerda, weight=2)
        
        ttk.Label(frame_esquerda, text="Compromissos Recorrentes Cadastrados", 
                font=('TkDefaultFont', 11, 'bold')).pack(pady=(0, 10))
        
        frame_tree = ttk.Frame(frame_esquerda)
        frame_tree.pack(fill='both', expand=True)
        
        colunas = ('Nome', 'Dia Venc.', 'Recorrência', 'Categoria', 'Valor Est.', 'Status')
        self.tree_compromissos = ttk.Treeview(frame_tree, columns=colunas, show='headings', height=20)
        
        self.tree_compromissos.heading('Nome', text='Nome')
        self.tree_compromissos.heading('Dia Venc.', text='Dia Venc.')
        self.tree_compromissos.heading('Recorrência', text='Recorrência')
        self.tree_compromissos.heading('Categoria', text='Categoria')
        self.tree_compromissos.heading('Valor Est.', text='Valor Est.')
        self.tree_compromissos.heading('Status', text='Status')
        
        self.tree_compromissos.column('Nome', width=180)
        self.tree_compromissos.column('Dia Venc.', width=70, anchor='center')
        self.tree_compromissos.column('Recorrência', width=120, anchor='center')
        self.tree_compromissos.column('Categoria', width=80, anchor='center')
        self.tree_compromissos.column('Valor Est.', width=100, anchor='e')
        self.tree_compromissos.column('Status', width=80, anchor='center')
        
        scrollbar = ttk.Scrollbar(frame_tree, orient='vertical', command=self.tree_compromissos.yview)
        self.tree_compromissos.configure(yscrollcommand=scrollbar.set)
        
        self.tree_compromissos.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        self.tree_compromissos.tag_configure('ativo', background='#e8f5e8')
        self.tree_compromissos.tag_configure('inativo', background='#ffe4e1')
        
        self.tree_compromissos.bind('<<TreeviewSelect>>', self.on_select_compromisso)
        
        # ========================================================================
        # SEÇÃO DIREITA: Formulários
        # ========================================================================
        
        frame_direita = ttk.Frame(paned)
        paned.add(frame_direita, weight=1)
        
        # -----------------------------------------------------------------------
        # FORMULÁRIO: NOVO COMPROMISSO
        # -----------------------------------------------------------------------
        
        frame_novo = ttk.LabelFrame(frame_direita, text="Novo Compromisso", padding="10")
        frame_novo.pack(fill='x', pady=(0, 15))
        
        self.campos_novo = {}
        
        row = 0
        
        # Nome
        ttk.Label(frame_novo, text="Nome:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['nome'] = ttk.Entry(frame_novo, width=25)
        self.campos_novo['nome'].grid(row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Dia de vencimento
        ttk.Label(frame_novo, text="Dia Vencimento:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['dia_vencimento'] = ttk.Spinbox(frame_novo, from_=1, to=31, width=10)
        self.campos_novo['dia_vencimento'].set('5')
        self.campos_novo['dia_vencimento'].grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        
        # Recorrência
        ttk.Label(frame_novo, text="Recorrência:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['recorrencia'] = ttk.Combobox(frame_novo, 
                                                    values=['mensal', 'bimestral', 'trimestral', 
                                                            'semestral', 'anual'],
                                                    state='readonly', width=22)
        self.campos_novo['recorrencia'].set('mensal')
        self.campos_novo['recorrencia'].grid(row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        row += 1
        
        # ========================================================================
        # SEPARADOR ANTES DOS CAMPOS NOVOS
        # ========================================================================
        
        ttk.Separator(frame_novo, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=10
        )
        row += 1
        
        # Label indicativo (será atualizado dinamicamente)
        self.label_info_recorrencia = ttk.Label(frame_novo, 
                            text="📅 Para recorrências não-mensais:", 
                            font=('TkDefaultFont', 9, 'italic'), 
                            foreground='#666666')
        self.label_info_recorrencia.grid(row=row, column=0, columnspan=3, sticky='w', padx=5, pady=(0, 5))
        row += 1
        
        # ========================================================================
        # MÊS DE REFERÊNCIA - LINHA COMPLETA
        # ========================================================================
        
        ttk.Label(frame_novo, text="Mês de Referência:").grid(
            row=row, column=0, padx=5, pady=5, sticky='w'
        )
        
        self.campos_novo['mes_referencia'] = ttk.Combobox(frame_novo,
            values=['', '1-Jan', '2-Fev', '3-Mar', '4-Abr', '5-Mai', '6-Jun',
                    '7-Jul', '8-Ago', '9-Set', '10-Out', '11-Nov', '12-Dez'],
            state='disabled', width=22)
        self.campos_novo['mes_referencia'].grid(
            row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew'
        )
        row += 1
        
        # Tooltip mês de referência
        label_help_mes = ttk.Label(frame_novo, 
                            text="Mês da recorrência - Obrigatório se habilitado.",
                            font=('TkDefaultFont', 8), 
                            foreground='gray',
                            cursor='hand2')
        label_help_mes.grid(row=row, column=0, columnspan=3, sticky='w', padx=20, pady=(0, 5))
        
        def show_help_mes_ref(event):
            from tkinter import messagebox
            messagebox.showinfo("Ajuda - Mês de Referência",
                "Define quando começa a recorrência.\n\n"
                "Exemplos:\n"
                "• 13º Salário: 11-Nov (começa em novembro)\n"
                "• IPTU: 1-Jan (vence em janeiro)\n"
                "• Seguro: 3-Mar (renova em março)\n"
                "• Trimestral: 1-Jan (jan, abr, jul, out)")
        
        label_help_mes.bind('<Button-1>', show_help_mes_ref)
        row += 1
        
        # ========================================================================
        # MESES DE OCORRÊNCIA - LINHA COMPLETA
        # ========================================================================
        
        ttk.Label(frame_novo, text="Meses de Ocorrência:").grid(
            row=row, column=0, padx=5, pady=5, sticky='w'
        )
        
        self.campos_novo['meses_ocorrencias'] = ttk.Entry(frame_novo, width=25, state='disabled')
        self.campos_novo['meses_ocorrencias'].grid(
            row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew'
        )
        row += 1
        
        # Tooltip meses de ocorrência
        label_help_meses = ttk.Label(frame_novo,
                            text="Indicar os meses da recorrência. Ex. 6,12",
                            font=('TkDefaultFont', 8),
                            foreground='gray',
                            cursor='hand2')
        label_help_meses.grid(row=row, column=0, columnspan=3, sticky='w', padx=20, pady=(0, 10))
        
        def show_help_meses_ocorr(event):
            from tkinter import messagebox
            messagebox.showinfo("Ajuda - Meses de Ocorrência",
                "Opcional. Use para múltiplas ocorrências no mesmo período.\n\n"
                "Formato: números dos meses separados por vírgula\n\n"
                "Exemplos:\n"
                "• 13º Salário: 11,12 (novembro e dezembro)\n"
                "• Seguro (2 parcelas): 3,9 (março e setembro)\n"
                "• IPTU (única parcela): deixe vazio\n"
                "• Trimestral: deixe vazio (calcula automaticamente)")
        
        label_help_meses.bind('<Button-1>', show_help_meses_ocorr)
        row += 1
        
        # ========================================================================
        # SEPARADOR ANTES DOS CAMPOS BÁSICOS
        # ========================================================================
        
        ttk.Separator(frame_novo, orient='horizontal').grid(
            row=row, column=0, columnspan=3, sticky='ew', pady=10
        )
        row += 1
        
        # ========================================================================
        # CAMPOS BÁSICOS (CATEGORIA, TIPO, VALOR, OBSERVAÇÃO)
        # ========================================================================
        
        # Categoria
        ttk.Label(frame_novo, text="Categoria:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['categoria'] = ttk.Combobox(frame_novo, 
                                                    values=['ADM', 'DIV', 'LOC', 'MAT', 'MO', 'SERV', 'TP'],
                                                    state='readonly', width=22)
        self.campos_novo['categoria'].set('MO')
        self.campos_novo['categoria'].grid(row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        row += 1
        
        # Tipo de despesa
        ttk.Label(frame_novo, text="Tipo Despesa:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['tipo_despesa'] = ttk.Combobox(frame_novo, 
                                                        values=['2', '3', '5', '6', '7'],
                                                        state='readonly', width=10)
        self.campos_novo['tipo_despesa'].set('3')
        self.campos_novo['tipo_despesa'].grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        
        # Valor estimado
        ttk.Label(frame_novo, text="Valor Estimado (R$):").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['valor_estimado'] = ttk.Entry(frame_novo, width=15)
        self.campos_novo['valor_estimado'].insert(0, "0,00")
        self.campos_novo['valor_estimado'].grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
        
        # Observação
        ttk.Label(frame_novo, text="Observação:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['observacao'] = ttk.Entry(frame_novo, width=25)
        self.campos_novo['observacao'].grid(row=row, column=1, columnspan=2, padx=5, pady=5, sticky='ew')
        row += 1

        # Tipo de referencia
        ttk.Label(frame_novo, text="Tipo de Referência:").grid(row=row, column=0, padx=5, pady=5, sticky='w')
        self.campos_novo['tipo_referencia_mes'] = ttk.Combobox(frame_novo,
                                                        values=['anterior', 'atual'],
                                                        state='readonly', width=10) 
        self.campos_novo['tipo_referencia_mes'].set('anterior')
        self.campos_novo['tipo_referencia_mes'].grid(row=row, column=1, padx=5, pady=5, sticky='w')
        row += 1
                                                               
                                        
        
        # ========================================================================
        # BOTÃO ADICIONAR - APÓS TODOS OS CAMPOS
        # ========================================================================
        
        ttk.Button(frame_novo, text="Adicionar Compromisso", 
                command=self.adicionar_compromisso_recorrente).grid(
                    row=row, column=0, columnspan=3, pady=15
                )
        
        # ========================================================================
        # EVENTO: Mostrar/ocultar campos baseado na recorrência
        # ========================================================================
        
        def on_recorrencia_change(event=None):
            """Mostra campos adicionais apenas para recorrências não-mensais"""
            recorrencia = self.campos_novo['recorrencia'].get()
            
            if recorrencia in ['anual', 'trimestral', 'semestral', 'bimestral']:
                # Habilitar campos
                self.campos_novo['mes_referencia'].config(state='readonly')
                self.campos_novo['meses_ocorrencias'].config(state='normal')
                
                # Destacar visualmente
                self.label_info_recorrencia.config(
                    text="📅 Configure quando ocorrerá:",
                    foreground='#0066cc', 
                    font=('TkDefaultFont', 9, 'bold')
                )
            else:
                # Desabilitar e limpar campos
                self.campos_novo['mes_referencia'].config(state='disabled')
                self.campos_novo['mes_referencia'].set('')
                self.campos_novo['meses_ocorrencias'].config(state='disabled')
                self.campos_novo['meses_ocorrencias'].delete(0, tk.END)
                
                # Voltar estilo normal
                self.label_info_recorrencia.config(
                    text="📅 Para recorrências não-mensais:",
                    foreground='#666666', 
                    font=('TkDefaultFont', 9, 'italic')
                )
        
        # Vincular evento
        self.campos_novo['recorrencia'].bind('<<ComboboxSelected>>', on_recorrencia_change)
        
        # -----------------------------------------------------------------------
        # FORMULÁRIO: EDITAR COMPROMISSO (Similar ao novo, mas simplificado)
        # -----------------------------------------------------------------------
        
        frame_editar = ttk.LabelFrame(frame_direita, text="Editar Compromisso Selecionado", padding="10")
        frame_editar.pack(fill='x', pady=(0, 10))
        
        self.campos_editar = {}
        
        row_edit = 0
        
        ttk.Label(frame_editar, text="Nome:").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['nome'] = ttk.Entry(frame_editar, width=25)
        self.campos_editar['nome'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='ew')
        row_edit += 1
        
        ttk.Label(frame_editar, text="Dia Venc.:").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['dia_vencimento'] = ttk.Spinbox(frame_editar, from_=1, to=31, width=10)
        self.campos_editar['dia_vencimento'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='w')
        row_edit += 1
        
        ttk.Label(frame_editar, text="Recorrência:").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['recorrencia'] = ttk.Combobox(frame_editar,
                                                        values=['mensal', 'bimestral', 'trimestral',
                                                                'semestral', 'anual'],
                                                        state='readonly', width=22)
        self.campos_editar['recorrencia'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='ew')
        row_edit += 1
        
        # Campos novos de edição
        ttk.Label(frame_editar, text="Mês Ref.:").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['mes_referencia'] = ttk.Combobox(frame_editar,
            values=['', '1-Jan', '2-Fev', '3-Mar', '4-Abr', '5-Mai', '6-Jun',
                    '7-Jul', '8-Ago', '9-Set', '10-Out', '11-Nov', '12-Dez'],
            state='readonly', width=10)
        self.campos_editar['mes_referencia'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='w')
        row_edit += 1
        
        ttk.Label(frame_editar, text="Meses Ocorr.:").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['meses_ocorrencias'] = ttk.Entry(frame_editar, width=15)
        self.campos_editar['meses_ocorrencias'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='w')
        row_edit += 1
        
        ttk.Label(frame_editar, text="Valor (R$):").grid(row=row_edit, column=0, padx=5, pady=3, sticky='w')
        self.campos_editar['valor_estimado'] = ttk.Entry(frame_editar, width=15)
        self.campos_editar['valor_estimado'].grid(row=row_edit, column=1, padx=5, pady=3, sticky='w')
        row_edit += 1
        
        # Separador visual (opcional, mas recomendado)
        ttk.Separator(frame_editar, orient='horizontal').grid(
            row=row_edit, column=0, columnspan=2, sticky='ew', pady=5
        )
        row_edit += 1
        
        # Label + Campo
        ttk.Label(frame_editar, text="Tipo Ref.:").grid(
            row=row_edit, column=0, padx=5, pady=3, sticky='w'
        )
        self.campos_editar['tipo_referencia_mes'] = ttk.Combobox(frame_editar,
            values=['anterior', 'atual'],
            state='readonly', width=15)
        self.campos_editar['tipo_referencia_mes'].set('anterior')  # Padrão
        self.campos_editar['tipo_referencia_mes'].grid(
            row=row_edit, column=1, padx=5, pady=3, sticky='w'
        )
        row_edit += 1
        
        # Texto de ajuda pequeno
        label_help_edit = ttk.Label(frame_editar,
            text="ant. = mês passado | atual = mês do relatório",
            font=('TkDefaultFont', 7),
            foreground='gray')
        label_help_edit.grid(
            row=row_edit, column=0, columnspan=2, sticky='w', padx=20
        )
        row_edit += 1

        # Botões de ação (já existentes)
        frame_botoes_edit = ttk.Frame(frame_editar)
        frame_botoes_edit.grid(row=row_edit, column=0, columnspan=2, pady=10)
        
        ttk.Button(frame_botoes_edit, text="Salvar", 
                command=self.salvar_alteracoes_compromisso).pack(side='left', padx=2)
        ttk.Button(frame_botoes_edit, text="Ativar/Desativar", 
                command=self.toggle_compromisso_status).pack(side='left', padx=2)
        ttk.Button(frame_botoes_edit, text="Remover", 
                command=self.remover_compromisso_recorrente).pack(side='left', padx=2)
        
        # Configurar expansão de colunas
        frame_novo.columnconfigure(1, weight=1)
        frame_editar.columnconfigure(1, weight=1)
        
        # Carregar dados iniciais
        self.carregar_compromissos_tree()


    # ==============================================================================
    # MODIFICAÇÃO 2: Atualizar adicionar_compromisso_recorrente
    # ==============================================================================

    def adicionar_compromisso_recorrente(self):
        """
        Adiciona novo compromisso recorrente com suporte a mês de referência
        e múltiplas ocorrências
        """
        from tkinter import messagebox
        import json
        
        try:
            # Validações básicas
            nome = self.campos_novo['nome'].get().strip().upper()
            if not nome:
                messagebox.showerror("Erro", "Nome é obrigatório!")
                return
            
            # Coletar dados básicos
            dia_vencimento = int(self.campos_novo['dia_vencimento'].get())
            recorrencia = self.campos_novo['recorrencia'].get()
            categoria = self.campos_novo['categoria'].get()
            tipo_despesa = int(self.campos_novo['tipo_despesa'].get())
            observacao = self.campos_novo['observacao'].get().strip()
            
            # Processar valor
            valor_str = self.campos_novo['valor_estimado'].get().replace(',', '.')
            try:
                valor_estimado = float(valor_str) if valor_str else 0.0
            except ValueError:
                messagebox.showerror("Erro", "Valor inválido!")
                return
            
            # ========================================================================
            # PROCESSAR NOVOS CAMPOS
            # ========================================================================
            
            # Mês de referência
            mes_referencia = None
            if recorrencia in ['anual', 'trimestral', 'semestral']:
                mes_ref_str = self.campos_novo['mes_referencia'].get()
                if mes_ref_str:
                    # Extrair número do mês (ex: "11-Nov" -> 11)
                    try:
                        mes_referencia = int(mes_ref_str.split('-')[0])
                    except:
                        pass
            
            # Meses de ocorrências
            meses_ocorrencias = None
            meses_ocorr_str = self.campos_novo['meses_ocorrencias'].get().strip()
            if meses_ocorr_str:
                try:
                    # Converter string "11,12" em lista [11, 12]
                    meses_ocorrencias = [int(m.strip()) for m in meses_ocorr_str.split(',')]
                    
                    # Validar meses (1-12)
                    if not all(1 <= m <= 12 for m in meses_ocorrencias):
                        messagebox.showerror("Erro", "Meses de ocorrência devem estar entre 1 e 12!")
                        return
                        
                except ValueError:
                    messagebox.showerror("Erro", 
                        "Formato inválido para meses de ocorrência!\n"
                        "Use números separados por vírgula (ex: 11,12)")
                    return
            
            # ========================================================================
            # VALIDAÇÕES ESPECÍFICAS
            # ========================================================================
            
            # Para recorrências não-mensais, mês de referência é recomendado
            if recorrencia in ['anual', 'trimestral', 'semestral'] and not mes_referencia:
                resposta = messagebox.askyesno("Atenção",
                    f"Compromisso '{recorrencia}' sem mês de referência.\n\n"
                    f"Recomendamos definir o mês de referência para "
                    f"controlar quando o compromisso deve aparecer.\n\n"
                    f"Deseja continuar mesmo assim?")
                if not resposta:
                    return
            
            # ========================================================================
            # CRIAR COMPROMISSO
            # ========================================================================
            
            novo_compromisso = {
                'nome': nome,
                'dia_vencimento': dia_vencimento,
                'recorrencia': recorrencia,
                'valor_estimado': valor_estimado,
                'categoria': categoria,
                'tipo_despesa': tipo_despesa,
                'ativo': True,
                'observacao': observacao,
                'mes_referencia': mes_referencia,  # NOVO
                'meses_ocorrencias': meses_ocorrencias  # NOVO
            }
            
            # ========================================================================
            # SALVAR NO JSON
            # ========================================================================
            
            # Carregar configuração atual
            config = self.carregar_configuracoes()
            
            if 'compromissos_recorrentes' not in config:
                config['compromissos_recorrentes'] = {'lista': [], 'historico_alteracoes': []}
            
            # Verificar duplicatas
            if any(c['nome'] == nome for c in config['compromissos_recorrentes']['lista']):
                messagebox.showerror("Erro", "Já existe um compromisso com este nome!")
                return
            
            # Adicionar à lista
            config['compromissos_recorrentes']['lista'].append(novo_compromisso)
            
            # Salvar
            config_path = self.config_path
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            # Atualizar cache
            self._atualizar_cache(config)
            
            # ========================================================================
            # LIMPAR FORMULÁRIO
            # ========================================================================
            
            self.campos_novo['nome'].delete(0, tk.END)
            self.campos_novo['dia_vencimento'].set('5')
            self.campos_novo['recorrencia'].set('mensal')
            self.campos_novo['mes_referencia'].set('')
            self.campos_novo['meses_ocorrencias'].delete(0, tk.END)
            self.campos_novo['categoria'].set('MO')
            self.campos_novo['tipo_despesa'].set('3')
            self.campos_novo['valor_estimado'].delete(0, tk.END)
            self.campos_novo['valor_estimado'].insert(0, "0,00")
            self.campos_novo['observacao'].delete(0, tk.END)
            
            # Recarregar lista
            self.carregar_compromissos_tree()
            
            # Mensagem de sucesso
            msg_sucesso = f"Compromisso '{nome}' adicionado com sucesso!"
            if mes_referencia:
                meses_dict = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun',
                            7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
                msg_sucesso += f"\n\nMês de referência: {meses_dict.get(mes_referencia, mes_referencia)}"
            if meses_ocorrencias:
                msg_sucesso += f"\nOcorrências: {', '.join(map(str, meses_ocorrencias))}"
            
            messagebox.showinfo("Sucesso", msg_sucesso)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao adicionar compromisso:\n{str(e)}")
            import traceback
            traceback.print_exc()


    # ==============================================================================
    # MODIFICAÇÃO 3: Novo método on_select_compromisso
    # ==============================================================================

    def on_select_compromisso(self, event=None):
        """Preenche campos de edição quando um compromisso é selecionado"""
        try:
            selecionado = self.tree_compromissos.selection()
            if not selecionado:
                return
            
            # Obter nome do compromisso selecionado
            valores = self.tree_compromissos.item(selecionado[0])['values']
            nome_compromisso = valores[0]
            
            # Buscar compromisso completo
            config = self.carregar_configuracoes()
            compromisso = None
            
            for comp in config.get('compromissos_recorrentes', {}).get('lista', []):
                if comp['nome'] == nome_compromisso:
                    compromisso = comp
                    break
            
            if not compromisso:
                return
            
            # Preencher campos de edição
            self.campos_editar['nome'].delete(0, tk.END)
            self.campos_editar['nome'].insert(0, compromisso['nome'])
            
            self.campos_editar['dia_vencimento'].delete(0, tk.END)
            self.campos_editar['dia_vencimento'].insert(0, str(compromisso['dia_vencimento']))
            
            self.campos_editar['recorrencia'].set(compromisso['recorrencia'])
            
            # NOVOS CAMPOS
            # Mês de referência
            mes_ref = compromisso.get('mes_referencia')
            if mes_ref:
                meses_dict = {1:'1-Jan', 2:'2-Fev', 3:'3-Mar', 4:'4-Abr', 5:'5-Mai', 6:'6-Jun',
                            7:'7-Jul', 8:'8-Ago', 9:'9-Set', 10:'10-Out', 11:'11-Nov', 12:'12-Dez'}
                self.campos_editar['mes_referencia'].set(meses_dict.get(mes_ref, ''))
            else:
                self.campos_editar['mes_referencia'].set('')
            
            # Meses de ocorrências
            self.campos_editar['meses_ocorrencias'].delete(0, tk.END)
            meses_ocorr = compromisso.get('meses_ocorrencias')
            if meses_ocorr:
                self.campos_editar['meses_ocorrencias'].insert(0, ','.join(map(str, meses_ocorr)))
            
            # Valor
            self.campos_editar['valor_estimado'].delete(0, tk.END)
            valor_formatado = f"{compromisso['valor_estimado']:.2f}".replace('.', ',')
            self.campos_editar['valor_estimado'].insert(0, valor_formatado)

            # Tipo_referencia_mes (anterior/atual)
            tipo_ref = compromisso.get('tipo_referencia_mes', 'anterior')
            self.campos_editar['tipo_referencia_mes'].set(tipo_ref)
            
        except Exception as e:
            print(f"Erro ao selecionar compromisso: {e}")


    # ==============================================================================
    # MODIFICAÇÃO 4: Novo método salvar_alteracoes_compromisso
    # ==============================================================================

    def salvar_alteracoes_compromisso(self):
        """Salva alterações no compromisso selecionado"""
        from tkinter import messagebox
        import json
        
        selecionado = self.tree_compromissos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um compromisso para editar!")
            return
        
        try:
            # Obter nome original
            valores = self.tree_compromissos.item(selecionado[0])['values']
            nome_original = valores[0]
            
            # Obter novos valores
            novo_nome = self.campos_editar['nome'].get().strip().upper()
            novo_dia = int(self.campos_editar['dia_vencimento'].get())
            nova_recorrencia = self.campos_editar['recorrencia'].get()
            
            # Processar mês de referência
            novo_mes_ref = None
            mes_ref_str = self.campos_editar['mes_referencia'].get()
            if mes_ref_str:
                try:
                    novo_mes_ref = int(mes_ref_str.split('-')[0])
                except:
                    pass
            
            # Processar meses de ocorrências
            novos_meses_ocorr = None
            meses_ocorr_str = self.campos_editar['meses_ocorrencias'].get().strip()
            if meses_ocorr_str:
                try:
                    novos_meses_ocorr = [int(m.strip()) for m in meses_ocorr_str.split(',')]
                    if not all(1 <= m <= 12 for m in novos_meses_ocorr):
                        messagebox.showerror("Erro", "Meses devem estar entre 1 e 12!")
                        return
                except ValueError:
                    messagebox.showerror("Erro", "Formato inválido para meses!")
                    return
            
            # Processar valor
            valor_str = self.campos_editar['valor_estimado'].get().replace(',', '.')
            novo_valor = float(valor_str) if valor_str else 0.0

            novo_tipo_ref = self.campos_editar['tipo_referencia_mes'].get()

            
            # Carregar e atualizar configuração
            config = self.carregar_configuracoes()
            
            for comp in config['compromissos_recorrentes']['lista']:
                if comp['nome'] == nome_original:
                    comp['nome'] = novo_nome
                    comp['dia_vencimento'] = novo_dia
                    comp['recorrencia'] = nova_recorrencia
                    comp['valor_estimado'] = novo_valor
                    comp['mes_referencia'] = novo_mes_ref  # NOVO
                    comp['meses_ocorrencias'] = novos_meses_ocorr  # NOVO
                    comp['tipo_referencia_mes'] = novo_tipo_ref
                    break
            
            # Salvar
            config_path = self.config_path
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4, ensure_ascii=False)
            
            self._atualizar_cache(config)
            self.carregar_compromissos_tree()
            
            messagebox.showinfo("Sucesso", "Compromisso atualizado com sucesso!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar: {str(e)}")


    # ==============================================================================
    # MODIFICAÇÃO 5: Novo método carregar_compromissos_tree
    # ==============================================================================

    def carregar_compromissos_tree(self):
        """Carrega compromissos no treeview"""
        # Limpar tree
        for item in self.tree_compromissos.get_children():
            self.tree_compromissos.delete(item)
        
        try:
            config = self.carregar_configuracoes()
            compromissos = config.get('compromissos_recorrentes', {}).get('lista', [])
            
            for comp in compromissos:
                status = "ATIVO" if comp.get('ativo', True) else "INATIVO"
                tag = 'ativo' if comp.get('ativo', True) else 'inativo'
                
                # Formatar recorrência com indicador de mês de referência
                recorrencia_display = comp['recorrencia']
                if comp.get('mes_referencia'):
                    meses_abrev = {1:'Jan', 2:'Fev', 3:'Mar', 4:'Abr', 5:'Mai', 6:'Jun',
                                7:'Jul', 8:'Ago', 9:'Set', 10:'Out', 11:'Nov', 12:'Dez'}
                    mes_abrev = meses_abrev.get(comp['mes_referencia'], '')
                    recorrencia_display += f" ({mes_abrev})"
                
                if comp.get('meses_ocorrencias'):
                    recorrencia_display += f" *{len(comp['meses_ocorrencias'])}x"
                
                valores = (
                    comp['nome'],
                    comp['dia_vencimento'],
                    recorrencia_display,
                    comp.get('categoria', ''),
                    f"R$ {comp['valor_estimado']:.2f}",
                    status
                )
                
                self.tree_compromissos.insert('', 'end', values=valores, tags=(tag,))
            
        except Exception as e:
            print(f"Erro ao carregar compromissos: {e}")

    def toggle_compromisso_status(self):
        """Ativa/desativa compromisso selecionado"""
        selecionado = self.tree_compromissos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um compromisso!")
            return
        
        nome = self.tree_compromissos.item(selecionado[0])['values'][0]
        
        # Buscar e alterar status
        for c in self.config['compromissos_recorrentes']['lista']:
            if c['nome'] == nome:
                c['ativo'] = not c.get('ativo', True)
                status = "ATIVADO" if c['ativo'] else "DESATIVADO"
                
                # Registrar alteração
                self.config['compromissos_recorrentes']['historico_alteracoes'].append({
                    'acao': status,
                    'compromisso': nome,
                    'data': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
                })
                
                break
        
        self.salvar_configuracoes()
        self.atualizar_lista_compromissos_recorrentes()
        
        messagebox.showinfo("Sucesso", f"Status do compromisso alterado!")

    def remover_compromisso_recorrente(self):
        """Remove compromisso selecionado"""
        selecionado = self.tree_compromissos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um compromisso para remover!")
            return
        
        nome = self.tree_compromissos.item(selecionado[0])['values'][0]
        
        if messagebox.askyesno("Confirmar", f"Deseja remover o compromisso '{nome}'?"):
            # Remover da lista
            self.config['compromissos_recorrentes']['lista'] = [
                c for c in self.config['compromissos_recorrentes']['lista'] 
                if c['nome'] != nome
            ]
            
            # Registrar remoção
            self.config['compromissos_recorrentes']['historico_alteracoes'].append({
                'acao': 'REMOVER',
                'compromisso': nome,
                'data': datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            })
            
            self.salvar_configuracoes()
            self.atualizar_lista_compromissos_recorrentes()
            
            # Limpar campos de edição
            self.entry_edit_comp_nome.delete(0, tk.END)
            self.entry_edit_comp_dia.delete(0, tk.END)
            self.entry_edit_comp_valor.delete(0, tk.END)
            
            messagebox.showinfo("Sucesso", "Compromisso removido com sucesso!")

    def atualizar_lista_compromissos_recorrentes(self):
        """Atualiza a exibição da lista de compromissos recorrentes"""
        # Limpar tree
        for item in self.tree_compromissos.get_children():
            self.tree_compromissos.delete(item)
        
        # Verificar se existe a seção
        if 'compromissos_recorrentes' not in self.config:
            return
        
        # Inserir compromissos
        for compromisso in self.config['compromissos_recorrentes']['lista']:
            status = "ATIVO" if compromisso.get('ativo', True) else "INATIVO"
            tag = 'ativo' if compromisso.get('ativo', True) else 'inativo'
            
            self.tree_compromissos.insert('', 'end', 
                values=(
                    compromisso['nome'],
                    compromisso['dia_vencimento'],
                    compromisso['recorrencia'],
                    compromisso['categoria'],
                    f"R$ {compromisso['valor_estimado']:.2f}",
                    status
                ),
                tags=(tag,)
            )
        
        # Configurar tags de cor
        self.tree_compromissos.tag_configure('ativo', background='#e8f5e8')
        self.tree_compromissos.tag_configure('inativo', background='#ffe4e1')

    # Adicionar aba de insumos (método completo):
    def setup_aba_insumos(self):
        """Configura a aba de insumos"""
        frame_insumos = ttk.Frame(self.notebook)
        self.notebook.add(frame_insumos, text='Insumos')
        
        # Frame para adicionar novo insumo
        frame_novo = ttk.LabelFrame(frame_insumos, text="Adicionar Novo Insumo")
        frame_novo.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo, text="Nome do Insumo:").grid(row=0, column=0, padx=5, pady=5)
        self.novo_insumo = ttk.Entry(frame_novo, width=40)
        self.novo_insumo.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo, text="Adicionar",
                command=self.adicionar_insumo).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de insumos
        frame_lista = ttk.LabelFrame(frame_insumos, text="Insumos Cadastrados")
        frame_lista.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.tree_insumos = ttk.Treeview(frame_lista, columns=('Insumo',), show='headings')
        self.tree_insumos.heading('Insumo', text='Insumo')
        self.tree_insumos.pack(fill='both', expand=True, padx=5, pady=5)
        
        frame_botoes_insumos = ttk.Frame(frame_lista)
        frame_botoes_insumos.pack(fill='x', pady=5)
        
        ttk.Button(frame_botoes_insumos, text="Mover para Cima",
                command=self.mover_insumo_cima).pack(side='left', padx=5)
        ttk.Button(frame_botoes_insumos, text="Mover para Baixo",
                command=self.mover_insumo_baixo).pack(side='left', padx=5)
        ttk.Button(frame_botoes_insumos, text="Remover Selecionado",
                command=self.remover_insumo).pack(side='right', padx=5)
        
        self.atualizar_lista_insumos()

    def adicionar_insumo(self):
        """Adiciona um novo insumo à lista"""
        insumo = self.novo_insumo.get().strip().upper()
        if not insumo:
            messagebox.showerror("Erro", "Digite o nome do insumo!")
            return
            
        if insumo in self.config['insumos']['lista']:
            messagebox.showerror("Erro", "Este insumo já está cadastrado!")
            return
            
        self.config['insumos']['lista'].append(insumo)
        self.salvar_configuracoes()
        
        self.novo_insumo.delete(0, tk.END)
        self.atualizar_lista_insumos()
        messagebox.showinfo("Sucesso", "Insumo adicionado com sucesso!")

    def remover_insumo(self):
        """Remove o insumo selecionado"""
        selecionado = self.tree_insumos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um insumo para remover!")
            return
            
        insumo = self.tree_insumos.item(selecionado)['values'][0]
        if messagebox.askyesno("Confirmar", f"Deseja remover o insumo '{insumo}'?"):
            self.config['insumos']['lista'].remove(insumo)
            self.salvar_configuracoes()
            self.atualizar_lista_insumos()

    def mover_insumo_cima(self):
        """Move o insumo selecionado para cima na lista"""
        selecionado = self.tree_insumos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um insumo!")
            return
        
        insumo = self.tree_insumos.item(selecionado)['values'][0]
        lista = self.config['insumos']['lista']
        indice = lista.index(insumo)
        
        if indice > 0:
            lista[indice], lista[indice-1] = lista[indice-1], lista[indice]
            self.salvar_configuracoes()
            self.atualizar_lista_insumos()

    def mover_insumo_baixo(self):
        """Move o insumo selecionado para baixo na lista"""
        selecionado = self.tree_insumos.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um insumo!")
            return
        
        insumo = self.tree_insumos.item(selecionado)['values'][0]
        lista = self.config['insumos']['lista']
        indice = lista.index(insumo)
        
        if indice < len(lista) - 1:
            lista[indice], lista[indice+1] = lista[indice+1], lista[indice]
            self.salvar_configuracoes()
            self.atualizar_lista_insumos()

    def atualizar_lista_insumos(self):
        """Atualiza a exibição da lista de insumos"""
        for item in self.tree_insumos.get_children():
            self.tree_insumos.delete(item)
            
        for insumo in self.config['insumos']['lista']:
            self.tree_insumos.insert('', 'end', values=(insumo,))

    def setup_aba_indices_correcao(self):
        """Configura a aba de índices de correção monetária"""
        frame_indices = ttk.Frame(self.notebook)
        self.notebook.add(frame_indices, text='Correção Monetária')
        
        # Configurações gerais
        frame_config = ttk.LabelFrame(frame_indices, text="Configurações Gerais")
        frame_config.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_config, text="Índice Padrão:").grid(row=0, column=0, padx=5, pady=5)
        self.combo_indice_padrao = ttk.Combobox(frame_config, state='readonly')
        self.combo_indice_padrao['values'] = ['IGPM', 'IPCA', 'INPC']
        
        # Verificar se existe configuração de índices
        indices_config = self.config.get('indices_correcao', {})
        self.combo_indice_padrao.set(indices_config.get('indice_padrao', 'IGPM'))
        self.combo_indice_padrao.grid(row=0, column=1, padx=5, pady=5)
        
        # Correção automática
        frame_auto = ttk.LabelFrame(frame_indices, text="Correção Automática")
        frame_auto.pack(fill='x', padx=5, pady=5)
        
        self.var_correcao_ativa = tk.BooleanVar()
        correcao_config = self.config.get('correcao_automatica', {})
        self.var_correcao_ativa.set(correcao_config.get('ativa', True))
        
        ttk.Checkbutton(frame_auto, text="Ativar correção automática",
                    variable=self.var_correcao_ativa).grid(row=0, column=0, padx=5, pady=5, sticky='w')
        
        ttk.Label(frame_auto, text="Dia do mês para cálculo:").grid(row=1, column=0, padx=5, pady=5, sticky='w')
        self.entry_dia_calculo = ttk.Entry(frame_auto, width=5)
        self.entry_dia_calculo.insert(0, str(correcao_config.get('dia_calculo', 15)))
        self.entry_dia_calculo.grid(row=1, column=1, padx=5, pady=5, sticky='w')
        
        # Botão para abrir gerenciador completo
        ttk.Button(frame_indices, text="Abrir Gerenciador Completo de Índices",
                command=self.abrir_gerenciador_indices).pack(pady=20)

    def abrir_gerenciador_indices(self):
        """Abre o gerenciador completo de índices"""
        try:
            from src.correcao_monetaria import InterfaceIndicesCorrecao
            interface = InterfaceIndicesCorrecao(self.root)
        except ImportError as e:
            messagebox.showerror("Erro", f"Erro ao importar módulo de correção: {str(e)}")

    def salvar_todas_alteracoes(self):
        """Salva todas as alterações feitas nas configurações"""
        try:
            # Salvar configurações de correção monetária
            if hasattr(self, 'combo_indice_padrao'):
                if 'indices_correcao' not in self.config:
                    self.config['indices_correcao'] = {'indices_disponiveis': {
                        'IGPM': {'nome_completo': 'Índice Geral de Preços do Mercado', 'historico': []},
                        'IPCA': {'nome_completo': 'Índice Nacional de Preços ao Consumidor Amplo', 'historico': []},
                        'INPC': {'nome_completo': 'Índice Nacional de Preços ao Consumidor', 'historico': []}
                    }}
                self.config['indices_correcao']['indice_padrao'] = self.combo_indice_padrao.get()
            
            if hasattr(self, 'var_correcao_ativa'):
                if 'correcao_automatica' not in self.config:
                    self.config['correcao_automatica'] = {}
                self.config['correcao_automatica']['ativa'] = self.var_correcao_ativa.get()
                
                if hasattr(self, 'entry_dia_calculo'):
                    try:
                        dia = int(self.entry_dia_calculo.get())
                        if 1 <= dia <= 31:
                            self.config['correcao_automatica']['dia_calculo'] = dia
                    except ValueError:
                        pass
            
            # Salvar arquivo
            self.salvar_configuracoes()
            messagebox.showinfo("Sucesso", "Todas as alterações foram salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar alterações: {str(e)}")

    def setup_aba_materiais(self):
        """Configura a aba de parâmetros de materiais"""
        frame_materiais = ttk.Frame(self.notebook)
        self.notebook.add(frame_materiais, text='Parâmetros de Materiais')
        
        # Criar notebook interno para as subseções de materiais
        notebook_materiais = ttk.Notebook(frame_materiais)
        notebook_materiais.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Sub-aba: Categorias de Materiais
        self.setup_subaba_categorias_materiais(notebook_materiais)
        
        # Sub-aba: Ambientes
        self.setup_subaba_ambientes(notebook_materiais)
        
        # Sub-aba: Status de Instalação
        self.setup_subaba_status_instalacao(notebook_materiais)
        
        # Sub-aba: Unidades
        self.setup_subaba_unidades(notebook_materiais)

    def setup_subaba_categorias_materiais(self, parent_notebook):
        """Configura a sub-aba de categorias de materiais"""
        frame_cat_materiais = ttk.Frame(parent_notebook)
        parent_notebook.add(frame_cat_materiais, text='Categorias de Materiais')
        
        # Frame principal dividido em duas colunas
        main_frame = ttk.Frame(frame_cat_materiais)
        main_frame.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Coluna esquerda - Lista de categorias
        frame_esquerda = ttk.Frame(main_frame)
        frame_esquerda.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        # Lista de categorias existentes
        frame_lista_cat = ttk.LabelFrame(frame_esquerda, text="Categorias Existentes")
        frame_lista_cat.pack(fill='both', expand=True)
        
        self.tree_cat_materiais = ttk.Treeview(frame_lista_cat, columns=('Categoria', 'Cor'), show='headings', height=15)
        self.tree_cat_materiais.heading('Categoria', text='Categoria')
        self.tree_cat_materiais.heading('Cor', text='Cor')
        self.tree_cat_materiais.column('Categoria', width=200)
        self.tree_cat_materiais.column('Cor', width=100)
        
        scrollbar_cat = ttk.Scrollbar(frame_lista_cat, orient='vertical', command=self.tree_cat_materiais.yview)
        self.tree_cat_materiais.configure(yscrollcommand=scrollbar_cat.set)
        
        self.tree_cat_materiais.pack(side='left', fill='both', expand=True)
        scrollbar_cat.pack(side='right', fill='y')
        
        # Bind para seleção
        self.tree_cat_materiais.bind('<<TreeviewSelect>>', self.on_categoria_material_select)
        
        # Coluna direita - Detalhes e edição
        frame_direita = ttk.Frame(main_frame)
        frame_direita.pack(side='right', fill='y', padx=(5, 0))
        
        # Frame para nova categoria
        frame_nova_cat = ttk.LabelFrame(frame_direita, text="Nova Categoria")
        frame_nova_cat.pack(fill='x', pady=(0, 5))
        
        ttk.Label(frame_nova_cat, text="Nome:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.entry_nova_cat_material = ttk.Entry(frame_nova_cat, width=25)
        self.entry_nova_cat_material.grid(row=0, column=1, padx=5, pady=5)
                
        ttk.Button(frame_nova_cat, text="Adicionar Categoria",
                  command=self.adicionar_categoria_material).grid(row=2, column=0, columnspan=2, pady=10)
        
        # Frame para editar categoria selecionada
        frame_editar_cat = ttk.LabelFrame(frame_direita, text="Editar Categoria Selecionada")
        frame_editar_cat.pack(fill='x', pady=(5, 5))
        
        ttk.Label(frame_editar_cat, text="Nome:").grid(row=0, column=0, padx=5, pady=5, sticky='w')
        self.entry_editar_cat_material = ttk.Entry(frame_editar_cat, width=25)
        self.entry_editar_cat_material.grid(row=0, column=1, padx=5, pady=5)
               
        frame_botoes_cat = ttk.Frame(frame_editar_cat)
        frame_botoes_cat.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(frame_botoes_cat, text="Salvar Alterações",
                  command=self.salvar_categoria_material).pack(side='left', padx=5)
        ttk.Button(frame_botoes_cat, text="Remover Categoria",
                  command=self.remover_categoria_material).pack(side='left', padx=5)
        
        # Frame para subcategorias
        frame_subcategorias = ttk.LabelFrame(frame_direita, text="Subcategorias")
        frame_subcategorias.pack(fill='both', expand=True, pady=(5, 0))
        
        self.listbox_subcategorias = tk.Listbox(frame_subcategorias, height=8)
        scrollbar_sub = ttk.Scrollbar(frame_subcategorias, orient='vertical', command=self.listbox_subcategorias.yview)
        self.listbox_subcategorias.configure(yscrollcommand=scrollbar_sub.set)
        
        self.listbox_subcategorias.pack(side='left', fill='both', expand=True, padx=(5, 0), pady=5)
        scrollbar_sub.pack(side='right', fill='y', pady=5)
        
        # Frame para gerenciar subcategorias
        frame_ger_sub = ttk.Frame(frame_subcategorias)
        frame_ger_sub.pack(fill='x', padx=5, pady=5)
        
        self.entry_nova_subcategoria = ttk.Entry(frame_ger_sub, width=25)
        self.entry_nova_subcategoria.pack(side='top', pady=(0, 5))
        
        frame_btn_sub = ttk.Frame(frame_ger_sub)
        frame_btn_sub.pack(side='top')
        
        ttk.Button(frame_btn_sub, text="Adicionar",
                  command=self.adicionar_subcategoria).pack(side='left', padx=2)
        ttk.Button(frame_btn_sub, text="Remover",
                  command=self.remover_subcategoria).pack(side='left', padx=2)
        
        self.atualizar_lista_categorias_materiais()

    def setup_subaba_ambientes(self, parent_notebook):
        """Configura a sub-aba de ambientes"""
        frame_ambientes = ttk.Frame(parent_notebook)
        parent_notebook.add(frame_ambientes, text='Ambientes')
        
        # Frame para adicionar novo ambiente
        frame_novo_amb = ttk.LabelFrame(frame_ambientes, text="Adicionar Novo Ambiente")
        frame_novo_amb.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo_amb, text="Nome do Ambiente:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_novo_ambiente = ttk.Entry(frame_novo_amb, width=30)
        self.entry_novo_ambiente.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo_amb, text="Adicionar",
                  command=self.adicionar_ambiente).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de ambientes
        frame_lista_amb = ttk.LabelFrame(frame_ambientes, text="Ambientes Cadastrados")
        frame_lista_amb.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.listbox_ambientes = tk.Listbox(frame_lista_amb)
        scrollbar_amb = ttk.Scrollbar(frame_lista_amb, orient='vertical', command=self.listbox_ambientes.yview)
        self.listbox_ambientes.configure(yscrollcommand=scrollbar_amb.set)
        
        self.listbox_ambientes.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar_amb.pack(side='right', fill='y', pady=5)
        
        ttk.Button(frame_lista_amb, text="Remover Selecionado",
                  command=self.remover_ambiente).pack(pady=5)
        
        self.atualizar_lista_ambientes()

    def setup_subaba_status_instalacao(self, parent_notebook):
        """Configura a sub-aba de status de instalação"""
        frame_status = ttk.Frame(parent_notebook)
        parent_notebook.add(frame_status, text='Status de Instalação')
        
        # Frame para adicionar novo status
        frame_novo_status = ttk.LabelFrame(frame_status, text="Adicionar Novo Status")
        frame_novo_status.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_novo_status, text="Nome do Status:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_novo_status = ttk.Entry(frame_novo_status, width=30)
        self.entry_novo_status.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_novo_status, text="Adicionar",
                  command=self.adicionar_status).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de status
        frame_lista_status = ttk.LabelFrame(frame_status, text="Status Cadastrados")
        frame_lista_status.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.listbox_status = tk.Listbox(frame_lista_status)
        scrollbar_status = ttk.Scrollbar(frame_lista_status, orient='vertical', command=self.listbox_status.yview)
        self.listbox_status.configure(yscrollcommand=scrollbar_status.set)
        
        self.listbox_status.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar_status.pack(side='right', fill='y', pady=5)
        
        ttk.Button(frame_lista_status, text="Remover Selecionado",
                  command=self.remover_status).pack(pady=5)
        
        self.atualizar_lista_status()

    def setup_subaba_unidades(self, parent_notebook):
        """Configura a sub-aba de unidades"""
        frame_unidades = ttk.Frame(parent_notebook)
        parent_notebook.add(frame_unidades, text='Unidades')
        
        # Frame para adicionar nova unidade
        frame_nova_unidade = ttk.LabelFrame(frame_unidades, text="Adicionar Nova Unidade")
        frame_nova_unidade.pack(fill='x', padx=5, pady=5)
        
        ttk.Label(frame_nova_unidade, text="Sigla da Unidade:").grid(row=0, column=0, padx=5, pady=5)
        self.entry_nova_unidade = ttk.Entry(frame_nova_unidade, width=10)
        self.entry_nova_unidade.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(frame_nova_unidade, text="Adicionar",
                  command=self.adicionar_unidade).grid(row=1, column=0, columnspan=2, pady=10)
        
        # Lista de unidades
        frame_lista_unidades = ttk.LabelFrame(frame_unidades, text="Unidades Cadastradas")
        frame_lista_unidades.pack(fill='both', expand=True, padx=5, pady=5)
        
        self.listbox_unidades = tk.Listbox(frame_lista_unidades)
        scrollbar_unidades = ttk.Scrollbar(frame_lista_unidades, orient='vertical', command=self.listbox_unidades.yview)
        self.listbox_unidades.configure(yscrollcommand=scrollbar_unidades.set)
        
        self.listbox_unidades.pack(side='left', fill='both', expand=True, padx=5, pady=5)
        scrollbar_unidades.pack(side='right', fill='y', pady=5)
        
        ttk.Button(frame_lista_unidades, text="Remover Selecionada",
                  command=self.remover_unidade).pack(pady=5)
        
        self.atualizar_lista_unidades()

    # =================================
    # MÉTODOS PARA CATEGORIAS DE MATERIAIS
    # =================================
    
    def adicionar_categoria_material(self):
        """Adiciona uma nova categoria de material"""
        nome = self.entry_nova_cat_material.get().strip().upper()
        if not nome:
            messagebox.showerror("Erro", "Digite o nome da categoria!")
            return
        
        if nome in self.materiais_config['categorias_materiais']:
            messagebox.showerror("Erro", "Esta categoria já existe!")
            return
        
        # Adicionar nova categoria
        self.materiais_config['categorias_materiais'][nome] = {
            'subcategorias': []
        }
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_categorias_materiais()
        
        # Limpar campos
        self.entry_nova_cat_material.delete(0, tk.END)
               
        messagebox.showinfo("Sucesso", "Categoria adicionada com sucesso!")

    def on_categoria_material_select(self, event):
        """Evento de seleção de categoria de material"""
        selecionado = self.tree_cat_materiais.selection()
        if not selecionado:
            return
        
        categoria = self.tree_cat_materiais.item(selecionado)['values'][0]
        
        # Preencher campos de edição
        self.entry_editar_cat_material.delete(0, tk.END)
        self.entry_editar_cat_material.insert(0, categoria)
        
        # Atualizar lista de subcategorias
        self.atualizar_lista_subcategorias(categoria)

    def salvar_categoria_material(self):
        """Salva alterações na categoria selecionada"""
        selecionado = self.tree_cat_materiais.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma categoria para editar!")
            return
        
        categoria_antiga = self.tree_cat_materiais.item(selecionado)['values'][0]
        categoria_nova = self.entry_editar_cat_material.get().strip().upper()
        
        if not categoria_nova:
            messagebox.showerror("Erro", "Digite o nome da categoria!")
            return
        
        # Se mudou o nome, verificar se o novo nome já existe
        if categoria_nova != categoria_antiga and categoria_nova in self.materiais_config['categorias_materiais']:
            messagebox.showerror("Erro", "Já existe uma categoria com este nome!")
            return
        
        # Salvar dados da categoria
        dados_categoria = self.materiais_config['categorias_materiais'][categoria_antiga].copy()
        
        # Se mudou o nome, remover a antiga e adicionar a nova
        if categoria_nova != categoria_antiga:
            del self.materiais_config['categorias_materiais'][categoria_antiga]
        
        self.materiais_config['categorias_materiais'][categoria_nova] = dados_categoria
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_categorias_materiais()
        
        messagebox.showinfo("Sucesso", "Categoria atualizada com sucesso!")

    def remover_categoria_material(self):
        """Remove a categoria selecionada"""
        selecionado = self.tree_cat_materiais.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma categoria para remover!")
            return
        
        categoria = self.tree_cat_materiais.item(selecionado)['values'][0]
        
        if messagebox.askyesno("Confirmar", f"Deseja remover a categoria '{categoria}' e todas as suas subcategorias?"):
            del self.materiais_config['categorias_materiais'][categoria]
            self.salvar_configuracoes_materiais()
            self.atualizar_lista_categorias_materiais()
            
            # Limpar campos de edição
            self.entry_editar_cat_material.delete(0, tk.END)
            self.listbox_subcategorias.delete(0, tk.END)
            
            messagebox.showinfo("Sucesso", "Categoria removida com sucesso!")

    def adicionar_subcategoria(self):
        """Adiciona uma subcategoria à categoria selecionada"""
        selecionado = self.tree_cat_materiais.selection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma categoria primeiro!")
            return
        
        categoria = self.tree_cat_materiais.item(selecionado)['values'][0]
        subcategoria = self.entry_nova_subcategoria.get().strip().upper()
        
        if not subcategoria:
            messagebox.showerror("Erro", "Digite o nome da subcategoria!")
            return
        
        if subcategoria in self.materiais_config['categorias_materiais'][categoria]['subcategorias']:
            messagebox.showerror("Erro", "Esta subcategoria já existe!")
            return
        
        # Adicionar subcategoria
        self.materiais_config['categorias_materiais'][categoria]['subcategorias'].append(subcategoria)
        self.materiais_config['categorias_materiais'][categoria]['subcategorias'].sort()
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_subcategorias(categoria)
        
        # Limpar campo
        self.entry_nova_subcategoria.delete(0, tk.END)
        
        messagebox.showinfo("Sucesso", "Subcategoria adicionada com sucesso!")

    def remover_subcategoria(self):
        """Remove a subcategoria selecionada"""
        selecionado_cat = self.tree_cat_materiais.selection()
        if not selecionado_cat:
            messagebox.showwarning("Aviso", "Selecione uma categoria primeiro!")
            return
        
        selecionado_sub = self.listbox_subcategorias.curselection()
        if not selecionado_sub:
            messagebox.showwarning("Aviso", "Selecione uma subcategoria para remover!")
            return
        
        categoria = self.tree_cat_materiais.item(selecionado_cat)['values'][0]
        subcategoria = self.listbox_subcategorias.get(selecionado_sub[0])
        
        if messagebox.askyesno("Confirmar", f"Deseja remover a subcategoria '{subcategoria}'?"):
            self.materiais_config['categorias_materiais'][categoria]['subcategorias'].remove(subcategoria)
            self.salvar_configuracoes_materiais()
            self.atualizar_lista_subcategorias(categoria)
            
            messagebox.showinfo("Sucesso", "Subcategoria removida com sucesso!")

    def atualizar_lista_categorias_materiais(self):
        """Atualiza a exibição da lista de categorias de materiais"""
        for item in self.tree_cat_materiais.get_children():
            self.tree_cat_materiais.delete(item)
        
        for categoria, dados in sorted(self.materiais_config['categorias_materiais'].items()):
            self.tree_cat_materiais.insert('', 'end', values=(categoria))

    def atualizar_lista_subcategorias(self, categoria):
        """Atualiza a lista de subcategorias para a categoria selecionada"""
        self.listbox_subcategorias.delete(0, tk.END)
        
        subcategorias = self.materiais_config['categorias_materiais'][categoria]['subcategorias']
        for subcategoria in sorted(subcategorias):
            self.listbox_subcategorias.insert(tk.END, subcategoria)

    # =================================
    # MÉTODOS PARA AMBIENTES
    # =================================
    
    def adicionar_ambiente(self):
        """Adiciona um novo ambiente"""
        ambiente = self.entry_novo_ambiente.get().strip().upper()
        if not ambiente:
            messagebox.showerror("Erro", "Digite o nome do ambiente!")
            return
        
        if ambiente in self.materiais_config['ambientes']:
            messagebox.showerror("Erro", "Este ambiente já existe!")
            return
        
        self.materiais_config['ambientes'].append(ambiente)
        self.materiais_config['ambientes'].sort()
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_ambientes()
        
        self.entry_novo_ambiente.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Ambiente adicionado com sucesso!")

    def remover_ambiente(self):
        """Remove o ambiente selecionado"""
        selecionado = self.listbox_ambientes.curselection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um ambiente para remover!")
            return
        
        ambiente = self.listbox_ambientes.get(selecionado[0])
        
        if messagebox.askyesno("Confirmar", f"Deseja remover o ambiente '{ambiente}'?"):
            self.materiais_config['ambientes'].remove(ambiente)
            self.salvar_configuracoes_materiais()
            self.atualizar_lista_ambientes()
            
            messagebox.showinfo("Sucesso", "Ambiente removido com sucesso!")

    def atualizar_lista_ambientes(self):
        """Atualiza a exibição da lista de ambientes"""
        self.listbox_ambientes.delete(0, tk.END)
        for ambiente in sorted(self.materiais_config['ambientes']):
            self.listbox_ambientes.insert(tk.END, ambiente)

    # =================================
    # MÉTODOS PARA STATUS DE INSTALAÇÃO
    # =================================
    
    def adicionar_status(self):
        """Adiciona um novo status de instalação"""
        status = self.entry_novo_status.get().strip().upper()
        if not status:
            messagebox.showerror("Erro", "Digite o nome do status!")
            return
        
        if status in self.materiais_config['status_instalacao']:
            messagebox.showerror("Erro", "Este status já existe!")
            return
        
        self.materiais_config['status_instalacao'].append(status)
        self.materiais_config['status_instalacao'].sort()
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_status()
        
        self.entry_novo_status.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Status adicionado com sucesso!")

    def remover_status(self):
        """Remove o status selecionado"""
        selecionado = self.listbox_status.curselection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione um status para remover!")
            return
        
        status = self.listbox_status.get(selecionado[0])
        
        if messagebox.askyesno("Confirmar", f"Deseja remover o status '{status}'?"):
            self.materiais_config['status_instalacao'].remove(status)
            self.salvar_configuracoes_materiais()
            self.atualizar_lista_status()
            
            messagebox.showinfo("Sucesso", "Status removido com sucesso!")

    def atualizar_lista_status(self):
        """Atualiza a exibição da lista de status"""
        self.listbox_status.delete(0, tk.END)
        for status in sorted(self.materiais_config['status_instalacao']):
            self.listbox_status.insert(tk.END, status)

    # =================================
    # MÉTODOS PARA UNIDADES
    # =================================
    
    def adicionar_unidade(self):
        """Adiciona uma nova unidade"""
        unidade = self.entry_nova_unidade.get().strip().upper()
        if not unidade:
            messagebox.showerror("Erro", "Digite a sigla da unidade!")
            return
        
        if unidade in self.materiais_config['unidades']:
            messagebox.showerror("Erro", "Esta unidade já existe!")
            return
        
        self.materiais_config['unidades'].append(unidade)
        self.materiais_config['unidades'].sort()
        
        self.salvar_configuracoes_materiais()
        self.atualizar_lista_unidades()
        
        self.entry_nova_unidade.delete(0, tk.END)
        messagebox.showinfo("Sucesso", "Unidade adicionada com sucesso!")

    def remover_unidade(self):
        """Remove a unidade selecionada"""
        selecionado = self.listbox_unidades.curselection()
        if not selecionado:
            messagebox.showwarning("Aviso", "Selecione uma unidade para remover!")
            return
        
        unidade = self.listbox_unidades.get(selecionado[0])
        
        if messagebox.askyesno("Confirmar", f"Deseja remover a unidade '{unidade}'?"):
            self.materiais_config['unidades'].remove(unidade)
            self.salvar_configuracoes_materiais()
            self.atualizar_lista_unidades()
            
            messagebox.showinfo("Sucesso", "Unidade removida com sucesso!")

    def atualizar_lista_unidades(self):
        """Atualiza a exibição da lista de unidades"""
        self.listbox_unidades.delete(0, tk.END)
        for unidade in sorted(self.materiais_config['unidades']):
            self.listbox_unidades.insert(tk.END, unidade)

    def criar_aba_servicos_construcao(self):
        """Cria a aba de gerenciamento de serviços de construção"""
        from pathlib import Path
        import json
        
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Serviços de Construção")
        
        # Título
        ttk.Label(frame, text="Gerenciamento de Serviços de Construção", 
                font=('Arial', 14, 'bold')).grid(row=0, column=0, columnspan=2, pady=(0, 20))
        
        # Frame principal dividido em duas colunas
        main_container = ttk.Frame(frame)
        main_container.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # ===== COLUNA ESQUERDA: CATEGORIAS =====
        left_frame = ttk.LabelFrame(main_container, text="Categorias", padding="10")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Treeview de categorias
        cols_cat = ('ID', 'Nome', 'Qtd')
        self.tree_servicos_cat = ttk.Treeview(left_frame, columns=cols_cat, show='headings', height=15)
        
        self.tree_servicos_cat.heading('ID', text='ID')
        self.tree_servicos_cat.heading('Nome', text='Nome')
        self.tree_servicos_cat.heading('Qtd', text='Serviços')
        
        self.tree_servicos_cat.column('ID', width=120)
        self.tree_servicos_cat.column('Nome', width=180)
        self.tree_servicos_cat.column('Qtd', width=60)
        
        scrollbar_cat = ttk.Scrollbar(left_frame, orient=tk.VERTICAL, command=self.tree_servicos_cat.yview)
        self.tree_servicos_cat.configure(yscrollcommand=scrollbar_cat.set)
        
        self.tree_servicos_cat.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_cat.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Botões de categoria
        btn_frame_cat = ttk.Frame(left_frame)
        btn_frame_cat.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame_cat, text="Nova Categoria", 
                command=self.nova_categoria_servico).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame_cat, text="Editar", 
                command=self.editar_categoria_servico).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame_cat, text="Excluir", 
                command=self.excluir_categoria_servico).pack(side=tk.LEFT, padx=2)
        
        left_frame.columnconfigure(0, weight=1)
        left_frame.rowconfigure(0, weight=1)
        
        # ===== COLUNA DIREITA: SERVIÇOS =====
        right_frame = ttk.LabelFrame(main_container, text="Serviços", padding="10")
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(5, 0))
        
        # Treeview de serviços
        cols_serv = ('Nome', 'Ambientes')
        self.tree_servicos_list = ttk.Treeview(right_frame, columns=cols_serv, show='headings', height=15)
        
        self.tree_servicos_list.heading('Nome', text='Nome do Serviço')
        self.tree_servicos_list.heading('Ambientes', text='Ambientes')
        
        self.tree_servicos_list.column('Nome', width=250)
        self.tree_servicos_list.column('Ambientes', width=200)
        
        scrollbar_serv = ttk.Scrollbar(right_frame, orient=tk.VERTICAL, command=self.tree_servicos_list.yview)
        self.tree_servicos_list.configure(yscrollcommand=scrollbar_serv.set)
        
        self.tree_servicos_list.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_serv.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # Botões de serviços
        btn_frame_serv = ttk.Frame(right_frame)
        btn_frame_serv.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(btn_frame_serv, text="Novo Serviço", 
                command=self.novo_servico_construcao).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame_serv, text="Editar", 
                command=self.editar_servico_construcao).pack(side=tk.LEFT, padx=2)
        ttk.Button(btn_frame_serv, text="Excluir", 
                command=self.excluir_servico_construcao).pack(side=tk.LEFT, padx=2)
        
        right_frame.columnconfigure(0, weight=1)
        right_frame.rowconfigure(0, weight=1)
        
        # Configurar grid do container principal
        main_container.columnconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=2)
        main_container.rowconfigure(0, weight=1)
        
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)
        
        # Bind para seleção de categoria
        self.tree_servicos_cat.bind('<<TreeviewSelect>>', self.on_categoria_servico_selecionada)
        
        # Carregar dados iniciais
        self.carregar_servicos_config()
        self.atualizar_lista_categorias_servicos()

    # =================================
    # MÉTODOS PARA SERVIÇOS - CONTRATO
    # =================================

    def carregar_servicos_config(self):
        """Carrega configurações de serviços"""
        if self.SERVICOS_JSON_PATH.exists():
            try:
                with open(self.SERVICOS_JSON_PATH, 'r', encoding='utf-8') as f:
                    self.servicos_config = json.load(f)
            except Exception as e:
                logger.error(f"Erro ao carregar serviços: {e}")
                self.servicos_config = {"categorias": {}}
        else:
            self.servicos_config = {"categorias": {}}
        
        # Garantir estrutura
        self._garantir_estrutura_servicos()
    
    def _garantir_estrutura_servicos(self):
        """Garante estrutura completa dos serviços"""
        if 'categorias' not in self.servicos_config:
            self.servicos_config['categorias'] = {}
        
        for cat_id, cat_data in self.servicos_config['categorias'].items():
            if 'nome' not in cat_data:
                cat_data['nome'] = cat_id.title()
            
            if 'servicos' not in cat_data:
                cat_data['servicos'] = []
            
            # Converter serviços antigos (string) para novo formato (dict)
            servicos_atualizados = []
            for servico in cat_data['servicos']:
                if isinstance(servico, str):
                    servicos_atualizados.append({
                        'nome': servico,
                        'ambientes': [],
                        'descricao': ''
                    })
                elif isinstance(servico, dict):
                    if 'nome' not in servico:
                        continue
                    if 'ambientes' not in servico:
                        servico['ambientes'] = []
                    if 'descricao' not in servico:
                        servico['descricao'] = ''
                    servicos_atualizados.append(servico)
            
            cat_data['servicos'] = servicos_atualizados
    
    def salvar_servicos_config(self):
        """Salva configurações de serviços"""
        try:
            self.SERVICOS_JSON_PATH.parent.mkdir(parents=True, exist_ok=True)
            with open(self.SERVICOS_JSON_PATH, 'w', encoding='utf-8') as f:
                json.dump(self.servicos_config, f, indent=2, ensure_ascii=False)
            logger.info("Serviços salvos com sucesso")
            return True
        except Exception as e:
            logger.error(f"Erro ao salvar serviços: {e}")
            messagebox.showerror("Erro", f"Erro ao salvar serviços: {e}")
            return False
    
    @staticmethod
    def listar_todos_servicos():
        """Método estático para listar serviços - usado pelo combobox"""
        try:
            from src.config.config import BASE_PATH
            json_path = BASE_PATH / "servicos_construcao.json"
            
            if not json_path.exists():
                return []
            
            with open(json_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
            
            servicos = []
            for cat_data in config.get('categorias', {}).values():
                for servico in cat_data.get('servicos', []):
                    if isinstance(servico, dict):
                        servicos.append(servico['nome'])
                    elif isinstance(servico, str):
                        servicos.append(servico)
            
            return sorted(set(servicos))
        except Exception as e:
            print(f"Erro ao listar serviços: {e}")
            return []
    
    @staticmethod
    def adicionar_servico_rapido(nome_servico):
        """Adiciona serviço rapidamente - usado pelo combobox"""
        try:
            from src.config.config import BASE_PATH
            json_path = BASE_PATH / "servicos_construcao.json"
            
            # Carregar ou criar config
            if json_path.exists():
                with open(json_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
            else:
                config = {"categorias": {}}
            
            # Adicionar em "diversos" se não especificado
            if "diversos" not in config['categorias']:
                config['categorias']['diversos'] = {
                    'nome': 'Serviços Diversos',
                    'servicos': []
                }
            
            # Adicionar serviço
            novo_servico = {
                'nome': nome_servico,
                'descricao': '',
                'ambientes': []
            }
            
            config['categorias']['diversos']['servicos'].append(novo_servico)
            
            # Salvar
            json_path.parent.mkdir(parents=True, exist_ok=True)
            with open(json_path, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
            
            return True
        except Exception as e:
            print(f"Erro ao adicionar serviço: {e}")
            return False


# ==============================================================================
# 4. ADICIONE ESTA NOVA ABA NO MÉTODO criar_interface()
#    (Procure onde cria as outras abas e adicione esta junto)
# ==============================================================================

    def criar_aba_servicos_construcao(self):
        """Cria aba simplificada para gerenciar serviços"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Serviços")
        
        # Carregar dados
        self.carregar_servicos_config()
        
        ttk.Label(frame, text="Gerenciamento Rápido de Serviços", 
                 font=('Arial', 12, 'bold')).pack(pady=10)
        
        # Frame para adicionar
        add_frame = ttk.LabelFrame(frame, text="Adicionar Serviço", padding="10")
        add_frame.pack(fill=tk.X, padx=10, pady=10)
        
        ttk.Label(add_frame, text="Nome do Serviço:").grid(row=0, column=0, sticky=tk.W, padx=5)
        self.entry_servico = ttk.Entry(add_frame, width=50)
        self.entry_servico.grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Button(add_frame, text="➕ Adicionar", 
                  command=self.add_servico_simples).grid(row=0, column=2, padx=5)
        
        # Lista de serviços
        lista_frame = ttk.LabelFrame(frame, text="Serviços Cadastrados", padding="10")
        lista_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.listbox_servicos = tk.Listbox(lista_frame, height=15)
        scrollbar = ttk.Scrollbar(lista_frame, orient=tk.VERTICAL, 
                                  command=self.listbox_servicos.yview)
        self.listbox_servicos.config(yscrollcommand=scrollbar.set)
        
        self.listbox_servicos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botão remover
        ttk.Button(frame, text="➖ Remover Selecionado", 
                  command=self.remover_servico_simples).pack(pady=5)
        
        # Atualizar lista
        self.atualizar_lista_servicos_simples()
    
    def add_servico_simples(self):
        """Adiciona serviço simples"""
        nome = self.entry_servico.get().strip()
        if not nome:
            messagebox.showwarning("Aviso", "Digite o nome do serviço!")
            return
        
        # Verificar duplicado
        servicos_existentes = self.listar_todos_servicos()
        if nome in servicos_existentes:
            messagebox.showwarning("Aviso", "Este serviço já existe!")
            return
        
        # Adicionar
        if self.adicionar_servico_rapido(nome):
            self.entry_servico.delete(0, tk.END)
            self.carregar_servicos_config()
            self.atualizar_lista_servicos_simples()
            messagebox.showinfo("Sucesso", f"Serviço '{nome}' adicionado!")
        else:
            messagebox.showerror("Erro", "Não foi possível adicionar o serviço!")
    
    def remover_servico_simples(self):
        """Remove serviço selecionado"""
        sel = self.listbox_servicos.curselection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecione um serviço!")
            return
        
        nome_servico = self.listbox_servicos.get(sel[0])
        
        if not messagebox.askyesno("Confirmar", f"Remover '{nome_servico}'?"):
            return
        
        # Remover do JSON
        for cat_data in self.servicos_config['categorias'].values():
            cat_data['servicos'] = [
                s for s in cat_data['servicos'] 
                if s.get('nome') != nome_servico
            ]
        
        if self.salvar_servicos_config():
            self.atualizar_lista_servicos_simples()
            messagebox.showinfo("Sucesso", "Serviço removido!")
    
    def atualizar_lista_servicos_simples(self):
        """Atualiza lista de serviços"""
        self.listbox_servicos.delete(0, tk.END)
        for servico in sorted(self.listar_todos_servicos()):
            self.listbox_servicos.insert(tk.END, servico)

    def voltar_menu_local(self):  
        if hasattr(self, 'menu_principal') and self.menu_principal is not None:
            self.menu_principal.deiconify()  # Reexibe o menu principal
        self.root.destroy()  # Fecha a janela de configurações

    def run(self):
        """Inicia a execução do sistema de configurações"""
        self.root.mainloop()


if __name__ == "__main__":
    app = GerenciadorConfiguracoes()
    app.run()