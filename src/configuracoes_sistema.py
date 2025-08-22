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
        self.root.geometry("800x600")
        
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
        self.setup_aba_etapas_obra()
        self.setup_aba_insumos()
        self.setup_aba_indices_correcao()
        self.setup_aba_materiais()
        
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