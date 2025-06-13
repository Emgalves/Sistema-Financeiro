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
    
    # Cache de configurações para acesso rápido
    _config_cache = None
    
    @staticmethod
    def _atualizar_cache(config):
        """Atualiza o cache de configurações"""
        GerenciadorConfiguracoes._config_cache = config
    
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
    def _garantir_estrutura_completa(config):
        """Garante que arquivos antigos tenham a estrutura completa"""
        estrutura_padrao = GerenciadorConfiguracoes._obter_configuracoes_padrao_estaticas()
        
        # Adicionar seções faltantes
        for secao, valores_padrao in estrutura_padrao.items():
            if secao not in config:
                config[secao] = valores_padrao
                print(f"Adicionada seção faltante: {secao}")
            elif isinstance(valores_padrao, dict):
                # Verificar subseções
                for subsecao, sub_valores in valores_padrao.items():
                    if subsecao not in config[secao]:
                        config[secao][subsecao] = sub_valores
                        print(f"Adicionada subseção faltante: {secao}.{subsecao}")
        
        return config

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

    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Configurações do Sistema")
        self.root.geometry("800x600")
        
        # Usar o caminho da variável de classe 
        self.config_path = GerenciadorConfiguracoes.CONFIG_PATH
        
        # Carregar ou criar configurações iniciais
        self.carregar_configuracoes_locais()
        
        # Setup da interface
        self.setup_gui()

    def carregar_configuracoes_locais(self):
        """Carrega as configurações do sistema com suporte completo a correção monetária"""
        self.config = GerenciadorConfiguracoes.carregar_configuracoes()
        
        # Se não foi possível carregar, criar configurações padrão COMPLETAS
        if self.config is None:
            self.config = self._obter_configuracoes_padrao_completas()
            self.salvar_configuracoes()

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
            # NOVAS CONFIGURAÇÕES: Índices de correção monetária
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

    def setup_gui(self):
        """Configura a interface gráfica"""
        # Notebook para diferentes seções
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Abas
        self.setup_aba_cafe()
        self.setup_aba_bancos()
        self.setup_aba_categorias()
        self.setup_aba_indices_correcao()
        
        # Botões globais
        frame_botoes = ttk.Frame(self.root)
        frame_botoes.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(frame_botoes, text="Salvar Todas Alterações",
                  command=self.salvar_todas_alteracoes).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Voltar ao Menu Principal", 
                  command=self.voltar_menu_local).pack(side='left', padx=5, expand=True)
        ttk.Button(frame_botoes, text="Fechar",
                  command=self.root.destroy).pack(side='right', padx=5)

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