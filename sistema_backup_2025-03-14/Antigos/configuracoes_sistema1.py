import tkinter as tk
from tkinter import ttk, messagebox
import json
from datetime import datetime
from pathlib import Path
import os
from openpyxl import load_workbook, Workbook
from config.logger_config import system_logger, log_action

logger = system_logger.get_logger()

class GerenciadorConfiguracoes:
    @staticmethod
    @log_action("Carregar configurações")
    def carregar_configuracoes():
        """
        Método estático para carregar configurações do sistema
        """
        config_path = Path(__file__).parent / 'config' / 'parametros_sistema.json'
        
        if config_path.exists():
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                logger.error(f"Erro ao carregar configurações: {e}")
                return None
        
        return None

    def __init__(self, parent=None):
        self.root = tk.Toplevel(parent) if parent else tk.Tk()
        self.root.title("Configurações do Sistema")
        self.root.geometry("800x600")
        
        # Caminho para o arquivo de configurações
        self.config_path = Path(__file__).parent / 'config' / 'parametros_sistema.json'
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        
        # Carregar ou criar configurações iniciais
        self.carregar_configuracoes()
        
        # Setup da interface
        self.setup_gui()


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

    def carregar_configuracoes(self):
        """Carrega ou cria as configurações do sistema"""
        default_config = {
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
            }
        }
        
        try:
            if self.config_path.exists():
                with open(self.config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = default_config
                self.salvar_configuracoes()
        except Exception:
            self.config = default_config
            self.salvar_configuracoes()

    @log_action("Salvar configurações")
    def salvar_configuracoes(self):
        """Salva as configurações no arquivo"""
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(self.config, f, indent=4, ensure_ascii=False)

    def setup_gui(self):
        """Configura a interface gráfica"""
        # Notebook para diferentes seções
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Abas
        self.setup_aba_cafe()
        self.setup_aba_bancos()
        self.setup_aba_categorias()
        
        # Botões globais
        frame_botoes = ttk.Frame(self.root)
        frame_botoes.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(frame_botoes, text="Salvar Todas Alterações",
                  command=self.salvar_todas_alteracoes).pack(side='left', padx=5)
        ttk.Button(frame_botoes, text="Voltar ao Menu Principal", 
                  command=self.voltar_menu_local).pack(side='left', padx=5, expand=True)  # Centraliza)
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

    def salvar_todas_alteracoes(self):
        """Salva todas as alterações feitas nas configurações"""
        try:
            self.salvar_configuracoes()
            messagebox.showinfo("Sucesso", "Todas as alterações foram salvas com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar alterações: {str(e)}")

    def voltar_menu_local(self):  
        if hasattr(self, 'menu_principal') and self.menu_principal is not None:
            self.menu_principal.deiconify()  # Reexibe o menu principal
        self.root.destroy()  # Fecha a janela de configurações

    @staticmethod
    def get_configuracoes():
        """Método estático para obter as configurações atuais"""
        config_path = Path(__file__).parent / 'config' / 'parametros_sistema.json'
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Erro ao carregar configurações: {str(e)}")
            return None

    def run(self):
        """Inicia a execução do sistema de configurações"""
        self.root.mainloop()


if __name__ == "__main__":
    app = GerenciadorConfiguracoes()
    app.run()
