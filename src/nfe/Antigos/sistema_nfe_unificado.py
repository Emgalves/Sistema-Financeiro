# -*- coding: utf-8 -*-
"""
SISTEMA NFe UNIFICADO - VERSÃO FINAL
Combina processamento híbrido com integração perfeita ao sistema principal
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xml.etree.ElementTree as ET
from datetime import datetime, timedelta
from pathlib import Path
import re
import json


class SistemaNFeUnificado:
    """
    Sistema único que combina:
    - Processamento de XML local
    - Consulta por chave SEFAZ (futuro)
    - Integração perfeita com SistemaEntradaDados
    """
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.dados_nfe_atual = None
        
        print("🚀 Sistema NFe Unificado inicializado")
    
    def criar_interface_importacao(self):
        """Interface única e simplificada para importação"""
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("Importação NFe - Sistema Unificado")
        self.janela.geometry("800x600")
        self.janela.grab_set()
        
        # FRAME PRINCIPAL
        main_frame = ttk.Frame(self.janela)
        main_frame.pack(fill='both', expand=True, padx=20, pady=20)
        
        # TÍTULO
        titulo = tk.Label(main_frame, 
                         text="📄 IMPORTAÇÃO DE NOTA FISCAL ELETRÔNICA",
                         font=('Arial', 14, 'bold'))
        titulo.pack(pady=10)
        
        # SEÇÃO SELEÇÃO DE ARQUIVO
        self.criar_secao_arquivo(main_frame)
        
        # SEÇÃO DADOS EXTRAÍDOS (inicialmente oculta)
        self.frame_dados = ttk.LabelFrame(main_frame, text="Dados Extraídos", padding=10)
        
        # SEÇÃO OPÇÕES DE IMPORTAÇÃO (inicialmente oculta)
        self.frame_opcoes = ttk.LabelFrame(main_frame, text="Configurar Importação", padding=10)
        
        # BOTÕES PRINCIPAIS
        self.criar_botoes_principais(main_frame)
    
    def criar_secao_arquivo(self, parent):
        """Seção para seleção de arquivo XML"""
        frame_arquivo = ttk.LabelFrame(parent, text="1. Selecionar Arquivo XML", padding=10)
        frame_arquivo.pack(fill='x', pady=10)
        
        # STATUS DO ARQUIVO
        self.label_arquivo = tk.Label(frame_arquivo, 
                                     text="📁 Nenhum arquivo selecionado", 
                                     fg='gray', font=('Arial', 10))
        self.label_arquivo.pack(anchor='w', pady=5)
        
        # BOTÕES DE SELEÇÃO
        frame_btns = ttk.Frame(frame_arquivo)
        frame_btns.pack(fill='x', pady=5)
        
        ttk.Button(frame_btns, text="📁 Selecionar XML", 
                  command=self.selecionar_xml).pack(side='left', padx=5)
        
        ttk.Button(frame_btns, text="📧 Extrair de Email", 
                  command=self.extrair_de_email).pack(side='left', padx=5)
        
        # INFORMAÇÕES ÚTEIS
        info_text = """
💡 Dicas:
• Aceita XMLs salvos diretamente do email
• Processa automaticamente todos os dados da NFe
• Classifica produtos em categorias de material
• Cria lançamentos financeiros no formato do sistema
        """.strip()
        
        tk.Label(frame_arquivo, text=info_text, justify='left', 
                fg='blue', font=('Arial', 8)).pack(anchor='w', pady=5)
    
    def criar_botoes_principais(self, parent):
        """Botões principais da interface"""
        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill='x', pady=20)
        
        # BOTÃO PROCESSAR (inicialmente desabilitado)
        self.btn_processar = ttk.Button(frame_botoes, 
                                       text="2. 🔄 Processar XML", 
                                       command=self.processar_xml,
                                       state='disabled')
        self.btn_processar.pack(side='left', padx=5)
        
        # BOTÃO IMPORTAR (inicialmente desabilitado)
        self.btn_importar = ttk.Button(frame_botoes, 
                                      text="3. 📥 Importar para Sistema", 
                                      command=self.importar_para_sistema,
                                      state='disabled')
        self.btn_importar.pack(side='left', padx=5)
        
        # BOTÃO FECHAR
        ttk.Button(frame_botoes, text="❌ Fechar", 
                  command=self.janela.destroy).pack(side='right', padx=5)
    
    def selecionar_xml(self):
        """Seleciona arquivo XML"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar XML da NFe",
            filetypes=[
                ("Arquivos XML", "*.xml"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if arquivo:
            self.arquivo_xml_atual = arquivo
            nome_arquivo = Path(arquivo).name
            self.label_arquivo.config(
                text=f"✅ Selecionado: {nome_arquivo}", 
                fg='green'
            )
            self.btn_processar.config(state='normal')
    
    def extrair_de_email(self):
        """Placeholder para extração de email"""
        messagebox.showinfo("Em Desenvolvimento", 
            "Funcionalidade em desenvolvimento.\n\n"
            "Por enquanto:\n"
            "1. Abra o email com a NFe\n"
            "2. Baixe o anexo XML\n"
            "3. Use 'Selecionar XML'")
    
    def processar_xml(self):
        """Processa o XML selecionado"""
        try:
            if not hasattr(self, 'arquivo_xml_atual'):
                messagebox.showerror("Erro", "Selecione um arquivo XML primeiro!")
                return
            
            # MOSTRAR PROGRESSO
            self.label_arquivo.config(text="🔄 Processando XML...", fg='blue')
            self.janela.update()
            
            # PROCESSAR XML
            self.dados_nfe_atual = self.processar_xml_nfe(self.arquivo_xml_atual)
            
            if self.dados_nfe_atual:
                # MOSTRAR DADOS EXTRAÍDOS
                self.exibir_dados_extraidos()
                
                # MOSTRAR OPÇÕES
                self.criar_opcoes_importacao()
                
                # HABILITAR IMPORTAÇÃO
                self.btn_importar.config(state='normal')
                
                # ATUALIZAR STATUS
                self.label_arquivo.config(
                    text=f"✅ XML processado com sucesso!", 
                    fg='green'
                )
            else:
                self.label_arquivo.config(
                    text="❌ Erro ao processar XML", 
                    fg='red'
                )
                
        except Exception as e:
            self.label_arquivo.config(text="❌ Erro no processamento", fg='red')
            messagebox.showerror("Erro", f"Erro ao processar XML:\n{str(e)}")
    
    def processar_xml_nfe(self, caminho_arquivo):
        """Processa arquivo XML da NFe"""
        try:
            print(f"📄 Processando XML: {caminho_arquivo}")
            
            # CARREGAR XML
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # NAMESPACE NFE
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # ESTRUTURA BÁSICA
            dados = {
                'fonte_dados': 'XML Local',
                'arquivo_origem': caminho_arquivo,
                'chave_acesso': '',
                'numero_nf': '',
                'serie': '1',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'cnpj_emitente': '',
                'razao_social_emitente': '',
                'valor_total': 0.0,
                'valor_produtos': 0.0,
                'produtos': []
            }
            
            # BUSCAR DADOS NO XML
            inf_nfe = root.find('.//nfe:infNFe', ns)
            if inf_nfe is None:
                raise Exception("Estrutura XML inválida - infNFe não encontrado")
            
            # CHAVE DE ACESSO
            dados['chave_acesso'] = inf_nfe.get('Id', '').replace('NFe', '')
            
            # DADOS IDENTIFICAÇÃO
            ide = inf_nfe.find('nfe:ide', ns)
            if ide is not None:
                dados['numero_nf'] = self.get_xml_text(ide.find('nfe:nNF', ns))
                dados['serie'] = self.get_xml_text(ide.find('nfe:serie', ns))
                
                # DATA DE EMISSÃO
                dh_emi = self.get_xml_text(ide.find('nfe:dhEmi', ns))
                if dh_emi:
                    dados['data_emissao'] = self.formatar_data_xml(dh_emi)
            
            # DADOS DO EMITENTE
            emit = inf_nfe.find('nfe:emit', ns)
            if emit is not None:
                dados['cnpj_emitente'] = self.get_xml_text(emit.find('nfe:CNPJ', ns))
                dados['razao_social_emitente'] = self.get_xml_text(emit.find('nfe:xNome', ns))
            
            # TOTAIS
            total = inf_nfe.find('.//nfe:total/nfe:ICMSTot', ns)
            if total is not None:
                dados['valor_total'] = float(self.get_xml_text(total.find('nfe:vNF', ns)) or 0)
                dados['valor_produtos'] = float(self.get_xml_text(total.find('nfe:vProd', ns)) or 0)
            
            # PRODUTOS
            dados['produtos'] = self.extrair_produtos_xml(inf_nfe, ns)
            
            print(f"✅ XML processado com sucesso!")
            print(f"📄 NFe {dados['numero_nf']} - {dados['razao_social_emitente']}")
            print(f"💰 R$ {dados['valor_total']:,.2f} - {len(dados['produtos'])} produtos")
            
            return dados
            
        except Exception as e:
            print(f"❌ Erro ao processar XML: {e}")
            raise e
    
    def extrair_produtos_xml(self, inf_nfe, ns):
        """Extrai produtos do XML"""
        produtos = []
        
        try:
            itens = inf_nfe.findall('nfe:det', ns)
            
            for item in itens:
                prod = item.find('nfe:prod', ns)
                if prod is None:
                    continue
                
                produto = {
                    'numero_item': item.get('nItem', ''),
                    'codigo': self.get_xml_text(prod.find('nfe:cProd', ns)),
                    'descricao': self.get_xml_text(prod.find('nfe:xProd', ns)),
                    'ncm': self.get_xml_text(prod.find('nfe:NCM', ns)),
                    'cfop': self.get_xml_text(prod.find('nfe:CFOP', ns)),
                    'unidade': self.get_xml_text(prod.find('nfe:uCom', ns)),
                    'quantidade': float(self.get_xml_text(prod.find('nfe:qCom', ns)) or 0),
                    'valor_unitario': float(self.get_xml_text(prod.find('nfe:vUnCom', ns)) or 0),
                    'valor_total': float(self.get_xml_text(prod.find('nfe:vProd', ns)) or 0)
                }
                
                # CLASSIFICAÇÃO AUTOMÁTICA
                produto['categoria_sugerida'] = self.classificar_produto(produto['descricao'])
                produto['subcategoria_sugerida'] = self.sugerir_subcategoria(produto['descricao'])
                
                produtos.append(produto)
                
        except Exception as e:
            print(f"⚠️ Erro ao extrair produtos: {e}")
        
        return produtos
    
    def exibir_dados_extraidos(self):
        """Exibe dados extraídos da NFe"""
        self.frame_dados.pack(fill='x', pady=10)
        
        # LIMPAR CONTEÚDO ANTERIOR
        for widget in self.frame_dados.winfo_children():
            widget.destroy()
        
        dados = self.dados_nfe_atual
        
        # INFORMAÇÕES PRINCIPAIS
        info_frame = ttk.Frame(self.frame_dados)
        info_frame.pack(fill='x')
        
        # PRIMEIRA LINHA
        linha1 = ttk.Frame(info_frame)
        linha1.pack(fill='x', pady=2)
        
        tk.Label(linha1, text="📄 NFe:", font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(linha1, text=f"{dados.get('numero_nf', '')} (Série {dados.get('serie', '')})").pack(side='left', padx=5)
        
        tk.Label(linha1, text="📅 Data:", font=('Arial', 9, 'bold')).pack(side='left', padx=(20,0))
        tk.Label(linha1, text=dados.get('data_emissao', '')).pack(side='left', padx=5)
        
        # SEGUNDA LINHA
        linha2 = ttk.Frame(info_frame)
        linha2.pack(fill='x', pady=2)
        
        tk.Label(linha2, text="🏢 Fornecedor:", font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(linha2, text=dados.get('razao_social_emitente', '')[:50]).pack(side='left', padx=5)
        
        # TERCEIRA LINHA
        linha3 = ttk.Frame(info_frame)
        linha3.pack(fill='x', pady=2)
        
        tk.Label(linha3, text="💰 Valor Total:", font=('Arial', 9, 'bold')).pack(side='left')
        tk.Label(linha3, text=f"R$ {dados.get('valor_total', 0):,.2f}").pack(side='left', padx=5)
        
        tk.Label(linha3, text="📦 Produtos:", font=('Arial', 9, 'bold')).pack(side='left', padx=(20,0))
        tk.Label(linha3, text=str(len(dados.get('produtos', [])))).pack(side='left', padx=5)
        
        # BOTÃO PARA VER PRODUTOS
        if dados.get('produtos'):
            ttk.Button(linha3, text="👁️ Ver Produtos", 
                      command=self.mostrar_produtos).pack(side='right', padx=5)
    
    def mostrar_produtos(self):
        """Mostra lista detalhada de produtos"""
        janela_produtos = tk.Toplevel(self.janela)
        janela_produtos.title("Produtos da NFe")
        janela_produtos.geometry("900x500")
        janela_produtos.grab_set()
        
        # FRAME PRINCIPAL
        frame_main = ttk.Frame(janela_produtos)
        frame_main.pack(fill='both', expand=True, padx=10, pady=10)
        
        # TREEVIEW
        colunas = ('Item', 'Código', 'Descrição', 'Categoria', 'Qtd', 'Un', 'Vl Unit', 'Total')
        tree = ttk.Treeview(frame_main, columns=colunas, show='headings', height=15)
        
        # CONFIGURAR COLUNAS
        tree.heading('Item', text='#')
        tree.heading('Código', text='Código')
        tree.heading('Descrição', text='Descrição')
        tree.heading('Categoria', text='Categoria Sugerida')
        tree.heading('Qtd', text='Qtd')
        tree.heading('Un', text='Un')
        tree.heading('Vl Unit', text='Vl Unit')
        tree.heading('Total', text='Total')
        
        tree.column('Item', width=40)
        tree.column('Código', width=80)
        tree.column('Descrição', width=250)
        tree.column('Categoria', width=120)
        tree.column('Qtd', width=60)
        tree.column('Un', width=40)
        tree.column('Vl Unit', width=80)
        tree.column('Total', width=80)
        
        # PREENCHER DADOS
        produtos = self.dados_nfe_atual.get('produtos', [])
        for i, produto in enumerate(produtos, 1):
            tree.insert('', 'end', values=(
                i,
                produto.get('codigo', '')[:12],
                produto.get('descricao', '')[:35],
                produto.get('categoria_sugerida', ''),
                produto.get('quantidade', ''),
                produto.get('unidade', ''),
                f"R$ {produto.get('valor_unitario', 0):.2f}",
                f"R$ {produto.get('valor_total', 0):.2f}"
            ))
        
        # SCROLLBAR
        scrollbar = ttk.Scrollbar(frame_main, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # BOTÃO FECHAR
        ttk.Button(janela_produtos, text="Fechar", 
                  command=janela_produtos.destroy).pack(pady=10)
    
    def criar_opcoes_importacao(self):
        """Cria opções de importação"""
        self.frame_opcoes.pack(fill='x', pady=10)
        
        # LIMPAR CONTEÚDO ANTERIOR
        for widget in self.frame_opcoes.winfo_children():
            widget.destroy()
        
        # VARIÁVEIS DE CONTROLE
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        # TÍTULO
        tk.Label(self.frame_opcoes, text="Selecione o que importar:", 
                font=('Arial', 10, 'bold')).pack(anchor='w', pady=5)
        
        # CHECKBOXES PRINCIPAIS
        frame_checks = ttk.Frame(self.frame_opcoes)
        frame_checks.pack(fill='x', pady=5)
        
        cb_financeiro = tk.Checkbutton(
            frame_checks, 
            text="💰 Dados Financeiros (despesa no sistema)", 
            variable=self.importar_financeiro,
            font=('Arial', 10),
            command=self.toggle_opcoes_financeiro
        )
        cb_financeiro.pack(anchor='w', pady=2)
        
        cb_materiais = tk.Checkbutton(
            frame_checks, 
            text="📦 Materiais da Obra (banco de dados para manual)", 
            variable=self.importar_materiais,
            font=('Arial', 10),
            command=self.toggle_opcoes_materiais
        )
        cb_materiais.pack(anchor='w', pady=2)
        
        # OPÇÕES FINANCEIRAS
        self.frame_opcoes_financeiro = ttk.LabelFrame(self.frame_opcoes, 
                                                     text="Configurações Financeiras", 
                                                     padding=10)
        self.frame_opcoes_financeiro.pack(fill='x', pady=5)
        self.criar_opcoes_financeiro()
        
        # OPÇÕES MATERIAIS
        self.frame_opcoes_materiais = ttk.LabelFrame(self.frame_opcoes, 
                                                    text="Configurações Materiais", 
                                                    padding=10)
        self.frame_opcoes_materiais.pack(fill='x', pady=5)
        self.criar_opcoes_materiais()
    
    def criar_opcoes_financeiro(self):
        """Cria opções específicas para dados financeiros"""
        # PRIMEIRA LINHA
        linha1 = ttk.Frame(self.frame_opcoes_financeiro)
        linha1.pack(fill='x', pady=2)
        
        tk.Label(linha1, text="Tipo Despesa:").pack(side='left')
        self.tipo_despesa = ttk.Combobox(linha1, width=15, state='readonly')
        self.tipo_despesa['values'] = ['1', '2', '3', '4', '5', '6', '7']
        self.tipo_despesa.set('3')  # Material
        self.tipo_despesa.pack(side='left', padx=5)
        
        tk.Label(linha1, text="Categoria:").pack(side='left', padx=(20,0))
        self.categoria_financeira = tk.Entry(linha1, width=10)
        self.categoria_financeira.insert(0, 'MAT')
        self.categoria_financeira.pack(side='left', padx=5)
        
        # SEGUNDA LINHA
        linha2 = ttk.Frame(self.frame_opcoes_financeiro)
        linha2.pack(fill='x', pady=2)
        
        tk.Label(linha2, text="Etapa Obra:").pack(side='left')
        self.etapa_obra = tk.Entry(linha2, width=20)
        self.etapa_obra.insert(0, 'MATERIAIS')
        self.etapa_obra.pack(side='left', padx=5)
        
        tk.Label(linha2, text="Forma Pgto:").pack(side='left', padx=(20,0))
        self.forma_pagamento = ttk.Combobox(linha2, width=15, state='readonly')
        self.forma_pagamento['values'] = ['A_VISTA', 'A_PRAZO', 'CARTAO', 'PIX']
        self.forma_pagamento.set('A_PRAZO')
        self.forma_pagamento.pack(side='left', padx=5)
    
    def criar_opcoes_materiais(self):
        """Cria opções específicas para materiais"""
        # PRIMEIRA LINHA
        linha1 = ttk.Frame(self.frame_opcoes_materiais)
        linha1.pack(fill='x', pady=2)
        
        tk.Label(linha1, text="Ambiente Padrão:").pack(side='left')
        self.ambiente_padrao = ttk.Combobox(linha1, width=25, state='readonly')
        
        # CARREGAR AMBIENTES DO SISTEMA DE MATERIAIS
        ambientes = ['', 'GERAL', 'INSTALAÇÃO DA OBRA', 'MATERIAIS']
        try:
            if hasattr(self.sistema, 'gerenciador_materiais'):
                ambientes = self.sistema.gerenciador_materiais.parametros.get('ambientes', ambientes)
        except:
            pass
        
        self.ambiente_padrao['values'] = ambientes
        self.ambiente_padrao.pack(side='left', padx=5)
        
        tk.Label(linha1, text="Garantia:").pack(side='left', padx=(20,0))
        self.garantia_meses = tk.Entry(linha1, width=5)
        self.garantia_meses.insert(0, '12')
        self.garantia_meses.pack(side='left', padx=5)
        tk.Label(linha1, text="meses").pack(side='left')
    
    def toggle_opcoes_financeiro(self):
        """Habilita/desabilita opções financeiras"""
        estado = 'normal' if self.importar_financeiro.get() else 'disabled'
        
        for widget in self.frame_opcoes_financeiro.winfo_children():
            if isinstance(widget, ttk.Frame):
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, (tk.Entry, ttk.Combobox)):
                        subwidget.config(state=estado)
    
    def toggle_opcoes_materiais(self):
        """Habilita/desabilita opções de materiais"""
        estado = 'normal' if self.importar_materiais.get() else 'disabled'
        
        for widget in self.frame_opcoes_materiais.winfo_children():
            if isinstance(widget, ttk.Frame):
                for subwidget in widget.winfo_children():
                    if isinstance(subwidget, (tk.Entry, ttk.Combobox)):
                        subwidget.config(state=estado)
    
    def importar_para_sistema(self):
        """Importa dados para o sistema principal"""
        try:
            if not self.dados_nfe_atual:
                messagebox.showerror("Erro", "Nenhum dado carregado!")
                return
            
            # VALIDAR SELEÇÕES
            if not self.importar_financeiro.get() and not self.importar_materiais.get():
                messagebox.showwarning("Aviso", "Selecione pelo menos uma opção!")
                return
            
            resultados = []
            
            # IMPORTAR DADOS FINANCEIROS
            if self.importar_financeiro.get():
                resultado = self.criar_lancamento_financeiro()
                resultados.append(f"💰 Financeiro: {resultado}")
            
            # IMPORTAR MATERIAIS
            if self.importar_materiais.get():
                resultado = self.criar_materiais_obra()
                resultados.append(f"📦 Materiais: {resultado}")
            
            # MOSTRAR RESULTADO
            self.mostrar_resultado_importacao(resultados)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na importação:\n{str(e)}")
    
    def criar_lancamento_financeiro(self):
        """Cria lançamento financeiro no formato do sistema"""
        try:
            dados_nfe = self.dados_nfe_atual
            
            # MAPEAR PARA FORMATO DO SISTEMA
            dados_financeiros = {
                'data': dados_nfe.get('data_emissao', ''),
                'cnpj_cpf': re.sub(r'[^0-9]', '', dados_nfe.get('cnpj_emitente', '')),
                'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                'categoria': self.categoria_financeira.get().upper(),
                'tp_desp': self.tipo_despesa.get(),
                'referencia': f"NFE {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')[:20]}".upper(),
                'etapa_obra': self.etapa_obra.get().upper(),
                'nf': dados_nfe.get('numero_nf', ''),
                'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dias': 1,
                'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dt_vencto': dados_nfe.get('data_emissao', ''),
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - CHAVE: {dados_nfe.get('chave_acesso', '')[:20]}...".upper(),
                'forma_pagamento': self.forma_pagamento.get()
            }
            
            # ADICIONAR À LISTA DO SISTEMA
            if not hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir = []
            
            self.sistema.dados_para_incluir.append(dados_financeiros)
            
            return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            
        except Exception as e:
            raise Exception(f"Erro ao criar lançamento: {str(e)}")
    
    def criar_materiais_obra(self):
        """Cria materiais da obra"""
        try:
            produtos = self.dados_nfe_atual.get('produtos', [])
            if not produtos:
                return "Nenhum produto encontrado"
            
            # VERIFICAR SISTEMA DE MATERIAIS
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                return f"{len(produtos)} produtos (sistema materiais não inicializado)"
            
            materiais_criados = 0
            dados_nfe = self.dados_nfe_atual
            
            for produto in produtos:
                try:
                    # DADOS DO MATERIAL
                    dados_material = {
                        'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                        'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                        'Subcategoria': produto.get('subcategoria_sugerida', ''),
                        'Codigo_Produto': produto.get('codigo', ''),
                        'Descricao_Completa': produto.get('descricao', ''),
                        'Marca': dados_nfe.get('razao_social_emitente', '')[:20],
                        'Modelo': '',
                        'Cor_Acabamento': '',
                        'Dimensoes': '',
                        'Especificacoes_Tecnicas': self.gerar_especificacoes(produto),
                        'Ambiente_Aplicacao': self.ambiente_padrao.get(),
                        'Localizacao_Especifica': '',
                        'Data_Instalacao': '',
                        'Instalador': '',
                        'Status_Instalacao': 'PENDENTE',
                        'Garantia_Meses': int(self.garantia_meses.get() or 12),
                        'Observacoes': f"Importado NFe {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('razao_social_emitente', '')}",
                        
                        # DADOS DE COMPRA
                        'Tem_Dados_Compra': True,
                        'Data_Compra': dados_nfe.get('data_emissao', ''),
                        'CNPJ_Fornecedor': dados_nfe.get('cnpj_emitente', ''),
                        'Nome_Fornecedor': dados_nfe.get('razao_social_emitente', ''),
                        'Numero_NF': dados_nfe.get('numero_nf', ''),
                        'Item_NF': produto.get('numero_item', ''),
                        'Quantidade': produto.get('quantidade', 1),
                        'Unidade': produto.get('unidade', 'UN'),
                        'Valor_Unitario': produto.get('valor_unitario', 0),
                        'Valor_Total': produto.get('valor_total', 0),
                        'Origem_Dados': 'IMPORTACAO_NFE'
                    }
                    
                    # SALVAR MATERIAL
                    material_id = self.sistema.gerenciador_materiais.salvar_material(dados_material)
                    materiais_criados += 1
                    
                except Exception as e:
                    print(f"⚠️ Erro ao criar material {produto.get('descricao', '')}: {e}")
                    continue
            
            return f"{materiais_criados} de {len(produtos)} produtos"
            
        except Exception as e:
            raise Exception(f"Erro ao criar materiais: {str(e)}")
    
    def mostrar_resultado_importacao(self, resultados):
        """Mostra resultado da importação"""
        janela_resultado = tk.Toplevel(self.janela)
        janela_resultado.title("Importação Concluída")
        janela_resultado.geometry("500x400")
        janela_resultado.grab_set()
        
        # FRAME PRINCIPAL
        frame_main = ttk.Frame(janela_resultado)
        frame_main.pack(fill='both', expand=True, padx=20, pady=20)
        
        # TÍTULO
        titulo = tk.Label(frame_main, 
                         text="✅ IMPORTAÇÃO CONCLUÍDA", 
                         font=('Arial', 14, 'bold'),
                         fg='green')
        titulo.pack(pady=10)
        
        # RESUMO NFE
        dados_nfe = self.dados_nfe_atual
        resumo_text = f"""
📄 NFe: {dados_nfe.get('numero_nf', '')}
🏢 Fornecedor: {dados_nfe.get('razao_social_emitente', '')}
📅 Data: {dados_nfe.get('data_emissao', '')}
💰 Valor: R$ {dados_nfe.get('valor_total', 0):,.2f}
📦 Produtos: {len(dados_nfe.get('produtos', []))}
        """.strip()
        
        frame_resumo = ttk.LabelFrame(frame_main, text="NFe Processada", padding=10)
        frame_resumo.pack(fill='x', pady=5)
        tk.Label(frame_resumo, text=resumo_text, justify='left').pack(anchor='w')
        
        # RESULTADOS
        frame_resultados = ttk.LabelFrame(frame_main, text="Dados Importados", padding=10)
        frame_resultados.pack(fill='x', pady=5)
        
        for resultado in resultados:
            tk.Label(frame_resultados, text=resultado, fg='blue', 
                    font=('Arial', 10)).pack(anchor='w', pady=1)
        
        # PRÓXIMOS PASSOS
        frame_passos = ttk.LabelFrame(frame_main, text="Próximos Passos", padding=10)
        frame_passos.pack(fill='x', pady=5)
        
        passos_text = """
1. ✅ Dados foram adicionados ao sistema
2. 📊 Use 'Enviar Dados' para salvar na planilha
3. 📦 Confira materiais em 'Consultar Materiais'
4. 📄 Gere 'Manual do Proprietário' quando pronto
        """.strip()
        
        tk.Label(frame_passos, text=passos_text, justify='left').pack(anchor='w')
        
        # BOTÕES
        frame_botoes = ttk.Frame(frame_main)
        frame_botoes.pack(fill='x', pady=10)
        
        ttk.Button(frame_botoes, 
                  text="📊 Processar no Sistema", 
                  command=lambda: self.processar_no_sistema(janela_resultado)).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="👁️ Ver Dados", 
                  command=self.visualizar_dados_sistema).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ Fechar", 
                  command=janela_resultado.destroy).pack(side='right', padx=5)
    
    def processar_no_sistema(self, janela_resultado):
        """Chama enviar_dados() do sistema principal"""
        try:
            janela_resultado.destroy()
            self.janela.destroy()
            
            # VERIFICAR SE HÁ DADOS
            if not hasattr(self.sistema, 'dados_para_incluir') or not self.sistema.dados_para_incluir:
                messagebox.showwarning("Aviso", "Não há dados para processar!")
                return
            
            # CHAMAR MÉTODO DO SISTEMA
            self.sistema.enviar_dados()
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar: {str(e)}")
    
    def visualizar_dados_sistema(self):
        """Abre visualizador de dados"""
        try:
            if hasattr(self.sistema, 'visualizar_dados'):
                self.sistema.visualizar_dados()
            else:
                messagebox.showinfo("Info", "Visualizador não disponível")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {str(e)}")
    
    # MÉTODOS UTILITÁRIOS
    
    def get_xml_text(self, element):
        """Extrai texto de elemento XML"""
        return element.text if element is not None else ''
    
    def formatar_data_xml(self, data_str):
        """Formata data do XML para dd/mm/yyyy"""
        try:
            if 'T' in data_str:
                dt = datetime.fromisoformat(data_str.replace('Z', '+00:00'))
            else:
                dt = datetime.strptime(data_str, '%Y-%m-%d')
            return dt.strftime('%d/%m/%Y')
        except:
            return data_str
    
    def classificar_produto(self, descricao):
        """Classifica produto em categoria"""
        if not descricao:
            return 'OUTROS'
        
        desc_upper = descricao.upper()
        
        categorias = {
            'ACABAMENTOS': [
                'CERAMICA', 'PORCELANATO', 'AZULEJO', 'PASTILHA', 'REVESTIMENTO', 'PISO',
                'RODAPE', 'MOLDURA', 'REJUNTE', 'GESSO', 'FORRO', 'LAMINADO'
            ],
            'TINTAS': [
                'TINTA', 'VERNIZ', 'ESMALTE', 'PRIMER', 'SELADOR', 'MASSA CORRIDA',
                'TEXTURA', 'RESINA'
            ],
            'ELETRICO': [
                'FIO', 'CABO', 'TOMADA', 'INTERRUPTOR', 'LAMPADA', 'LED', 'ELETRICO',
                'DISJUNTOR', 'QUADRO', 'CONDULETE', 'ELETRODUTO'
            ],
            'HIDRAULICO': [
                'TUBO', 'CONEXAO', 'REGISTRO', 'TORNEIRA', 'VALVULA', 'HIDRAULICO',
                'CANO', 'CHUVEIRO', 'VASO SANITARIO', 'PIA', 'SIFAO'
            ],
            'ESTRUTURAL': [
                'CIMENTO', 'CONCRETO', 'FERRO', 'AÇO', 'TIJOLO', 'BLOCO', 'VIGA',
                'AREIA', 'BRITA', 'CAL', 'ARGAMASSA'
            ],
            'ESQUADRIAS': [
                'PORTA', 'JANELA', 'FECHADURA', 'DOBRADIÇA', 'VIDRO', 'MARCO',
                'FERRAGEM', 'TRINCO', 'BATENTE'
            ],
            'FERRAGENS': [
                'PARAFUSO', 'PREGO', 'BUCHA', 'CHAVE', 'CADEADO', 'REBITE',
                'PORCA', 'ARRUELA', 'GANCHO'
            ]
        }
        
        for categoria, palavras_chave in categorias.items():
            if any(palavra in desc_upper for palavra in palavras_chave):
                return categoria
        
        return 'OUTROS'
    
    def sugerir_subcategoria(self, descricao):
        """Sugere subcategoria baseada na descrição"""
        if not descricao:
            return ''
        
        desc_upper = descricao.upper()
        
        subcategorias = {
            'CERAMICA': ['CERAMICA'],
            'PORCELANATO': ['PORCELANATO'],
            'PISO_LAMINADO': ['LAMINADO'],
            'RODAPE': ['RODAPE'],
            'REJUNTE': ['REJUNTE'],
            'GESSO': ['GESSO', 'FORRO'],
            'CABOS': ['FIO', 'CABO'],
            'TOMADAS': ['TOMADA'],
            'INTERRUPTORES': ['INTERRUPTOR'],
            'ILUMINACAO': ['LAMPADA', 'LED'],
            'PROTECAO': ['DISJUNTOR'],
            'TUBULACAO': ['TUBO', 'CANO'],
            'CONEXOES': ['CONEXAO', 'JOELHO', 'TE'],
            'METAIS': ['TORNEIRA', 'REGISTRO', 'CHUVEIRO'],
            'LOUÇAS': ['VASO', 'PIA', 'TANQUE']
        }
        
        for subcategoria, palavras in subcategorias.items():
            if any(palavra in desc_upper for palavra in palavras):
                return subcategoria
        
        return ''
    
    def gerar_especificacoes(self, produto):
        """Gera especificações técnicas do produto"""
        specs = []
        
        if produto.get('ncm'):
            specs.append(f"NCM: {produto['ncm']}")
        
        if produto.get('cfop'):
            specs.append(f"CFOP: {produto['cfop']}")
        
        if produto.get('unidade'):
            specs.append(f"Unidade: {produto['unidade']}")
        
        return " | ".join(specs)


# FUNÇÃO PARA SUBSTITUIR OS SISTEMAS DUPLICADOS
def substituir_sistemas_nfe_por_unificado(sistema_principal):
    """
    Remove sistemas duplicados e instala o sistema unificado
    """
    try:
        print("🔄 Substituindo sistemas NFe por versão unificada...")
        
        # REMOVER REFERÊNCIAS ANTIGAS
        if hasattr(sistema_principal, 'processador_nfe'):
            delattr(sistema_principal, 'processador_nfe')
        
        if hasattr(sistema_principal, 'integrador_nfe'):
            delattr(sistema_principal, 'integrador_nfe')
        
        if hasattr(sistema_principal, 'importar_nfe_xml'):
            delattr(sistema_principal, 'importar_nfe_xml')
        
        if hasattr(sistema_principal, 'importar_nfe_com_interface'):
            delattr(sistema_principal, 'importar_nfe_com_interface')
        
        # INSTALAR SISTEMA UNIFICADO
        sistema_nfe = SistemaNFeUnificado(sistema_principal)
        sistema_principal.sistema_nfe_unificado = sistema_nfe
        
        # MÉTODO DE CONVENIÊNCIA
        def abrir_importacao_nfe():
            """Abre interface de importação NFe"""
            sistema_nfe.criar_interface_importacao()
        
        sistema_principal.abrir_importacao_nfe = abrir_importacao_nfe
        
        # ADICIONAR BOTÃO ÚNICO NA INTERFACE
        adicionar_botao_unificado_na_interface(sistema_principal, abrir_importacao_nfe)
        
        print("✅ Sistema NFe unificado instalado!")
        print("📌 Método disponível: sistema.abrir_importacao_nfe()")
        
        return sistema_nfe
        
    except Exception as e:
        print(f"❌ Erro ao instalar sistema unificado: {e}")
        return None


def adicionar_botao_unificado_na_interface(sistema, callback_importar):
    """
    Adiciona APENAS UM botão na interface, removendo duplicações
    """
    try:
        if not hasattr(sistema, 'aba_fornecedor'):
            return
        
        # REMOVER BOTÕES/SEÇÕES NFE EXISTENTES
        widgets_para_remover = []
        for widget in sistema.aba_fornecedor.winfo_children():
            if isinstance(widget, ttk.LabelFrame):
                texto = widget['text']
                if any(palavra in texto.lower() for palavra in ['nfe', 'nf-e', 'importação nfe']):
                    widgets_para_remover.append(widget)
        
        for widget in widgets_para_remover:
            widget.destroy()
        
        # ENCONTRAR SEÇÃO DE MATERIAIS
        frame_materiais = None
        for widget in sistema.aba_fornecedor.winfo_children():
            if isinstance(widget, ttk.LabelFrame) and 'Materiais' in widget['text']:
                frame_materiais = widget
                break
        
        if frame_materiais:
            # ADICIONAR BOTÃO NA SEÇÃO DE MATERIAIS
            for subwidget in frame_materiais.winfo_children():
                if isinstance(subwidget, ttk.Frame):
                    # VERIFICAR SE JÁ EXISTE BOTÃO NFE
                    botoes_existentes = [w for w in subwidget.winfo_children() 
                                       if isinstance(w, ttk.Button) and 'NFe' in w['text']]
                    
                    if not botoes_existentes:
                        ttk.Button(
                            subwidget,
                            text="📄 Importar NFe",
                            command=callback_importar,
                            style='Medium.TButton'
                        ).pack(side='left', padx=5)
                    break
        else:
            # CRIAR SEÇÃO PRÓPRIA SE NÃO ENCONTRAR MATERIAIS
            frame_nfe = ttk.LabelFrame(sistema.aba_fornecedor, 
                                      text="📄 Importação NFe", 
                                      padding=10)
            frame_nfe.pack(fill='x', padx=10, pady=5)
            
            ttk.Button(frame_nfe, 
                      text="📁 Importar XML NFe", 
                      command=callback_importar).pack(pady=5)
        
        print("✅ Botão NFe unificado adicionado")
        
    except Exception as e:
        print(f"❌ Erro ao adicionar botão: {e}")


def limpar_sistemas_nfe_duplicados(sistema_principal):
    """
    Remove todos os sistemas NFe duplicados e suas interfaces
    """
    try:
        print("🧹 Limpando sistemas NFe duplicados...")
        
        # LISTA DE ATRIBUTOS A REMOVER
        atributos_para_remover = [
            'processador_nfe',
            'integrador_nfe', 
            'importar_nfe_xml',
            'importar_nfe_com_interface'
        ]
        
        for atributo in atributos_para_remover:
            if hasattr(sistema_principal, atributo):
                delattr(sistema_principal, atributo)
                print(f"  ✅ Removido: {atributo}")
        
        # REMOVER WIDGETS DUPLICADOS DA INTERFACE
        if hasattr(sistema_principal, 'aba_fornecedor'):
            widgets_nfe = []
            
            for widget in sistema_principal.aba_fornecedor.winfo_children():
                if isinstance(widget, ttk.LabelFrame):
                    texto = widget['text'].lower()
                    if any(palavra in texto for palavra in ['nfe', 'nf-e', 'importação nfe', 'híbrido']):
                        widgets_nfe.append(widget)
            
            for widget in widgets_nfe:
                widget.destroy()
                print(f"  ✅ Removido widget: {widget}")
        
        print("✅ Limpeza concluída!")
        
    except Exception as e:
        print(f"❌ Erro na limpeza: {e}")


# EXEMPLO DE USO PARA SUBSTITUIR SISTEMAS DUPLICADOS
"""
PARA RESOLVER A DUPLICAÇÃO, USE:

# No __init__ do seu SistemaEntradaDados, SUBSTITUA as linhas antigas por:

# ❌ REMOVER ESTAS LINHAS:
# from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
# from src.nfe.integracao_nfe_sistema import integrar_nfe_no_sistema
# inicializar_sistema_nfe_hibrido(self)
# integrar_nfe_no_sistema(self)

# ✅ ADICIONAR APENAS ESTA LINHA:
from src.nfe.sistema_nfe_unificado import substituir_sistemas_nfe_por_unificado
substituir_sistemas_nfe_por_unificado(self)

RESULTADO:
- ✅ Interface limpa com apenas UM botão NFe
- ✅ Sistema completo (XML local + futuro webservice)
- ✅ Integração perfeita com seu sistema atual
- ✅ Funcionalidade completa mantida
"""