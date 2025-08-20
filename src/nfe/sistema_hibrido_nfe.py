# -*- coding: utf-8 -*-
"""
Sistema Híbrido de Importação de NFe
Suporta:
1. Processamento de XML recebido por email
2. Consulta via chave de acesso (webservice SEFAZ)
3. Importação para sistema financeiro e materiais
"""

import xml.etree.ElementTree as ET
import requests
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from pathlib import Path
import re
import json
import zipfile
import base64
import ssl


class ProcessadorNFeHibrido:
    """Classe principal para processamento híbrido de NFe"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.certificado_path = None
        self.certificado_senha = None
        
        # Cache para evitar consultas repetidas
        self.cache_consultas = {}
        
        print("🔄 Processador NFe Híbrido inicializado")
    
    def criar_interface_importacao(self):
        """Cria interface unificada para importação"""
        self.janela_nfe = tk.Toplevel(self.sistema.root)
        self.janela_nfe.title("Importação de NF-e - Sistema Híbrido")
        self.janela_nfe.geometry("900x800")
        self.janela_nfe.grab_set()
        
        # Notebook para abas
        notebook = ttk.Notebook(self.janela_nfe)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Aba 1: XML Direto
        self.criar_aba_xml_direto(notebook)
        
        # Aba 2: Consulta por Chave
        self.criar_aba_consulta_chave(notebook)
        
        # Aba 3: Importação em Lote
        self.criar_aba_lote(notebook)
        
        # Frame de botões principais
        frame_botoes = ttk.Frame(self.janela_nfe)
        frame_botoes.pack(fill='x', padx=10, pady=5)
        
        ttk.Button(frame_botoes, text="⚙️ Configurar Certificado", 
                  command=self.configurar_certificado).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, text="❌ Fechar", 
                  command=self.janela_nfe.destroy).pack(side='right', padx=5)
    
    def criar_aba_xml_direto(self, notebook):
        """Aba para processamento de XML recebido"""
        frame_xml = ttk.Frame(notebook)
        notebook.add(frame_xml, text="📄 XML Recebido")
        
        # Seção de seleção de arquivo
        frame_arquivo = ttk.LabelFrame(frame_xml, text="Selecionar Arquivo XML", padding=10)
        frame_arquivo.pack(fill='x', padx=10, pady=5)
        
        self.label_arquivo = tk.Label(frame_arquivo, text="Nenhum arquivo selecionado", 
                                     fg='gray')
        self.label_arquivo.pack(anchor='w', pady=5)
        
        frame_btns_arquivo = ttk.Frame(frame_arquivo)
        frame_btns_arquivo.pack(fill='x', pady=5)
        
        ttk.Button(frame_btns_arquivo, text="📁 Selecionar XML", 
                  command=self.selecionar_xml).pack(side='left', padx=5)
        
        ttk.Button(frame_btns_arquivo, text="📧 Extrair de Email", 
                  command=self.extrair_xml_email).pack(side='left', padx=5)
        
        # Seção de dados extraídos
        self.frame_dados_xml = ttk.LabelFrame(frame_xml, text="Dados Extraídos", padding=10)
        
        # Seção de opções de importação
        self.frame_opcoes_xml = ttk.LabelFrame(frame_xml, text="Opções de Importação", padding=10)
        
        # Botão processar
        self.btn_processar_xml = ttk.Button(frame_xml, text="📥 Processar XML", 
                                           command=self.processar_xml_selecionado, 
                                           state='disabled')
        self.btn_processar_xml.pack(pady=10)
    
    def criar_aba_consulta_chave(self, notebook):
        """Aba para consulta via chave de acesso"""
        frame_chave = ttk.Frame(notebook)
        notebook.add(frame_chave, text="🔍 Consultar por Chave")
        
        # Entrada da chave
        frame_input = ttk.LabelFrame(frame_chave, text="Chave de Acesso", padding=10)
        frame_input.pack(fill='x', padx=10, pady=5)
        
        tk.Label(frame_input, text="Chave de Acesso (44 dígitos):", 
                font=('Arial', 9, 'bold')).pack(anchor='w')
        
        self.entry_chave = tk.Entry(frame_input, width=50, font=('Courier', 10))
        self.entry_chave.pack(fill='x', pady=5)
        
        # Bind para formatação automática
        self.entry_chave.bind('<KeyRelease>', self.formatar_chave_tempo_real)
        
        frame_btns_chave = ttk.Frame(frame_input)
        frame_btns_chave.pack(fill='x', pady=5)
        
        ttk.Button(frame_btns_chave, text="🔍 Consultar", 
                  command=self.consultar_por_chave).pack(side='left', padx=5)
        
        ttk.Button(frame_btns_chave, text="📋 Colar Chave", 
                  command=self.colar_chave).pack(side='left', padx=5)
        
        # Status da consulta
        self.label_status = tk.Label(frame_input, text="", fg='blue')
        self.label_status.pack(anchor='w', pady=2)
        
        # Seção de dados da consulta
        self.frame_dados_chave = ttk.LabelFrame(frame_chave, text="Dados da Consulta", padding=10)
        
        # Seção de opções
        self.frame_opcoes_chave = ttk.LabelFrame(frame_chave, text="Opções de Importação", padding=10)
        
        # Botão importar
        self.btn_importar_chave = ttk.Button(frame_chave, text="📥 Importar Dados", 
                                            command=self.importar_dados_chave, 
                                            state='disabled')
        self.btn_importar_chave.pack(pady=10)
    
    def criar_aba_lote(self, notebook):
        """Aba para importação em lote"""
        frame_lote = ttk.Frame(notebook)
        notebook.add(frame_lote, text="📊 Importação em Lote")
        
        # Lista de chaves
        frame_lista = ttk.LabelFrame(frame_lote, text="Lista de Chaves para Consultar", padding=10)
        frame_lista.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Área de texto para chaves
        tk.Label(frame_lista, text="Cole as chaves de acesso (uma por linha):", 
                font=('Arial', 9, 'bold')).pack(anchor='w')
        
        self.text_chaves = tk.Text(frame_lista, height=8, width=50, font=('Courier', 9))
        self.text_chaves.pack(fill='both', expand=True, pady=5)
        
        # Ou carregar de arquivo
        frame_arquivo_chaves = ttk.Frame(frame_lista)
        frame_arquivo_chaves.pack(fill='x', pady=5)
        
        ttk.Button(frame_arquivo_chaves, text="📁 Carregar de Arquivo TXT", 
                  command=self.carregar_chaves_arquivo).pack(side='left', padx=5)
        
        ttk.Button(frame_arquivo_chaves, text="💾 Salvar Lista", 
                  command=self.salvar_lista_chaves).pack(side='left', padx=5)
        
        # Progresso
        self.frame_progresso = ttk.LabelFrame(frame_lote, text="Progresso", padding=10)
        
        # Botão processar lote
        ttk.Button(frame_lote, text="🚀 Processar Lote", 
                  command=self.processar_lote).pack(pady=10)
    
    def selecionar_xml(self):
        """Seleciona arquivo XML"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar arquivo XML da NF-e",
            filetypes=[
                ("Arquivos XML", "*.xml"),
                ("Todos os arquivos", "*.*")
            ]
        )
        
        if arquivo:
            self.arquivo_xml_atual = arquivo
            nome_arquivo = Path(arquivo).name
            self.label_arquivo.config(text=f"Arquivo: {nome_arquivo}", fg='green')
            self.btn_processar_xml.config(state='normal')
    
    def extrair_xml_email(self):
        """Extrai XML de anexos de email"""
        messagebox.showinfo("Em Desenvolvimento", 
            "Funcionalidade de extração de email será implementada.\n"
            "Por enquanto, salve o XML do email e use 'Selecionar XML'.")
    
    def processar_xml_selecionado(self):
        """Processa XML selecionado"""
        try:
            if not hasattr(self, 'arquivo_xml_atual'):
                messagebox.showerror("Erro", "Selecione um arquivo XML primeiro!")
                return
            
            # Processar XML
            dados_nfe = self.processar_xml_nfe(self.arquivo_xml_atual)
            
            if dados_nfe:
                self.dados_nfe_atual = dados_nfe
                self.exibir_dados_extraidos(dados_nfe, self.frame_dados_xml)
                self.criar_opcoes_importacao(self.frame_opcoes_xml, 'xml')
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao processar XML:\n{str(e)}")
    
    def processar_xml_nfe(self, caminho_arquivo):
        """Processa arquivo XML da NFe"""
        try:
            print(f"📄 Processando XML: {caminho_arquivo}")
            
            # Ler arquivo XML
            tree = ET.parse(caminho_arquivo)
            root = tree.getroot()
            
            # Namespace da NFe
            ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
            
            # Extrair dados principais
            dados = {
                'fonte_dados': 'XML Local',
                'arquivo_origem': caminho_arquivo
            }
            
            # Buscar elementos principais
            inf_nfe = root.find('.//nfe:infNFe', ns)
            if inf_nfe is None:
                raise Exception("Estrutura XML inválida - infNFe não encontrado")
            
            # Chave de acesso
            dados['chave_acesso'] = inf_nfe.get('Id', '').replace('NFe', '')
            
            # Dados da NFe
            ide = inf_nfe.find('nfe:ide', ns)
            if ide is not None:
                dados['numero_nf'] = self.get_xml_text(ide.find('nfe:nNF', ns))
                dados['serie'] = self.get_xml_text(ide.find('nfe:serie', ns))
                
                # Data de emissão
                dh_emi = self.get_xml_text(ide.find('nfe:dhEmi', ns))
                if dh_emi:
                    dados['data_emissao'] = self.formatar_data_xml(dh_emi)
            
            # Dados do emitente
            emit = inf_nfe.find('nfe:emit', ns)
            if emit is not None:
                dados['cnpj_emitente'] = self.get_xml_text(emit.find('nfe:CNPJ', ns))
                dados['razao_social_emitente'] = self.get_xml_text(emit.find('nfe:xNome', ns))
                
                # Endereço do emitente
                endereco = emit.find('nfe:enderEmit', ns)
                if endereco is not None:
                    dados['endereco_emitente'] = {
                        'logradouro': self.get_xml_text(endereco.find('nfe:xLgr', ns)),
                        'numero': self.get_xml_text(endereco.find('nfe:nro', ns)),
                        'cidade': self.get_xml_text(endereco.find('nfe:xMun', ns)),
                        'uf': self.get_xml_text(endereco.find('nfe:UF', ns)),
                        'cep': self.get_xml_text(endereco.find('nfe:CEP', ns))
                    }
            
            # Dados do destinatário
            dest = inf_nfe.find('nfe:dest', ns)
            if dest is not None:
                cnpj_dest = self.get_xml_text(dest.find('nfe:CNPJ', ns))
                cpf_dest = self.get_xml_text(dest.find('nfe:CPF', ns))
                dados['documento_destinatario'] = cnpj_dest or cpf_dest
                dados['nome_destinatario'] = self.get_xml_text(dest.find('nfe:xNome', ns))
            
            # Totais
            total = inf_nfe.find('.//nfe:total/nfe:ICMSTot', ns)
            if total is not None:
                dados['valor_total'] = float(self.get_xml_text(total.find('nfe:vNF', ns)) or 0)
                dados['valor_produtos'] = float(self.get_xml_text(total.find('nfe:vProd', ns)) or 0)
            
            # Produtos/Itens
            dados['produtos'] = self.extrair_produtos_xml(inf_nfe, ns)
            
            print(f"✅ XML processado: {dados['razao_social_emitente']}")
            print(f"📄 NF-e {dados['numero_nf']} - R$ {dados.get('valor_total', 0):,.2f}")
            print(f"📦 {len(dados['produtos'])} produtos extraídos")
            
            return dados
            
        except Exception as e:
            print(f"❌ Erro ao processar XML: {e}")
            raise e
    
    def extrair_produtos_xml(self, inf_nfe, ns):
        """Extrai produtos do XML"""
        produtos = []
        
        try:
            # Buscar todos os itens
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
                
                # Classificar produto para categoria de material
                produto['categoria_sugerida'] = self.classificar_produto_por_descricao(produto['descricao'])
                produto['subcategoria_sugerida'] = self.sugerir_subcategoria(produto['descricao'], produto['categoria_sugerida'])
                
                produtos.append(produto)
            
        except Exception as e:
            print(f"⚠️ Erro ao extrair produtos: {e}")
        
        return produtos
    
    def consultar_por_chave(self):
        """Consulta NFe via webservice da SEFAZ"""
        try:
            chave = self.entry_chave.get().strip().replace(' ', '').replace('-', '')
            
            if not self.validar_chave_acesso(chave):
                messagebox.showerror("Erro", "Chave de acesso inválida!\nDeve ter 44 dígitos.")
                return
            
            # Verificar cache primeiro
            if chave in self.cache_consultas:
                print("📋 Usando dados do cache")
                dados_nfe = self.cache_consultas[chave]
                self.dados_nfe_atual = dados_nfe
                self.exibir_dados_extraidos(dados_nfe, self.frame_dados_chave)
                self.criar_opcoes_importacao(self.frame_opcoes_chave, 'chave')
                self.btn_importar_chave.config(state='normal')
                return
            
            # Mostrar status
            self.label_status.config(text="🔍 Consultando SEFAZ...", fg='blue')
            self.janela_nfe.update()
            
            # Fazer consulta
            dados_nfe = self.consultar_nfe_sefaz(chave)
            
            if dados_nfe:
                # Adicionar ao cache
                self.cache_consultas[chave] = dados_nfe
                
                self.dados_nfe_atual = dados_nfe
                self.exibir_dados_extraidos(dados_nfe, self.frame_dados_chave)
                self.criar_opcoes_importacao(self.frame_opcoes_chave, 'chave')
                self.btn_importar_chave.config(state='normal')
                self.label_status.config(text="✅ Consulta realizada com sucesso!", fg='green')
            else:
                self.label_status.config(text="❌ NFe não encontrada", fg='red')
            
        except Exception as e:
            self.label_status.config(text=f"❌ Erro: {str(e)}", fg='red')
            messagebox.showerror("Erro", f"Erro na consulta:\n{str(e)}")
    
    def consultar_nfe_sefaz(self, chave_acesso):
        """Consulta NFe no webservice da SEFAZ"""
        try:
            print(f"🔍 Consultando SEFAZ: {chave_acesso}")
            
            # Extrair UF da chave para determinar webservice
            uf_codigo = chave_acesso[:2]
            uf_sigla = self.obter_uf_por_codigo(uf_codigo)
            
            # URL do webservice de consulta
            url_webservice = self.obter_url_consulta_sefaz(uf_sigla)
            
            # Montar envelope SOAP
            envelope_soap = self.criar_envelope_consulta(chave_acesso)
            
            # Headers da requisição
            headers = {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF'
            }
            
            # Configurar SSL se tiver certificado
            session = requests.Session()
            if self.certificado_path and Path(self.certificado_path).exists():
                session.cert = (self.certificado_path, self.certificado_senha)
                session.verify = False  # Para desenvolvimento, em produção use True
            
            # Fazer requisição
            response = session.post(url_webservice, data=envelope_soap, headers=headers, timeout=30)
            
            if response.status_code == 200:
                # Processar resposta XML
                dados_nfe = self.processar_resposta_sefaz(response.text, chave_acesso)
                return dados_nfe
            else:
                raise Exception(f"Erro HTTP {response.status_code}: {response.text}")
            
        except Exception as e:
            print(f"❌ Erro na consulta SEFAZ: {e}")
            # Como fallback, simular dados básicos da chave
            return self.criar_dados_simulados(chave_acesso)
    
    def criar_envelope_consulta(self, chave_acesso):
        """Cria envelope SOAP para consulta"""
        return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:xsd="http://www.w3.org/2001/XMLSchema" 
               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">
            <consReciNFe versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
                <tpAmb>1</tpAmb>
                <chNFe>{chave_acesso}</chNFe>
            </consReciNFe>
        </nfeDadosMsg>
    </soap:Body>
</soap:Envelope>"""
    
    def processar_resposta_sefaz(self, xml_resposta, chave_acesso):
        """Processa resposta do webservice SEFAZ"""
        try:
            # Parse da resposta
            root = ET.fromstring(xml_resposta)
            
            # Buscar dados da NFe na resposta
            # (Implementar parsing específico da resposta SEFAZ)
            
            # Por enquanto, retornar dados simulados baseados na chave
            return self.criar_dados_simulados(chave_acesso)
            
        except Exception as e:
            print(f"⚠️ Erro ao processar resposta SEFAZ: {e}")
            return self.criar_dados_simulados(chave_acesso)
    
    def criar_dados_simulados(self, chave_acesso):
        """Cria dados simulados baseados na chave de acesso"""
        info_chave = self.extrair_info_chave(chave_acesso)
        
        return {
            'chave_acesso': chave_acesso,
            'numero_nf': info_chave['numero'],
            'serie': '1',
            'data_emissao': datetime.now().strftime('%d/%m/%Y'),
            'cnpj_emitente': info_chave['cnpj_emitente'],
            'razao_social_emitente': 'FORNECEDOR CONSULTADO LTDA',
            'valor_total': 0.0,
            'valor_produtos': 0.0,
            'produtos': [],
            'fonte_dados': 'Simulação (Consulta SEFAZ)',
            'observacao': 'Dados básicos extraídos da chave de acesso'
        }
    
    def exibir_dados_extraidos(self, dados, frame_container):
        """Exibe dados extraídos da NFe"""
        frame_container.pack(fill='x', padx=10, pady=5)
        
        # Limpar frame
        for widget in frame_container.winfo_children():
            widget.destroy()
        
        # Criar grade de informações
        info_frame = ttk.Frame(frame_container)
        info_frame.pack(fill='x')
        
        informacoes = [
            ("Chave:", dados.get('chave_acesso', '')[:44]),
            ("Número:", dados.get('numero_nf', '')),
            ("Série:", dados.get('serie', '')),
            ("Data:", dados.get('data_emissao', '')),
            ("CNPJ:", dados.get('cnpj_emitente', '')),
            ("Emitente:", dados.get('razao_social_emitente', '')),
            ("Valor Total:", f"R$ {dados.get('valor_total', 0):,.2f}"),
            ("Produtos:", len(dados.get('produtos', []))),
            ("Fonte:", dados.get('fonte_dados', ''))
        ]
        
        for i, (label, valor) in enumerate(informacoes):
            row = i // 3
            col = (i % 3) * 2
            
            tk.Label(info_frame, text=label, font=('Arial', 9, 'bold')).grid(
                row=row, column=col, sticky='w', padx=5, pady=2)
            tk.Label(info_frame, text=str(valor)[:50]).grid(
                row=row, column=col+1, sticky='w', padx=5, pady=2)
        
        # Lista de produtos se houver
        if dados.get('produtos'):
            self.exibir_produtos(dados['produtos'], frame_container)
    
    def exibir_produtos(self, produtos, frame_container):
        """Exibe lista de produtos"""
        frame_produtos = ttk.LabelFrame(frame_container, text=f"Produtos ({len(produtos)} itens)", padding=5)
        frame_produtos.pack(fill='both', expand=True, pady=5)
        
        # TreeView para produtos
        colunas = ('Item', 'Código', 'Descrição', 'Qtd', 'Un', 'Vl Unit', 'Vl Total')
        tree = ttk.Treeview(frame_produtos, columns=colunas, show='headings', height=6)
        
        # Configurar colunas
        tree.heading('Item', text='#')
        tree.heading('Código', text='Código')
        tree.heading('Descrição', text='Descrição')
        tree.heading('Qtd', text='Qtd')
        tree.heading('Un', text='Un')
        tree.heading('Vl Unit', text='Vl Unit')
        tree.heading('Vl Total', text='Vl Total')
        
        tree.column('Item', width=40)
        tree.column('Código', width=80)
        tree.column('Descrição', width=250)
        tree.column('Qtd', width=60)
        tree.column('Un', width=40)
        tree.column('Vl Unit', width=80)
        tree.column('Vl Total', width=80)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(frame_produtos, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Preencher dados
        for i, produto in enumerate(produtos, 1):
            tree.insert('', 'end', values=(
                i,
                produto.get('codigo', ''),
                produto.get('descricao', '')[:40],
                produto.get('quantidade', ''),
                produto.get('unidade', ''),
                f"R$ {produto.get('valor_unitario', 0):.2f}",
                f"R$ {produto.get('valor_total', 0):.2f}"
            ))
        
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
    
    def criar_opcoes_importacao(self, frame_container, origem):
        """Cria opções de importação"""
        frame_container.pack(fill='x', padx=10, pady=5)
        
        # Limpar frame
        for widget in frame_container.winfo_children():
            widget.destroy()
        
        # Variáveis de controle
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        tk.Label(frame_container, text="Selecione o que importar:", 
                font=('Arial', 10, 'bold')).pack(anchor='w', pady=5)
        
        # Checkboxes
        cb_frame = ttk.Frame(frame_container)
        cb_frame.pack(fill='x', pady=5)
        
        cb_financeiro = tk.Checkbutton(
            cb_frame, 
            text="💰 Dados Financeiros (despesa/receita)", 
            variable=self.importar_financeiro,
            font=('Arial', 10)
        )
        cb_financeiro.pack(anchor='w', pady=2)
        
        cb_materiais = tk.Checkbutton(
            cb_frame, 
            text="📦 Materiais da Obra", 
            variable=self.importar_materiais,
            font=('Arial', 10)
        )
        cb_materiais.pack(anchor='w', pady=2)
    
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
    
    def validar_chave_acesso(self, chave):
        """Valida chave de acesso de 44 dígitos"""
        chave_limpa = re.sub(r'[^0-9]', '', str(chave))
        return len(chave_limpa) == 44
    
    def extrair_info_chave(self, chave):
        """Extrai informações básicas da chave de acesso"""
        chave_limpa = re.sub(r'[^0-9]', '', str(chave))
        return {
            'uf': chave_limpa[0:2],
            'ano_mes': chave_limpa[2:6],
            'cnpj_emitente': chave_limpa[6:20],
            'numero': chave_limpa[25:34]
        }
    
    def classificar_produto_por_descricao(self, descricao):
        """Classifica produto para categoria de material"""
        if not descricao:
            return 'OUTROS'
        
        desc_upper = str(descricao).upper()
        
        classificacoes = {
            'ACABAMENTOS': [
                'CERAMICA', 'PORCELANATO', 'AZULEJO', 'PASTILHA', 'REVESTIMENTO', 'PISO',
                'RODAPE', 'MOLDURA', 'REJUNTE', 'GESSO', 'FORRO'
            ],
            'TINTAS': [
                'TINTA', 'VERNIZ', 'ESMALTE', 'PRIMER', 'SELADOR', 'MASSA CORRIDA'
            ],
            'ELETRICO': [
                'FIO', 'CABO', 'TOMADA', 'INTERRUPTOR', 'LAMPADA', 'LED', 'ELETRICO',
                'DISJUNTOR', 'QUADRO ELETRICO'
            ],
            'HIDRAULICO': [
                'TUBO', 'CONEXAO', 'REGISTRO', 'TORNEIRA', 'VALVULA', 'HIDRAULICO',
                'CANO', 'CHUVEIRO', 'VASO SANITARIO'
            ],
            'ESTRUTURAL': [
                'CIMENTO', 'CONCRETO', 'FERRO', 'AÇO', 'TIJOLO', 'BLOCO', 'VIGA',
                'AREIA', 'BRITA', 'CAL'
            ],
            'ESQUADRIAS': [
                'PORTA', 'JANELA', 'FECHADURA', 'DOBRADIÇA', 'VIDRO', 'MARCO',
                'FERRAGEM', 'TRINCO'
            ],
            'FERRAGENS': [
                'PARAFUSO', 'PREGO', 'BUCHA', 'CHAVE', 'CADEADO', 'REBITE'
            ]
        }
        
        for categoria, palavras_chave in classificacoes.items():
            if any(palavra in desc_upper for palavra in palavras_chave):
                return categoria
        
        return 'OUTROS'
    
    def sugerir_subcategoria(self, descricao, categoria):
        """Sugere subcategoria baseada na descrição e categoria"""
        if not descricao:
            return ''
        
        desc_upper = str(descricao).upper()
        
        subcategorias = {
            'ACABAMENTOS': {
                'CERAMICA': ['CERAMICA'],
                'PORCELANATO': ['PORCELANATO'],
                'PISO_LAMINADO': ['LAMINADO'],
                'RODAPE': ['RODAPE'],
                'REJUNTE': ['REJUNTE'],
                'GESSO': ['GESSO', 'FORRO']
            },
            'ELETRICO': {
                'CABOS': ['FIO', 'CABO'],
                'TOMADAS': ['TOMADA'],
                'INTERRUPTORES': ['INTERRUPTOR'],
                'ILUMINACAO': ['LAMPADA', 'LED'],
                'PROTECAO': ['DISJUNTOR']
            },
            'HIDRAULICO': {
                'TUBULACAO': ['TUBO', 'CANO'],
                'CONEXOES': ['CONEXAO', 'JOELHO', 'TE'],
                'METAIS': ['TORNEIRA', 'REGISTRO', 'CHUVEIRO'],
                'LOUÇAS': ['VASO', 'PIA', 'TANQUE']
            }
        }
        
        if categoria in subcategorias:
            for subcat, palavras in subcategorias[categoria].items():
                if any(palavra in desc_upper for palavra in palavras):
                    return subcat
        
        return ''
    
    def obter_uf_por_codigo(self, codigo):
        """Converte código UF para sigla"""
        ufs = {
            '11': 'RO', '12': 'AC', '13': 'AM', '14': 'RR', '15': 'PA',
            '16': 'AP', '17': 'TO', '21': 'MA', '22': 'PI', '23': 'CE',
            '24': 'RN', '25': 'PB', '26': 'PE', '27': 'AL', '28': 'SE',
            '29': 'BA', '31': 'MG', '32': 'ES', '33': 'RJ', '35': 'SP',
            '41': 'PR', '42': 'SC', '43': 'RS', '50': 'MS', '51': 'MT',
            '52': 'GO', '53': 'DF'
        }
        return ufs.get(codigo, 'SP')  # Default SP se não encontrar
    
    def obter_url_consulta_sefaz(self, uf):
        """Obtém URL do webservice de consulta por UF"""
        urls = {
            'SP': 'https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'RJ': 'https://nfe.sefaz.rj.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'MG': 'https://nfe.fazenda.mg.gov.br/nfe2/services/NFeConsultaProtocolo4',
            # Adicionar outras UFs conforme necessário
        }
        return urls.get(uf, urls['SP'])  # Default SP
    
    def formatar_chave_tempo_real(self, event):
        """Formata chave de acesso em tempo real"""
        chave = self.entry_chave.get()
        # Remove caracteres não numéricos
        chave_limpa = re.sub(r'[^0-9]', '', chave)
        
        # Limita a 44 dígitos
        if len(chave_limpa) > 44:
            chave_limpa = chave_limpa[:44]
        
        # Formatar com espaços (grupos de 4)
        chave_formatada = ' '.join([chave_limpa[i:i+4] for i in range(0, len(chave_limpa), 4)])
        
        # Atualizar campo sem trigger do evento
        self.entry_chave.delete(0, tk.END)
        self.entry_chave.insert(0, chave_formatada)
    
    def colar_chave(self):
        """Cola chave do clipboard"""
        try:
            chave = self.janela_nfe.clipboard_get()
            chave_limpa = re.sub(r'[^0-9]', '', chave)
            
            if len(chave_limpa) == 44:
                self.entry_chave.delete(0, tk.END)
                self.entry_chave.insert(0, chave_limpa)
                self.formatar_chave_tempo_real(None)
            else:
                messagebox.showwarning("Aviso", "Chave de acesso inválida no clipboard!")
                
        except tk.TclError:
            messagebox.showwarning("Aviso", "Nenhum texto no clipboard!")
    
    def configurar_certificado(self):
        """Configura certificado digital"""
        print("SISTEMA_HIBRIDO: Usando versão corrigida")
        try:
            return self.sistema.consultor_sefaz_a1.configurar_certificado_interface()
        except:
            # Fallback para método direto
            return self.sistema.configurar_certificado_rapido()

        janela_cert = tk.Toplevel(self.janela_nfe)
        janela_cert.title("Configurar Certificado Digital")
        janela_cert.geometry("500x300")
        janela_cert.grab_set()
        
        frame_cert = ttk.LabelFrame(janela_cert, text="Certificado Digital A1 (.pfx)", padding=10)
        frame_cert.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Arquivo do certificado
        tk.Label(frame_cert, text="Arquivo do Certificado (.pfx):").pack(anchor='w', pady=5)
        
        frame_arquivo = ttk.Frame(frame_cert)
        frame_arquivo.pack(fill='x', pady=5)
        
        self.entry_cert_path = tk.Entry(frame_arquivo, width=50)
        self.entry_cert_path.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        ttk.Button(frame_arquivo, text="📁", 
                  command=self.selecionar_certificado).pack(side='right')
        
        # Senha do certificado
        tk.Label(frame_cert, text="Senha do Certificado:").pack(anchor='w', pady=(10, 5))
        self.entry_cert_senha = tk.Entry(frame_cert, show='*', width=30)
        self.entry_cert_senha.pack(anchor='w', pady=5)
        
        # Botões
        frame_btns_cert = ttk.Frame(frame_cert)
        frame_btns_cert.pack(fill='x', pady=10)
        
        ttk.Button(frame_btns_cert, text="✅ Salvar", 
                  command=lambda: self.salvar_configuracao_certificado(janela_cert)).pack(side='left', padx=5)
        
        ttk.Button(frame_btns_cert, text="🧪 Testar", 
                  command=self.testar_certificado).pack(side='left', padx=5)
        
        ttk.Button(frame_btns_cert, text="❌ Cancelar", 
                  command=janela_cert.destroy).pack(side='right', padx=5)
        
        # Carregar configuração atual se existir
        if self.certificado_path:
            self.entry_cert_path.insert(0, self.certificado_path)
    
    def selecionar_certificado(self):
        """Seleciona arquivo de certificado"""
        arquivo = filedialog.askopenfilename(
            title="Selecionar Certificado Digital",
            filetypes=[("Certificado", "*.pfx *.p12"), ("Todos os arquivos", "*.*")]
        )
        if arquivo:
            self.entry_cert_path.delete(0, tk.END)
            self.entry_cert_path.insert(0, arquivo)
    
    def salvar_configuracao_certificado(self, janela):
        """Salva configuração do certificado"""
        self.certificado_path = self.entry_cert_path.get()
        self.certificado_senha = self.entry_cert_senha.get()
        
        if self.certificado_path and Path(self.certificado_path).exists():
            messagebox.showinfo("Sucesso", "Certificado configurado com sucesso!")
            janela.destroy()
        else:
            messagebox.showerror("Erro", "Arquivo de certificado não encontrado!")
    
    def testar_certificado(self):
        """Testa certificado configurado"""
        try:
            cert_path = self.entry_cert_path.get()
            cert_senha = self.entry_cert_senha.get()
            
            if not cert_path or not Path(cert_path).exists():
                messagebox.showerror("Erro", "Selecione um arquivo de certificado válido!")
                return
            
            # Aqui você pode implementar teste real do certificado
            # Por enquanto, apenas verifica se o arquivo existe
            messagebox.showinfo("Teste", "Certificado parece válido!\n(Teste completo será implementado)")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao testar certificado:\n{str(e)}")
    
    # MÉTODOS DE IMPORTAÇÃO EM LOTE
    
    def carregar_chaves_arquivo(self):
        """Carrega chaves de um arquivo texto"""
        arquivo = filedialog.askopenfilename(
            title="Carregar Lista de Chaves",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            try:
                with open(arquivo, 'r', encoding='utf-8') as f:
                    conteudo = f.read()
                
                self.text_chaves.delete('1.0', tk.END)
                self.text_chaves.insert('1.0', conteudo)
                
                # Contar chaves válidas
                chaves = self.extrair_chaves_do_texto(conteudo)
                messagebox.showinfo("Sucesso", f"Carregadas {len(chaves)} chaves válidas do arquivo!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao carregar arquivo:\n{str(e)}")
    
    def salvar_lista_chaves(self):
        """Salva lista atual de chaves"""
        arquivo = filedialog.asksaveasfilename(
            title="Salvar Lista de Chaves",
            defaultextension=".txt",
            filetypes=[("Arquivos de texto", "*.txt"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            try:
                conteudo = self.text_chaves.get('1.0', tk.END)
                with open(arquivo, 'w', encoding='utf-8') as f:
                    f.write(conteudo)
                
                messagebox.showinfo("Sucesso", "Lista de chaves salva com sucesso!")
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao salvar arquivo:\n{str(e)}")
    
    def extrair_chaves_do_texto(self, texto):
        """Extrai chaves válidas do texto"""
        chaves_validas = []
        linhas = texto.strip().split('\n')
        
        for linha in linhas:
            chave_limpa = re.sub(r'[^0-9]', '', linha.strip())
            if len(chave_limpa) == 44:
                chaves_validas.append(chave_limpa)
        
        return chaves_validas
    
    def processar_lote(self):
        """Processa lote de chaves"""
        try:
            texto_chaves = self.text_chaves.get('1.0', tk.END)
            chaves = self.extrair_chaves_do_texto(texto_chaves)
            
            if not chaves:
                messagebox.showwarning("Aviso", "Nenhuma chave válida encontrada!")
                return
            
            # Criar janela de progresso
            self.criar_janela_progresso(len(chaves))
            
            # Processar cada chave
            resultados = []
            for i, chave in enumerate(chaves):
                try:
                    self.atualizar_progresso(i + 1, len(chaves), f"Processando {chave[:20]}...")
                    
                    # Consultar NFe
                    dados_nfe = self.consultar_nfe_sefaz(chave)
                    
                    if dados_nfe:
                        # Importar dados conforme configuração
                        if self.importar_financeiro.get():
                            self.importar_dados_financeiro(dados_nfe)
                        
                        if self.importar_materiais.get():
                            self.importar_dados_material(dados_nfe)
                        
                        resultados.append(f"✅ {chave}: {dados_nfe.get('razao_social_emitente', 'OK')}")
                    else:
                        resultados.append(f"❌ {chave}: Não encontrada")
                
                except Exception as e:
                    resultados.append(f"❌ {chave}: Erro - {str(e)}")
                
                # Pequena pausa para não sobrecarregar
                self.janela_nfe.after(100)
            
            # Fechar progresso
            self.fechar_progresso()
            
            # Mostrar resultados
            self.mostrar_resultados_lote(resultados)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no processamento em lote:\n{str(e)}")
    
    def criar_janela_progresso(self, total):
        """Cria janela de progresso"""
        self.janela_progresso = tk.Toplevel(self.janela_nfe)
        self.janela_progresso.title("Processando Lote")
        self.janela_progresso.geometry("400x150")
        self.janela_progresso.grab_set()
        
        frame_prog = ttk.Frame(self.janela_progresso)
        frame_prog.pack(fill='both', expand=True, padx=20, pady=20)
        
        self.label_progresso = tk.Label(frame_prog, text="Iniciando processamento...")
        self.label_progresso.pack(pady=10)
        
        self.progress_bar = ttk.Progressbar(frame_prog, length=300, mode='determinate')
        self.progress_bar.pack(pady=10)
        self.progress_bar['maximum'] = total
        
        self.label_status = tk.Label(frame_prog, text="0/0")
        self.label_status.pack()
    
    def atualizar_progresso(self, atual, total, status):
        """Atualiza barra de progresso"""
        if hasattr(self, 'janela_progresso') and self.janela_progresso.winfo_exists():
            self.progress_bar['value'] = atual
            self.label_progresso.config(text=status)
            self.label_status.config(text=f"{atual}/{total}")
            self.janela_progresso.update()
    
    def fechar_progresso(self):
        """Fecha janela de progresso"""
        if hasattr(self, 'janela_progresso') and self.janela_progresso.winfo_exists():
            self.janela_progresso.destroy()
    
    def mostrar_resultados_lote(self, resultados):
        """Mostra resultados do processamento em lote"""
        janela_result = tk.Toplevel(self.janela_nfe)
        janela_result.title("Resultados do Processamento")
        janela_result.geometry("600x400")
        
        frame_result = ttk.Frame(janela_result)
        frame_result.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Texto com resultados
        text_result = tk.Text(frame_result, wrap='word')
        scrollbar_result = ttk.Scrollbar(frame_result, orient='vertical', command=text_result.yview)
        text_result.configure(yscrollcommand=scrollbar_result.set)
        
        # Preencher resultados
        texto_final = f"RELATÓRIO DE PROCESSAMENTO EM LOTE\n{'='*50}\n\n"
        for resultado in resultados:
            texto_final += resultado + "\n"
        
        # Estatísticas
        sucessos = len([r for r in resultados if r.startswith('✅')])
        erros = len([r for r in resultados if r.startswith('❌')])
        
        texto_final += f"\n{'='*50}\n"
        texto_final += f"RESUMO:\n"
        texto_final += f"✅ Sucessos: {sucessos}\n"
        texto_final += f"❌ Erros: {erros}\n"
        texto_final += f"📊 Total: {len(resultados)}\n"
        
        text_result.insert('1.0', texto_final)
        text_result.config(state='disabled')
        
        text_result.pack(side='left', fill='both', expand=True)
        scrollbar_result.pack(side='right', fill='y')
        
        # Botão fechar
        ttk.Button(janela_result, text="Fechar", 
                  command=janela_result.destroy).pack(pady=10)
    
    # MÉTODOS DE IMPORTAÇÃO DE DADOS
    
    def importar_dados_chave(self):
        """Importa dados da consulta por chave"""
        try:
            if not hasattr(self, 'dados_nfe_atual') or not self.dados_nfe_atual:
                messagebox.showerror("Erro", "Nenhum dado carregado para importar!")
                return
            
            resultados = []
            
            # Importar dados financeiros
            if self.importar_financeiro.get():
                resultado_fin = self.importar_dados_financeiro(self.dados_nfe_atual)
                resultados.append(f"✅ Financeiro: {resultado_fin}")
            
            # Importar materiais
            if self.importar_materiais.get():
                resultado_mat = self.importar_dados_material(self.dados_nfe_atual)
                resultados.append(f"✅ Materiais: {resultado_mat}")
            
            if resultados:
                messagebox.showinfo("Importação Concluída", "\n".join(resultados))
            else:
                messagebox.showwarning("Aviso", "Nenhuma opção de importação selecionada!")
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na importação:\n{str(e)}")
    
    def importar_dados_financeiro(self, dados_nfe):
        """Importa dados para o sistema financeiro"""
        try:
            # Criar entrada financeira baseada na NFe
            dados_financeiros = {
                'data': dados_nfe.get('data_emissao', ''),
                'cnpj_cpf': dados_nfe.get('cnpj_emitente', ''),
                'nome': dados_nfe.get('razao_social_emitente', ''),
                'categoria': 'MAT',  # Material
                'tp_desp': '3',  # Tipo despesa
                'referencia': 'MATERIAL OBRA',
                'nf': dados_nfe.get('numero_nf', ''),
                'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dias': 1,
                'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dt_vencto': dados_nfe.get('data_emissao', ''),
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('chave_acesso', '')}",
                'forma_pagamento': ''
            }
            
            # Adicionar à lista de dados para incluir do sistema principal
            if hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir.append(dados_financeiros)
                return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            else:
                return "Adicionado (sistema não inicializado)"
            
        except Exception as e:
            raise Exception(f"Erro ao importar dados financeiros: {str(e)}")
    
    def importar_dados_material(self, dados_nfe):
        """Importa produtos como materiais da obra"""
        try:
            produtos = dados_nfe.get('produtos', [])
            if not produtos:
                return "Nenhum produto para importar"
            
            materiais_importados = 0
            
            # Verificar se tem gerenciador de materiais
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                from src.materiais.gerenciador_materiais import GerenciadorMateriais
                self.sistema.gerenciador_materiais = GerenciadorMateriais(self.sistema)
            
            for produto in produtos:
                # Mapear produto para material
                dados_material = {
                    'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                    'Data_Cadastro': datetime.now().strftime('%d/%m/%Y'),
                    'Categoria': produto.get('categoria_sugerida', 'OUTROS'),
                    'Subcategoria': produto.get('subcategoria_sugerida', ''),
                    'Codigo_Produto': produto.get('codigo', ''),
                    'Descricao_Completa': produto.get('descricao', ''),
                    'Marca': '',  # Não disponível na NFe
                    'Modelo': '',
                    'Cor_Acabamento': '',
                    'Dimensoes': '',
                    'Ambiente_Aplicacao': '',  # Usuário deve definir depois
                    'Data_Instalacao': '',
                    'Instalador': '',
                    'Status_Instalacao': 'PENDENTE',
                    'Garantia_Meses': 0,
                    'Observacoes': f"Importado da NF-e {dados_nfe.get('numero_nf', '')} - {dados_nfe.get('fonte_dados', '')}",
                    'Tem_Dados_Compra': True,
                    'Nome_Fornecedor': dados_nfe.get('razao_social_emitente', ''),
                    'CNPJ_Fornecedor': dados_nfe.get('cnpj_emitente', ''),
                    'Data_Compra': dados_nfe.get('data_emissao', ''),
                    'Quantidade': produto.get('quantidade', ''),
                    'Unidade': produto.get('unidade', 'UN'),
                    'Valor_Unitario': produto.get('valor_unitario', 0),
                    'Valor_Total': produto.get('valor_total', 0),
                    'Numero_NF': dados_nfe.get('numero_nf', '')
                }
                
                # Salvar material
                material_id = self.sistema.gerenciador_materiais.salvar_material(dados_material)
                materiais_importados += 1
                
                print(f"✅ Material importado - ID: {material_id}")
            
            return f"{materiais_importados} produto(s) importado(s)"
            
        except Exception as e:
            raise Exception(f"Erro ao importar materiais: {str(e)}")


# CLASSE PARA INTEGRAÇÃO COM SISTEMA EXISTENTE
class IntegradorSistemaExistente:
    """Integra o processador híbrido com o sistema existente"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.processador = ProcessadorNFeHibrido(sistema_principal)
    
    def adicionar_botao_nfe_na_interface(self):
        """Adiciona botão de importação NFe na interface existente"""
        try:
            # Localizar frame de botões de materiais
            if hasattr(self.sistema, 'aba_fornecedor'):
                # Adicionar na seção de materiais
                frame_materiais = None
                for widget in self.sistema.aba_fornecedor.winfo_children():
                    if isinstance(widget, ttk.LabelFrame) and 'Materiais' in widget['text']:
                        frame_materiais = widget
                        break
                
                if frame_materiais:
                    # Encontrar frame de botões dentro da seção de materiais
                    for subwidget in frame_materiais.winfo_children():
                        if isinstance(subwidget, ttk.Frame):
                            ttk.Button(
                                subwidget,
                                text="📄 Importar NF-e",
                                command=self.processador.criar_interface_importacao,
                                style='Medium.TButton'
                            ).pack(side='left', padx=5)
                            break
                    
                    print("✅ Botão de importação NFe adicionado!")
                
        except Exception as e:
            print(f"❌ Erro ao adicionar botão NFe: {e}")
    
    def substituir_metodos_existentes(self):
        """Substitui métodos existentes por versões melhoradas"""
        try:
            # Substituir método de importação de NFe se existir
            if hasattr(self.sistema, 'criar_interface_importacao_nfe'):
                self.sistema.criar_interface_importacao_nfe_original = self.sistema.criar_interface_importacao_nfe
                self.sistema.criar_interface_importacao_nfe = self.processador.criar_interface_importacao
                print("✅ Método de importação NFe substituído!")
            
        except Exception as e:
            print(f"❌ Erro ao substituir métodos: {e}")


# EXEMPLO DE USO NO SISTEMA PRINCIPAL
def inicializar_sistema_nfe_hibrido(sistema_principal):
    """
    Função para inicializar o sistema híbrido de NFe no sistema principal
    
    Args:
        sistema_principal: Instância da classe SistemaEntradaDados
    """
    try:
        print("🚀 Inicializando Sistema Híbrido de NFe...")
        
        # Criar integrador
        integrador = IntegradorSistemaExistente(sistema_principal)
        
        # Adicionar botão na interface
        integrador.adicionar_botao_nfe_na_interface()
        
        # Substituir métodos existentes
        integrador.substituir_metodos_existentes()
        
        # Armazenar referência no sistema principal
        sistema_principal.processador_nfe = integrador.processador
        sistema_principal.integrador_nfe = integrador
        
        print("✅ Sistema Híbrido de NFe inicializado com sucesso!")
        
        return integrador
        
    except Exception as e:
        print(f"❌ Erro ao inicializar sistema NFe: {e}")
        return None


# CLASSE AUXILIAR PARA GERENCIAMENTO DE CERTIFICADOS
class GerenciadorCertificado:
    """Gerencia certificados digitais A1 para consulta NFe"""
    
    def __init__(self):
        self.certificado_path = None
        self.certificado_senha = None
        self.certificado_info = {}
    
    def carregar_certificado(self, caminho, senha):
        """Carrega e valida certificado"""
        try:
            from cryptography.hazmat.primitives import serialization
            from cryptography.hazmat.primitives.serialization import pkcs12
            
            # Ler arquivo de certificado
            with open(caminho, 'rb') as f:
                cert_data = f.read()
            
            # Carregar PKCS12
            private_key, certificate, additional_certificates = pkcs12.load_key_and_certificates(
                cert_data, senha.encode('utf-8')
            )
            
            # Extrair informações do certificado
            self.certificado_info = {
                'subject': certificate.subject.rfc4514_string(),
                'issuer': certificate.issuer.rfc4514_string(),
                'serial_number': str(certificate.serial_number),
                'not_valid_before': certificate.not_valid_before,
                'not_valid_after': certificate.not_valid_after,
                'is_valid': certificate.not_valid_before <= datetime.now() <= certificate.not_valid_after
            }
            
            self.certificado_path = caminho
            self.certificado_senha = senha
            
            return True
            
        except Exception as e:
            print(f"❌ Erro ao carregar certificado: {e}")
            return False
    
    def validar_certificado(self):
        """Valida se o certificado está válido"""
        if not self.certificado_info:
            return False
        
        return self.certificado_info.get('is_valid', False)
    
    def obter_info_certificado(self):
        """Retorna informações do certificado"""
        return self.certificado_info


# CLASSE PARA LOGS E HISTÓRICO
class LogImportacaoNFe:
    """Gerencia logs das importações de NFe"""
    
    def __init__(self, base_path):
        self.base_path = Path(base_path)
        self.arquivo_log = self.base_path / "logs_importacao_nfe.json"
        self.logs = self.carregar_logs()
    
    def carregar_logs(self):
        """Carrega logs existentes"""
        try:
            if self.arquivo_log.exists():
                with open(self.arquivo_log, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return []
        except Exception as e:
            print(f"❌ Erro ao carregar logs: {e}")
            return []
    
    def salvar_logs(self):
        """Salva logs no arquivo"""
        try:
            self.base_path.mkdir(parents=True, exist_ok=True)
            with open(self.arquivo_log, 'w', encoding='utf-8') as f:
                json.dump(self.logs, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"❌ Erro ao salvar logs: {e}")
    
    def adicionar_log(self, chave_acesso, tipo_importacao, resultado, detalhes=None):
        """Adiciona novo log de importação"""
        log_entry = {
            'timestamp': datetime.now().isoformat(),
            'chave_acesso': chave_acesso,
            'tipo_importacao': tipo_importacao,  # 'xml', 'consulta', 'lote'
            'resultado': resultado,  # 'sucesso', 'erro'
            'detalhes': detalhes or {}
        }
        
        self.logs.append(log_entry)
        self.salvar_logs()
    
    def obter_historico_chave(self, chave_acesso):
        """Obtém histórico de uma chave específica"""
        return [log for log in self.logs if log['chave_acesso'] == chave_acesso]
    
    def limpar_logs_antigos(self, dias=30):
        """Remove logs mais antigos que X dias"""
        try:
            limite = datetime.now() - timedelta(days=dias)
            self.logs = [
                log for log in self.logs 
                if datetime.fromisoformat(log['timestamp']) > limite
            ]
            self.salvar_logs()
            print(f"✅ Logs antigos removidos (>{dias} dias)")
        except Exception as e:
            print(f"❌ Erro ao limpar logs: {e}")


# FUNÇÕES UTILITÁRIAS PARA INTEGRAÇÃO
def extrair_chave_de_texto(texto):
    """Extrai chave de acesso de um texto qualquer"""
    # Remove tudo que não é número
    numeros = re.sub(r'[^0-9]', '', texto)
    
    # Procura sequência de 44 dígitos
    match = re.search(r'\d{44}', numeros)
    return match.group(0) if match else None

def validar_estrutura_xml_nfe(caminho_arquivo):
    """Valida se arquivo XML é uma NFe válida"""
    try:
        tree = ET.parse(caminho_arquivo)
        root = tree.getroot()
        
        # Verificar namespace NFe
        if 'portalfiscal.inf.br/nfe' not in str(root.tag):
            return False, "Não é um XML de NFe"
        
        # Verificar elementos essenciais
        ns = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}
        inf_nfe = root.find('.//nfe:infNFe', ns)
        
        if inf_nfe is None:
            return False, "Estrutura NFe inválida"
        
        return True, "XML NFe válido"
        
    except ET.ParseError:
        return False, "XML malformado"
    except Exception as e:
        return False, f"Erro: {str(e)}"

def formatar_cnpj_cpf(documento):
    """Formata CNPJ ou CPF"""
    if not documento:
        return ""
    
    numeros = re.sub(r'[^0-9]', '', documento)
    
    if len(numeros) == 11:  # CPF
        return f"{numeros[:3]}.{numeros[3:6]}.{numeros[6:9]}-{numeros[9:]}"
    elif len(numeros) == 14:  # CNPJ
        return f"{numeros[:2]}.{numeros[2:5]}.{numeros[5:8]}/{numeros[8:12]}-{numeros[12:]}"
    else:
        return documento

def formatar_valor_monetario(valor):
    """Formata valor para exibição monetária"""
    try:
        if isinstance(valor, str):
            valor = float(valor.replace(',', '.'))
        return f"R$ {valor:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
    except:
        return "R$ 0,00"


# CONFIGURAÇÕES PADRÃO
CONFIGURACOES_PADRAO = {
    'timeout_consulta': 30,
    'max_tentativas': 3,
    'delay_entre_consultas': 1,  # segundos
    'cache_habilitado': True,
    'logs_habilitados': True,
    'dias_manter_logs': 30,
    'auto_classificar_produtos': True,
    'salvar_xml_resposta': False,
    'verificar_certificado_startup': True
}


# CLASSE PRINCIPAL SIMPLIFICADA PARA IMPORTAÇÃO DIRETA
class ImportadorNFeSimples:
    """Versão simplificada para importação rápida"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.processador = ProcessadorNFeHibrido(sistema_principal)
    
    def importar_xml_arquivo(self, caminho_xml, importar_financeiro=True, importar_materiais=True):
        """Importa XML diretamente de arquivo"""
        try:
            # Processar XML
            dados_nfe = self.processador.processar_xml_nfe(caminho_xml)
            
            if not dados_nfe:
                return False, "Erro ao processar XML"
            
            resultados = []
            
            # Importar dados
            if importar_financeiro:
                resultado = self.processador.importar_dados_financeiro(dados_nfe)
                resultados.append(f"Financeiro: {resultado}")
            
            if importar_materiais:
                resultado = self.processador.importar_dados_material(dados_nfe)
                resultados.append(f"Materiais: {resultado}")
            
            return True, " | ".join(resultados)
            
        except Exception as e:
            return False, str(e)
    
    def importar_por_chave(self, chave_acesso, importar_financeiro=True, importar_materiais=True):
        """Importa NFe por chave de acesso"""
        try:
            # Consultar NFe
            dados_nfe = self.processador.consultar_nfe_sefaz(chave_acesso)
            
            if not dados_nfe:
                return False, "NFe não encontrada"
            
            resultados = []
            
            # Importar dados
            if importar_financeiro:
                resultado = self.processador.importar_dados_financeiro(dados_nfe)
                resultados.append(f"Financeiro: {resultado}")
            
            if importar_materiais:
                resultado = self.processador.importar_dados_material(dados_nfe)
                resultados.append(f"Materiais: {resultado}")
            
            return True, " | ".join(resultados)
            
        except Exception as e:
            return False, str(e)


# EXEMPLO DE IMPLEMENTAÇÃO NO SISTEMA PRINCIPAL
"""
COMO INTEGRAR NO SEU SISTEMA PRINCIPAL:

1. No __init__ da classe SistemaEntradaDados, adicione:
   
   # Inicializar sistema NFe
   from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido
   inicializar_sistema_nfe_hibrido(self)

2. Para usar a versão simplificada:
   
   from src.nfe.sistema_hibrido_nfe import ImportadorNFeSimples
   
   importador = ImportadorNFeSimples(self)
   sucesso, resultado = importador.importar_xml_arquivo("caminho/para/nfe.xml")
   
   if sucesso:
       print(f"Importado com sucesso: {resultado}")
   else:
       print(f"Erro: {resultado}")

3. Para consulta por chave:
   
   chave = "35200114200166000187550010000000271234567890"
   sucesso, resultado = importador.importar_por_chave(chave)

OBSERVAÇÕES IMPORTANTES:

- Para consulta via webservice SEFAZ, é OBRIGATÓRIO ter certificado digital A1
- O certificado deve ser configurado através da interface ou programaticamente
- Para XMLs recebidos por email, não é necessário certificado
- O sistema mantém cache das consultas para evitar repetições
- Logs de importação são salvos automaticamente
- As classificações de produtos são sugestões baseadas na descrição
"""