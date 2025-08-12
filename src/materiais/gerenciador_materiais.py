# -*- coding: utf-8 -*-
"""
Gerenciador de Materiais Corrigido
Integrado com o sistema híbrido de importação de NFe
"""

import pandas as pd
import json
from pathlib import Path
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox


class GerenciadorMateriais:
    """Classe para gerenciar materiais da obra integrada ao sistema existente"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        
        # CORREÇÃO: Usar as variáveis globais corretas
        try:
            # Tentar usar BASE_PATH global se existir
            if hasattr(sistema_principal, 'BASE_PATH'):
                self.base_path = sistema_principal.BASE_PATH
                print(f"✅ Usando BASE_PATH do sistema: {self.base_path}")
            else:
                # Usar caminho padrão baseado no sistema
                self.base_path = Path.cwd() / "data" / "materiais"
                print(f"⚠️ Usando caminho padrão: {self.base_path}")
                
        except Exception as e:
            print(f"⚠️ Erro ao determinar BASE_PATH: {e}")
            self.base_path = Path.cwd() / "data" / "materiais"
        
        # Criar pasta se não existir
        self.base_path.mkdir(parents=True, exist_ok=True)
        
        # Definir arquivos
        self.arquivo_materiais = self.base_path / "materiais_obra.xlsx"
        self.arquivo_parametros = self.base_path / "parametros_materiais.json"
        
        print(f"📁 Arquivo materiais: {self.arquivo_materiais}")
        print(f"⚙️ Arquivo parâmetros: {self.arquivo_parametros}")
        
        # Carregar configurações
        self.carregar_parametros()
        
        # Inicializar planilha
        self.inicializar_planilha_materiais()
    
    def carregar_parametros(self):
        """Carrega parâmetros de categorias e configurações"""
        try:
            if self.arquivo_parametros.exists():
                with open(self.arquivo_parametros, 'r', encoding='utf-8') as f:
                    self.parametros = json.load(f)
                print("✅ Parâmetros carregados do arquivo")
            else:
                self.criar_parametros_padrao()
                print("✅ Parâmetros padrão criados")
                
        except Exception as e:
            print(f"❌ Erro ao carregar parâmetros: {e}")
            self.criar_parametros_padrao()
    
    def criar_parametros_padrao(self):
        """Cria arquivo de parâmetros padrão"""
        self.parametros = {
            "categorias_materiais": {
                "ACABAMENTOS": {
                    "subcategorias": [
                        "CERAMICA", "PORCELANATO", "PISO_LAMINADO", "RODAPE", 
                        "MOLDURA", "REJUNTE", "GESSO", "FORRO"
                    ],
                    "cor": "#4CAF50"
                },
                "TINTAS": {
                    "subcategorias": [
                        "TINTA_LATEX", "TINTA_ACRILICA", "VERNIZ", "ESMALTE", 
                        "PRIMER", "SELADOR", "MASSA_CORRIDA"
                    ],
                    "cor": "#FF9800"
                },
                "ELETRICO": {
                    "subcategorias": [
                        "CABOS", "TOMADAS", "INTERRUPTORES", "ILUMINACAO", 
                        "PROTECAO", "CONDUTOS"
                    ],
                    "cor": "#2196F3"
                },
                "HIDRAULICO": {
                    "subcategorias": [
                        "TUBULACAO", "CONEXOES", "METAIS", "LOUÇAS", 
                        "VALVULAS", "BOMBAS"
                    ],
                    "cor": "#00BCD4"
                },
                "ESTRUTURAL": {
                    "subcategorias": [
                        "CONCRETO", "FERRAGEM", "ALVENARIA", "MADEIRAMENTO", 
                        "METALICO", "FUNDACAO"
                    ],
                    "cor": "#795548"
                },
                "ESQUADRIAS": {
                    "subcategorias": [
                        "PORTAS", "JANELAS", "FERRAGENS", "VIDROS", 
                        "MARCOS", "BATENTES"
                    ],
                    "cor": "#607D8B"
                },
                "PAISAGISMO": {
                    "subcategorias": [
                        "PLANTAS", "TERRA_SUBSTRATO", "IRRIGACAO", 
                        "DECORATIVO", "FERRAMENTAS"
                    ],
                    "cor": "#8BC34A"
                },
                "OUTROS": {
                    "subcategorias": [
                        "DIVERSOS", "FERRAMENTAS", "LIMPEZA", "SEGURANCA"
                    ],
                    "cor": "#9E9E9E"
                }
            },
            "ambientes": [
                "INSTALAÇÃO DA OBRA",
                "SALA DE ESTAR", 
                "SALA DE JANTAR",
                "COZINHA",
                "DORMITÓRIO SUITE",
                "DORMITÓRIO 1",
                "DORMITÓRIO 2", 
                "DORMITÓRIO 3",
                "BANHEIRO SUITE",
                "BANHEIRO SOCIAL",
                "LAVABO",
                "ÁREA DE SERVIÇO",
                "VARANDA",
                "ÁREA EXTERNA",
                "GARAGEM",
                "DEPÓSITO",
                "GERAL"
            ],
            "unidades": [
                "UN", "M", "M²", "M³", "KG", "G", "L", "ML", 
                "PC", "CX", "SC", "PAR", "JG", "KIT", "ROL", "LT"
            ],
            "status_instalacao": [
                "PENDENTE",
                "INSTALADO", 
                "EM_INSTALACAO",
                "AGUARDANDO_MATERIAL",
                "AGUARDANDO_INSTALADOR",
                "DEFEITO",
                "SUBSTITUIDO",
                "CANCELADO"
            ],
            "config": {
                "auto_backup": True,
                "validar_dados": True,
                "log_alteracoes": True
            }
        }
        
        # Salvar arquivo
        try:
            with open(self.arquivo_parametros, 'w', encoding='utf-8') as f:
                json.dump(self.parametros, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"❌ Erro ao salvar parâmetros: {e}")
    
    def inicializar_planilha_materiais(self):
        """Cria planilha de materiais se não existir"""
        if not self.arquivo_materiais.exists():
            try:
                # Estrutura completa da planilha
                colunas = [
                    # Identificação
                    'ID', 'Cliente', 'Data_Cadastro', 'Data_Ultima_Atualizacao',
                    
                    # Classificação
                    'Categoria', 'Subcategoria', 'Codigo_Produto',
                    
                    # Descrição
                    'Descricao_Completa', 'Marca', 'Modelo', 'Cor_Acabamento', 
                    'Dimensoes', 'Especificacoes_Tecnicas',
                    
                    # Dados de compra
                    'Tem_Dados_Compra', 'Data_Compra', 'CNPJ_Fornecedor', 
                    'Nome_Fornecedor', 'Numero_NF', 'Item_NF',
                    
                    # Quantidades e valores
                    'Quantidade', 'Unidade', 'Valor_Unitario', 'Valor_Total',
                    
                    # Localização e instalação
                    'Ambiente_Aplicacao', 'Localizacao_Especifica',
                    'Data_Instalacao', 'Instalador', 'Status_Instalacao',
                    
                    # Garantia e manutenção
                    'Garantia_Meses', 'Data_Fim_Garantia', 'Manutencao_Preventiva',
                    
                    # Documentação
                    'Observacoes', 'Foto_Produto', 'Manual_Fabricante', 
                    'Certificados', 'Origem_Dados'
                ]
                
                df = pd.DataFrame(columns=colunas)
                df.to_excel(self.arquivo_materiais, index=False)
                
                print(f"✅ Planilha de materiais criada: {self.arquivo_materiais}")
                
            except Exception as e:
                print(f"❌ Erro ao criar planilha de materiais: {e}")
                raise e
    
    def salvar_material(self, dados_material):
        """Salva um material na planilha"""
        try:
            # Carregar planilha existente
            try:
                df_existente = pd.read_excel(self.arquivo_materiais)
                print(f"📊 Planilha carregada: {len(df_existente)} registros")
            except FileNotFoundError:
                print("📊 Criando nova planilha...")
                self.inicializar_planilha_materiais()
                df_existente = pd.read_excel(self.arquivo_materiais)
            
            # Gerar ID único
            if len(df_existente) > 0 and 'ID' in df_existente.columns:
                ultimo_id = int(df_existente['ID'].max())
            else:
                ultimo_id = 0
            
            # Preparar dados do material
            novo_id = ultimo_id + 1
            dados_material['ID'] = novo_id
            dados_material['Data_Cadastro'] = datetime.now().strftime('%d/%m/%Y')
            dados_material['Data_Ultima_Atualizacao'] = datetime.now().strftime('%d/%m/%Y')
            dados_material['Origem_Dados'] = dados_material.get('Origem_Dados', 'CADASTRO_MANUAL')
            
            # Calcular data fim garantia se tiver garantia
            if dados_material.get('Garantia_Meses') and dados_material.get('Data_Instalacao'):
                try:
                    from dateutil.relativedelta import relativedelta
                    data_inst = datetime.strptime(dados_material['Data_Instalacao'], '%d/%m/%Y')
                    meses_garantia = int(dados_material['Garantia_Meses'])
                    data_fim = data_inst + relativedelta(months=meses_garantia)
                    dados_material['Data_Fim_Garantia'] = data_fim.strftime('%d/%m/%Y')
                except:
                    dados_material['Data_Fim_Garantia'] = ''
            
            # Validar dados essenciais
            if not dados_material.get('Categoria'):
                dados_material['Categoria'] = 'OUTROS'
            
            if not dados_material.get('Descricao_Completa'):
                raise ValueError("Descrição é obrigatória")
            
            # Adicionar nova linha
            nova_linha = pd.DataFrame([dados_material])
            df_atualizado = pd.concat([df_existente, nova_linha], ignore_index=True)
            
            # Salvar planilha
            df_atualizado.to_excel(self.arquivo_materiais, index=False)
            
            print(f"✅ Material salvo: ID {novo_id} - {dados_material.get('Descricao_Completa', '')[:50]}")
            
            return novo_id
            
        except Exception as e:
            print(f"❌ Erro ao salvar material: {e}")
            raise e
    
    def carregar_materiais_cliente(self, cliente):
        """Carrega materiais de um cliente específico"""
        try:
            if not self.arquivo_materiais.exists():
                print("⚠️ Arquivo de materiais não existe")
                return pd.DataFrame()
            
            df = pd.read_excel(self.arquivo_materiais)
            
            if 'Cliente' in df.columns:
                df_cliente = df[df['Cliente'] == cliente]
                print(f"📊 Carregados {len(df_cliente)} materiais do cliente {cliente}")
                return df_cliente
            else:
                print(f"📊 Carregados {len(df)} materiais (sem filtro de cliente)")
                return df
                
        except Exception as e:
            print(f"❌ Erro ao carregar materiais: {e}")
            return pd.DataFrame()
    
    def atualizar_material(self, material_id, dados_atualizacao):
        """Atualiza um material existente"""
        try:
            df = pd.read_excel(self.arquivo_materiais)
            
            # Encontrar material
            mask = df['ID'] == material_id
            if not mask.any():
                raise ValueError(f"Material ID {material_id} não encontrado")
            
            # Atualizar dados
            for campo, valor in dados_atualizacao.items():
                if campo in df.columns:
                    df.loc[mask, campo] = valor
            
            # Atualizar timestamp
            df.loc[mask, 'Data_Ultima_Atualizacao'] = datetime.now().strftime('%d/%m/%Y')
            
            # Salvar
            df.to_excel(self.arquivo_materiais, index=False)
            
            print(f"✅ Material ID {material_id} atualizado")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao atualizar material: {e}")
            return False
    
    def excluir_material(self, material_id):
        """Exclui um material (marca como excluído)"""
        try:
            return self.atualizar_material(material_id, {
                'Status_Instalacao': 'CANCELADO',
                'Observacoes': f"EXCLUÍDO EM {datetime.now().strftime('%d/%m/%Y')}"
            })
        except Exception as e:
            print(f"❌ Erro ao excluir material: {e}")
            return False
    
    def buscar_materiais(self, filtros):
        """Busca materiais com filtros"""
        try:
            df = pd.read_excel(self.arquivo_materiais)
            
            # Aplicar filtros
            for campo, valor in filtros.items():
                if campo in df.columns and valor:
                    if campo in ['Categoria', 'Status_Instalacao', 'Ambiente_Aplicacao']:
                        df = df[df[campo] == valor]
                    else:
                        # Busca textual (contains)
                        df = df[df[campo].astype(str).str.contains(str(valor), case=False, na=False)]
            
            return df
            
        except Exception as e:
            print(f"❌ Erro na busca: {e}")
            return pd.DataFrame()
    
    def gerar_relatorio_resumo(self, cliente=None):
        """Gera relatório resumo dos materiais"""
        try:
            if cliente:
                df = self.carregar_materiais_cliente(cliente)
            else:
                df = pd.read_excel(self.arquivo_materiais)
            
            if len(df) == 0:
                return {"erro": "Nenhum material encontrado"}
            
            resumo = {
                "total_materiais": len(df),
                "total_valor": df['Valor_Total'].fillna(0).sum() if 'Valor_Total' in df.columns else 0,
                "por_categoria": df['Categoria'].value_counts().to_dict() if 'Categoria' in df.columns else {},
                "por_status": df['Status_Instalacao'].value_counts().to_dict() if 'Status_Instalacao' in df.columns else {},
                "por_ambiente": df['Ambiente_Aplicacao'].value_counts().to_dict() if 'Ambiente_Aplicacao' in df.columns else {},
                "com_dados_nf": len(df[df['Tem_Dados_Compra'] == True]) if 'Tem_Dados_Compra' in df.columns else 0,
                "sem_dados_nf": len(df[df['Tem_Dados_Compra'] == False]) if 'Tem_Dados_Compra' in df.columns else 0
            }
            
            return resumo
            
        except Exception as e:
            print(f"❌ Erro ao gerar relatório: {e}")
            return {"erro": str(e)}
    
    def exportar_para_csv(self, cliente=None, caminho_arquivo=None):
        """Exporta materiais para CSV"""
        try:
            if cliente:
                df = self.carregar_materiais_cliente(cliente)
                nome_padrao = f"materiais_{cliente}_{datetime.now().strftime('%Y%m%d')}.csv"
            else:
                df = pd.read_excel(self.arquivo_materiais)
                nome_padrao = f"materiais_todos_{datetime.now().strftime('%Y%m%d')}.csv"
            
            if caminho_arquivo is None:
                caminho_arquivo = self.base_path / nome_padrao
            
            df.to_csv(caminho_arquivo, index=False, encoding='utf-8-sig', sep=';')
            
            print(f"✅ Dados exportados para: {caminho_arquivo}")
            return str(caminho_arquivo)
            
        except Exception as e:
            print(f"❌ Erro ao exportar: {e}")
            return None
    
    def fazer_backup(self):
        """Cria backup da planilha de materiais"""
        try:
            if not self.arquivo_materiais.exists():
                return False
            
            # Pasta de backup
            pasta_backup = self.base_path / "backups"
            pasta_backup.mkdir(exist_ok=True)
            
            # Nome do backup
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            nome_backup = f"materiais_backup_{timestamp}.xlsx"
            caminho_backup = pasta_backup / nome_backup
            
            # Copiar arquivo
            import shutil
            shutil.copy2(self.arquivo_materiais, caminho_backup)
            
            print(f"✅ Backup criado: {caminho_backup}")
            
            # Limpar backups antigos (manter últimos 10)
            self.limpar_backups_antigos(pasta_backup)
            
            return str(caminho_backup)
            
        except Exception as e:
            print(f"❌ Erro ao criar backup: {e}")
            return None
    
    def limpar_backups_antigos(self, pasta_backup, manter=10):
        """Remove backups antigos, mantendo apenas os últimos N"""
        try:
            backups = list(pasta_backup.glob("materiais_backup_*.xlsx"))
            if len(backups) > manter:
                # Ordenar por data de modificação
                backups.sort(key=lambda x: x.stat().st_mtime)
                
                # Remover os mais antigos
                for backup in backups[:-manter]:
                    backup.unlink()
                    print(f"🗑️ Backup antigo removido: {backup.name}")
                    
        except Exception as e:
            print(f"⚠️ Erro ao limpar backups: {e}")


class IntegradorMateriais:
    """Integra o gerenciador de materiais com o sistema principal"""
    
    def __init__(self, sistema_principal):
        self.sistema = sistema_principal
        self.gerenciador = GerenciadorMateriais(sistema_principal)
        
        # Adicionar referência no sistema principal
        sistema_principal.gerenciador_materiais = self.gerenciador
    
    def adicionar_botoes_interface(self):
        """Adiciona botões de materiais na interface principal"""
        try:
            # Verificar se existe aba fornecedor
            if hasattr(self.sistema, 'aba_fornecedor'):
                self.adicionar_secao_materiais()
                print("✅ Seção de materiais adicionada na aba fornecedor")
                
        except Exception as e:
            print(f"❌ Erro ao adicionar botões: {e}")
    
    def adicionar_secao_materiais(self):
        """Adiciona seção completa de materiais"""
        # Frame principal de materiais
        frame_materiais = ttk.LabelFrame(self.sistema.aba_fornecedor, 
                                        text="🏗️ Gestão de Materiais da Obra", 
                                        padding=10)
        frame_materiais.pack(fill='x', padx=10, pady=5)
        
        # Container para botões
        frame_botoes = ttk.Frame(frame_materiais)
        frame_botoes.pack(fill='x', pady=5)
        
        # Botões principais
        botoes = [
            ("📦 Novo Material", self.abrir_cadastro_material),
            ("📋 Consultar Materiais", self.abrir_consulta_materiais),
            ("📄 Manual do Proprietário", self.gerar_manual_proprietario),
            ("📊 Relatórios", self.abrir_relatorios),
            ("💾 Backup", self.fazer_backup_materiais)
        ]
        
        for texto, comando in botoes:
            ttk.Button(frame_botoes, text=texto, command=comando,
                      style='Medium.TButton').pack(side='left', padx=5)
        
        # Status/Info
        self.label_status_materiais = tk.Label(frame_materiais, 
                                             text="Sistema de materiais carregado", 
                                             fg='green', font=('Arial', 8))
        self.label_status_materiais.pack(anchor='w', pady=2)
        
        # Atualizar status
        self.atualizar_status_materiais()
    
    def atualizar_status_materiais(self):
        """Atualiza status dos materiais na interface"""
        try:
            cliente_atual = getattr(self.sistema, 'cliente_atual', None)
            if cliente_atual:
                resumo = self.gerenciador.gerar_relatorio_resumo(cliente_atual)
                if 'erro' not in resumo:
                    texto = f"Cliente: {cliente_atual} | {resumo['total_materiais']} materiais | R$ {resumo['total_valor']:,.2f}"
                    self.label_status_materiais.config(text=texto, fg='blue')
                else:
                    self.label_status_materiais.config(text="Nenhum material cadastrado", fg='gray')
            else:
                self.label_status_materiais.config(text="Nenhum cliente selecionado", fg='orange')
                
        except Exception as e:
            self.label_status_materiais.config(text=f"Erro: {str(e)}", fg='red')
    
    def abrir_cadastro_material(self):
        """Abre janela de cadastro de material"""
        try:
            CadastroMaterial(self.sistema, self.gerenciador)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir cadastro:\n{str(e)}")
    
    def abrir_consulta_materiais(self):
        """Abre janela de consulta de materiais"""
        try:
            ConsultaMateriais(self.sistema, self.gerenciador)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir consulta:\n{str(e)}")
    
    def gerar_manual_proprietario(self):
        """Gera manual do proprietário"""
        try:
            cliente_atual = getattr(self.sistema, 'cliente_atual', None)
            if not cliente_atual:
                messagebox.showwarning("Aviso", "Selecione um cliente primeiro!")
                return
            
            GeradorManual(self.sistema, self.gerenciador, cliente_atual)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao gerar manual:\n{str(e)}")
    
    def abrir_relatorios(self):
        """Abre janela de relatórios"""
        try:
            RelatoriosMateriais(self.sistema, self.gerenciador)
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao abrir relatórios:\n{str(e)}")
    
    def fazer_backup_materiais(self):
        """Faz backup dos materiais"""
        try:
            caminho_backup = self.gerenciador.fazer_backup()
            if caminho_backup:
                messagebox.showinfo("Backup", f"Backup criado com sucesso!\n\n{caminho_backup}")
            else:
                messagebox.showerror("Erro", "Erro ao criar backup!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro no backup:\n{str(e)}")


class CadastroMaterial:
    """Janela de cadastro de material"""
    
    def __init__(self, sistema, gerenciador):
        self.sistema = sistema
        self.gerenciador = gerenciador
        
        self.criar_janela()
        self.criar_interface()
        self.carregar_dados_iniciais()
    
    def criar_janela(self):
        """Cria janela principal"""
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("Cadastro de Material")
        self.janela.geometry("800x600")
        self.janela.grab_set()
        
        # Centralizar janela
        self.janela.transient(self.sistema.root)
    
    def criar_interface(self):
        """Cria interface do cadastro"""
        # Notebook para organizar abas
        notebook = ttk.Notebook(self.janela)
        notebook.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Aba 1: Dados Básicos
        self.criar_aba_dados_basicos(notebook)
        
        # Aba 2: Compra e Fornecedor
        self.criar_aba_compra(notebook)
        
        # Aba 3: Instalação
        self.criar_aba_instalacao(notebook)
        
        # Botões principais
        self.criar_botoes_principais()
    
    def criar_aba_dados_basicos(self, notebook):
        """Cria aba de dados básicos"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="📋 Dados Básicos")
        
        # Scroll para a aba
        canvas = tk.Canvas(frame)
        scrollbar = ttk.Scrollbar(frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)
        
        # Seção Identificação
        frame_id = ttk.LabelFrame(scrollable_frame, text="Identificação", padding=10)
        frame_id.pack(fill='x', padx=10, pady=5)
        
        self.campos = {}
        
        # Linha 1: Categoria e Subcategoria
        tk.Label(frame_id, text="Categoria:*", font=('Arial', 9, 'bold')).grid(row=0, column=0, sticky='w', pady=5)
        self.campos['categoria'] = ttk.Combobox(frame_id, width=20, state='readonly')
        self.campos['categoria'].grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        self.campos['categoria'].bind('<<ComboboxSelected>>', self.atualizar_subcategorias)
        
        tk.Label(frame_id, text="Subcategoria:").grid(row=0, column=2, sticky='w', pady=5)
        self.campos['subcategoria'] = ttk.Combobox(frame_id, width=20, state='readonly')
        self.campos['subcategoria'].grid(row=0, column=3, sticky='ew', padx=5, pady=5)
        
        # Linha 2: Código e Descrição
        tk.Label(frame_id, text="Código:").grid(row=1, column=0, sticky='w', pady=5)
        self.campos['codigo'] = tk.Entry(frame_id, width=15)
        self.campos['codigo'].grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(frame_id, text="Descrição:*", font=('Arial', 9, 'bold')).grid(row=1, column=2, sticky='w', pady=5)
        self.campos['descricao'] = tk.Entry(frame_id, width=40)
        self.campos['descricao'].grid(row=1, column=3, sticky='ew', padx=5, pady=5)
        
        # Linha 3: Marca e Modelo
        tk.Label(frame_id, text="Marca:").grid(row=2, column=0, sticky='w', pady=5)
        self.campos['marca'] = tk.Entry(frame_id, width=20)
        self.campos['marca'].grid(row=2, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(frame_id, text="Modelo:").grid(row=2, column=2, sticky='w', pady=5)
        self.campos['modelo'] = tk.Entry(frame_id, width=20)
        self.campos['modelo'].grid(row=2, column=3, sticky='ew', padx=5, pady=5)
        
        # Linha 4: Cor e Dimensões
        tk.Label(frame_id, text="Cor/Acabamento:").grid(row=3, column=0, sticky='w', pady=5)
        self.campos['cor'] = tk.Entry(frame_id, width=20)
        self.campos['cor'].grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(frame_id, text="Dimensões:").grid(row=3, column=2, sticky='w', pady=5)
        self.campos['dimensoes'] = tk.Entry(frame_id, width=20)
        self.campos['dimensoes'].grid(row=3, column=3, sticky='ew', padx=5, pady=5)
        
        # Seção Especificações
        frame_spec = ttk.LabelFrame(scrollable_frame, text="Especificações Técnicas", padding=10)
        frame_spec.pack(fill='both', expand=True, padx=10, pady=5)
        
        tk.Label(frame_spec, text="Especificações:").pack(anchor='w')
        self.campos['especificacoes'] = tk.Text(frame_spec, height=4, wrap='word')
        self.campos['especificacoes'].pack(fill='both', expand=True, padx=5, pady=5)
        
        # Configurar grid
        frame_id.columnconfigure(1, weight=1)
        frame_id.columnconfigure(3, weight=2)
        
        # Configurar scroll
        scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def criar_aba_compra(self, notebook):
        """Cria aba de dados de compra"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="💰 Compra")
        
        # Checkbox para habilitar dados de compra
        self.tem_dados_compra = tk.BooleanVar()
        cb_frame = ttk.Frame(frame)
        cb_frame.pack(fill='x', padx=10, pady=10)
        
        cb = tk.Checkbutton(cb_frame, text="Este material possui dados de compra/fornecedor", 
                           variable=self.tem_dados_compra, command=self.toggle_campos_compra,
                           font=('Arial', 10, 'bold'))
        cb.pack(anchor='w')
        
        # Frame para dados de compra
        self.frame_compra = ttk.LabelFrame(frame, text="Dados da Compra", padding=10)
        self.frame_compra.pack(fill='x', padx=10, pady=5)
        
        # Fornecedor
        tk.Label(self.frame_compra, text="Fornecedor:").grid(row=0, column=0, sticky='w', pady=5)
        self.campos['fornecedor'] = tk.Entry(self.frame_compra, width=40, state='disabled')
        self.campos['fornecedor'].grid(row=0, column=1, columnspan=2, sticky='ew', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="CNPJ/CPF:").grid(row=1, column=0, sticky='w', pady=5)
        self.campos['cnpj_fornecedor'] = tk.Entry(self.frame_compra, width=20, state='disabled')
        self.campos['cnpj_fornecedor'].grid(row=1, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="Data Compra:").grid(row=1, column=2, sticky='w', pady=5)
        self.campos['data_compra'] = tk.Entry(self.frame_compra, width=15, state='disabled')
        self.campos['data_compra'].grid(row=1, column=3, sticky='ew', padx=5, pady=5)
        
        # Valores
        tk.Label(self.frame_compra, text="Quantidade:").grid(row=2, column=0, sticky='w', pady=5)
        self.campos['quantidade'] = tk.Entry(self.frame_compra, width=10, state='disabled')
        self.campos['quantidade'].grid(row=2, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="Unidade:").grid(row=2, column=2, sticky='w', pady=5)
        self.campos['unidade'] = ttk.Combobox(self.frame_compra, width=10, state='disabled')
        self.campos['unidade'].grid(row=2, column=3, sticky='w', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="Valor Unit.:").grid(row=3, column=0, sticky='w', pady=5)
        self.campos['valor_unitario'] = tk.Entry(self.frame_compra, width=15, state='disabled')
        self.campos['valor_unitario'].grid(row=3, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="Valor Total:").grid(row=3, column=2, sticky='w', pady=5)
        self.campos['valor_total'] = tk.Entry(self.frame_compra, width=15, state='disabled')
        self.campos['valor_total'].grid(row=3, column=3, sticky='ew', padx=5, pady=5)
        
        # NF
        tk.Label(self.frame_compra, text="Número NF:").grid(row=4, column=0, sticky='w', pady=5)
        self.campos['numero_nf'] = tk.Entry(self.frame_compra, width=20, state='disabled')
        self.campos['numero_nf'].grid(row=4, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(self.frame_compra, text="Item NF:").grid(row=4, column=2, sticky='w', pady=5)
        self.campos['item_nf'] = tk.Entry(self.frame_compra, width=10, state='disabled')
        self.campos['item_nf'].grid(row=4, column=3, sticky='w', padx=5, pady=5)
        
        # Configurar grid
        self.frame_compra.columnconfigure(1, weight=1)
        self.frame_compra.columnconfigure(3, weight=1)
    
    def criar_aba_instalacao(self, notebook):
        """Cria aba de instalação"""
        frame = ttk.Frame(notebook)
        notebook.add(frame, text="🔧 Instalação")
        
        # Localização
        frame_local = ttk.LabelFrame(frame, text="Localização", padding=10)
        frame_local.pack(fill='x', padx=10, pady=5)
        
        tk.Label(frame_local, text="Ambiente:").grid(row=0, column=0, sticky='w', pady=5)
        self.campos['ambiente'] = ttk.Combobox(frame_local, width=25, state='readonly')
        self.campos['ambiente'].grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        
        tk.Label(frame_local, text="Localização Específica:").grid(row=0, column=2, sticky='w', pady=5)
        self.campos['localizacao_especifica'] = tk.Entry(frame_local, width=30)
        self.campos['localizacao_especifica'].grid(row=0, column=3, sticky='ew', padx=5, pady=5)
        
        # Instalação
        frame_inst = ttk.LabelFrame(frame, text="Dados da Instalação", padding=10)
        frame_inst.pack(fill='x', padx=10, pady=5)
        
        tk.Label(frame_inst, text="Data Instalação:").grid(row=0, column=0, sticky='w', pady=5)
        self.campos['data_instalacao'] = tk.Entry(frame_inst, width=15)
        self.campos['data_instalacao'].grid(row=0, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(frame_inst, text="Instalador:").grid(row=0, column=2, sticky='w', pady=5)
        self.campos['instalador'] = tk.Entry(frame_inst, width=30)
        self.campos['instalador'].grid(row=0, column=3, sticky='ew', padx=5, pady=5)
        
        tk.Label(frame_inst, text="Status:").grid(row=1, column=0, sticky='w', pady=5)
        self.campos['status'] = ttk.Combobox(frame_inst, width=15, state='readonly')
        self.campos['status'].grid(row=1, column=1, sticky='w', padx=5, pady=5)
        
        tk.Label(frame_inst, text="Garantia (meses):").grid(row=1, column=2, sticky='w', pady=5)
        self.campos['garantia_meses'] = tk.Entry(frame_inst, width=10)
        self.campos['garantia_meses'].grid(row=1, column=3, sticky='w', padx=5, pady=5)
        
        # Observações
        frame_obs = ttk.LabelFrame(frame, text="Observações", padding=10)
        frame_obs.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.campos['observacoes'] = tk.Text(frame_obs, height=6, wrap='word')
        self.campos['observacoes'].pack(fill='both', expand=True, padx=5, pady=5)
        
        # Configurar grids
        frame_local.columnconfigure(1, weight=1)
        frame_local.columnconfigure(3, weight=1)
        frame_inst.columnconfigure(3, weight=1)
    
    def criar_botoes_principais(self):
        """Cria botões principais"""
        frame_botoes = ttk.Frame(self.janela)
        frame_botoes.pack(fill='x', padx=10, pady=10)
        
        ttk.Button(frame_botoes, text="💾 Salvar Material", 
                  command=self.salvar_material).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, text="🧹 Limpar Campos", 
                  command=self.limpar_campos).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, text="❌ Cancelar", 
                  command=self.janela.destroy).pack(side='right', padx=5)
    
    def carregar_dados_iniciais(self):
        """Carrega dados iniciais nos comboboxes"""
        try:
            parametros = self.gerenciador.parametros
            
            # Categorias
            categorias = list(parametros['categorias_materiais'].keys())
            self.campos['categoria']['values'] = categorias
            
            # Ambientes
            self.campos['ambiente']['values'] = parametros['ambientes']
            
            # Unidades
            self.campos['unidade']['values'] = parametros['unidades']
            
            # Status
            self.campos['status']['values'] = parametros['status_instalacao']
            self.campos['status'].set('PENDENTE')
            
            # Valores padrão
            self.campos['quantidade'].insert(0, '1')
            self.campos['garantia_meses'].insert(0, '12')
            
        except Exception as e:
            print(f"❌ Erro ao carregar dados iniciais: {e}")
    
    def atualizar_subcategorias(self, event=None):
        """Atualiza subcategorias baseado na categoria"""
        try:
            categoria = self.campos['categoria'].get()
            parametros = self.gerenciador.parametros
            
            if categoria in parametros['categorias_materiais']:
                subcategorias = parametros['categorias_materiais'][categoria]['subcategorias']
                self.campos['subcategoria']['values'] = subcategorias
                self.campos['subcategoria'].set('')
            else:
                self.campos['subcategoria']['values'] = []
                self.campos['subcategoria'].set('')
                
        except Exception as e:
            print(f"❌ Erro ao atualizar subcategorias: {e}")
    
    def toggle_campos_compra(self):
        """Habilita/desabilita campos de compra"""
        estado = 'normal' if self.tem_dados_compra.get() else 'disabled'
        estado_combo = 'readonly' if self.tem_dados_compra.get() else 'disabled'
        
        campos_compra = [
            'fornecedor', 'cnpj_fornecedor', 'data_compra', 'quantidade',
            'valor_unitario', 'valor_total', 'numero_nf', 'item_nf'
        ]
        
        for campo in campos_compra:
            if campo in self.campos:
                self.campos[campo].config(state=estado)
        
        # Combobox tem estado diferente
        if 'unidade' in self.campos:
            self.campos['unidade'].config(state=estado_combo)
    
    def validar_dados(self):
        """Valida dados do formulário"""
        erros = []
        
        # Campos obrigatórios
        if not self.campos['categoria'].get().strip():
            erros.append("Categoria é obrigatória")
        
        if not self.campos['descricao'].get().strip():
            erros.append("Descrição é obrigatória")
        
        # Validações de dados de compra
        if self.tem_dados_compra.get():
            if self.campos['quantidade'].get() and not self.validar_numero(self.campos['quantidade'].get()):
                erros.append("Quantidade deve ser um número válido")
            
            if self.campos['valor_unitario'].get() and not self.validar_numero(self.campos['valor_unitario'].get()):
                erros.append("Valor unitário deve ser um número válido")
        
        return erros
    
    def validar_numero(self, valor):
        """Valida se valor é um número"""
        try:
            float(valor.replace(',', '.'))
            return True
        except:
            return False
    
    def salvar_material(self):
        """Salva o material"""
        try:
            # Validar dados
            erros = self.validar_dados()
            if erros:
                messagebox.showerror("Erro de Validação", "\n".join(erros))
                return
            
            # Preparar dados
            dados_material = {
                'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                'Categoria': self.campos['categoria'].get().strip(),
                'Subcategoria': self.campos['subcategoria'].get().strip(),
                'Codigo_Produto': self.campos['codigo'].get().strip(),
                'Descricao_Completa': self.campos['descricao'].get().strip(),
                'Marca': self.campos['marca'].get().strip(),
                'Modelo': self.campos['modelo'].get().strip(),
                'Cor_Acabamento': self.campos['cor'].get().strip(),
                'Dimensoes': self.campos['dimensoes'].get().strip(),
                'Especificacoes_Tecnicas': self.campos['especificacoes'].get('1.0', tk.END).strip(),
                'Ambiente_Aplicacao': self.campos['ambiente'].get().strip(),
                'Localizacao_Especifica': self.campos['localizacao_especifica'].get().strip(),
                'Data_Instalacao': self.campos['data_instalacao'].get().strip(),
                'Instalador': self.campos['instalador'].get().strip(),
                'Status_Instalacao': self.campos['status'].get() or 'PENDENTE',
                'Garantia_Meses': self.campos['garantia_meses'].get().strip() or '0',
                'Observacoes': self.campos['observacoes'].get('1.0', tk.END).strip(),
                'Tem_Dados_Compra': self.tem_dados_compra.get(),
                'Origem_Dados': 'CADASTRO_MANUAL'
            }
            
            # Adicionar dados de compra se habilitado
            if self.tem_dados_compra.get():
                dados_material.update({
                    'Nome_Fornecedor': self.campos['fornecedor'].get().strip(),
                    'CNPJ_Fornecedor': self.campos['cnpj_fornecedor'].get().strip(),
                    'Data_Compra': self.campos['data_compra'].get().strip(),
                    'Quantidade': self.campos['quantidade'].get().strip(),
                    'Unidade': self.campos['unidade'].get().strip(),
                    'Valor_Unitario': self.campos['valor_unitario'].get().strip(),
                    'Valor_Total': self.campos['valor_total'].get().strip(),
                    'Numero_NF': self.campos['numero_nf'].get().strip(),
                    'Item_NF': self.campos['item_nf'].get().strip()
                })
            
            # Salvar
            material_id = self.gerenciador.salvar_material(dados_material)
            
            messagebox.showinfo("Sucesso", 
                f"Material cadastrado com sucesso!\n\nID: {material_id}\nDescrição: {dados_material['Descricao_Completa']}")
            
            # Perguntar se quer cadastrar outro
            if messagebox.askyesno("Novo Material", "Deseja cadastrar outro material?"):
                self.limpar_campos()
            else:
                self.janela.destroy()
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar material:\n{str(e)}")
    
    def limpar_campos(self):
        """Limpa todos os campos do formulário"""
        try:
            # Limpar campos de texto
            for campo, widget in self.campos.items():
                if isinstance(widget, tk.Entry):
                    widget.delete(0, tk.END)
                elif isinstance(widget, ttk.Combobox):
                    widget.set('')
                elif isinstance(widget, tk.Text):
                    widget.delete('1.0', tk.END)
            
            # Resetar valores padrão
            self.campos['status'].set('PENDENTE')
            self.campos['quantidade'].insert(0, '1')
            self.campos['garantia_meses'].insert(0, '12')
            
            # Resetar checkbox
            self.tem_dados_compra.set(False)
            self.toggle_campos_compra()
            
        except Exception as e:
            print(f"❌ Erro ao limpar campos: {e}")


# FUNÇÃO PRINCIPAL DE INICIALIZAÇÃO
def inicializar_sistema_materiais_completo(sistema_principal):
    """
    Inicializa o sistema completo de materiais
    
    Args:
        sistema_principal: Instância da classe SistemaEntradaDados
    
    Returns:
        IntegradorMateriais: Instância do integrador
    """
    try:
        print("🚀 Inicializando Sistema Completo de Materiais...")
        
        # Criar integrador
        integrador = IntegradorMateriais(sistema_principal)
        
        # Adicionar interface na aba fornecedor
        integrador.adicionar_botoes_interface()
        
        # Fazer backup inicial se existir dados
        integrador.gerenciador.fazer_backup()
        
        print("✅ Sistema de Materiais inicializado com sucesso!")
        print(f"📁 Pasta base: {integrador.gerenciador.base_path}")
        print(f"📊 Arquivo de dados: {integrador.gerenciador.arquivo_materiais}")
        
        return integrador
        
    except Exception as e:
        print(f"❌ Erro ao inicializar sistema de materiais: {e}")
        return None