# -*- coding: utf-8 -*-
"""
PATCH FINAL OTIMIZADO - Interface Unificada
Corrige data (período ATUAL) e fluxo direto para configuração
"""

from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import json
from pathlib import Path

def aplicar_patch_final_otimizado(sistema_principal):
    """
    Aplica patch final com interface unificada e data correta
    """
    try:
        print("🔧 Aplicando patch final otimizado...")
        
        if not hasattr(sistema_principal, 'sistema_nfe_unificado'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        sistema_nfe = sistema_principal.sistema_nfe_unificado
        print("✅ Sistema NFe encontrado!")
        
        # PATCH PRINCIPAL: Substituir método processar_xml
        def processar_xml_direto_para_configuracao(self):
            """Processa XML e vai direto para configuração (sem tela intermediária)"""
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
                    # ESCONDER JANELA ATUAL
                    self.janela.withdraw()
                    
                    # ABRIR INTERFACE DE CONFIGURAÇÃO DIRETAMENTE
                    InterfaceConfiguracaoOtimizada(self.sistema, self.dados_nfe_atual, self.janela)
                else:
                    self.label_arquivo.config(text="❌ Erro ao processar XML", fg='red')
                    
            except Exception as e:
                self.label_arquivo.config(text="❌ Erro no processamento", fg='red')
                messagebox.showerror("Erro", f"Erro ao processar XML:\n{str(e)}")
        
        # APLICAR PATCH
        sistema_nfe.processar_xml = processar_xml_direto_para_configuracao.__get__(
            sistema_nfe, type(sistema_nfe)
        )
        
        print("✅ Patch otimizado aplicado com sucesso!")
        print("📌 Melhorias:")
        print("   🔄 Fluxo direto: Processar XML → Configuração")
        print("   📅 Data período ATUAL corrigida")
        print("   🎯 Interface unificada sem sobreposição")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar patch: {e}")
        import traceback
        print(f"📄 Traceback: {traceback.format_exc()}")
        return False


class InterfaceConfiguracaoOtimizada:
    """Interface de configuração unificada e otimizada"""
    
    def __init__(self, sistema_principal, dados_nfe, janela_anterior):
        self.sistema = sistema_principal
        self.dados_nfe = dados_nfe
        self.janela_anterior = janela_anterior
        
        # CALCULAR DATAS CORRETAS
        self.calcular_datas_periodo_atual()
        
        # CRIAR INTERFACE
        self.criar_interface()
    
    def calcular_datas_periodo_atual(self):
        """Calcula datas do período ATUAL (não da NFe)"""
        hoje = datetime.now()
        
        # DATA DE REFERÊNCIA = SEMPRE PERÍODO ATUAL
        if hoje.day <= 15:
            self.data_referencia = hoje.replace(day=5).strftime('%d/%m/%Y')
            self.periodo_nome = "PRIMEIRA QUINZENA"
            self.data_fim_periodo = hoje.replace(day=15).strftime('%d/%m/%Y')
        else:
            self.data_referencia = hoje.replace(day=20).strftime('%d/%m/%Y')
            self.periodo_nome = "SEGUNDA QUINZENA"
            # Último dia do mês
            import calendar
            ultimo_dia = calendar.monthrange(hoje.year, hoje.month)[1]
            self.data_fim_periodo = hoje.replace(day=ultimo_dia).strftime('%d/%m/%Y')
        
        # DATA DE VENCIMENTO = DATA ORIGINAL DA NFE
        self.data_vencimento = self.dados_nfe.get('data_emissao', hoje.strftime('%d/%m/%Y'))
        
        print(f"📅 DATAS CALCULADAS:")
        print(f"   🎯 Período: {self.periodo_nome}")
        print(f"   📊 Data Relatório: {self.data_referencia}")
        print(f"   📄 Data Vencimento: {self.data_vencimento}")
        print(f"   ⏰ Período vai até: {self.data_fim_periodo}")
    
    def criar_interface(self):
        """Cria interface unificada"""
        self.janela = tk.Toplevel(self.sistema.root)
        self.janela.title("⚙️ Configuração de Importação NFe")
        self.janela.geometry("700x600")
        self.janela.grab_set()
        
        # PROTOCOLO DE FECHAMENTO
        self.janela.protocol("WM_DELETE_WINDOW", self.fechar_interface)
        
        # FRAME PRINCIPAL
        main_frame = ttk.Frame(self.janela)
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # SEÇÕES
        self.criar_cabecalho(main_frame)
        self.criar_resumo_nfe(main_frame)
        self.criar_periodo_atual(main_frame)
        self.criar_configuracoes(main_frame)
        self.criar_botoes_finais(main_frame)
    
    def criar_cabecalho(self, parent):
        """Cria cabeçalho da interface"""
        frame_header = ttk.Frame(parent)
        frame_header.pack(fill='x', pady=(0,10))
        
        titulo = tk.Label(frame_header, 
                         text="⚙️ CONFIGURAÇÃO DE IMPORTAÇÃO NFe",
                         font=('Arial', 14, 'bold'),
                         fg='darkblue')
        titulo.pack()
        
        subtitulo = tk.Label(frame_header,
                           text="Configure os dados antes de importar para o sistema",
                           font=('Arial', 9),
                           fg='gray')
        subtitulo.pack()
    
    def criar_resumo_nfe(self, parent):
        """Cria resumo da NFe"""
        frame_nfe = ttk.LabelFrame(parent, text="📄 Dados da NFe", padding=10)
        frame_nfe.pack(fill='x', pady=5)
        
        # GRID DE INFORMAÇÕES
        info_frame = ttk.Frame(frame_nfe)
        info_frame.pack(fill='x')
        
        dados = [
            ("📄 Número:", self.dados_nfe.get('numero_nf', '')),
            ("📅 Data NFe:", self.dados_nfe.get('data_emissao', '')),
            ("🏢 Fornecedor:", self.dados_nfe.get('razao_social_emitente', '')[:40]),
            ("💰 Valor Total:", f"R$ {self.dados_nfe.get('valor_total', 0):,.2f}"),
            ("📦 Produtos:", str(len(self.dados_nfe.get('produtos', [])))),
        ]
        
        for i, (label, valor) in enumerate(dados):
            row = i // 2
            col = (i % 2) * 2
            
            tk.Label(info_frame, text=label, font=('Arial', 9, 'bold')).grid(
                row=row, column=col, sticky='w', padx=(0,5), pady=2)
            tk.Label(info_frame, text=valor, font=('Arial', 9)).grid(
                row=row, column=col+1, sticky='w', padx=(0,20), pady=2)
    
    def criar_periodo_atual(self, parent):
        """Cria seção do período atual (DESTAQUE)"""
        frame_periodo = ttk.LabelFrame(parent, text="📊 Período do Relatório", padding=10)
        frame_periodo.pack(fill='x', pady=5)
        
        # INFORMAÇÕES DO PERÍODO
        info_periodo = f"""🎯 Período: {self.periodo_nome}
📅 Data do Relatório: {self.data_referencia}
⏰ Período vai até: {self.data_fim_periodo}
📄 Data Vencimento (da NFe): {self.data_vencimento}"""
        
        tk.Label(frame_periodo, text=info_periodo, justify='left', 
                font=('Arial', 10), fg='darkgreen').pack(anchor='w')
        
        # AVISO IMPORTANTE
        aviso = f"💡 Esta NFe entrará no relatório de {self.data_referencia} para cálculo da taxa de administração"
        tk.Label(frame_periodo, text=aviso, justify='left', 
                font=('Arial', 9), fg='blue').pack(anchor='w', pady=(5,0))
    
    def criar_configuracoes(self, parent):
        """Cria seção de configurações"""
        frame_config = ttk.LabelFrame(parent, text="⚙️ Configurações", padding=10)
        frame_config.pack(fill='x', pady=5)
        
        # VARIÁVEIS DE CONTROLE
        self.importar_financeiro = tk.BooleanVar(value=True)
        self.importar_materiais = tk.BooleanVar(value=True)
        
        # CHECKBOXES
        cb_financeiro = tk.Checkbutton(
            frame_config,
            text="💰 Importar dados financeiros (lançamento no sistema)",
            variable=self.importar_financeiro,
            font=('Arial', 10, 'bold')
        )
        cb_financeiro.pack(anchor='w', pady=2)
        
        cb_materiais = tk.Checkbutton(
            frame_config,
            text="📦 Importar materiais da obra (banco de dados para manual)",
            variable=self.importar_materiais,
            font=('Arial', 10, 'bold')
        )
        cb_materiais.pack(anchor='w', pady=2)
        
        # CONFIGURAÇÕES ESPECÍFICAS
        self.criar_configuracoes_financeiras(frame_config)
        self.criar_configuracoes_materiais(frame_config)
    
    def criar_configuracoes_financeiras(self, parent):
        """Configurações financeiras"""
        frame_fin = ttk.LabelFrame(parent, text="💰 Detalhes Financeiros", padding=10)
        frame_fin.pack(fill='x', pady=5)
        
        # REFERÊNCIA EDITÁVEL (DESTAQUE)
        tk.Label(frame_fin, text="Referência para o relatório:", 
                font=('Arial', 9, 'bold'), fg='purple').pack(anchor='w')
        
        self.referencia_entry = tk.Entry(frame_fin, width=70, font=('Arial', 10))
        self.referencia_entry.pack(fill='x', pady=2)
        
        # REFERÊNCIA PADRÃO INTELIGENTE
        numero_nf = self.dados_nfe.get('numero_nf', '')
        fornecedor = self.dados_nfe.get('razao_social_emitente', '')[:30]
        ref_padrao = f"NFE {numero_nf} - {fornecedor}"
        self.referencia_entry.insert(0, ref_padrao)
        
        # OUTRAS CONFIGURAÇÕES EM LINHA
        linha_config = ttk.Frame(frame_fin)
        linha_config.pack(fill='x', pady=5)
        
        tk.Label(linha_config, text="Categoria:").pack(side='left')
        self.categoria_entry = tk.Entry(linha_config, width=8)
        self.categoria_entry.insert(0, 'MAT')
        self.categoria_entry.pack(side='left', padx=5)
        
        tk.Label(linha_config, text="Tipo:").pack(side='left', padx=(20,0))
        self.tipo_combo = ttk.Combobox(linha_config, width=8, state='readonly')
        self.tipo_combo['values'] = ['1', '2', '3', '4', '5', '6', '7']
        self.tipo_combo.set('3')
        self.tipo_combo.pack(side='left', padx=5)
        
        tk.Label(linha_config, text="Forma Pgto:").pack(side='left', padx=(20,0))
        self.forma_combo = ttk.Combobox(linha_config, width=12, state='readonly')
        self.forma_combo['values'] = ['A_VISTA', 'A_PRAZO', 'CARTAO', 'PIX']
        self.forma_combo.set('A_PRAZO')
        self.forma_combo.pack(side='left', padx=5)
    
    def criar_configuracoes_materiais(self, parent):
        """Configurações materiais"""
        frame_mat = ttk.LabelFrame(parent, text="📦 Detalhes Materiais", padding=10)
        frame_mat.pack(fill='x', pady=5)
        
        # LINHA DE CONFIGURAÇÕES
        linha_mat = ttk.Frame(frame_mat)
        linha_mat.pack(fill='x', pady=3)
        
        tk.Label(linha_mat, text="Ambiente:").pack(side='left')
        self.ambiente_combo = ttk.Combobox(linha_mat, width=20, state='readonly')
        self.ambiente_combo['values'] = self.carregar_ambientes()
        self.ambiente_combo.pack(side='left', padx=5)
        
        tk.Label(linha_mat, text="Status:").pack(side='left', padx=(20,0))
        self.status_combo = ttk.Combobox(linha_mat, width=15, state='readonly')
        self.status_combo['values'] = ['PENDENTE', 'INSTALADO', 'EM_INSTALACAO']
        self.status_combo.set('PENDENTE')
        self.status_combo.pack(side='left', padx=5)
        
        tk.Label(linha_mat, text="Garantia:").pack(side='left', padx=(20,0))
        self.garantia_entry = tk.Entry(linha_mat, width=5)
        self.garantia_entry.insert(0, '12')
        self.garantia_entry.pack(side='left', padx=5)
        tk.Label(linha_mat, text="meses").pack(side='left')
    
    def carregar_ambientes(self):
        """Carrega ambientes dos parâmetros"""
        try:
            if hasattr(self.sistema, 'gerenciador_materiais'):
                return self.sistema.gerenciador_materiais.parametros.get('ambientes', ['GERAL'])
        except:
            pass
        return ['GERAL', 'INSTALAÇÃO DA OBRA', 'SALA DE ESTAR', 'COZINHA']
    
    def criar_botoes_finais(self, parent):
        """Cria botões finais"""
        frame_botoes = ttk.Frame(parent)
        frame_botoes.pack(fill='x', pady=15)
        
        ttk.Button(frame_botoes, 
                  text="👁️ Preview", 
                  command=self.mostrar_preview).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ IMPORTAR DADOS", 
                  command=self.executar_importacao).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="❌ Cancelar", 
                  command=self.fechar_interface).pack(side='right', padx=5)
    
    def mostrar_preview(self):
        """Mostra preview dos dados"""
        opcoes = self.coletar_opcoes()
        
        preview_text = f"""
🎯 PREVIEW DA IMPORTAÇÃO
{'='*50}

📄 NFe: {self.dados_nfe.get('numero_nf', '')} - {self.dados_nfe.get('razao_social_emitente', '')}
💰 Valor: R$ {self.dados_nfe.get('valor_total', 0):,.2f}

📊 DADOS PARA O SISTEMA:
"""
        
        if opcoes['importar_financeiro']:
            preview_text += f"""
💰 LANÇAMENTO FINANCEIRO:
   📅 Data Relatório: {self.data_referencia} ({self.periodo_nome})
   📅 Data Vencimento: {self.data_vencimento} (original da NFe)
   📋 Referência: {opcoes['referencia']}
   🏷️ Categoria: {opcoes['categoria']} | Tipo: {opcoes['tipo']}
   💳 Forma Pgto: {opcoes['forma_pagamento']}
"""
        
        if opcoes['importar_materiais']:
            preview_text += f"""
📦 MATERIAIS ({len(self.dados_nfe.get('produtos', []))} produtos):
   🏠 Ambiente: {opcoes['ambiente']}
   ⚙️ Status: {opcoes['status']}
   🛡️ Garantia: {opcoes['garantia']} meses
"""
        
        preview_text += f"""
🎯 RELATÓRIO:
   Esta NFe entrará no relatório quinzenal de {self.data_referencia}
   Base para cálculo da taxa de administração: {self.periodo_nome}
"""
        
        # JANELA DE PREVIEW
        janela_preview = tk.Toplevel(self.janela)
        janela_preview.title("👁️ Preview da Importação")
        janela_preview.geometry("600x500")
        janela_preview.grab_set()
        
        text_widget = tk.Text(janela_preview, wrap='word', font=('Courier', 10))
        text_widget.pack(fill='both', expand=True, padx=10, pady=10)
        text_widget.insert('1.0', preview_text.strip())
        text_widget.config(state='disabled')
        
        ttk.Button(janela_preview, text="Fechar", 
                  command=janela_preview.destroy).pack(pady=10)
    
    def coletar_opcoes(self):
        """Coleta todas as opções"""
        return {
            'importar_financeiro': self.importar_financeiro.get(),
            'importar_materiais': self.importar_materiais.get(),
            'referencia': self.referencia_entry.get().strip(),
            'categoria': self.categoria_entry.get().strip(),
            'tipo': self.tipo_combo.get(),
            'forma_pagamento': self.forma_combo.get(),
            'ambiente': self.ambiente_combo.get(),
            'status': self.status_combo.get(),
            'garantia': int(self.garantia_entry.get() or 12)
        }
    
    def executar_importacao(self):
        """Executa a importação"""
        try:
            opcoes = self.coletar_opcoes()
            
            if not opcoes['importar_financeiro'] and not opcoes['importar_materiais']:
                messagebox.showwarning("Aviso", "Selecione pelo menos uma opção!")
                return
            
            resultados = []
            
            # IMPORTAR FINANCEIRO
            if opcoes['importar_financeiro']:
                resultado_fin = self.criar_lancamento_financeiro(opcoes)
                resultados.append(f"💰 Financeiro: {resultado_fin}")
            
            # IMPORTAR MATERIAIS
            if opcoes['importar_materiais']:
                resultado_mat = self.criar_materiais(opcoes)
                resultados.append(f"📦 Materiais: {resultado_mat}")
            
            # FECHAR E MOSTRAR RESULTADO
            self.fechar_interface()
            self.mostrar_resultado_final(resultados, opcoes)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na importação: {str(e)}")
    
    def criar_lancamento_financeiro(self, opcoes):
        """Cria lançamento financeiro com data correta"""
        try:
            dados_nfe = self.dados_nfe
            
            dados_financeiros = {
                'data': self.data_referencia,  # PERÍODO ATUAL (5 ou 20)
                'cnpj_cpf': ''.join(c for c in dados_nfe.get('cnpj_emitente', '') if c.isdigit()),
                'nome': dados_nfe.get('razao_social_emitente', '')[:50],
                'categoria': opcoes['categoria'].upper(),
                'tp_desp': opcoes['tipo'],
                'referencia': opcoes['referencia'].upper(),
                'etapa_obra': 'MATERIAIS',
                'nf': dados_nfe.get('numero_nf', ''),
                'vr_unit': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dias': 1,
                'valor': f"{dados_nfe.get('valor_total', 0):.2f}".replace('.', ','),
                'dt_vencto': self.data_vencimento,  # DATA ORIGINAL DA NFE
                'dados_bancarios': '',
                'observacao': f"IMPORTADO NFE {dados_nfe.get('numero_nf', '')} - {self.periodo_nome} {self.data_referencia}".upper(),
                'forma_pagamento': opcoes['forma_pagamento']
            }
            
            if not hasattr(self.sistema, 'dados_para_incluir'):
                self.sistema.dados_para_incluir = []
            
            self.sistema.dados_para_incluir.append(dados_financeiros)
            
            print(f"💰 LANÇAMENTO CRIADO:")
            print(f"   📅 Período: {self.periodo_nome}")
            print(f"   📊 Data Relatório: {self.data_referencia}")
            print(f"   📄 Data Vencimento: {self.data_vencimento}")
            print(f"   📋 Referência: {opcoes['referencia']}")
            
            return f"R$ {dados_nfe.get('valor_total', 0):,.2f}"
            
        except Exception as e:
            raise Exception(f"Erro ao criar lançamento: {str(e)}")
    
    def criar_materiais(self, opcoes):
        """Cria materiais da obra"""
        try:
            produtos = self.dados_nfe.get('produtos', [])
            if not produtos:
                return "Nenhum produto"
            
            if not hasattr(self.sistema, 'gerenciador_materiais'):
                return f"{len(produtos)} produtos (sistema materiais não inicializado)"
            
            materiais_criados = 0
            dados_nfe = self.dados_nfe
            
            for produto in produtos:
                try:
                    dados_material = {
                        'Cliente': getattr(self.sistema, 'cliente_atual', 'SEM_CLIENTE'),
                        'Categoria': self.classificar_produto(produto.get('descricao', '')),
                        'Subcategoria': '',
                        'Codigo_Produto': produto.get('codigo', ''),
                        'Descricao_Completa': produto.get('descricao', ''),
                        'Marca': dados_nfe.get('razao_social_emitente', '')[:20],
                        'Ambiente_Aplicacao': opcoes['ambiente'],
                        'Status_Instalacao': opcoes['status'],
                        'Garantia_Meses': opcoes['garantia'],
                        'Observacoes': f"Importado NFe {dados_nfe.get('numero_nf', '')} - {self.periodo_nome}",
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
                    
                    material_id = self.sistema.gerenciador_materiais.salvar_material(dados_material)
                    materiais_criados += 1
                    
                except Exception as e:
                    print(f"⚠️ Erro ao criar material: {e}")
                    continue
            
            return f"{materiais_criados} de {len(produtos)} produtos"
            
        except Exception as e:
            raise Exception(f"Erro ao criar materiais: {str(e)}")
    
    def classificar_produto(self, descricao):
        """Classificação básica de produtos"""
        if not descricao:
            return 'OUTROS'
        
        desc_upper = descricao.upper()
        
        if any(palavra in desc_upper for palavra in ['CERAMICA', 'PORCELANATO', 'PISO', 'AZULEJO']):
            return 'ACABAMENTOS'
        elif any(palavra in desc_upper for palavra in ['TINTA', 'VERNIZ', 'MASSA']):
            return 'TINTAS'
        elif any(palavra in desc_upper for palavra in ['FIO', 'CABO', 'LAMPADA', 'TOMADA']):
            return 'ELETRICO'
        elif any(palavra in desc_upper for palavra in ['TUBO', 'TORNEIRA', 'REGISTRO']):
            return 'HIDRAULICO'
        else:
            return 'OUTROS'
    
    def mostrar_resultado_final(self, resultados, opcoes):
        """Mostra resultado final"""
        janela_resultado = tk.Toplevel(self.sistema.root)
        janela_resultado.title("🎉 Importação Concluída")
        janela_resultado.geometry("600x400")
        janela_resultado.grab_set()
        
        frame_main = ttk.Frame(janela_resultado)
        frame_main.pack(fill='both', expand=True, padx=20, pady=20)
        
        # TÍTULO
        tk.Label(frame_main, 
                text="🎉 IMPORTAÇÃO CONCLUÍDA COM SUCESSO!", 
                font=('Arial', 14, 'bold'),
                fg='darkgreen').pack(pady=10)
        
        # RESULTADOS
        for resultado in resultados:
            tk.Label(frame_main, text=resultado, fg='blue', 
                    font=('Arial', 11, 'bold')).pack(anchor='w', pady=2)
        
        # PERÍODO
        periodo_text = f"""
📊 PERÍODO DO RELATÓRIO:
   🎯 {self.periodo_nome}
   📅 Data: {self.data_referencia}
   📋 Referência: {opcoes.get('referencia', 'N/A')}
"""
        tk.Label(frame_main, text=periodo_text.strip(), justify='left', 
                font=('Arial', 10), fg='darkblue').pack(anchor='w', pady=10)
        
        # BOTÕES
        frame_botoes = ttk.Frame(frame_main)
        frame_botoes.pack(fill='x', pady=15)
        
        ttk.Button(frame_botoes, 
                  text="📊 Processar no Sistema", 
                  command=lambda: self.processar_no_sistema(janela_resultado)).pack(side='left', padx=5)
        
        ttk.Button(frame_botoes, 
                  text="✅ Concluído", 
                  command=janela_resultado.destroy).pack(side='right', padx=5)
    
    def processar_no_sistema(self, janela_resultado):
        """Chama enviar_dados() do sistema"""
        try:
            janela_resultado.destroy()
            
            if hasattr(self.sistema, 'dados_para_incluir') and self.sistema.dados_para_incluir:
                print(f"📊 Enviando {len(self.sistema.dados_para_incluir)} lançamentos...")
                self.sistema.enviar_dados()
            else:
                messagebox.showwarning("Aviso", "Não há dados para processar!")
                
        except Exception as e:
            messagebox.showerror("Erro", f"Erro: {str(e)}")
    
    def fechar_interface(self):
        """Fecha interface e restaura anterior"""
        try:
            self.janela.destroy()
            if self.janela_anterior and self.janela_anterior.winfo_exists():
                self.janela_anterior.deiconify()
        except:
            pass

