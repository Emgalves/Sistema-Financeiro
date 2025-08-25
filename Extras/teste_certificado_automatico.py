# -*- coding: utf-8 -*-
"""
TESTE AUTOMÁTICO DO CERTIFICADO A1
Adicione este código ao final do __init__ do SistemaEntradaDados
"""

def verificar_certificado_a1_automatico(self):
    try:
        print("\n" + "="*50)
        print("VERIFICAÇÃO AUTOMÁTICA CERTIFICADO A1")
        print("="*50)
        
        # Verificações
        tem_config = hasattr(self, 'configurar_certificado_rapido')
        tem_consultor = hasattr(self, 'consultor_sefaz_a1')
        tem_processador = hasattr(self, 'processador_nfe')
        
        print(f"Configurar certificado: {'OK' if tem_config else 'FALTANDO'}")
        print(f"Consultor SEFAZ A1: {'OK' if tem_consultor else 'FALTANDO'}")
        print(f"Processador NFe: {'OK' if tem_processador else 'FALTANDO'}")
        
        if all([tem_config, tem_consultor, tem_processador]):
            print("SUCESSO: Certificado A1 integrado ao sistema!")
            return True
        else:
            print("ERRO: Certificado A1 não foi carregado corretamente")
            return False
        
    except Exception as e:
        print(f"ERRO na verificação: {e}")
        return False


def criar_botao_teste_certificado(self):
    """Cria botão para testar certificado na interface"""
    try:
        import tkinter as tk
        from tkinter import ttk, messagebox
        
        # Encontrar onde adicionar o botão
        if hasattr(self, 'aba_fornecedor'):
            # Criar frame para certificado A1
            frame_cert = ttk.LabelFrame(self.aba_fornecedor, text="Certificado A1 - SEFAZ", padding=10)
            frame_cert.pack(fill='x', padx=10, pady=5)
            
            # Frame para botões
            frame_botoes = ttk.Frame(frame_cert)
            frame_botoes.pack(fill='x')
            
            # Botão configurar
            btn_config = ttk.Button(
                frame_botoes,
                text="Configurar Certificado A1",
                command=self.configurar_certificado_interface
            )
            btn_config.pack(side='left', padx=5)
            
            # Botão testar
            btn_teste = ttk.Button(
                frame_botoes,
                text="Testar Conexão",
                command=self.testar_certificado_interface
            )
            btn_teste.pack(side='left', padx=5)
            
            # Botão consultar NFe
            btn_consultar = ttk.Button(
                frame_botoes,
                text="Consultar NFe",
                command=self.consultar_nfe_interface
            )
            btn_consultar.pack(side='left', padx=5)
            
            # Status
            self.label_cert_status = tk.Label(
                frame_cert,
                text="Status: Certificado não configurado",
                fg='red'
            )
            self.label_cert_status.pack(pady=5)
            
            print("Botões de certificado A1 adicionados à interface!")
            
    except Exception as e:
        print(f"Erro ao criar botões: {e}")


def configurar_certificado_interface(self):
    """Interface para configurar certificado"""
    try:
        from tkinter import messagebox
        
        # Usar o método já implementado
        sucesso = self.configurar_certificado_rapido()
        
        if sucesso:
            self.label_cert_status.config(
                text="Status: Certificado configurado e testado",
                fg='green'
            )
        else:
            self.label_cert_status.config(
                text="Status: Erro na configuração",
                fg='red'
            )
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao configurar: {e}")


def testar_certificado_interface(self):
    """Interface para testar certificado"""
    try:
        from tkinter import messagebox
        
        if not hasattr(self, 'consultor_sefaz_a1'):
            messagebox.showerror("Erro", "Consultor A1 não inicializado")
            return
        
        cert_info = self.consultor_sefaz_a1.obter_info_certificado()
        
        if not cert_info.get('is_valid'):
            messagebox.showwarning("Aviso", "Configure o certificado primeiro")
            return
        
        # Fazer teste
        sucesso, msg = self.consultor_sefaz_a1.testar_conexao()
        
        if sucesso:
            messagebox.showinfo("Teste OK", f"Sucesso: {msg}")
            self.label_cert_status.config(
                text="Status: Conexão OK",
                fg='green'
            )
        else:
            messagebox.showerror("Teste Falhou", f"Erro: {msg}")
            self.label_cert_status.config(
                text="Status: Erro na conexão",
                fg='orange'
            )
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro no teste: {e}")


def consultar_nfe_interface(self):
    """Interface para consultar NFe"""
    try:
        import tkinter as tk
        from tkinter import simpledialog, messagebox
        
        if not hasattr(self, 'consultor_sefaz_a1'):
            messagebox.showerror("Erro", "Consultor A1 não inicializado")
            return
        
        cert_info = self.consultor_sefaz_a1.obter_info_certificado()
        
        if not cert_info.get('is_valid'):
            messagebox.showwarning("Aviso", "Configure o certificado primeiro")
            return
        
        # Solicitar chave
        root = tk.Tk()
        root.withdraw()
        
        chave = simpledialog.askstring(
            "Consultar NFe",
            "Digite a chave de acesso da NFe (44 dígitos):\n\n"
            "Exemplo: 35200314200166000187550010000000271234567890"
        )
        
        root.destroy()
        
        if not chave:
            return
        
        # Limpar chave
        chave_limpa = ''.join(filter(str.isdigit, chave))
        
        if len(chave_limpa) != 44:
            messagebox.showerror("Erro", "Chave deve ter exatamente 44 dígitos")
            return
        
        # Consultar
        messagebox.showinfo("Info", "Consultando NFe... Aguarde.")
        
        try:
            dados = self.processador_nfe.consultar_nfe_sefaz(chave_limpa)
            
            # Mostrar resultado
            resultado = f"""
RESULTADO DA CONSULTA:

Chave: {dados.get('chave_acesso', '')}
NFe: {dados.get('numero_nf', '')}
Emitente: {dados.get('razao_social_emitente', '')}
CNPJ: {dados.get('cnpj_emitente', '')}
Data: {dados.get('data_emissao', '')}
Valor: R$ {dados.get('valor_total', 0):,.2f}
Status: {dados.get('status_sefaz', dados.get('fonte_dados', ''))}
"""
            
            if dados.get('observacao'):
                resultado += f"\nObservação: {dados['observacao']}"
            
            messagebox.showinfo("Resultado da Consulta", resultado)
            
            # Perguntar se quer processar
            if dados.get('valor_total', 0) > 0:
                processar = messagebox.askyesno(
                    "Processar NFe",
                    "NFe encontrada! Deseja processar para o sistema financeiro?"
                )
                
                if processar:
                    self.processar_nfe_consultada(dados)
            
        except Exception as e:
            messagebox.showerror("Erro", f"Erro na consulta: {e}")
            
    except Exception as e:
        messagebox.showerror("Erro", f"Erro: {e}")


def processar_nfe_consultada(self, dados_nfe):
    """Processa NFe consultada para o sistema"""
    try:
        from tkinter import messagebox
        
        if not hasattr(self, 'cliente_atual') or not self.cliente_atual:
            messagebox.showerror("Erro", "Selecione um cliente antes de processar")
            return
        
        # Criar dados financeiros
        valor = dados_nfe.get('valor_total', 0)
        
        if valor <= 0:
            messagebox.showwarning("Aviso", "NFe sem valor para processar")
            return
        
        dados_financeiro = {
            'data': dados_nfe.get('data_emissao', ''),
            'cnpj_cpf': dados_nfe.get('cnpj_emitente', ''),
            'nome': dados_nfe.get('razao_social_emitente', ''),
            'categoria': 'MAT',
            'tp_desp': '3',
            'referencia': f"NFE {dados_nfe.get('numero_nf', '')}",
            'etapa_obra': '',
            'nf': dados_nfe.get('numero_nf', ''),
            'vr_unit': f"{valor:.2f}",
            'dias': 1,
            'valor': f"{valor:.2f}",
            'dt_vencto': dados_nfe.get('data_emissao', ''),
            'dados_bancarios': '',
            'observacao': f"NFE SEFAZ - {dados_nfe.get('chave_acesso', '')}",
            'forma_pagamento': ''
        }
        
        # Adicionar aos dados do sistema
        self.dados_para_incluir = [dados_financeiro]
        
        # Processar
        self.enviar_dados()
        
        messagebox.showinfo(
            "Sucesso",
            f"NFe processada com sucesso!\n\n"
            f"Valor: R$ {valor:,.2f}\n"
            f"Fornecedor: {dados_nfe.get('razao_social_emitente', '')}\n"
            f"Adicionado ao sistema financeiro."
        )
        
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar: {e}")


# CÓDIGO PARA ADICIONAR AO __init__ DO SistemaEntradaDados:
"""
Adicione estas linhas no final do método __init__ da classe SistemaEntradaDados:

        # Verificação automática do certificado A1
        try:
            from teste_certificado_automatico import verificar_certificado_a1_automatico
            verificar_certificado_a1_automatico(self)
        except Exception as e:
            print(f"Erro na verificação do certificado: {e}")
        
        # Adicionar métodos de interface
        try:
            from teste_certificado_automatico import (
                criar_botao_teste_certificado,
                configurar_certificado_interface,
                testar_certificado_interface,
                consultar_nfe_interface,
                processar_nfe_consultada
            )
            
            self.criar_botao_teste_certificado = criar_botao_teste_certificado.__get__(self)
            self.configurar_certificado_interface = configurar_certificado_interface.__get__(self)
            self.testar_certificado_interface = testar_certificado_interface.__get__(self)
            self.consultar_nfe_interface = consultar_nfe_interface.__get__(self)
            self.processar_nfe_consultada = processar_nfe_consultada.__get__(self)
            
        except Exception as e:
            print(f"Erro ao adicionar métodos de interface: {e}")
"""
