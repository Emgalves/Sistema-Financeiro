# -*- coding: utf-8 -*-
"""
IMPLEMENTAÇÃO LIMPA DO CERTIFICADO A1
Execute este arquivo para aplicar as correções ao seu sistema
"""

import os
import tempfile
import shutil
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

try:
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.serialization import pkcs12
    import requests
    import xml.etree.ElementTree as ET
    DEPS_OK = True
except ImportError as e:
    DEPS_OK = False
    print(f"Dependências faltando: {e}")

import urllib3
urllib3.disable_warnings()


class CertificadoA1Corrigido:
    """Versão corrigida e limpa do gerenciador de certificado A1"""
    
    def __init__(self):
        self.cert_info = {}
        self.temp_dir = None
        self.cert_pem = None
        self.key_pem = None
        
        # URLs SEFAZ corretas
        self.urls_sefaz = {
            'SP': 'https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'RJ': 'https://nfe.sefaz.rj.gov.br/nfe/NFeConsultaProtocolo4',
            'MG': 'https://nfe.fazenda.mg.gov.br/nfe2/NFeConsultaProtocolo4',
            'PR': 'https://nfe.sefa.pr.gov.br/nfe/NFeConsultaProtocolo4',
            'RS': 'https://nfe.sefazrs.rs.gov.br/ws/NFeConsultaProtocolo/NFeConsultaProtocolo4.asmx'
        }
        
        self.codigos_uf = {
            '35': 'SP', '33': 'RJ', '31': 'MG', '41': 'PR', '43': 'RS',
            '42': 'SC', '29': 'BA', '23': 'CE', '52': 'GO', '50': 'MS'
        }
    
    def configurar_certificado(self, arquivo_pfx, senha):
        """Configura o certificado A1 com tratamento correto de senha"""
        try:
            print(f"Configurando certificado: {os.path.basename(arquivo_pfx)}")
            
            if not os.path.exists(arquivo_pfx):
                return False, "Arquivo não encontrado"
            
            # Ler arquivo
            with open(arquivo_pfx, 'rb') as f:
                cert_data = f.read()
            
            # Tentar diferentes formatos de senha
            formatos_senha = []
            if senha:
                formatos_senha = [
                    senha.encode('utf-8'),           # UTF-8
                    senha.encode('latin-1'),         # Latin-1
                    str(senha).encode('utf-8'),      # String para UTF-8
                    senha,                           # String direta
                ]
            formatos_senha.append(None)  # Sem senha
            
            certificate = None
            private_key = None
            
            for i, fmt_senha in enumerate(formatos_senha):
                try:
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, fmt_senha)
                    if certificate and private_key:
                        print(f"Sucesso com formato {i+1}")
                        break
                except Exception:
                    continue
            
            if not certificate:
                return False, "Senha incorreta ou certificado inválido"
            
            # Verificar validade
            agora = datetime.now()
            if certificate.not_valid_after < agora:
                return False, f"Certificado expirado em {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
            # Criar arquivos temporários
            self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_")
            
            cert_pem = certificate.public_bytes(serialization.Encoding.PEM)
            key_pem = private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            )
            
            # Salvar arquivos
            cert_file = os.path.join(self.temp_dir, "cert.pem")
            key_file = os.path.join(self.temp_dir, "key.pem")
            
            with open(cert_file, 'wb') as f:
                f.write(cert_pem)
            with open(key_file, 'wb') as f:
                f.write(key_pem)
            
            os.chmod(cert_file, 0o600)
            os.chmod(key_file, 0o600)
            
            # Armazenar dados
            self.cert_pem = cert_pem
            self.key_pem = key_pem
            self.cert_info = {
                'valido': True,
                'arquivo_cert': cert_file,
                'arquivo_key': key_file,
                'valido_ate': certificate.not_valid_after,
                'subject': certificate.subject.rfc4514_string()
            }
            
            return True, f"Certificado válido até {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
        except Exception as e:
            self.limpar()
            return False, f"Erro: {str(e)}"
    
    def consultar_nfe_sefaz(self, chave_acesso):
        """Consulta NFe na SEFAZ usando certificado A1"""
        try:
            if not self.cert_info.get('valido'):
                raise Exception("Certificado não configurado")
            
            # Verificar arquivos
            if not self._verificar_arquivos():
                self._recriar_arquivos()
            
            # Obter UF e URL
            uf_codigo = chave_acesso[:2]
            uf = self.codigos_uf.get(uf_codigo, 'SP')
            url = self.urls_sefaz.get(uf, self.urls_sefaz['SP'])
            
            print(f"Consultando SEFAZ {uf}: {chave_acesso}")
            
            # Envelope SOAP correto
            envelope = f"""<?xml version="1.0" encoding="utf-8"?>
<soap12:Envelope xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">
    <soap:Body>
        <nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">
            <consSitNFe versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
                <tpAmb>1</tpAmb>
                <xServ>CONSULTAR</xServ>
                <chNFe>{chave_acesso}</chNFe>
            </consSitNFe>
        </nfeDadosMsg>
    </soap:Body>
</soap:Envelope>"""
            
            headers = {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF'
            }
            
            # Fazer requisição
            response = requests.post(
                url,
                data=envelope,
                headers=headers,
                cert=(self.cert_info['arquivo_cert'], self.cert_info['arquivo_key']),
                verify=False,
                timeout=30
            )
            
            if response.status_code == 200:
                return self._processar_resposta_sefaz(response.text, chave_acesso)
            else:
                raise Exception(f"HTTP {response.status_code}")
                
        except Exception as e:
            # Retornar dados simulados em caso de erro
            return {
                'chave_acesso': chave_acesso,
                'numero_nf': str(int(chave_acesso[25:34])),
                'razao_social_emitente': 'ERRO NA CONSULTA SEFAZ',
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Erro',
                'erro': str(e)
            }
    
    def _processar_resposta_sefaz(self, xml_response, chave):
        """Processa resposta XML da SEFAZ"""
        try:
            root = ET.fromstring(xml_response)
            
            # Buscar status
            status = "000"
            motivo = "Processado"
            
            for elem in root.iter():
                if elem.tag.endswith('cStat'):
                    status = elem.text
                    break
            
            for elem in root.iter():
                if elem.tag.endswith('xMotivo'):
                    motivo = elem.text
                    break
            
            print(f"Status SEFAZ: {status} - {motivo}")
            
            # Dados básicos
            dados = {
                'chave_acesso': chave,
                'numero_nf': str(int(chave[25:34])),
                'cnpj_emitente': chave[6:20],
                'razao_social_emitente': 'CONSULTADO VIA SEFAZ',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Consulta SEFAZ A1',
                'status': f"{status} - {motivo}"
            }
            
            # Extrair dados se NFe autorizada
            if status == "100":
                try:
                    for elem in root.iter():
                        if elem.tag.endswith('xNome') and elem.text:
                            dados['razao_social_emitente'] = elem.text
                            break
                    
                    for elem in root.iter():
                        if elem.tag.endswith('vNF') and elem.text:
                            dados['valor_total'] = float(elem.text)
                            break
                except:
                    pass
            
            return dados
            
        except Exception as e:
            return {
                'chave_acesso': chave,
                'numero_nf': 'ERRO',
                'razao_social_emitente': 'ERRO NO PROCESSAMENTO',
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Erro XML',
                'erro': str(e)
            }
    
    def _verificar_arquivos(self):
        """Verifica se arquivos temporários existem"""
        cert_file = self.cert_info.get('arquivo_cert')
        key_file = self.cert_info.get('arquivo_key')
        return cert_file and key_file and os.path.exists(cert_file) and os.path.exists(key_file)
    
    def _recriar_arquivos(self):
        """Recria arquivos temporários"""
        if not self.cert_pem or not self.key_pem:
            return False
        
        if not self.temp_dir:
            self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_")
        
        cert_file = os.path.join(self.temp_dir, "cert.pem")
        key_file = os.path.join(self.temp_dir, "key.pem")
        
        with open(cert_file, 'wb') as f:
            f.write(self.cert_pem)
        with open(key_file, 'wb') as f:
            f.write(self.key_pem)
        
        self.cert_info['arquivo_cert'] = cert_file
        self.cert_info['arquivo_key'] = key_file
        return True
    
    def testar_conexao(self):
        """Testa conexão com SEFAZ"""
        try:
            chave_teste = "35200314200166000187550010000000271234567890"
            resultado = self.consultar_nfe_sefaz(chave_teste)
            
            if 'erro' in resultado:
                return False, resultado['erro']
            else:
                return True, "Conexão SEFAZ funcionando"
        except Exception as e:
            return False, str(e)
    
    def limpar(self):
        """Limpa arquivos temporários"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            shutil.rmtree(self.temp_dir, ignore_errors=True)
        self.temp_dir = None
        self.cert_info = {}


def aplicar_ao_seu_sistema(sistema_principal):
    """
    ESTA É A FUNÇÃO QUE VOCÊ DEVE CHAMAR
    Aplica o certificado A1 corrigido ao seu sistema
    """
    try:
        print("Aplicando certificado A1 corrigido ao sistema...")
        
        if not DEPS_OK:
            print("ERRO: Instale as dependências primeiro:")
            print("pip install cryptography requests")
            return False
        
        # Verificar se sistema NFe existe
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("Erro: Sistema NFe não encontrado")
            print("Primeiro inicialize o sistema NFe híbrido")
            return False
        
        # Criar instância do certificado corrigido
        cert_a1 = CertificadoA1Corrigido()
        
        # Salvar método original se existir
        if hasattr(sistema_principal.processador_nfe, 'consultar_nfe_sefaz'):
            sistema_principal.processador_nfe.consultar_nfe_sefaz_original = sistema_principal.processador_nfe.consultar_nfe_sefaz
        
        # SUBSTITUIR método de consulta
        sistema_principal.processador_nfe.consultar_nfe_sefaz = cert_a1.consultar_nfe_sefaz
        
        # ADICIONAR método de configuração
        def configurar_certificado_interface():
            """Interface para configurar certificado"""
            try:
                # Selecionar arquivo
                arquivo_cert = filedialog.askopenfilename(
                    title="Selecionar Certificado A1 (.pfx)",
                    filetypes=[("Certificado", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                if not arquivo_cert:
                    return False
                
                # Solicitar senha
                root = tk.Tk()
                root.withdraw()
                
                senha = simpledialog.askstring(
                    "Senha do Certificado",
                    "Digite a senha do certificado A1:\n\n"
                    "• Senha definida quando baixou o certificado\n"
                    "• Deixe em branco se não tem senha",
                    show='*'
                )
                
                root.destroy()
                
                if senha is None:  # Usuário cancelou
                    return False
                
                # Configurar certificado
                sucesso, mensagem = cert_a1.configurar_certificado(arquivo_cert, senha)
                
                if sucesso:
                    # Testar conexão
                    teste_ok, teste_msg = cert_a1.testar_conexao()
                    
                    resultado = f"Certificado: {mensagem}\nTeste: {teste_msg}"
                    
                    if teste_ok:
                        messagebox.showinfo("Sucesso", f"{resultado}\n\nSistema pronto para consultar NFe!")
                    else:
                        messagebox.showwarning("Parcial", f"{resultado}\n\nVerifique conectividade.")
                    
                    return True
                else:
                    messagebox.showerror("Erro", f"Falha: {mensagem}")
                    return False
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro: {str(e)}")
                return False
        
        # Adicionar métodos ao sistema
        sistema_principal.configurar_certificado_a1 = configurar_certificado_interface
        sistema_principal.certificado_a1 = cert_a1
        
        # Método de teste
        def testar_certificado():
            if cert_a1.cert_info.get('valido'):
                teste_ok, msg = cert_a1.testar_conexao()
                print(f"Teste certificado: {msg}")
                return teste_ok, msg
            else:
                return False, "Certificado não configurado"
        
        sistema_principal.testar_certificado_a1 = testar_certificado
        
        print("✓ Certificado A1 aplicado ao sistema com sucesso!")
        print("✓ Use: sistema_principal.configurar_certificado_a1()")
        print("✓ Teste: sistema_principal.testar_certificado_a1()")
        
        return True
        
    except Exception as e:
        print(f"Erro ao aplicar certificado: {e}")
        return False


def configuracao_rapida_certificado():
    """
    Configuração rápida standalone (se não tiver o sistema principal)
    """
    try:
        cert_a1 = CertificadoA1Corrigido()
        
        print("=== CONFIGURAÇÃO RÁPIDA CERTIFICADO A1 ===")
        
        # Selecionar arquivo
        arquivo = filedialog.askopenfilename(
            title="Selecionar Certificado A1",
            filetypes=[("Certificado", "*.pfx *.p12")]
        )
        
        if not arquivo:
            print("Nenhum arquivo selecionado")
            return False
        
        # Solicitar senha
        root = tk.Tk()
        root.withdraw()
        senha = simpledialog.askstring("Senha", "Digite a senha do certificado:", show='*')
        root.destroy()
        
        # Configurar
        sucesso, msg = cert_a1.configurar_certificado(arquivo, senha)
        print(f"Configuração: {msg}")
        
        if sucesso:
            # Testar
            teste_ok, teste_msg = cert_a1.testar_conexao()
            print(f"Teste: {teste_msg}")
            
            if teste_ok:
                print("✓ Certificado pronto para uso!")
                
                # Testar com chave real
                chave = input("Digite uma chave de NFe para testar (ou Enter para pular): ").strip()
                if chave and len(chave) == 44:
                    dados = cert_a1.consultar_nfe_sefaz(chave)
                    print(f"Resultado: {dados.get('razao_social_emitente', 'Erro')}")
        
        return sucesso
        
    except Exception as e:
        print(f"Erro: {e}")
        return False


if __name__ == "__main__":
    print("IMPLEMENTAÇÃO CERTIFICADO A1 - SEFAZ")
    print("=====================================")
    
    if not DEPS_OK:
        print("ERRO: Instale as dependências:")
        print("pip install cryptography requests")
        exit(1)
    
    print("1. Para aplicar ao sistema principal:")
    print("   from implementacao_certificado_a1 import aplicar_ao_seu_sistema")
    print("   aplicar_ao_seu_sistema(sistema_principal)")
    print()
    print("2. Para configuração rápida standalone:")
    print("   configuracao_rapida_certificado()")
    print()
    
    opcao = input("Executar configuração rápida agora? (s/n): ").lower()
    if opcao == 's':
        configuracao_rapida_certificado()
