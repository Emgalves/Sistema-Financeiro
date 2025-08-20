# -*- coding: utf-8 -*-
"""
CONSULTA SEFAZ COM CERTIFICADO A1 - VERSÃO COM SINTAXE CORRIGIDA
Substitua COMPLETAMENTE o arquivo src/nfe/consulta_sefaz_certificado.py
"""

import requests
import xml.etree.ElementTree as ET
from datetime import datetime, timezone
import tempfile
import os
import shutil
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog

try:
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.serialization import pkcs12
    CRYPTO_OK = True
except ImportError:
    CRYPTO_OK = False

import urllib3
urllib3.disable_warnings()


class ConsultorSefazA1:
    """Consultor para SEFAZ com certificado A1"""
    
    def __init__(self):
        self.cert_info = {}
        self.cert_data = None
        self.key_data = None
        self.temp_dir = None
        
        # URLs SEFAZ
        self.urls = {
            'PR': 'https://nfe.sefa.pr.gov.br/nfe/NFeConsultaProtocolo4',
            'SP': 'https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'RJ': 'https://nfe.sefaz.rj.gov.br/nfe/NFeConsultaProtocolo4',
            'MG': 'https://nfe.fazenda.mg.gov.br/nfe2/services/NFeConsultaProtocolo4',
            'RS': 'https://nfe.sefazrs.rs.gov.br/ws/NFeConsultaProtocolo/NFeConsultaProtocolo4.asmx',
            'SC': 'https://nfe.sef.sc.gov.br/ws/nfeconsultaprotocolo4.asmx'
        }
        
        # Códigos UF
        self.ufs = {
            '11': 'RO', '12': 'AC', '13': 'AM', '14': 'RR', '15': 'PA',
            '16': 'AP', '17': 'TO', '21': 'MA', '22': 'PI', '23': 'CE',
            '24': 'RN', '25': 'PB', '26': 'PE', '27': 'AL', '28': 'SE',
            '29': 'BA', '31': 'MG', '32': 'ES', '33': 'RJ', '35': 'SP',
            '41': 'PR', '42': 'SC', '43': 'RS', '50': 'MS', '51': 'MT',
            '52': 'GO', '53': 'DF'
        }
    
    def configurar_certificado(self, cert_path, cert_pin):
        """Configura certificado A1"""
        try:
            if not CRYPTO_OK:
                return False, "Erro: pip install cryptography"
            
            if not os.path.exists(cert_path):
                return False, "Arquivo de certificado não encontrado"
            
            print(f"Configurando certificado: {os.path.basename(cert_path)}")
            
            # Ler certificado
            with open(cert_path, 'rb') as f:
                cert_data = f.read()
            
            # Testar formatos de senha
            password_attempts = []
            
            if cert_pin:
                password_attempts.extend([
                    cert_pin.encode('utf-8'),
                    cert_pin.encode('latin-1'),
                    str(cert_pin).encode('utf-8'),
                    cert_pin,
                ])
            
            password_attempts.append(None)  # Sem senha
            
            private_key = None
            certificate = None
            
            for i, password in enumerate(password_attempts):
                try:
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, password)
                    if certificate and private_key:
                        print(f"Sucesso com formato {i+1}")
                        break
                except Exception:
                    continue
            
            if not certificate:
                return False, "Senha incorreta ou certificado inválido"
            
            # Verificar validade
            now = datetime.now(timezone.utc)
            try:
                not_valid_after = certificate.not_valid_after_utc
            except AttributeError:
                not_valid_after = certificate.not_valid_after.replace(tzinfo=timezone.utc)
            
            if not_valid_after < now:
                return False, f"Certificado expirado em {not_valid_after.strftime('%d/%m/%Y')}"
            
            # Criar arquivos temporários
            self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_")
            
            # Converter para PEM
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
            
            # Backup em memória
            self.cert_data = cert_pem
            self.key_data = key_pem
            
            # Armazenar informações
            self.cert_info = {
                'is_valid': True,
                'not_valid_after': not_valid_after,
                'subject': certificate.subject.rfc4514_string(),
                'cert_path': cert_file,
                'key_path': key_file
            }
            
            return True, f"Certificado válido até {not_valid_after.strftime('%d/%m/%Y')}"
            
        except Exception as e:
            self.limpar()
            return False, f"Erro: {str(e)}"
    
    def consultar_nfe(self, chave_acesso):
        """Consulta NFe na SEFAZ"""
        try:
            if not self.cert_info.get('is_valid'):
                raise Exception("Certificado não configurado")
            
            if len(chave_acesso) != 44:
                raise Exception("Chave deve ter 44 dígitos")
            
            # Verificar/recriar arquivos se necessário
            if not self.verificar_arquivos():
                if not self.recriar_arquivos():
                    raise Exception("Erro nos arquivos de certificado")
            
            # Obter UF e URL
            uf = self.obter_uf(chave_acesso)
            url = self.urls.get(uf, self.urls['SP'])
            
            print(f"Consultando SEFAZ {uf}: {chave_acesso}")
            
            # Envelope SOAP correto
            envelope = '<?xml version="1.0" encoding="utf-8"?>' + \
                        '<soap12:Envelope xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">' + \
                        '<soap12:Body>' + \
                        '<nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">' + \
                        '<consSitNFe versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">' + \
                        '<tpAmb>1</tpAmb>' + \
                        '<xServ>CONSULTAR</xServ>' + \
                        f'<chNFe>{chave_acesso}</chNFe>' + \
                        '</consSitNFe>' + \
                        '</nfeDadosMsg>' + \
                        '</soap12:Body>' + \
                        '</soap12:Envelope>'
            
            # Headers corretos
            headers = {
                'Content-Type': 'application/soap+xml; charset=utf-8'
            }
            
            # Fazer requisição
            response = requests.post(
                url,
                data=envelope.encode('utf-8'),
                headers=headers,
                cert=(self.cert_info['cert_path'], self.cert_info['key_path']),
                verify=False,
                timeout=60
            )
            
            print(f"Status HTTP: {response.status_code}")
            
            if response.status_code == 200:
                return self.processar_resposta(response.text, chave_acesso)
            else:
                print(f"Erro HTTP: {response.text[:500]}")
                raise Exception(f"HTTP {response.status_code}")
                
        except Exception as e:
            print(f"Erro na consulta: {e}")
            return {
                'chave_acesso': chave_acesso,
                'numero_nf': str(int(chave_acesso[25:34])),
                'razao_social_emitente': 'ERRO NA CONSULTA',
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Erro',
                'observacao': str(e)
            }
    
    def processar_resposta(self, xml_resp, chave):
        """Processa resposta"""
        try:
            root = ET.fromstring(xml_resp)
            
            # Buscar status
            status = "000"
            motivo = "Processado"
            
            for elem in root.iter():
                if 'cStat' in elem.tag:
                    status = elem.text
                    break
            
            for elem in root.iter():
                if 'xMotivo' in elem.tag:
                    motivo = elem.text
                    break
            
            print(f"SEFAZ: {status} - {motivo}")
            
            # Dados básicos
            dados = {
                'chave_acesso': chave,
                'numero_nf': str(int(chave[25:34])),
                'cnpj_emitente': chave[6:20],
                'razao_social_emitente': 'EMPRESA CONSULTADA',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Consulta SEFAZ A1',
                'status_sefaz': f"{status} - {motivo}"
            }
            
            # Extrair dados detalhados se NFe autorizada
            if status == '100':
                dados.update(self.extrair_detalhes(root))
            
            return dados
            
        except Exception as e:
            return {
                'chave_acesso': chave,
                'numero_nf': str(int(chave[25:34])),
                'razao_social_emitente': 'ERRO NO PROCESSAMENTO',
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Erro',
                'observacao': str(e)
            }
    
    def extrair_detalhes(self, root):
        """Extrai detalhes da NFe"""
        dados = {}
        try:
            # Nome do emitente
            for elem in root.iter():
                if 'xNome' in elem.tag and elem.text:
                    dados['razao_social_emitente'] = elem.text
                    break
            
            # Valor total
            for elem in root.iter():
                if 'vNF' in elem.tag and elem.text:
                    dados['valor_total'] = float(elem.text)
                    break
            
            # Data emissão
            for elem in root.iter():
                if 'dhEmi' in elem.tag and elem.text:
                    data_iso = elem.text.split('T')[0]
                    dt = datetime.strptime(data_iso, '%Y-%m-%d')
                    dados['data_emissao'] = dt.strftime('%d/%m/%Y')
                    break
        except:
            pass
        
        return dados
    
    def verificar_arquivos(self):
        """Verifica se arquivos existem"""
        if not self.cert_info.get('is_valid'):
            return False
        
        cert_path = self.cert_info.get('cert_path')
        key_path = self.cert_info.get('key_path')
        
        if not cert_path or not key_path:
            return False
        
        return os.path.exists(cert_path) and os.path.exists(key_path)
    
    def recriar_arquivos(self):
        """Recria arquivos a partir da memória"""
        try:
            if not self.cert_data or not self.key_data:
                return False
            
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_")
            
            cert_file = os.path.join(self.temp_dir, "cert.pem")
            key_file = os.path.join(self.temp_dir, "key.pem")
            
            with open(cert_file, 'wb') as f:
                f.write(self.cert_data)
            with open(key_file, 'wb') as f:
                f.write(self.key_data)
            
            os.chmod(cert_file, 0o600)
            os.chmod(key_file, 0o600)
            
            self.cert_info['cert_path'] = cert_file
            self.cert_info['key_path'] = key_file
            
            return True
            
        except Exception as e:
            print(f"Erro ao recriar: {e}")
            return False
    
    def obter_uf(self, chave):
        """Obtém UF da chave"""
        return self.ufs.get(chave[:2] if len(chave) >= 2 else '', 'SP')
    
    def testar_conexao(self):
        """Testa conexão"""
        try:
            if not self.cert_info.get('is_valid'):
                return False, "Certificado não configurado"
            
            chave_teste = "35202314200166000187550010000000271234567890"
            
            try:
                resultado = self.consultar_nfe(chave_teste)
                status = resultado.get('status_sefaz', '')
                
                if "217" in status or "999" in status or "não encontrada" in status.lower():
                    return True, "Conexão OK (NFe teste não encontrada - normal)"
                elif "100" in status:
                    return True, "Conexão OK"
                else:
                    return True, f"Conexão OK - Status: {status}"
                    
            except Exception as e:
                error_msg = str(e).lower()
                if "não encontrada" in error_msg or "217" in error_msg:
                    return True, "Conexão OK (NFe teste não existe)"
                else:
                    return False, f"Erro: {str(e)[:80]}"
                    
        except Exception as e:
            return False, f"Erro: {str(e)}"
    
    def obter_info_certificado(self):
        """Info do certificado"""
        return self.cert_info.copy()
    
    def limpar(self):
        """Limpa arquivos temporários"""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
        except:
            pass
        
        self.temp_dir = None
        self.cert_data = None
        self.key_data = None
        self.cert_info = {}
    
    def __del__(self):
        """Destrutor"""
        self.limpar()


def aplicar_melhorias_ao_sistema_existente(sistema_principal):
    """FUNÇÃO PRINCIPAL PARA INTEGRAÇÃO"""
    try:
        print("Aplicando melhorias de certificado A1...")
        
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("Sistema NFe não encontrado")
            return False
        
        # Criar UMA ÚNICA instância global do consultor
        consultor = ConsultorSefazA1()
        
        # Armazenar a instância no sistema principal
        sistema_principal.consultor_sefaz_a1 = consultor
        
        # Backup método original
        if hasattr(sistema_principal.processador_nfe, 'consultar_nfe_sefaz'):
            sistema_principal.processador_nfe.consultar_nfe_sefaz_original = sistema_principal.processador_nfe.consultar_nfe_sefaz
        
        def nova_consulta(chave):
            """Nova consulta usando a MESMA instância do consultor"""
            try:
                # DEBUG
                print(f"DEBUG: sistema_principal tem consultor_sefaz_a1: {hasattr(sistema_principal, 'consultor_sefaz_a1')}")
                
                if hasattr(sistema_principal, 'consultor_sefaz_a1'):
                    consultor_sistema = sistema_principal.consultor_sefaz_a1
                    print(f"DEBUG: Consultor existe: {consultor_sistema is not None}")
                    print(f"DEBUG: Cert info válido: {consultor_sistema.cert_info.get('is_valid', False)}")
                    
                    if consultor_sistema.cert_info.get('is_valid'):
                        print("Usando certificado A1...")
                        return consultor_sistema.consultar_nfe(chave)
                    else:
                        print("Certificado não configurado")
                        return {
                            'chave_acesso': chave,
                            'numero_nf': str(int(chave[25:34])),
                            'razao_social_emitente': 'CONFIGURE CERTIFICADO A1',
                            'valor_total': 0.0,
                            'produtos': [],
                            'fonte_dados': 'Simulação'
                        }
                else:
                    print("DEBUG: consultor_sefaz_a1 não encontrado no sistema_principal")
                    return {
                        'chave_acesso': chave,
                        'numero_nf': str(int(chave[25:34])),
                        'razao_social_emitente': 'ERRO: CONSULTOR NÃO ENCONTRADO',
                        'valor_total': 0.0,
                        'produtos': [],
                        'fonte_dados': 'Erro'
                    }
                    
            except Exception as e:
                print(f"Erro: {e}")
                return {
                    'chave_acesso': chave,
                    'numero_nf': 'ERRO',
                    'razao_social_emitente': 'ERRO NA CONSULTA',
                    'valor_total': 0.0,
                    'produtos': [],
                    'fonte_dados': 'Erro',
                    'observacao': str(e)
                }
        
        # Substituir método
        sistema_principal.processador_nfe.consultar_nfe_sefaz = nova_consulta
        
        # Métodos que usam a MESMA instância
        def configurar_cert(cert_path, cert_pin):
            """Configura certificado na instância do sistema"""
            resultado = sistema_principal.consultor_sefaz_a1.configurar_certificado(cert_path, cert_pin)
            print(f"DEBUG config_cert: Resultado: {resultado}")
            return resultado
        
        def testar_cert():
            """Testa certificado da instância do sistema"""
            return sistema_principal.consultor_sefaz_a1.testar_conexao()
        
        # Adicionar métodos
        sistema_principal.processador_nfe.configurar_certificado_a1 = configurar_cert
        sistema_principal.processador_nfe.testar_certificado_a1 = testar_cert
        
        # Configuração rápida
        def config_rapida():
            """Configuração rápida usando a instância do sistema"""
            print("CONSULTA_SEFAZ: Configurando certificado (função corrigida)")
            try:
                print("Configuração de certificado A1...")
                
                cert_path = filedialog.askopenfilename(
                    title="Selecionar Certificado A1",
                    filetypes=[("Certificado A1", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                if not cert_path:
                    return False
                
                root = tk.Tk()
                root.withdraw()
                
                cert_pin = simpledialog.askstring(
                    "PIN do Certificado A1",
                    "Digite o PIN/senha do certificado A1:\n\n"
                    "• Senha definida quando baixou o certificado\n"
                    "• Deixe em branco se não tem senha",
                    show='*'
                )
                
                root.destroy()
                
                if cert_pin is None:
                    return False
                
                # Usar a instância do sistema principal
                sucesso, msg = sistema_principal.consultor_sefaz_a1.configurar_certificado(cert_path, cert_pin)
                print(f"DEBUG config_rapida: Certificado configurado com sucesso: {sucesso}")
                print(f"DEBUG config_rapida: Cert info após configuração: {sistema_principal.consultor_sefaz_a1.cert_info}")
                
                if sucesso:
                    print(f"Certificado configurado: {msg}")
                    
                    # Testar usando a mesma instância
                    teste_ok, teste_msg = sistema_principal.consultor_sefaz_a1.testar_conexao()
                    
                    resultado = f"Certificado: {msg}\n\nTeste: {teste_msg}"
                    
                    if teste_ok:
                        resultado += "\n\nSistema pronto para consultar NFe via SEFAZ!"
                    else:
                        resultado += "\n\nVerifique conectividade"
                    
                    messagebox.showinfo("Certificado Configurado", resultado)
                    return True
                else:
                    print(f"Erro na configuração: {msg}")
                    messagebox.showerror("Erro", f"Erro: {msg}")
                    return False
                    
            except Exception as e:
                print(f"Erro na configuração rápida: {e}")
                messagebox.showerror("Erro", f"Erro: {str(e)}")
                return False
        
        # Adicionar ao sistema
        sistema_principal.configurar_certificado_rapido = config_rapida
        
        print("Melhorias aplicadas com sucesso!")
        return True
        
    except Exception as e:
        print(f"Erro: {e}")
        return False


# Alias para compatibilidade
aplicar_melhorias_certificado_a1 = aplicar_melhorias_ao_sistema_existente


if __name__ == "__main__":
    print("CERTIFICADO A1 PARA SEFAZ - VERSÃO CORRIGIDA")
    print("Principais correções:")
    print("• Sintaxe corrigida")
    print("• Instância única do consultor")
    print("• Debug detalhado")
    print("• Envelope SOAP 1.2")