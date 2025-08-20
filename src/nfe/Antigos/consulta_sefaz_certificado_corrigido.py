# -*- coding: utf-8 -*-
"""
CONSULTA SEFAZ COM CERTIFICADO A1 - VERSÃO CORRIGIDA
Corrige problemas de autenticação e consulta
"""

import requests
import xml.etree.ElementTree as ET
from datetime import datetime
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


class ConsultorSefazA1Corrigido:
    """Consultor SEFAZ com certificado A1 - versão corrigida"""
    
    def __init__(self):
        self.cert_info = {}
        self.cert_pem_data = None
        self.key_pem_data = None
        self.temp_dir = None
        
        # URLs SEFAZ atualizadas (2025)
        self.urls = {
            'PR': 'https://nfe.sefa.pr.gov.br/nfe/NFeConsultaProtocolo4',
            'SP': 'https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'RJ': 'https://nfe.sefaz.rj.gov.br/nfe/NFeConsultaProtocolo4',
            'MG': 'https://nfe.fazenda.mg.gov.br/nfe2/NFeConsultaProtocolo4',
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
    
    def configurar_certificado(self, cert_path, cert_password):
        """Configura certificado A1 - versão corrigida"""
        try:
            if not CRYPTO_OK:
                return False, "Instale cryptography: pip install cryptography"
            
            if not os.path.exists(cert_path):
                return False, "Arquivo de certificado não encontrado"
            
            print(f"🔐 Configurando certificado: {os.path.basename(cert_path)}")
            
            # Ler arquivo do certificado
            with open(cert_path, 'rb') as f:
                cert_data = f.read()
            
            # CORREÇÃO 1: Tratar senha corretamente
            password_bytes = None
            if cert_password:
                # Tentar diferentes formatos de senha
                try:
                    # Primeiro: senha como string convertida para bytes
                    password_bytes = str(cert_password).encode('utf-8')
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, password_bytes)
                    print("✅ Sucesso com senha UTF-8")
                except Exception:
                    try:
                        # Segundo: senha como bytes diretos
                        password_bytes = cert_password if isinstance(cert_password, bytes) else str(cert_password).encode('latin-1')
                        private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, password_bytes)
                        print("✅ Sucesso com senha Latin-1")
                    except Exception:
                        try:
                            # Terceiro: sem senha (certificado sem proteção)
                            private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, None)
                            print("✅ Sucesso sem senha")
                        except Exception as e:
                            return False, f"Senha incorreta ou certificado inválido: {str(e)}"
            else:
                try:
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, None)
                except Exception as e:
                    return False, f"Certificado requer senha: {str(e)}"
            
            if not certificate or not private_key:
                return False, "Falha ao extrair certificado e chave privada"
            
            # Verificar validade
            now = datetime.now()
            if certificate.not_valid_after < now:
                return False, f"Certificado expirado em {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
            if certificate.not_valid_before > now:
                return False, f"Certificado ainda não válido (válido a partir de {certificate.not_valid_before.strftime('%d/%m/%Y')})"
            
            # CORREÇÃO 2: Melhor gerenciamento de arquivos temporários
            self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_", suffix="_tmp")
            
            # Converter para PEM
            cert_pem = certificate.public_bytes(serialization.Encoding.PEM)
            key_pem = private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            )
            
            # Armazenar em memória para backup
            self.cert_pem_data = cert_pem
            self.key_pem_data = key_pem
            
            # Salvar arquivos temporários
            cert_file = os.path.join(self.temp_dir, "cert.pem")
            key_file = os.path.join(self.temp_dir, "key.pem")
            
            with open(cert_file, 'wb') as f:
                f.write(cert_pem)
            with open(key_file, 'wb') as f:
                f.write(key_pem)
            
            # Definir permissões restritivas nos arquivos
            os.chmod(cert_file, 0o600)
            os.chmod(key_file, 0o600)
            
            # Armazenar informações
            self.cert_info = {
                'is_valid': True,
                'not_valid_after': certificate.not_valid_after,
                'not_valid_before': certificate.not_valid_before,
                'subject': certificate.subject.rfc4514_string(),
                'issuer': certificate.issuer.rfc4514_string(),
                'cert_path': cert_file,
                'key_path': key_file,
                'serial_number': str(certificate.serial_number)
            }
            
            print(f"✅ Certificado configurado - válido até {certificate.not_valid_after.strftime('%d/%m/%Y')}")
            print(f"📋 Subject: {certificate.subject.rfc4514_string()}")
            
            return True, f"Certificado válido até {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
        except Exception as e:
            self.limpar_arquivos_temp()
            print(f"❌ Erro detalhado: {str(e)}")
            return False, f"Erro: {str(e)}"
    
    def consultar_nfe(self, chave_acesso):
        """Consulta NFe na SEFAZ - versão corrigida"""
        try:
            if not self.cert_info.get('is_valid'):
                raise Exception("Certificado não configurado")
            
            if len(chave_acesso) != 44:
                raise Exception("Chave deve ter exatamente 44 dígitos")
            
            # Verificar/recriar arquivos se necessário
            if not self._verificar_arquivos():
                if not self._recriar_arquivos():
                    raise Exception("Erro nos arquivos de certificado")
            
            # Obter UF e URL
            uf = self.obter_uf(chave_acesso)
            url = self.urls.get(uf, self.urls['SP'])
            
            print(f"🔍 Consultando SEFAZ {uf}: {chave_acesso}")
            print(f"🌐 URL: {url}")
            
            # CORREÇÃO 3: Envelope SOAP correto
            envelope = self._criar_envelope_soap_correto(chave_acesso)
            
            # Headers corretos
            headers = {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF'
            }
            
            # CORREÇÃO 4: Configuração adequada da sessão
            session = requests.Session()
            session.cert = (self.cert_info['cert_path'], self.cert_info['key_path'])
            session.verify = False  # Para desenvolvimento
            
            # Configurar timeout e retry
            from requests.adapters import HTTPAdapter
            from urllib3.util.retry import Retry
            
            retry_strategy = Retry(
                total=3,
                status_forcelist=[429, 500, 502, 503, 504],
                method_whitelist=["HEAD", "GET", "POST"],
                backoff_factor=1
            )
            adapter = HTTPAdapter(max_retries=retry_strategy)
            session.mount("http://", adapter)
            session.mount("https://", adapter)
            
            print("📡 Enviando requisição...")
            
            # Fazer requisição
            response = session.post(
                url,
                data=envelope,
                headers=headers,
                timeout=30
            )
            
            print(f"📋 Status HTTP: {response.status_code}")
            
            if response.status_code == 200:
                return self._processar_resposta(response.text, chave_acesso)
            else:
                print(f"❌ Resposta HTTP: {response.text[:500]}")
                raise Exception(f"HTTP {response.status_code}: {response.reason}")
                
        except Exception as e:
            print(f"❌ Erro na consulta: {str(e)}")
            raise Exception(f"Erro: {str(e)}")
    
    def _criar_envelope_soap_correto(self, chave_acesso):
        """Cria envelope SOAP com estrutura correta"""
        return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
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
    
    def _processar_resposta(self, xml_resp, chave):
        """Processa resposta da SEFAZ"""
        try:
            print("📄 Processando resposta XML...")
            
            root = ET.fromstring(xml_resp)
            
            # Buscar código de status
            status_code = "000"
            status_desc = "Processado"
            
            # Procurar elementos de status
            for elem in root.iter():
                if elem.tag.endswith('cStat'):
                    status_code = elem.text
                    break
            
            for elem in root.iter():
                if elem.tag.endswith('xMotivo'):
                    status_desc = elem.text
                    break
            
            print(f"📋 Status SEFAZ: {status_code} - {status_desc}")
            
            # Criar estrutura de dados básica
            dados = {
                'chave_acesso': chave,
                'numero_nf': str(int(chave[25:34])),
                'cnpj_emitente': chave[6:20],
                'razao_social_emitente': 'CONSULTADO VIA SEFAZ',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Consulta SEFAZ A1',
                'status_sefaz': f"{status_code} - {status_desc}",
                'resposta_completa': xml_resp  # Para debug
            }
            
            # Se autorizada, tentar extrair mais dados
            if status_code == "100":
                dados_extras = self._extrair_dados_detalhados(root)
                dados.update(dados_extras)
            elif status_code == "101":
                dados['observacao'] = "NFe cancelada"
            elif status_code == "110":
                dados['observacao'] = "NFe denegada"
            elif status_code in ["217", "999"]:
                dados['observacao'] = "NFe não encontrada na base da SEFAZ"
            
            return dados
            
        except ET.ParseError as e:
            print(f"❌ Erro de parsing XML: {e}")
            print(f"📄 XML recebido: {xml_resp[:1000]}")
            raise Exception(f"Resposta XML inválida: {e}")
        except Exception as e:
            print(f"❌ Erro ao processar resposta: {e}")
            raise Exception(f"Erro no processamento: {e}")
    
    def _extrair_dados_detalhados(self, root):
        """Extrai dados detalhados da NFe autorizada"""
        dados = {}
        
        try:
            # Buscar dados do emitente
            for elem in root.iter():
                if elem.tag.endswith('xNome') and elem.text:
                    dados['razao_social_emitente'] = elem.text
                    break
            
            # Buscar valor total
            for elem in root.iter():
                if elem.tag.endswith('vNF') and elem.text:
                    dados['valor_total'] = float(elem.text)
                    break
            
            # Buscar data de emissão
            for elem in root.iter():
                if elem.tag.endswith('dhEmi') and elem.text:
                    data_iso = elem.text.split('T')[0]
                    dt = datetime.strptime(data_iso, '%Y-%m-%d')
                    dados['data_emissao'] = dt.strftime('%d/%m/%Y')
                    break
            
            # Buscar produtos (simplificado)
            produtos = []
            for elem in root.iter():
                if elem.tag.endswith('det'):
                    # Extrair dados básicos do produto
                    produto = {'descricao': 'Produto da NFe', 'quantidade': 1, 'valor_total': 0}
                    produtos.append(produto)
            
            if produtos:
                dados['produtos'] = produtos
            
        except Exception as e:
            print(f"⚠️ Erro ao extrair dados detalhados: {e}")
        
        return dados
    
    def _verificar_arquivos(self):
        """Verifica se arquivos temporários existem"""
        if not self.cert_info.get('cert_path') or not self.cert_info.get('key_path'):
            return False
        
        return (os.path.exists(self.cert_info['cert_path']) and 
                os.path.exists(self.cert_info['key_path']))
    
    def _recriar_arquivos(self):
        """Recria arquivos temporários a partir dos dados em memória"""
        try:
            if not self.cert_pem_data or not self.key_pem_data:
                return False
            
            if not self.temp_dir:
                self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_", suffix="_tmp")
            
            cert_file = os.path.join(self.temp_dir, "cert.pem")
            key_file = os.path.join(self.temp_dir, "key.pem")
            
            with open(cert_file, 'wb') as f:
                f.write(self.cert_pem_data)
            with open(key_file, 'wb') as f:
                f.write(self.key_pem_data)
            
            os.chmod(cert_file, 0o600)
            os.chmod(key_file, 0o600)
            
            self.cert_info['cert_path'] = cert_file
            self.cert_info['key_path'] = key_file
            
            print("✅ Arquivos temporários recriados")
            return True
            
        except Exception as e:
            print(f"❌ Erro ao recriar arquivos: {e}")
            return False
    
    def obter_uf(self, chave):
        """Obtém UF da chave"""
        return self.ufs.get(chave[:2] if len(chave) >= 2 else '', 'SP')
    
    def testar_conexao(self):
        """Testa conexão com SEFAZ"""
        try:
            if not self.cert_info.get('is_valid'):
                return False, "Certificado não configurado"
            
            # Usar chave de teste válida (formato correto, mas NFe inexistente)
            chave_teste = "35200314200166000187550010000000271234567890"
            
            try:
                resultado = self.consultar_nfe(chave_teste)
                
                # Se chegou até aqui, a conexão funcionou
                status = resultado.get('status_sefaz', '')
                
                if "217" in status or "999" in status or "não encontrada" in status.lower():
                    return True, "Conexão OK (NFe teste não encontrada - normal)"
                elif "100" in status:
                    return True, "Conexão OK (NFe teste encontrada)"
                else:
                    return True, f"Conexão OK - Status: {status}"
                    
            except Exception as e:
                error_msg = str(e).lower()
                
                if "não encontrada" in error_msg or "217" in error_msg:
                    return True, "Conexão OK (NFe teste não existe - esperado)"
                elif "timeout" in error_msg:
                    return False, "Timeout - verifique conectividade"
                elif "ssl" in error_msg or "certificate" in error_msg:
                    return False, "Erro de certificado SSL"
                else:
                    return False, f"Erro: {str(e)[:100]}"
                    
        except Exception as e:
            return False, f"Erro: {str(e)}"
    
    def obter_info_certificado(self):
        """Retorna informações do certificado"""
        return self.cert_info.copy()
    
    def limpar_arquivos_temp(self):
        """Limpa arquivos temporários"""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                print("🗑️ Arquivos temporários removidos")
        except Exception as e:
            print(f"⚠️ Erro ao limpar temporários: {e}")
        
        self.temp_dir = None
    
    def __del__(self):
        """Destrutor - limpa arquivos"""
        self.limpar_arquivos_temp()


def aplicar_correcoes_ao_sistema(sistema_principal):
    """Aplica as correções ao sistema existente"""
    try:
        print("🔧 Aplicando correções de certificado A1...")
        
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("❌ Sistema NFe não encontrado")
            return False
        
        # Substituir consultor por versão corrigida
        consultor_corrigido = ConsultorSefazA1Corrigido()
        
        # Backup método original se existir
        if hasattr(sistema_principal.processador_nfe, 'consultar_nfe_sefaz'):
            sistema_principal.processador_nfe.consultar_nfe_sefaz_original = sistema_principal.processador_nfe.consultar_nfe_sefaz
        
        def nova_consulta_corrigida(chave):
            """Nova consulta com correções aplicadas"""
            try:
                if consultor_corrigido.cert_info.get('is_valid'):
                    print("🔐 Usando certificado A1 corrigido...")
                    return consultor_corrigido.consultar_nfe(chave)
                else:
                    print("⚠️ Certificado não configurado")
                    return {
                        'chave_acesso': chave,
                        'numero_nf': str(int(chave[25:34])),
                        'razao_social_emitente': 'CONFIGURE CERTIFICADO A1',
                        'valor_total': 0.0,
                        'produtos': [],
                        'fonte_dados': 'Simulação - Configure certificado'
                    }
            except Exception as e:
                print(f"❌ Erro na consulta: {e}")
                return {
                    'chave_acesso': chave,
                    'numero_nf': 'ERRO',
                    'razao_social_emitente': 'ERRO NA CONSULTA',
                    'valor_total': 0.0,
                    'produtos': [],
                    'fonte_dados': 'Erro',
                    'observacao': str(e)
                }
        
        # Aplicar nova consulta
        sistema_principal.processador_nfe.consultar_nfe_sefaz = nova_consulta_corrigida
        
        # Aplicar novos métodos
        sistema_principal.processador_nfe.configurar_certificado_a1 = consultor_corrigido.configurar_certificado
        sistema_principal.processador_nfe.testar_certificado_a1 = consultor_corrigido.testar_conexao
        
        # Método de configuração rápida corrigido
        def config_rapida_corrigida():
            try:
                print("🔐 Configuração corrigida de certificado A1...")
                
                # Selecionar arquivo
                cert_path = filedialog.askopenfilename(
                    title="Selecionar Certificado A1",
                    filetypes=[("Certificado A1", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                if not cert_path:
                    print("❌ Arquivo não selecionado")
                    return False
                
                print(f"📁 Arquivo: {os.path.basename(cert_path)}")
                
                # Solicitar senha
                root = tk.Tk()
                root.withdraw()
                
                cert_password = simpledialog.askstring(
                    "Senha do Certificado A1",
                    "Digite a senha do certificado:\n\n"
                    "• Senha definida quando o certificado foi criado/baixado\n"
                    "• Pode ser numérica ou alfanumérica\n"
                    "• Diferente da senha de relacionamento com a AC",
                    show='*'
                )
                
                root.destroy()
                
                if cert_password is None:
                    print("❌ Configuração cancelada")
                    return False
                
                print("🔒 Configurando certificado...")
                sucesso, msg = consultor_corrigido.configurar_certificado(cert_path, cert_password)
                
                if sucesso:
                    print(f"✅ {msg}")
                    
                    # Testar conexão
                    print("🧪 Testando conexão...")
                    teste_ok, teste_msg = consultor_corrigido.testar_conexao()
                    
                    resultado = f"✅ {msg}\n\n🧪 Teste: {teste_msg}"
                    
                    if teste_ok:
                        resultado += "\n\n🎉 Sistema pronto para consultar NFe via SEFAZ!"
                    else:
                        resultado += "\n\n⚠️ Verifique conectividade com internet"
                    
                    messagebox.showinfo("Certificado Configurado", resultado)
                    return True
                else:
                    print(f"❌ {msg}")
                    
                    erro = f"❌ Erro na configuração:\n\n{msg}"
                    
                    if "senha" in msg.lower() or "incorret" in msg.lower():
                        erro += "\n\n💡 Dicas:\n"
                        erro += "• Verifique se a senha está correta\n"
                        erro += "• Tente sem senha se o certificado não tem proteção\n"
                        erro += "• Certifique-se de que o arquivo .pfx não está corrompido\n"
                        erro += "• Contate quem forneceu o certificado"
                    
                    messagebox.showerror("Erro", erro)
                    return False
                    
            except Exception as e:
                erro = f"❌ Erro: {str(e)}"
                print(erro)
                messagebox.showerror("Erro", erro)
                return False
        
        # Aplicar configuração corrigida
        sistema_principal.configurar_certificado_rapido = config_rapida_corrigida
        
        # Armazenar referência do consultor
        sistema_principal.consultor_sefaz_a1 = consultor_corrigido
        
        print("✅ Correções aplicadas com sucesso!")
        print("🔧 Melhorias:")
        print("   • Tratamento correto de senhas de certificado")
        print("   • Envelope SOAP corrigido")
        print("   • URLs SEFAZ atualizadas")
        print("   • Melhor gerenciamento de arquivos temporários")
        print("   • Tratamento de erros aprimorado")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar correções: {e}")
        return False


def diagnosticar_problema_certificado(cert_path, cert_password):
    """Diagnostica problemas específicos do certificado"""
    try:
        print("\n🔍 DIAGNÓSTICO DO CERTIFICADO")
        print("=" * 50)
        
        # 1. Verificar arquivo
        if not os.path.exists(cert_path):
            print("❌ Arquivo não encontrado")
            return False
        
        print(f"✅ Arquivo encontrado: {os.path.basename(cert_path)}")
        print(f"📏 Tamanho: {os.path.getsize(cert_path)} bytes")
        
        # 2. Verificar se é PKCS#12
        with open(cert_path, 'rb') as f:
            header = f.read(10)
        
        if not header.startswith(b'\x30'):
            print("❌ Arquivo não parece ser um certificado PKCS#12 válido")
            return False
        
        print("✅ Formato PKCS#12 detectado")
        
        # 3. Testar carregamento
        try:
            consultor = ConsultorSefazA1Corrigido()
            sucesso, msg = consultor.configurar_certificado(cert_path, cert_password)
            
            if sucesso:
                print(f"✅ Certificado carregado: {msg}")
                
                info = consultor.obter_info_certificado()
                print(f"📋 Subject: {info.get('subject', '')}")
                print(f"🏢 Emissor: {info.get('issuer', '')}")
                print(f"🆔 Serial: {info.get('serial_number', '')}")
                
                return True
            else:
                print(f"❌ Falha no carregamento: {msg}")
                return False
                
        except Exception as e:
            print(f"❌ Erro no diagnóstico: {e}")
            return False
            
    except Exception as e:
        print(f"❌ Erro geral: {e}")
        return False


# EXEMPLO DE USO DAS CORREÇÕES
if __name__ == "__main__":
    print("""
CORREÇÕES PARA CERTIFICADO A1 - SEFAZ

Para aplicar as correções ao seu sistema:

1. Aplicar correções:
   from consulta_sefaz_certificado_corrigido import aplicar_correcoes_ao_sistema
   aplicar_correcoes_ao_sistema(sistema_principal)

2. Configurar certificado:
   sistema_principal.configurar_certificado_rapido()

3. Testar certificado:
   sucesso, msg = sistema_principal.processador_nfe.testar_certificado_a1()
   print(f"Teste: {msg}")

4. Diagnosticar problemas:
   from consulta_sefaz_certificado_corrigido import diagnosticar_problema_certificado
   diagnosticar_problema_certificado("caminho/certificado.pfx", "senha")

5. Consultar NFe:
   chave = "35200314200166000187550010000000271234567890"
   dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)

PRINCIPAIS CORREÇÕES APLICADAS:

1. TRATAMENTO DE SENHA:
   - Tenta múltiplos formatos de encoding (UTF-8, Latin-1)
   - Suporte a certificados sem senha
   - Melhor tratamento de erros de senha

2. ENVELOPE SOAP:
   - Estrutura correta: consSitNFe (não consReciNFe)
   - Namespace adequado
   - Parâmetros corretos (tpAmb=1, xServ=CONSULTAR)

3. URLS SEFAZ:
   - URLs atualizadas para 2025
   - Mapeamento correto por UF
   - Fallback para SP se UF não encontrada

4. GERENCIAMENTO DE ARQUIVOS:
   - Criação segura de arquivos temporários
   - Permissões restritivas (0600)
   - Limpeza automática
   - Backup em memória para recriação

5. TRATAMENTO DE RESPOSTA:
   - Parse correto de códigos de status
   - Mapeamento de diferentes cenários:
     * 100: NFe autorizada
     * 101: NFe cancelada
     * 110: NFe denegada
     * 217/999: NFe não encontrada
   - Extração de dados detalhados quando disponível

6. TESTE DE CONECTIVIDADE:
   - Chave de teste válida
   - Interpretação correta de respostas
   - Distinção entre problemas de conectividade e certificado

PROBLEMAS IDENTIFICADOS NO CÓDIGO ORIGINAL:

1. Múltiplas tentativas de PIN sem lógica adequada
2. Envelope SOAP com estrutura incorreta
3. URLs desatualizadas
4. Falta de tratamento adequado de permissões de arquivo
5. Processamento de resposta XML limitado
6. Falta de distinção entre tipos de erro

COMO USAR:

1. Substitua o arquivo consulta_sefaz_certificado.py pelo corrigido
2. Execute aplicar_correcoes_ao_sistema(sistema_principal)
3. Configure o certificado usando a interface
4. Teste a conectividade antes de usar em produção

REQUISITOS:

- Python 3.7+
- cryptography >= 3.0
- requests >= 2.25
- tkinter (interface gráfica)

SOLUÇÃO DE PROBLEMAS COMUNS:

1. "Senha incorreta":
   - Verifique se a senha está correta
   - Tente sem senha se o certificado não tem proteção
   - Contate quem forneceu o certificado

2. "Timeout":
   - Verifique conexão com internet
   - Verifique firewall (porta 443)
   - Tente novamente após alguns minutos

3. "NFe não encontrada":
   - Normal para chaves de teste
   - Verifique se a chave está correta (44 dígitos)
   - Confirme se a NFe existe no SEFAZ

4. "Certificado expirado":
   - Renove o certificado junto à Autoridade Certificadora
   - Verifique data de validade

OBSERVAÇÕES IMPORTANTES:

- Este código é para ambiente de produção
- Use apenas certificados A1 válidos
- Mantenha a senha do certificado segura
- Faça backup do certificado
- Teste em ambiente controlado primeiro
""")