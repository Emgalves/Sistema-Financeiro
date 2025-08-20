# -*- coding: utf-8 -*-
"""
CORREÇÃO COMPLETA DO SISTEMA DE CERTIFICADO A1
Salve como: src/nfe/correcao_certificado_a1.py
"""

import requests
import xml.etree.ElementTree as ET
from datetime import datetime
import tempfile
import os
import shutil
import tkinter as tk
from tkinter import messagebox, filedialog, simpledialog
from pathlib import Path

# Importar cryptography com tratamento de erro
try:
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.serialization import pkcs12
    from cryptography import x509
    CRYPTO_OK = True
except ImportError:
    CRYPTO_OK = False

# Desabilitar warnings SSL
import urllib3
urllib3.disable_warnings()


class ConsultorSefazA1Corrigido:
    """Consultor SEFAZ com certificado A1 - VERSÃO CORRIGIDA"""
    
    def __init__(self):
        self.cert_info = {}
        self.cert_data = None
        self.key_data = None
        self.temp_dir = None
        self.certificado_valido = False
        
        # URLs SEFAZ atualizadas 2025
        self.urls_sefaz = {
            'PR': 'https://nfe.sefa.pr.gov.br/ws/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx',
            'SP': 'https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx', 
            'RJ': 'https://nfe.sefaz.rj.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'MG': 'https://nfe.fazenda.mg.gov.br/nfe2/services/NFeConsultaProtocolo4',
            'RS': 'https://nfe.sefazrs.rs.gov.br/ws/NfeConsulta/NfeConsulta4.asmx',
            'SC': 'https://nfe.sef.sc.gov.br/ws/nfeconsultaprotocolo4.asmx',
            'GO': 'https://nfe.sefaz.go.gov.br/nfe/services/NFeConsultaProtocolo4',
            'MT': 'https://nfe.sefaz.mt.gov.br/nfews/v2/services/NfeConsulta4',
            'MS': 'https://nfe.sefaz.ms.gov.br/ws/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx',
            'BA': 'https://nfe.sefaz.ba.gov.br/ws/NFeConsultaProtocolo4/NFeConsultaProtocolo4.asmx',
            'CE': 'https://nfe.sefaz.ce.gov.br/nfe2/services/NFeConsultaProtocolo4',
            'PE': 'https://nfe.sefaz.pe.gov.br/nfe-service/services/NFeConsultaProtocolo4',
            'DF': 'https://dec.fazenda.df.gov.br/ws/nfeconsultaprotocolo4.asmx'
        }
        
        # Códigos UF
        self.codigos_uf = {
            '11': 'RO', '12': 'AC', '13': 'AM', '14': 'RR', '15': 'PA', '16': 'AP', '17': 'TO',
            '21': 'MA', '22': 'PI', '23': 'CE', '24': 'RN', '25': 'PB', '26': 'PE', '27': 'AL', 
            '28': 'SE', '29': 'BA', '31': 'MG', '32': 'ES', '33': 'RJ', '35': 'SP', '41': 'PR', 
            '42': 'SC', '43': 'RS', '50': 'MS', '51': 'MT', '52': 'GO', '53': 'DF'
        }
    
    def configurar_certificado(self, cert_path, cert_password):
        """Configura certificado A1 com validação completa"""
        try:
            print(f"🔑 Configurando certificado: {os.path.basename(cert_path)}")
            
            if not CRYPTO_OK:
                return False, "❌ Biblioteca cryptography não encontrada. Execute: pip install cryptography"
            
            if not os.path.exists(cert_path):
                return False, f"❌ Arquivo não encontrado: {cert_path}"
            
            # Ler arquivo do certificado
            with open(cert_path, 'rb') as f:
                cert_data = f.read()
            
            # Testar diferentes formatos de senha
            senhas_teste = self._gerar_variacoes_senha(cert_password)
            
            private_key = None
            certificate = None
            
            # Tentar carregar com diferentes senhas
            for i, senha in enumerate(senhas_teste):
                try:
                    print(f"🧪 Testando senha formato {i+1}/{len(senhas_teste)}...")
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, senha)
                    if certificate and private_key:
                        print(f"✅ Sucesso com formato de senha {i+1}!")
                        break
                except Exception as e:
                    if i == len(senhas_teste) - 1:  # Última tentativa
                        print(f"❌ Última tentativa falhou: {e}")
                    continue
            
            if not certificate or not private_key:
                return False, "❌ Senha incorreta ou certificado inválido. Verifique:\n• PIN de 6 dígitos\n• Senha de relacionamento\n• Arquivo .pfx correto"
            
            # Verificar validade do certificado
            agora = datetime.now()
            if certificate.not_valid_after < agora:
                return False, f"❌ Certificado expirado em {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
            if certificate.not_valid_before > agora:
                return False, f"❌ Certificado ainda não válido (válido a partir de {certificate.not_valid_before.strftime('%d/%m/%Y')})"
            
            # Criar diretório temporário seguro
            self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_", suffix="_secure")
            os.chmod(self.temp_dir, 0o700)  # Apenas proprietário pode acessar
            
            # Converter para PEM
            cert_pem = certificate.public_bytes(serialization.Encoding.PEM)
            key_pem = private_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.PKCS8,
                encryption_algorithm=serialization.NoEncryption()
            )
            
            # Salvar arquivos temporários
            cert_file = os.path.join(self.temp_dir, "certificado.pem")
            key_file = os.path.join(self.temp_dir, "chave_privada.pem")
            
            with open(cert_file, 'wb') as f:
                f.write(cert_pem)
            with open(key_file, 'wb') as f:
                f.write(key_pem)
            
            # Definir permissões restritivas
            os.chmod(cert_file, 0o600)
            os.chmod(key_file, 0o600)
            
            # Backup em memória
            self.cert_data = cert_pem
            self.key_data = key_pem
            
            # Extrair informações do certificado
            subject_info = self._extrair_subject_info(certificate)
            
            # Armazenar informações
            self.cert_info = {
                'is_valid': True,
                'not_valid_after': certificate.not_valid_after,
                'not_valid_before': certificate.not_valid_before,
                'subject': certificate.subject.rfc4514_string(),
                'subject_info': subject_info,
                'serial_number': str(certificate.serial_number),
                'cert_path': cert_file,
                'key_path': key_file,
                'issuer': certificate.issuer.rfc4514_string()
            }
            
            self.certificado_valido = True
            
            print(f"✅ Certificado configurado com sucesso!")
            print(f"📋 Proprietário: {subject_info.get('CN', 'N/A')}")
            print(f"📅 Válido até: {certificate.not_valid_after.strftime('%d/%m/%Y %H:%M')}")
            
            return True, f"Certificado válido até {certificate.not_valid_after.strftime('%d/%m/%Y')}"
            
        except Exception as e:
            self._limpar_arquivos_temp()
            return False, f"❌ Erro na configuração: {str(e)}"
    
    def _gerar_variacoes_senha(self, senha_original):
        """Gera variações da senha para teste"""
        senhas = []
        senha_str = str(senha_original)
        
        # Senha original
        senhas.append(senha_str.encode('utf-8'))
        
        # Senha como bytes direto
        if isinstance(senha_original, str):
            senhas.append(senha_original.encode('utf-8'))
        
        # Sem senha (alguns certificados)
        senhas.append(b'')
        
        # Se tem 6 dígitos, testar variações
        if len(senha_str) == 6 and senha_str.isdigit():
            # Com zeros à direita
            senhas.append(senha_str.ljust(8, '0').encode('utf-8'))
            # Com zeros à esquerda
            senhas.append(senha_str.zfill(8).encode('utf-8'))
            # Repetir 2x (alguns certificados usam isso)
            senhas.append((senha_str + senha_str).encode('utf-8'))
        
        # Senha em maiúsculo/minúsculo
        senhas.append(senha_str.upper().encode('utf-8'))
        senhas.append(senha_str.lower().encode('utf-8'))
        
        # Remover duplicatas mantendo ordem
        senhas_unicas = []
        for senha in senhas:
            if senha not in senhas_unicas:
                senhas_unicas.append(senha)
        
        return senhas_unicas
    
    def _extrair_subject_info(self, certificate):
        """Extrai informações do subject do certificado"""
        subject_info = {}
        
        for attribute in certificate.subject:
            if attribute.oid._name == 'commonName':
                subject_info['CN'] = attribute.value
            elif attribute.oid._name == 'organizationName':
                subject_info['O'] = attribute.value
            elif attribute.oid._name == 'countryName':
                subject_info['C'] = attribute.value
            elif attribute.oid._name == 'stateOrProvinceName':
                subject_info['ST'] = attribute.value
            elif attribute.oid._name == 'localityName':
                subject_info['L'] = attribute.value
        
        return subject_info
    
    def verificar_certificado_configurado(self):
        """Verifica se certificado está configurado e válido"""
        if not self.certificado_valido or not self.cert_info.get('is_valid'):
            return False, "❌ Certificado não configurado"
        
        # Verificar se arquivos ainda existem
        cert_path = self.cert_info.get('cert_path')
        key_path = self.cert_info.get('key_path')
        
        if not cert_path or not key_path:
            return False, "❌ Caminhos de arquivos não definidos"
        
        if not os.path.exists(cert_path) or not os.path.exists(key_path):
            print("⚠️ Arquivos temporários perdidos, recriando...")
            if self._recriar_arquivos_temporarios():
                return True, "✅ Certificado válido (arquivos recriados)"
            else:
                return False, "❌ Falha ao recriar arquivos temporários"
        
        # Verificar validade temporal
        if self.cert_info['not_valid_after'] < datetime.now():
            return False, f"❌ Certificado expirado em {self.cert_info['not_valid_after'].strftime('%d/%m/%Y')}"
        
        return True, "✅ Certificado configurado e válido"
    
    def _recriar_arquivos_temporarios(self):
        """Recria arquivos temporários a partir da memória"""
        try:
            if not self.cert_data or not self.key_data:
                return False
            
            if not self.temp_dir or not os.path.exists(self.temp_dir):
                self.temp_dir = tempfile.mkdtemp(prefix="nfe_cert_", suffix="_secure")
                os.chmod(self.temp_dir, 0o700)
            
            cert_file = os.path.join(self.temp_dir, "certificado.pem")
            key_file = os.path.join(self.temp_dir, "chave_privada.pem")
            
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
            print(f"❌ Erro ao recriar arquivos: {e}")
            return False
    
    def consultar_nfe(self, chave_acesso):
        """Consulta NFe na SEFAZ com certificado A1"""
        try:
            # Verificar certificado
            cert_ok, cert_msg = self.verificar_certificado_configurado()
            if not cert_ok:
                raise Exception(cert_msg)
            
            # Validar chave
            if len(chave_acesso) != 44 or not chave_acesso.isdigit():
                raise Exception("❌ Chave de acesso deve ter 44 dígitos")
            
            # Determinar UF e URL
            uf_codigo = chave_acesso[:2]
            uf_sigla = self.codigos_uf.get(uf_codigo, 'SP')
            url_webservice = self.urls_sefaz.get(uf_sigla, self.urls_sefaz['SP'])
            
            print(f"🔍 Consultando SEFAZ {uf_sigla}: {chave_acesso}")
            print(f"🌐 URL: {url_webservice}")
            
            # Criar envelope SOAP
            envelope = self._criar_envelope_consulta(chave_acesso)
            
            # Headers
            headers = {
                'Content-Type': 'text/xml; charset=utf-8',
                'SOAPAction': 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF'
            }
            
            # Fazer requisição com certificado
            response = requests.post(
                url_webservice,
                data=envelope,
                headers=headers,
                cert=(self.cert_info['cert_path'], self.cert_info['key_path']),
                verify=False,  # Em produção, usar True com certificados de CA válidos
                timeout=30
            )
            
            print(f"📡 Status HTTP: {response.status_code}")
            
            if response.status_code == 200:
                return self._processar_resposta_sefaz(response.text, chave_acesso)
            else:
                raise Exception(f"❌ Erro HTTP {response.status_code}: {response.text[:200]}")
                
        except Exception as e:
            print(f"❌ Erro na consulta: {e}")
            raise Exception(f"Falha na consulta SEFAZ: {str(e)}")
    
    def _criar_envelope_consulta(self, chave_acesso):
        """Cria envelope SOAP para consulta"""
        return f"""<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" 
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
               xmlns:xsd="http://www.w3.org/2001/XMLSchema">
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
    
    def _processar_resposta_sefaz(self, xml_resposta, chave_acesso):
        """Processa resposta da SEFAZ"""
        try:
            root = ET.fromstring(xml_resposta)
            
            # Buscar código de status
            status_code = "000"
            status_msg = "Processado"
            
            # Procurar elementos de status
            for elem in root.iter():
                tag_name = elem.tag.lower()
                if 'cstat' in tag_name and elem.text:
                    status_code = elem.text
                    break
            
            for elem in root.iter():
                tag_name = elem.tag.lower()
                if 'xmotivo' in tag_name and elem.text:
                    status_msg = elem.text
                    break
            
            print(f"📋 Status SEFAZ: {status_code} - {status_msg}")
            
            # Dados básicos da NFe
            dados_nfe = {
                'chave_acesso': chave_acesso,
                'numero_nf': str(int(chave_acesso[25:34])),
                'cnpj_emitente': self._formatar_cnpj(chave_acesso[6:20]),
                'razao_social_emitente': 'CONSULTADO VIA SEFAZ',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Consulta SEFAZ com Certificado A1',
                'status_sefaz': f"{status_code} - {status_msg}",
                'uf_emitente': self.codigos_uf.get(chave_acesso[:2], 'SP')
            }
            
            # Se NFe foi encontrada (status 100), extrair mais dados
            if status_code == "100":
                dados_nfe.update(self._extrair_dados_detalhados(root))
            elif status_code == "101":
                dados_nfe['observacao'] = "NFe cancelada"
            elif status_code == "110":
                dados_nfe['observacao'] = "NFe denegada"
            elif status_code == "217":
                dados_nfe['observacao'] = "NFe não encontrada na base de dados"
            
            return dados_nfe
            
        except Exception as e:
            print(f"⚠️ Erro ao processar resposta: {e}")
            # Retornar dados básicos em caso de erro
            return {
                'chave_acesso': chave_acesso,
                'numero_nf': str(int(chave_acesso[25:34])),
                'cnpj_emitente': self._formatar_cnpj(chave_acesso[6:20]),
                'razao_social_emitente': 'ERRO NO PROCESSAMENTO',
                'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                'valor_total': 0.0,
                'produtos': [],
                'fonte_dados': 'Erro na Consulta',
                'observacao': f"Erro ao processar resposta: {str(e)}"
            }
    
    def _extrair_dados_detalhados(self, root):
        """Extrai dados detalhados da NFe da resposta SEFAZ"""
        dados = {}
        
        try:
            # Namespace da NFe
            namespaces = {
                'nfe': 'http://www.portalfiscal.inf.br/nfe'
            }
            
            # Buscar dados do emitente
            emit = root.find('.//nfe:emit', namespaces)
            if emit is not None:
                nome_elem = emit.find('nfe:xNome', namespaces)
                if nome_elem is not None and nome_elem.text:
                    dados['razao_social_emitente'] = nome_elem.text
            
            # Buscar data de emissão
            ide = root.find('.//nfe:ide', namespaces)
            if ide is not None:
                dh_emi = ide.find('nfe:dhEmi', namespaces)
                if dh_emi is not None and dh_emi.text:
                    try:
                        dt = datetime.fromisoformat(dh_emi.text.replace('Z', '+00:00'))
                        dados['data_emissao'] = dt.strftime('%d/%m/%Y')
                    except:
                        pass
            
            # Buscar valor total
            total = root.find('.//nfe:total/nfe:ICMSTot', namespaces)
            if total is not None:
                v_nf = total.find('nfe:vNF', namespaces)
                if v_nf is not None and v_nf.text:
                    try:
                        dados['valor_total'] = float(v_nf.text)
                    except:
                        pass
            
            # Buscar produtos (simplificado)
            itens = root.findall('.//nfe:det', namespaces)
            produtos = []
            
            for item in itens[:5]:  # Limitar a 5 primeiros produtos
                prod = item.find('nfe:prod', namespaces)
                if prod is not None:
                    produto = {
                        'codigo': self._get_text_safe(prod.find('nfe:cProd', namespaces)),
                        'descricao': self._get_text_safe(prod.find('nfe:xProd', namespaces)),
                        'quantidade': self._get_float_safe(prod.find('nfe:qCom', namespaces)),
                        'unidade': self._get_text_safe(prod.find('nfe:uCom', namespaces)),
                        'valor_unitario': self._get_float_safe(prod.find('nfe:vUnCom', namespaces)),
                        'valor_total': self._get_float_safe(prod.find('nfe:vProd', namespaces))
                    }
                    produtos.append(produto)
            
            dados['produtos'] = produtos
            
        except Exception as e:
            print(f"⚠️ Erro ao extrair dados detalhados: {e}")
        
        return dados
    
    def _get_text_safe(self, element):
        """Extrai texto de elemento XML com segurança"""
        return element.text if element is not None and element.text else ""
    
    def _get_float_safe(self, element):
        """Extrai float de elemento XML com segurança"""
        try:
            return float(element.text) if element is not None and element.text else 0.0
        except:
            return 0.0
    
    def _formatar_cnpj(self, cnpj):
        """Formata CNPJ"""
        if len(cnpj) == 14:
            return f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
        return cnpj
    
    def testar_conectividade(self):
        """Testa conectividade com SEFAZ"""
        try:
            cert_ok, cert_msg = self.verificar_certificado_configurado()
            if not cert_ok:
                return False, cert_msg
            
            # Chave de teste (formato válido)
            chave_teste = "35210714200166000187550010000000271234567890"
            
            print("🧪 Testando conectividade com SEFAZ...")
            
            try:
                resultado = self.consultar_nfe(chave_teste)
                
                if resultado:
                    status = resultado.get('status_sefaz', '')
                    if '217' in status or 'não encontrada' in status.lower():
                        return True, "✅ Conectividade OK (NFe teste não encontrada - comportamento esperado)"
                    else:
                        return True, f"✅ Conectividade OK - Status: {status}"
                else:
                    return False, "❌ Resposta vazia da SEFAZ"
                    
            except Exception as e:
                error_msg = str(e).lower()
                if 'não encontrada' in error_msg or '217' in error_msg:
                    return True, "✅ Conectividade OK (erro esperado com chave teste)"
                elif 'timeout' in error_msg:
                    return False, "❌ Timeout - Verifique conexão com internet"
                elif 'certificate' in error_msg:
                    return False, "❌ Problema com certificado - Reconfigure"
                else:
                    return False, f"❌ Erro: {str(e)[:100]}..."
                    
        except Exception as e:
            return False, f"❌ Erro no teste: {str(e)}"
    
    def obter_info_certificado(self):
        """Retorna informações do certificado"""
        return self.cert_info.copy()
    
    def _limpar_arquivos_temp(self):
        """Limpa arquivos temporários"""
        try:
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
                print("🧹 Arquivos temporários limpos")
        except Exception as e:
            print(f"⚠️ Erro ao limpar arquivos: {e}")
        
        self.temp_dir = None
    
    def __del__(self):
        """Destrutor - limpa arquivos automaticamente"""
        self._limpar_arquivos_temp()


def corrigir_sistema_certificado_a1(sistema_principal):
    """Corrige e atualiza o sistema de certificado A1"""
    try:
        print("\n🔧 CORREÇÃO DO SISTEMA DE CERTIFICADO A1")
        print("=" * 50)
        
        # Verificar se sistema híbrido existe
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("❌ Sistema híbrido NFe não encontrado!")
            return False
        
        # Criar consultor corrigido
        consultor_corrigido = ConsultorSefazA1Corrigido()
        
        # Substituir consultor antigo
        sistema_principal.consultor_sefaz_a1 = consultor_corrigido
        
        # Backup do método original se existir
        processador = sistema_principal.processador_nfe
        if hasattr(processador, 'consultar_nfe_sefaz'):
            processador.consultar_nfe_sefaz_backup = processador.consultar_nfe_sefaz
        
        # Substituir método de consulta
        def consulta_corrigida(chave):
            """Método de consulta corrigido"""
            try:
                if consultor_corrigido.certificado_valido:
                    print("🔐 Usando certificado A1 (versão corrigida)...")
                    return consultor_corrigido.consultar_nfe(chave)
                else:
                    print("⚠️ Certificado não configurado - usando dados simulados")
                    return {
                        'chave_acesso': chave,
                        'numero_nf': str(int(chave[25:34])),
                        'cnpj_emitente': consultor_corrigido._formatar_cnpj(chave[6:20]),
                        'razao_social_emitente': 'CONFIGURE CERTIFICADO A1',
                        'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                        'valor_total': 0.0,
                        'produtos': [],
                        'fonte_dados': 'Simulação (Certificado não configurado)',
                        'observacao': 'Execute sistema_principal.configurar_certificado_rapido() para configurar'
                    }
            except Exception as e:
                print(f"❌ Erro na consulta: {e}")
                return {
                    'chave_acesso': chave,
                    'numero_nf': 'ERRO',
                    'cnpj_emitente': '00.000.000/0000-00',
                    'razao_social_emitente': 'ERRO NA CONSULTA',
                    'data_emissao': datetime.now().strftime('%d/%m/%Y'),
                    'valor_total': 0.0,
                    'produtos': [],
                    'fonte_dados': 'Erro',
                    'observacao': str(e)
                }
        
        # Aplicar método corrigido
        processador.consultar_nfe_sefaz = consulta_corrigida
        
        # Método corrigido de configuração
        def configurar_certificado_corrigido(cert_path, cert_password):
            """Configuração corrigida de certificado"""
            return consultor_corrigido.configurar_certificado(cert_path, cert_password)
        
        # Método corrigido de teste
        def testar_certificado_corrigido():
            """Teste corrigido de certificado"""
            return consultor_corrigido.testar_conectividade()
        
        # Aplicar métodos corrigidos
        processador.configurar_certificado_a1 = configurar_certificado_corrigido
        processador.testar_certificado_a1 = testar_certificado_corrigido
        
        # Interface corrigida de configuração rápida
        def configuracao_rapida_corrigida():
            """Interface corrigida de configuração rápida"""
            try:
                print("🔑 Configuração de Certificado A1 (Versão Corrigida)")
                
                # Criar interface
                root = tk.Tk()
                root.title("Configuração Certificado A1 - Corrigida")
                root.geometry("600x500")
                root.configure(bg='#f0f0f0')
                
                # Frame principal
                main_frame = tk.Frame(root, bg='#f0f0f0', padx=20, pady=20)
                main_frame.pack(fill='both', expand=True)
                
                # Título
                title_label = tk.Label(
                    main_frame,
                    text="🔐 Configuração de Certificado A1",
                    font=('Arial', 16, 'bold'),
                    fg='#0056b3',
                    bg='#f0f0f0'
                )
                title_label.pack(pady=(0, 20))
                
                # Instruções
                instructions_frame = tk.LabelFrame(
                    main_frame, 
                    text="📋 Instruções Importantes",
                    font=('Arial', 10, 'bold'),
                    padx=10, 
                    pady=10
                )
                instructions_frame.pack(fill='x', pady=(0, 20))
                
                instructions_text = """
✅ CERTIFICADO DIGITAL A1 (.pfx ou .p12)
• Arquivo baixado de Autoridade Certificadora
• NÃO é token/cartão (esses são A3)
• Deve estar dentro da validade

🔑 SENHA/PIN:
• PIN de 6 dígitos criado no download
• OU primeiros 6 dígitos da senha de relacionamento
• OU senha completa de relacionamento

⚠️ PROBLEMAS COMUNS:
• "Senha incorreta": Tente PIN de 6 dígitos
• "Arquivo não encontrado": Verifique caminho
• "Certificado expirado": Renove na AC
                """
                
                tk.Label(
                    instructions_frame, 
                    text=instructions_text.strip(),
                    justify='left',
                    font=('Arial', 9),
                    wraplength=500
                ).pack(anchor='w')
                
                # Seleção de arquivo
                file_frame = tk.LabelFrame(
                    main_frame,
                    text="📁 Selecionar Certificado",
                    font=('Arial', 10, 'bold'),
                    padx=10,
                    pady=10
                )
                file_frame.pack(fill='x', pady=(0, 15))
                
                # Variável para caminho do arquivo
                file_path_var = tk.StringVar()
                
                # Entry e botão de seleção
                file_entry_frame = tk.Frame(file_frame)
                file_entry_frame.pack(fill='x', pady=5)
                
                file_entry = tk.Entry(
                    file_entry_frame, 
                    textvariable=file_path_var,
                    font=('Arial', 9),
                    width=50
                )
                file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10))
                
                def selecionar_arquivo():
                    arquivo = filedialog.askopenfilename(
                        title="Selecionar Certificado A1",
                        filetypes=[
                            ("Certificado A1", "*.pfx *.p12"),
                            ("Todos os arquivos", "*.*")
                        ]
                    )
                    if arquivo:
                        file_path_var.set(arquivo)
                        print(f"📁 Arquivo selecionado: {os.path.basename(arquivo)}")
                
                tk.Button(
                    file_entry_frame,
                    text="🔍 Procurar",
                    command=selecionar_arquivo,
                    font=('Arial', 9)
                ).pack(side='right')
                
                # Senha
                password_frame = tk.LabelFrame(
                    main_frame,
                    text="🔑 Senha do Certificado",
                    font=('Arial', 10, 'bold'),
                    padx=10,
                    pady=10
                )
                password_frame.pack(fill='x', pady=(0, 15))
                
                password_var = tk.StringVar()
                
                tk.Label(
                    password_frame,
                    text="Digite a senha/PIN do certificado:",
                    font=('Arial', 9)
                ).pack(anchor='w', pady=(0, 5))
                
                password_entry = tk.Entry(
                    password_frame,
                    textvariable=password_var,
                    show='*',
                    font=('Arial', 10),
                    width=30
                )
                password_entry.pack(anchor='w', pady=(0, 10))
                
                # Checkbox para mostrar senha
                show_password_var = tk.BooleanVar()
                
                def toggle_password():
                    if show_password_var.get():
                        password_entry.config(show='')
                    else:
                        password_entry.config(show='*')
                
                tk.Checkbutton(
                    password_frame,
                    text="Mostrar senha",
                    variable=show_password_var,
                    command=toggle_password,
                    font=('Arial', 9)
                ).pack(anchor='w')
                
                # Status
                status_frame = tk.Frame(main_frame)
                status_frame.pack(fill='x', pady=(0, 15))
                
                status_label = tk.Label(
                    status_frame,
                    text="Aguardando configuração...",
                    font=('Arial', 10),
                    fg='gray'
                )
                status_label.pack(anchor='w')
                
                # Função de configurar
                def configurar():
                    cert_path = file_path_var.get().strip()
                    cert_password = password_var.get()
                    
                    if not cert_path:
                        messagebox.showerror("Erro", "Selecione o arquivo do certificado!")
                        return
                    
                    if not cert_password:
                        messagebox.showerror("Erro", "Digite a senha do certificado!")
                        return
                    
                    # Atualizar status
                    status_label.config(text="🔄 Configurando certificado...", fg='blue')
                    root.update()
                    
                    try:
                        # Configurar certificado
                        sucesso, mensagem = consultor_corrigido.configurar_certificado(cert_path, cert_password)
                        
                        if sucesso:
                            status_label.config(text="✅ Certificado configurado!", fg='green')
                            
                            # Testar conectividade
                            status_label.config(text="🧪 Testando conectividade...", fg='blue')
                            root.update()
                            
                            teste_ok, teste_msg = consultor_corrigido.testar_conectividade()
                            
                            if teste_ok:
                                resultado = f"✅ SUCESSO!\n\n📋 {mensagem}\n🌐 {teste_msg}\n\n🎉 Sistema pronto para consultar NFe!"
                                messagebox.showinfo("Certificado Configurado", resultado)
                                
                                # Salvar configuração
                                salvar = messagebox.askyesno(
                                    "Salvar Configuração",
                                    "Deseja salvar o caminho do certificado?\n(A senha NÃO será salva por segurança)"
                                )
                                
                                if salvar:
                                    try:
                                        import json
                                        config_data = {
                                            'certificado_path': cert_path,
                                            'data_configuracao': datetime.now().isoformat(),
                                            'info_certificado': consultor_corrigido.obter_info_certificado()
                                        }
                                        
                                        with open('config_certificado_a1.json', 'w') as f:
                                            json.dump(config_data, f, indent=2, default=str)
                                        
                                        print("💾 Configuração salva em config_certificado_a1.json")
                                    except Exception as e:
                                        print(f"⚠️ Erro ao salvar configuração: {e}")
                                
                                root.destroy()
                            else:
                                status_label.config(text="⚠️ Configurado, mas com problemas de conectividade", fg='orange')
                                resultado = f"⚠️ CERTIFICADO CONFIGURADO\n\n📋 {mensagem}\n\n❌ Conectividade: {teste_msg}\n\nO certificado foi configurado, mas há problemas de conexão com a SEFAZ."
                                messagebox.showwarning("Aviso", resultado)
                        else:
                            status_label.config(text="❌ Falha na configuração", fg='red')
                            messagebox.showerror("Erro", f"❌ Falha na configuração:\n\n{mensagem}")
                    
                    except Exception as e:
                        status_label.config(text="❌ Erro durante configuração", fg='red')
                        messagebox.showerror("Erro", f"❌ Erro durante configuração:\n\n{str(e)}")
                
                # Função de diagnóstico
                def diagnosticar():
                    """Executa diagnóstico do sistema"""
                    try:
                        # Verificar certificado atual
                        cert_info = consultor_corrigido.obter_info_certificado()
                        
                        if cert_info.get('is_valid'):
                            # Teste completo
                            teste_ok, teste_msg = consultor_corrigido.testar_conectividade()
                            
                            info_text = f"""
📋 DIAGNÓSTICO DO CERTIFICADO A1

✅ Status: CONFIGURADO
📅 Válido até: {cert_info['not_valid_after'].strftime('%d/%m/%Y %H:%M')}
👤 Proprietário: {cert_info.get('subject_info', {}).get('CN', 'N/A')}
🔢 Serial: {cert_info.get('serial_number', 'N/A')}

🌐 Teste de Conectividade:
{teste_msg}

📁 Arquivos temporários: {'✅ OK' if consultor_corrigido.temp_dir and os.path.exists(consultor_corrigido.temp_dir) else '⚠️ Ausentes'}
🔐 Certificado em memória: {'✅ OK' if consultor_corrigido.cert_data else '❌ Não'}
                            """
                        else:
                            info_text = """
📋 DIAGNÓSTICO DO CERTIFICADO A1

❌ Status: NÃO CONFIGURADO

Para configurar:
1. Clique em "Configurar" abaixo
2. Selecione arquivo .pfx do certificado
3. Digite senha/PIN do certificado
4. Aguarde validação e teste
                            """
                        
                        messagebox.showinfo("Diagnóstico", info_text.strip())
                        
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro no diagnóstico:\n{str(e)}")
                
                # Botões
                button_frame = tk.Frame(main_frame)
                button_frame.pack(fill='x', pady=20)
                
                tk.Button(
                    button_frame,
                    text="🔧 Configurar",
                    command=configurar,
                    font=('Arial', 11, 'bold'),
                    bg='#0056b3',
                    fg='white',
                    padx=20,
                    pady=5
                ).pack(side='left', padx=(0, 10))
                
                tk.Button(
                    button_frame,
                    text="🔍 Diagnóstico",
                    command=diagnosticar,
                    font=('Arial', 10),
                    padx=15,
                    pady=5
                ).pack(side='left', padx=(0, 10))
                
                tk.Button(
                    button_frame,
                    text="❌ Cancelar",
                    command=root.destroy,
                    font=('Arial', 10),
                    padx=15,
                    pady=5
                ).pack(side='right')
                
                # Verificar se já tem certificado configurado
                cert_info = consultor_corrigido.obter_info_certificado()
                if cert_info.get('is_valid'):
                    status_label.config(
                        text=f"✅ Certificado atual válido até {cert_info['not_valid_after'].strftime('%d/%m/%Y')}",
                        fg='green'
                    )
                
                root.mainloop()
                
            except Exception as e:
                error_msg = f"❌ Erro na interface: {str(e)}"
                print(error_msg)
                messagebox.showerror("Erro", error_msg)
        
        # Aplicar configuração corrigida
        sistema_principal.configurar_certificado_rapido = configuracao_rapida_corrigida
        
        # Diagnóstico corrigido
        def diagnostico_corrigido():
            """Diagnóstico corrigido do sistema"""
            print("\n🔍 DIAGNÓSTICO DO SISTEMA NFe (VERSÃO CORRIGIDA)")
            print("=" * 55)
            
            # Sistema híbrido
            if hasattr(sistema_principal, 'processador_nfe'):
                print("✅ Sistema híbrido NFe: ATIVO")
                
                # Consultor corrigido
                if hasattr(sistema_principal, 'consultor_sefaz_a1'):
                    print("✅ Consultor SEFAZ A1: PRESENTE (versão corrigida)")
                    
                    # Info do certificado
                    cert_info = consultor_corrigido.obter_info_certificado()
                    if cert_info.get('is_valid'):
                        print(f"✅ Certificado: CONFIGURADO")
                        print(f"   📅 Válido até: {cert_info['not_valid_after'].strftime('%d/%m/%Y %H:%M')}")
                        print(f"   👤 Proprietário: {cert_info.get('subject_info', {}).get('CN', 'N/A')}")
                        
                        # Verificar arquivos
                        cert_ok, cert_msg = consultor_corrigido.verificar_certificado_configurado()
                        print(f"   📁 Arquivos: {cert_msg}")
                        
                        # Teste de conectividade
                        try:
                            teste_ok, teste_msg = consultor_corrigido.testar_conectividade()
                            print(f"   🌐 Conectividade: {teste_msg}")
                        except Exception as e:
                            print(f"   ❌ Erro no teste: {str(e)[:50]}...")
                    else:
                        print("⚠️ Certificado: NÃO CONFIGURADO")
                        print("   💡 Execute: sistema_principal.configurar_certificado_rapido()")
                
                # Dependências
                print(f"\n📦 Dependências:")
                print(f"   Cryptography: {'✅ OK' if CRYPTO_OK else '❌ AUSENTE'}")
                print(f"   Requests: ✅ OK")
                print(f"   Tkinter: ✅ OK")
                
                # URLs SEFAZ
                print(f"\n🌐 URLs SEFAZ configuradas: {len(consultor_corrigido.urls_sefaz)} estados")
                
            else:
                print("❌ Sistema híbrido NFe: INATIVO")
            
            print("=" * 55)
        
        # Aplicar diagnóstico corrigido
        sistema_principal.diagnosticar_nfe = diagnostico_corrigido
        
        print("✅ Sistema de certificado A1 corrigido com sucesso!")
        print("\n🔧 MELHORIAS APLICADAS:")
        print("   • ✅ Validação robusta de certificados")
        print("   • ✅ Múltiplos formatos de senha suportados")
        print("   • ✅ Gerenciamento seguro de arquivos temporários")
        print("   • ✅ URLs SEFAZ atualizadas para 2025")
        print("   • ✅ Tratamento melhorado de erros")
        print("   • ✅ Interface de configuração reformulada")
        print("   • ✅ Testes de conectividade aprimorados")
        print("   • ✅ Fallback automático para dados simulados")
        
        print("\n🎯 PRÓXIMOS PASSOS:")
        print("1. Configure o certificado:")
        print("   sistema_principal.configurar_certificado_rapido()")
        print("\n2. Para diagnóstico:")
        print("   sistema_principal.diagnosticar_nfe()")
        print("\n3. Para testar consulta:")
        print("   sistema_principal.processador_nfe.consultar_nfe_sefaz('chave_44_digitos')")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro na correção: {e}")
        import traceback
        traceback.print_exc()
        return False


def testar_certificado_manual(cert_path, cert_password, chave_teste=None):
    """Função para testar certificado manualmente"""
    try:
        print("\n🧪 TESTE MANUAL DE CERTIFICADO A1")
        print("=" * 40)
        
        if not chave_teste:
            chave_teste = "35210714200166000187550010000000271234567890"
        
        # Criar consultor
        consultor = ConsultorSefazA1Corrigido()
        
        # Configurar certificado
        print("🔑 Configurando certificado...")
        sucesso, msg = consultor.configurar_certificado(cert_path, cert_password)
        
        if not sucesso:
            print(f"❌ Falha na configuração: {msg}")
            return False
        
        print(f"✅ {msg}")
        
        # Testar conectividade
        print("\n🧪 Testando conectividade...")
        teste_ok, teste_msg = consultor.testar_conectividade()
        print(f"🌐 {teste_msg}")
        
        if teste_ok:
            # Testar consulta real
            print(f"\n🔍 Testando consulta com chave: {chave_teste}")
            resultado = consultor.consultar_nfe(chave_teste)
            
            if resultado:
                print("✅ Consulta realizada com sucesso!")
                print(f"   📋 NFe: {resultado.get('numero_nf', 'N/A')}")
                print(f"   🏢 Emitente: {resultado.get('razao_social_emitente', 'N/A')}")
                print(f"   💰 Valor: R$ {resultado.get('valor_total', 0):,.2f}")
                print(f"   📊 Status: {resultado.get('status_sefaz', 'N/A')}")
                return True
            else:
                print("❌ Consulta retornou resultado vazio")
                return False
        
        return teste_ok
        
    except Exception as e:
        print(f"❌ Erro no teste manual: {e}")
        return False
    finally:
        # Limpar recursos
        if 'consultor' in locals():
            consultor._limpar_arquivos_temp()


def diagnosticar_problema_certificado():
    """Diagnostica problemas comuns com certificados"""
    print("\n🔍 DIAGNÓSTICO DE PROBLEMAS COMUNS")
    print("=" * 45)
    
    # Verificar dependências
    print("📦 Verificando dependências...")
    
    try:
        import cryptography
        print(f"   ✅ Cryptography: {cryptography.__version__}")
    except ImportError:
        print("   ❌ Cryptography: NÃO INSTALADO")
        print("      💡 Solução: pip install cryptography")
    
    try:
        import requests
        print(f"   ✅ Requests: OK")
    except ImportError:
        print("   ❌ Requests: NÃO INSTALADO")
        print("      💡 Solução: pip install requests")
    
    # Verificar conectividade básica
    print("\n🌐 Verificando conectividade...")
    
    try:
        response = requests.get("https://www.google.com", timeout=5)
        print("   ✅ Internet: OK")
    except:
        print("   ❌ Internet: PROBLEMA")
        print("      💡 Verifique conexão com a internet")
    
    # Testar SEFAZ
    print("\n🏛️ Testando acesso à SEFAZ...")
    
    urls_teste = [
        "https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx",
        "https://nfe.sefaz.rj.gov.br/ws/nfeconsultaprotocolo4.asmx"
    ]
    
    for url in urls_teste:
        try:
            response = requests.get(url, timeout=10, verify=False)
            print(f"   ✅ {url.split('/')[2]}: Acessível")
        except Exception as e:
            print(f"   ❌ {url.split('/')[2]}: {str(e)[:30]}...")
    
    print("\n💡 DICAS PARA PROBLEMAS COMUNS:")
    print("=" * 45)
    print("🔑 SENHA INCORRETA:")
    print("   • Tente PIN de 6 dígitos do certificado")
    print("   • Tente primeiros 6 dígitos da senha de relacionamento")
    print("   • Tente senha completa de relacionamento")
    print("   • Verifique se não há espaços no início/fim")
    
    print("\n📁 ARQUIVO NÃO ENCONTRADO:")
    print("   • Verifique se arquivo .pfx existe")
    print("   • Verifique permissões de leitura")
    print("   • Tente copiar arquivo para área de trabalho")
    
    print("\n⏰ CERTIFICADO EXPIRADO:")
    print("   • Verifique data de validade")
    print("   • Renove na Autoridade Certificadora")
    print("   • Baixe novo certificado A1")
    
    print("\n🌐 PROBLEMAS DE CONECTIVIDADE:")
    print("   • Verifique firewall (porta 443)")
    print("   • Verifique proxy corporativo")
    print("   • Teste em horário comercial")
    print("   • Aguarde alguns minutos e tente novamente")


# Função para aplicar correção automaticamente
def aplicar_correcao_automatica(sistema_principal):
    """Aplica correção automaticamente no sistema"""
    try:
        print("\n🚀 APLICANDO CORREÇÃO AUTOMÁTICA")
        print("=" * 40)
        
        # Aplicar correção
        sucesso = corrigir_sistema_certificado_a1(sistema_principal)
        
        if sucesso:
            print("\n✅ CORREÇÃO APLICADA COM SUCESSO!")
            print("\n🎯 TESTE AGORA:")
            print("sistema_principal.configurar_certificado_rapido()")
            
            return True
        else:
            print("\n❌ FALHA NA CORREÇÃO")
            print("💡 Execute diagnóstico:")
            print("diagnosticar_problema_certificado()")
            
            return False
            
    except Exception as e:
        print(f"❌ Erro na aplicação: {e}")
        return False


if __name__ == "__main__":
    print("""
🔧 CORREÇÃO DE CERTIFICADO A1 - SISTEMA NFe

COMO USAR:

1. APLICAR CORREÇÃO NO SEU SISTEMA:
   from src.nfe.correcao_certificado_a1 import aplicar_correcao_automatica
   aplicar_correcao_automatica(sistema_principal)

2. CONFIGURAR CERTIFICADO:
   sistema_principal.configurar_certificado_rapido()

3. TESTAR MANUALMENTE:
   from src.nfe.correcao_certificado_a1 import testar_certificado_manual
   testar_certificado_manual("/caminho/cert.pfx", "123456")

4. DIAGNOSTICAR PROBLEMAS:
   from src.nfe.correcao_certificado_a1 import diagnosticar_problema_certificado
   diagnosticar_problema_certificado()

PRINCIPAIS CORREÇÕES:
✅ Validação robusta de certificados
✅ Múltiplos formatos de senha
✅ Gerenciamento seguro de arquivos
✅ URLs SEFAZ atualizadas
✅ Tratamento de erros melhorado
✅ Interface de configuração reformulada
    """)
