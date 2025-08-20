# -*- coding: utf-8 -*-
"""
SCRIPT DE TESTE PARA CERTIFICADO A1 - DIAGNÓSTICO COMPLETO
Execute este script para identificar e corrigir problemas de autenticação
"""

import os
import tempfile
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, simpledialog, messagebox

try:
    from cryptography.hazmat.primitives import serialization
    from cryptography.hazmat.primitives.serialization import pkcs12
    CRYPTO_OK = True
    print("✅ Cryptography disponível")
except ImportError:
    CRYPTO_OK = False
    print("❌ Instale cryptography: pip install cryptography")

try:
    import requests
    print("✅ Requests disponível")
except ImportError:
    print("❌ Instale requests: pip install requests")


def testar_certificado_step_by_step():
    """Teste passo a passo do certificado A1"""
    
    print("\n" + "="*60)
    print("🔍 TESTE COMPLETO DE CERTIFICADO A1 PARA SEFAZ")
    print("="*60)
    
    if not CRYPTO_OK:
        print("❌ Biblioteca cryptography não encontrada!")
        print("💡 Execute: pip install cryptography")
        return False
    
    # Passo 1: Selecionar certificado
    print("\n📁 PASSO 1: Selecionando certificado...")
    
    root = tk.Tk()
    root.withdraw()
    
    cert_path = filedialog.askopenfilename(
        title="Selecione seu certificado A1",
        filetypes=[
            ("Certificado PKCS#12", "*.pfx *.p12"),
            ("Todos os arquivos", "*.*")
        ]
    )
    
    if not cert_path:
        print("❌ Nenhum certificado selecionado")
        root.destroy()
        return False
    
    print(f"✅ Arquivo selecionado: {os.path.basename(cert_path)}")
    print(f"📏 Tamanho: {os.path.getsize(cert_path):,} bytes")
    
    # Passo 2: Verificar formato
    print("\n🔍 PASSO 2: Verificando formato do arquivo...")
    
    try:
        with open(cert_path, 'rb') as f:
            header = f.read(50)
        
        # Verificar assinatura PKCS#12
        if header.startswith(b'\x30'):
            print("✅ Arquivo tem assinatura PKCS#12 válida")
        else:
            print("❌ Arquivo não parece ser um certificado PKCS#12")
            print(f"Header: {header[:20].hex()}")
            return False
    except Exception as e:
        print(f"❌ Erro ao ler arquivo: {e}")
        return False
    
    # Passo 3: Obter senha
    print("\n🔒 PASSO 3: Obtendo senha do certificado...")
    
    cert_password = simpledialog.askstring(
        "Senha do Certificado",
        "Digite a senha do certificado A1:\n\n"
        "• Senha definida quando baixou o certificado\n"
        "• Pode ser numérica ou alfanumérica\n"
        "• Deixe em branco se não tem senha\n"
        "• Diferente da senha de relacionamento",
        show='*'
    )
    
    root.destroy()
    
    if cert_password is None:
        print("❌ Teste cancelado pelo usuário")
        return False
    
    if cert_password == "":
        print("ℹ️ Tentando carregar sem senha")
        cert_password = None
    else:
        print(f"🔑 Senha fornecida: {'*' * len(cert_password)}")
    
    # Passo 4: Tentar carregar certificado
    print("\n🔓 PASSO 4: Carregando certificado...")
    
    try:
        with open(cert_path, 'rb') as f:
            cert_data = f.read()
        
        print("✅ Arquivo lido com sucesso")
        
        # Tentar diferentes formatos de senha
        tentativas = []
        
        if cert_password is None:
            tentativas = [None]
        else:
            tentativas = [
                cert_password.encode('utf-8'),
                cert_password.encode('latin-1'),
                cert_password.encode('cp1252'),
                str(cert_password).encode('utf-8'),
                cert_password,  # String direta
                None  # Sem senha
            ]
        
        private_key = None
        certificate = None
        senha_usada = None
        
        for i, senha_teste in enumerate(tentativas):
            try:
                print(f"🧪 Tentativa {i+1}: ", end="")
                
                if senha_teste is None:
                    print("sem senha")
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, None)
                    senha_usada = "sem senha"
                else:
                    formato = type(senha_teste).__name__
                    print(f"senha como {formato}")
                    private_key, certificate, _ = pkcs12.load_key_and_certificates(cert_data, senha_teste)
                    senha_usada = f"senha formato {formato}"
                
                if certificate and private_key:
                    print(f"✅ SUCESSO com {senha_usada}!")
                    break
                    
            except Exception as e:
                print(f"❌ Falhou: {str(e)[:50]}")
                continue
        
        if not certificate or not private_key:
            print("\n❌ FALHA: Não foi possível carregar o certificado")
            print("💡 Possíveis causas:")
            print("   • Senha incorreta")
            print("   • Arquivo corrompido")
            print("   • Formato não suportado")
            print("   • Certificado protegido por hardware")
            return False
        
        print(f"\n✅ SUCESSO: Certificado carregado com {senha_usada}")
        
    except Exception as e:
        print(f"❌ Erro ao carregar: {e}")
        return False
    
    # Passo 5: Verificar validade
    print("\n📅 PASSO 5: Verificando validade...")
    
    now = datetime.now()
    
    print(f"📅 Data atual: {now.strftime('%d/%m/%Y %H:%M')}")
    print(f"🟢 Válido desde: {certificate.not_valid_before.strftime('%d/%m/%Y %H:%M')}")
    print(f"🔴 Válido até: {certificate.not_valid_after.strftime('%d/%m/%Y %H:%M')}")
    
    if certificate.not_valid_after < now:
        print("❌ CERTIFICADO EXPIRADO!")
        print("💡 Renove o certificado junto à Autoridade Certificadora")
        return False
    
    if certificate.not_valid_before > now:
        print("❌ Certificado ainda não é válido!")
        return False
    
    dias_restantes = (certificate.not_valid_after - now).days
    print(f"✅ Certificado válido por mais {dias_restantes} dias")
    
    if dias_restantes < 30:
        print("⚠️ ATENÇÃO: Certificado vence em menos de 30 dias!")
    
    # Passo 6: Informações do certificado
    print("\n📋 PASSO 6: Informações do certificado...")
    
    try:
        subject = certificate.subject.rfc4514_string()
        issuer = certificate.issuer.rfc4514_string()
        serial = str(certificate.serial_number)
        
        print(f"👤 Titular: {subject}")
        print(f"🏢 Emissor: {issuer}")
        print(f"🆔 Serial: {serial}")
        
        # Extrair informações importantes
        import re
        
        # Tentar extrair CNPJ/CPF
        cnpj_match = re.search(r'(\d{14})', subject)
        cpf_match = re.search(r'(\d{11})', subject)
        
        if cnpj_match:
            cnpj = cnpj_match.group(1)
            print(f"🏭 CNPJ: {cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}")
        elif cpf_match:
            cpf = cpf_match.group(1)
            print(f"👨 CPF: {cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}")
        
    except Exception as e:
        print(f"⚠️ Erro ao extrair informações: {e}")
    
    # Passo 7: Criar arquivos temporários
    print("\n💾 PASSO 7: Criando arquivos temporários...")
    
    try:
        temp_dir = tempfile.mkdtemp(prefix="teste_cert_")
        print(f"📁 Diretório temporário: {temp_dir}")
        
        # Converter para PEM
        cert_pem = certificate.public_bytes(serialization.Encoding.PEM)
        key_pem = private_key.private_bytes(
            encoding=serialization.Encoding.PEM,
            format=serialization.PrivateFormat.PKCS8,
            encryption_algorithm=serialization.NoEncryption()
        )
        
        # Salvar arquivos
        cert_file = os.path.join(temp_dir, "cert.pem")
        key_file = os.path.join(temp_dir, "key.pem")
        
        with open(cert_file, 'wb') as f:
            f.write(cert_pem)
        with open(key_file, 'wb') as f:
            f.write(key_pem)
        
        # Definir permissões
        os.chmod(cert_file, 0o600)
        os.chmod(key_file, 0o600)
        
        print("✅ Arquivos PEM criados com sucesso")
        print(f"📄 Certificado: {cert_file}")
        print(f"🔑 Chave: {key_file}")
        
    except Exception as e:
        print(f"❌ Erro ao criar arquivos: {e}")
        return False
    
    # Passo 8: Testar conexão HTTPS
    print("\n🌐 PASSO 8: Testando conexão com SEFAZ...")
    
    try:
        import urllib3
        urllib3.disable_warnings()
        
        # URL de teste da SEFAZ-SP
        url_teste = "https://nfe.fazenda.sp.gov.br/ws/nfeconsultaprotocolo4.asmx"
        
        # Envelope SOAP básico
        envelope = """<?xml version="1.0" encoding="utf-8"?>
<soap:Envelope xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">
    <soap:Body>
        <nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4">
            <consSitNFe versao="4.00" xmlns="http://www.portalfiscal.inf.br/nfe">
                <tpAmb>1</tpAmb>
                <xServ>CONSULTAR</xServ>
                <chNFe>35200314200166000187550010000000271234567890</chNFe>
            </consSitNFe>
        </nfeDadosMsg>
    </soap:Body>
</soap:Envelope>"""
        
        headers = {
            'Content-Type': 'text/xml; charset=utf-8',
            'SOAPAction': 'http://www.portalfiscal.inf.br/nfe/wsdl/NFeConsultaProtocolo4/nfeConsultaNF'
        }
        
        print("📡 Enviando requisição de teste...")
        
        response = requests.post(
            url_teste,
            data=envelope,
            headers=headers,
            cert=(cert_file, key_file),
            verify=False,
            timeout=30
        )
        
        print(f"📊 Status HTTP: {response.status_code}")
        
        if response.status_code == 200:
            print("✅ CONEXÃO SEFAZ: SUCESSO!")
            
            # Analisar resposta
            import xml.etree.ElementTree as ET
            
            try:
                root = ET.fromstring(response.text)
                
                # Buscar código de status
                for elem in root.iter():
                    if elem.tag.endswith('cStat'):
                        codigo = elem.text
                        break
                else:
                    codigo = "?"
                
                # Buscar motivo
                for elem in root.iter():
                    if elem.tag.endswith('xMotivo'):
                        motivo = elem.text
                        break
                else:
                    motivo = "?"
                
                print(f"📋 Resposta SEFAZ: {codigo} - {motivo}")
                
                if codigo in ["217", "999"]:
                    print("✅ Resposta esperada (NFe teste não existe)")
                elif codigo == "100":
                    print("✅ NFe encontrada (inesperado mas OK)")
                else:
                    print(f"ℹ️ Resposta: {codigo} - {motivo}")
                
            except Exception as e:
                print(f"⚠️ Erro ao analisar resposta XML: {e}")
                print(f"📄 Primeiros 500 chars: {response.text[:500]}")
            
        else:
            print(f"❌ ERRO HTTP: {response.status_code}")
            print(f"📄 Resposta: {response.text[:500]}")
            
            if response.status_code == 403:
                print("💡 Erro 403: Possível problema de certificado ou permissão")
            elif response.status_code == 500:
                print("💡 Erro 500: Problema no servidor SEFAZ ou dados inválidos")
            
    except requests.exceptions.Timeout:
        print("❌ TIMEOUT: Conexão demorou muito")
        print("💡 Verifique sua conexão com a internet")
    except requests.exceptions.SSLError as e:
        print(f"❌ ERRO SSL: {e}")
        print("💡 Problema com certificado ou configuração SSL")
    except Exception as e:
        print(f"❌ ERRO: {e}")
    
    # Passo 9: Limpeza
    print("\n🧹 PASSO 9: Limpeza...")
    
    try:
        import shutil
        shutil.rmtree(temp_dir, ignore_errors=True)
        print("✅ Arquivos temporários removidos")
    except Exception as e:
        print(f"⚠️ Erro na limpeza: {e}")
    
    # Resumo final
    print("\n" + "="*60)
    print("📊 RESUMO DO TESTE")
    print("="*60)
    
    print("✅ Certificado carregado com sucesso")
    print("✅ Certificado válido")
    print("✅ Arquivos temporários criados")
    
    if 'response' in locals() and response.status_code == 200:
        print("✅ Conexão SEFAZ funcionando")
        print("\n🎉 CERTIFICADO ESTÁ PRONTO PARA USO!")
    else:
        print("❌ Problema na conexão SEFAZ")
        print("\n⚠️ CERTIFICADO OK, MAS VERIFIQUE CONECTIVIDADE")
    
    print("\n💡 Próximos passos:")
    print("1. Configure o certificado no seu sistema")
    print("2. Teste com chave real de NFe")
    print("3. Integre com sistema financeiro/materiais")
    
    return True


def aplicar_correcoes_rapidas(sistema_principal):
    """Aplica correções rápidas ao sistema existente"""
    try:
        print("\n🔧 Aplicando correções rápidas...")
        
        # Importar consultor corrigido
        from consulta_sefaz_certificado_corrigido import ConsultorSefazA1Corrigido, aplicar_correcoes_ao_sistema
        
        # Aplicar correções
        sucesso = aplicar_correcoes_ao_sistema(sistema_principal)
        
        if sucesso:
            print("✅ Correções aplicadas!")
            print("💡 Use: sistema_principal.configurar_certificado_rapido()")
            return True
        else:
            print("❌ Falha ao aplicar correções")
            return False
            
    except ImportError:
        print("❌ Arquivo de correções não encontrado")
        print("💡 Certifique-se de ter o arquivo consulta_sefaz_certificado_corrigido.py")
        return False
    except Exception as e:
        print(f"❌ Erro: {e}")
        return False


if __name__ == "__main__":
    print("🚀 SISTEMA DE TESTE PARA CERTIFICADO A1")
    print("Este script irá diagnosticar problemas com seu certificado")
    
    try:
        testar_certificado_step_by_step()
    except KeyboardInterrupt:
        print("\n❌ Teste interrompido pelo usuário")
    except Exception as e:
        print(f"\n❌ Erro inesperado: {e}")
        import traceback
        traceback.print_exc()
    
    input("\nPressione Enter para sair...")
