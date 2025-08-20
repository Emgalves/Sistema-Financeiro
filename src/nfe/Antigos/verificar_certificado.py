# -*- coding: utf-8 -*-
"""
Script para verificar o tipo de certificado digital
Execute para identificar se é A1 ou A3
"""

import os
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

def verificar_certificado():
    """Verifica tipo e características do certificado"""
    
    print("🔍 VERIFICADOR DE CERTIFICADO DIGITAL")
    print("=" * 45)
    
    # Verificar se tem arquivo .pfx/.p12
    print("1. PROCURANDO ARQUIVOS .pfx/.p12...")
    
    # Diretórios comuns onde certificados ficam
    diretorios_comuns = [
        Path.home() / "Downloads",
        Path.home() / "Documents",
        Path.home() / "Desktop",
        Path("C:/") if os.name == 'nt' else Path("/"),
    ]
    
    arquivos_encontrados = []
    for diretorio in diretorios_comuns:
        if diretorio.exists():
            for arquivo in diretorio.rglob("*.pfx"):
                arquivos_encontrados.append(arquivo)
            for arquivo in diretorio.rglob("*.p12"):
                arquivos_encontrados.append(arquivo)
    
    if arquivos_encontrados:
        print("✅ Arquivos de certificado A1 encontrados:")
        for i, arquivo in enumerate(arquivos_encontrados[:5], 1):
            print(f"   {i}. {arquivo}")
        if len(arquivos_encontrados) > 5:
            print(f"   ... e mais {len(arquivos_encontrados) - 5} arquivos")
    else:
        print("⚠️ Nenhum arquivo .pfx/.p12 encontrado nos diretórios comuns")
    
    print("\n2. VERIFICANDO TOKENS/SMARTCARDS (A3)...")
    
    # Verificar se há drivers de token instalados
    drivers_a3 = [
        "C:/Windows/System32/drivers/ukbfltr.sys",  # SafeNet
        "C:/Windows/System32/eTPKCS11.dll",        # eToken
        "C:/Windows/System32/ngp11v211.dll",       # Gemalto
    ]
    
    tokens_detectados = []
    for driver in drivers_a3:
        if os.path.exists(driver):
            tokens_detectados.append(driver)
    
    if tokens_detectados:
        print("⚠️ Drivers de token A3 detectados:")
        for driver in tokens_detectados:
            print(f"   • {driver}")
        print("   Isso indica que você pode ter certificado A3 (token/cartão)")
    else:
        print("ℹ️ Nenhum driver de token A3 comum detectado")
    
    print("\n3. IDENTIFICAÇÃO DO SEU CERTIFICADO:")
    print("-" * 35)
    
    if arquivos_encontrados:
        print("✅ VOCÊ TEM CERTIFICADO A1:")
        print("   • Tipo: Arquivo (.pfx ou .p12)")
        print("   • Senha: A senha definida quando o arquivo foi criado")
        print("   • Compatível: SIM com nosso sistema")
        print("   • Como usar: Selecione o arquivo .pfx/.p12 e digite a SENHA (não PIN)")
        
        # Oferecer teste
        try:
            root = tk.Tk()
            root.withdraw()
            
            testar = messagebox.askyesno(
                "Testar Certificado A1",
                "Deseja testar um dos arquivos .pfx/.p12 encontrados?"
            )
            
            if testar:
                arquivo = filedialog.askopenfilename(
                    title="Selecionar Certificado A1",
                    filetypes=[("Certificado A1", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                if arquivo:
                    testar_arquivo_certificado(arquivo)
            
            root.destroy()
        except:
            print("💡 Para testar: use sistema_principal.configurar_certificado_rapido()")
    
    if tokens_detectados:
        print("\n⚠️ VOCÊ PODE TER CERTIFICADO A3:")
        print("   • Tipo: Token USB ou Cartão Smart")
        print("   • Acesso: PIN para desbloquear hardware")
        print("   • Compatível: NÃO com nosso sistema atual")
        print("   • Solução: Precisa exportar para A1 ou usar sistema diferente")
    
    if not arquivos_encontrados and not tokens_detectados:
        print("\n❓ CERTIFICADO NÃO IDENTIFICADO:")
        print("   • Verifique se tem certificado digital instalado")
        print("   • Consulte sua Autoridade Certificadora")
        print("   • Ou baixe o certificado como arquivo A1")
    
    print("\n" + "=" * 45)
    
    return len(arquivos_encontrados) > 0


def testar_arquivo_certificado(caminho_arquivo):
    """Testa um arquivo de certificado"""
    try:
        print(f"\n🧪 TESTANDO ARQUIVO: {caminho_arquivo}")
        print("-" * 50)
        
        # Verificar se arquivo existe
        if not os.path.exists(caminho_arquivo):
            print("❌ Arquivo não encontrado")
            return False
        
        # Verificar tamanho
        tamanho = os.path.getsize(caminho_arquivo)
        print(f"📁 Tamanho: {tamanho:,} bytes")
        
        if tamanho < 1000:
            print("⚠️ Arquivo muito pequeno - pode estar corrompido")
        elif tamanho > 50000:
            print("⚠️ Arquivo muito grande - pode não ser certificado")
        else:
            print("✅ Tamanho adequado")
        
        # Tentar carregar com cryptography
        try:
            from cryptography.hazmat.primitives.serialization import pkcs12
            
            with open(caminho_arquivo, 'rb') as f:
                cert_data = f.read()
            
            print("📦 Arquivo lido com sucesso")
            
            # Testar sem senha (alguns certificados não têm senha)
            try:
                private_key, certificate, additional = pkcs12.load_key_and_certificates(cert_data, b'')
                if certificate:
                    print("✅ Certificado SEM SENHA detectado!")
                    mostrar_info_certificado(certificate)
                    return True
            except:
                print("🔐 Certificado protegido por senha")
            
            # Solicitar senha
            from tkinter import simpledialog
            import tkinter as tk
            
            root = tk.Tk()
            root.withdraw()
            
            senha = simpledialog.askstring(
                "Senha do Certificado",
                f"Digite a SENHA (não PIN) do certificado:\n{os.path.basename(caminho_arquivo)}",
                show='*'
            )
            
            root.destroy()
            
            if not senha:
                print("❌ Senha não informada")
                return False
            
            # Testar com senha
            try:
                private_key, certificate, additional = pkcs12.load_key_and_certificates(
                    cert_data, senha.encode('utf-8')
                )
                
                if certificate:
                    print("✅ Certificado carregado com sucesso!")
                    mostrar_info_certificado(certificate)
                    return True
                else:
                    print("❌ Certificado não encontrado no arquivo")
                    return False
                    
            except Exception as e:
                error_msg = str(e).lower()
                if "invalid" in error_msg or "wrong" in error_msg or "incorrect" in error_msg:
                    print("❌ SENHA INCORRETA")
                    print("💡 Dicas:")
                    print("   • Não use o PIN do token")
                    print("   • Use a senha definida quando o arquivo foi criado")
                    print("   • Verifique se não há caps lock ativado")
                    print("   • Consulte quem criou/instalou o certificado")
                else:
                    print(f"❌ Erro ao carregar: {e}")
                return False
                
        except ImportError:
            print("❌ Biblioteca 'cryptography' não instalada")
            print("💡 Execute: pip install cryptography")
            return False
        
    except Exception as e:
        print(f"❌ Erro no teste: {e}")
        return False


def mostrar_info_certificado(certificate):
    """Mostra informações do certificado"""
    try:
        from datetime import datetime
        
        print("\n📋 INFORMAÇÕES DO CERTIFICADO:")
        print("-" * 30)
        
        # Subject (titular)
        subject = certificate.subject.rfc4514_string()
        print(f"👤 Titular: {subject}")
        
        # Emissor
        issuer = certificate.issuer.rfc4514_string()
        print(f"🏛️ Emissor: {issuer}")
        
        # Validade
        inicio = certificate.not_valid_before
        fim = certificate.not_valid_after
        agora = datetime.now()
        
        print(f"📅 Válido de: {inicio.strftime('%d/%m/%Y')}")
        print(f"📅 Válido até: {fim.strftime('%d/%m/%Y')}")
        
        if fim < agora:
            print("❌ CERTIFICADO EXPIRADO!")
        elif inicio > agora:
            print("⚠️ Certificado ainda não é válido")
        else:
            dias_restantes = (fim - agora).days
            print(f"✅ Certificado válido ({dias_restantes} dias restantes)")
        
        # Serial
        print(f"🔢 Serial: {certificate.serial_number}")
        
        print("-" * 30)
        
    except Exception as e:
        print(f"⚠️ Erro ao extrair informações: {e}")


# Função para diagnosticar problemas comuns
def diagnosticar_problemas_certificado():
    """Diagnostica problemas comuns com certificado"""
    
    print("\n🔧 DIAGNÓSTICO DE PROBLEMAS COMUNS")
    print("=" * 40)
    
    problemas_comuns = [
        {
            'sintoma': 'Erro "senha incorreta" mesmo com senha certa',
            'causas': [
                'Usando PIN do token em vez da senha do arquivo',
                'Certificado A3 em vez de A1',
                'Arquivo corrompido',
                'Caps Lock ativado'
            ],
            'solucoes': [
                'Verificar se é arquivo .pfx/.p12 (A1)',
                'Solicitar senha do arquivo para quem instalou',
                'Baixar novo arquivo A1 da AC',
                'Verificar se não é token/cartão (A3)'
            ]
        },
        {
            'sintoma': 'Erro "arquivo não encontrado" ou "invalid path"',
            'causas': [
                'Certificado A3 (token) sendo usado como A1',
                'Arquivo temporário perdido',
                'Caminho com caracteres especiais'
            ],
            'solucoes': [
                'Verificar se tem arquivo .pfx/.p12',
                'Salvar arquivo em local simples (sem acentos)',
                'Reconfigurar certificado'
            ]
        },
        {
            'sintoma': 'Sistema não encontra certificado',
            'causas': [
                'Certificado apenas no navegador',
                'Certificado A3 não exportado',
                'Arquivo em local protegido'
            ],
            'solucoes': [
                'Exportar certificado do navegador para arquivo',
                'Solicitar arquivo A1 da Autoridade Certificadora',
                'Mover arquivo para pasta acessível'
            ]
        }
    ]
    
    for i, problema in enumerate(problemas_comuns, 1):
        print(f"\n{i}. PROBLEMA: {problema['sintoma']}")
        print("   CAUSAS POSSÍVEIS:")
        for causa in problema['causas']:
            print(f"   • {causa}")
        print("   SOLUÇÕES:")
        for solucao in problema['solucoes']:
            print(f"   ✓ {solucao}")
    
    print("\n" + "=" * 40)


if __name__ == "__main__":
    print("🔍 VERIFICADOR DE CERTIFICADO DIGITAL")
    print("Este script ajuda a identificar problemas com certificados")
    print()
    
    # Executar verificação
    tem_a1 = verificar_certificado()
    
    # Mostrar diagnóstico se não tem A1
    if not tem_a1:
        diagnosticar_problemas_certificado()
    
    print("\n💡 RESUMO:")
    if tem_a1:
        print("✅ Você tem certificado A1 - use a SENHA do arquivo")
    else:
        print("⚠️ Certificado A1 não encontrado - verifique o tipo")
    
    print("\n🎯 Para usar no sistema:")
    print("   sistema_principal.configurar_certificado_rapido()")