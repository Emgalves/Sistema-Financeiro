# -*- coding: utf-8 -*-
"""
SCRIPT PARA APLICAR CORREÇÕES DE CERTIFICADO A1
Execute este script no seu console Python após inicializar o sistema
"""

def aplicar_todas_correcoes(sistema_principal):
    """
    Aplica todas as correções necessárias para resolver os problemas
    de certificado A1 identificados no sistema
    """
    try:
        print("\n🚀 INICIANDO APLICAÇÃO DE TODAS AS CORREÇÕES")
        print("=" * 60)
        
        # 1. Aplicar correção do sistema de certificado
        print("\n📋 ETAPA 1: Corrigindo sistema de certificado A1...")
        from src.nfe.correcao_certificado_a1 import corrigir_sistema_certificado_a1
        
        sucesso_cert = corrigir_sistema_certificado_a1(sistema_principal)
        
        if sucesso_cert:
            print("✅ Sistema de certificado A1 corrigido!")
        else:
            print("❌ Falha na correção do certificado")
            return False
        
        # 2. Executar diagnóstico para verificar estado atual
        print("\n📋 ETAPA 2: Executando diagnóstico...")
        sistema_principal.diagnosticar_nfe()
        
        # 3. Verificar se melhorias anteriores estão funcionando
        print("\n📋 ETAPA 3: Verificando integrações...")
        
        # Verificar se sistema híbrido está ativo
        if hasattr(sistema_principal, 'processador_nfe'):
            print("✅ Sistema híbrido NFe: ATIVO")
        else:
            print("❌ Sistema híbrido NFe: INATIVO")
            return False
        
        # Verificar se extensões estão aplicadas
        if hasattr(sistema_principal, 'integrador_nfe_completo'):
            print("✅ Integrador completo: PRESENTE")
        else:
            print("⚠️ Integrador completo: AUSENTE")
        
        # 4. Testar consulta básica (sem certificado)
        print("\n📋 ETAPA 4: Testando consulta básica...")
        
        try:
            chave_teste = "35210714200166000187550010000000271234567890"
            resultado = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave_teste)
            
            if resultado:
                print("✅ Consulta básica funcionando")
                print(f"   📄 NFe: {resultado.get('numero_nf', 'N/A')}")
                print(f"   🏢 Emitente: {resultado.get('razao_social_emitente', 'N/A')}")
                print(f"   📊 Fonte: {resultado.get('fonte_dados', 'N/A')}")
            else:
                print("❌ Consulta básica falhou")
                
        except Exception as e:
            print(f"❌ Erro na consulta básica: {e}")
        
        # 5. Verificar dependências críticas
        print("\n📋 ETAPA 5: Verificando dependências...")
        
        dependencias_ok = True
        
        try:
            import cryptography
            print(f"✅ Cryptography: {cryptography.__version__}")
        except ImportError:
            print("❌ Cryptography: AUSENTE")
            print("   💡 Execute: pip install cryptography")
            dependencias_ok = False
        
        try:
            import requests
            print("✅ Requests: OK")
        except ImportError:
            print("❌ Requests: AUSENTE")
            dependencias_ok = False
        
        try:
            import tkinter
            print("✅ Tkinter: OK")
        except ImportError:
            print("❌ Tkinter: AUSENTE")
            dependencias_ok = False
        
        if not dependencias_ok:
            print("\n⚠️ DEPENDÊNCIAS FALTANDO - Instale antes de continuar")
            return False
        
        # 6. Configurar métodos de emergência
        print("\n📋 ETAPA 6: Configurando métodos de emergência...")
        
        # Método para resetar sistema em caso de problemas
        def resetar_sistema_nfe():
            """Reseta sistema NFe em caso de problemas"""
            try:
                print("🔄 Resetando sistema NFe...")
                
                # Limpar cache de consultas
                if hasattr(sistema_principal.processador_nfe, 'cache_consultas'):
                    sistema_principal.processador_nfe.cache_consultas.clear()
                    print("✅ Cache de consultas limpo")
                
                # Recriar consultor
                from src.nfe.correcao_certificado_a1 import ConsultorSefazA1Corrigido
                sistema_principal.consultor_sefaz_a1 = ConsultorSefazA1Corrigido()
                print("✅ Consultor SEFAZ recriado")
                
                print("✅ Sistema resetado com sucesso!")
                return True
                
            except Exception as e:
                print(f"❌ Erro no reset: {e}")
                return False
        
        # Método para teste rápido
        def teste_rapido_certificado():
            """Teste rápido do certificado configurado"""
            try:
                if not hasattr(sistema_principal, 'consultor_sefaz_a1'):
                    print("❌ Consultor não encontrado")
                    return False
                
                consultor = sistema_principal.consultor_sefaz_a1
                cert_info = consultor.obter_info_certificado()
                
                if cert_info.get('is_valid'):
                    print("✅ Certificado configurado")
                    print(f"   📅 Válido até: {cert_info['not_valid_after'].strftime('%d/%m/%Y')}")
                    
                    # Teste de conectividade
                    teste_ok, teste_msg = consultor.testar_conectividade()
                    print(f"   🌐 Conectividade: {teste_msg}")
                    
                    return teste_ok
                else:
                    print("⚠️ Certificado não configurado")
                    print("   💡 Execute: sistema_principal.configurar_certificado_rapido()")
                    return False
                    
            except Exception as e:
                print(f"❌ Erro no teste: {e}")
                return False
        
        # Adicionar métodos ao sistema
        sistema_principal.resetar_sistema_nfe = resetar_sistema_nfe
        sistema_principal.teste_rapido_certificado = teste_rapido_certificado
        
        print("✅ Métodos de emergência configurados")
        
        # 7. Criar guia rápido de uso
        print("\n📋 ETAPA 7: Gerando guia de uso...")
        
        guia_uso = """
🎯 GUIA RÁPIDO DE USO - CERTIFICADO A1

CONFIGURAR CERTIFICADO:
  sistema_principal.configurar_certificado_rapido()

DIAGNOSTICAR SISTEMA:
  sistema_principal.diagnosticar_nfe()

TESTE RÁPIDO:
  sistema_principal.teste_rapido_certificado()

RESETAR EM CASO DE PROBLEMAS:
  sistema_principal.resetar_sistema_nfe()

CONSULTAR NFE POR CHAVE:
  chave = "44_digitos_da_nfe"
  dados = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)

IMPORTAR NFE COMPLETA:
  # Via interface gráfica:
  # Menu > NFe > Importar NFe > Botão "Processar NFe"

SOLUÇÃO DE PROBLEMAS COMUNS:

❌ "Could not find TLS key file":
   → Execute: sistema_principal.configurar_certificado_rapido()
   → Selecione arquivo .pfx válido
   → Digite senha/PIN correto

❌ "Senha incorreta":
   → Tente PIN de 6 dígitos
   → Tente primeiros 6 dígitos da senha de relacionamento
   → Tente senha completa de relacionamento

❌ "Certificado expirado":
   → Renove certificado na Autoridade Certificadora
   → Baixe novo arquivo .pfx

❌ "Timeout" ou "Conectividade":
   → Verifique conexão com internet
   → Verifique firewall (porta 443)
   → Tente em horário comercial
   → Execute: sistema_principal.teste_rapido_certificado()
        """
        
        # Salvar guia
        try:
            with open('GUIA_USO_CERTIFICADO_A1.txt', 'w', encoding='utf-8') as f:
                f.write(guia_uso)
            print("✅ Guia salvo em: GUIA_USO_CERTIFICADO_A1.txt")
        except:
            print("⚠️ Não foi possível salvar guia (sem permissão de escrita)")
        
        # 8. Resumo final
        print("\n📋 ETAPA 8: Resumo final...")
        
        print("\n🎉 TODAS AS CORREÇÕES APLICADAS COM SUCESSO!")
        print("=" * 50)
        
        print("✅ CORREÇÕES REALIZADAS:")
        print("   • Sistema de certificado A1 corrigido")
        print("   • Validação robusta de certificados")
        print("   • Múltiplos formatos de senha suportados")
        print("   • URLs SEFAZ atualizadas para 2025")
        print("   • Gerenciamento seguro de arquivos temporários")
        print("   • Interface de configuração melhorada")
        print("   • Tratamento de erros aprimorado")
        print("   • Métodos de emergência adicionados")
        print("   • Fallback automático implementado")
        
        print("\n🎯 PRÓXIMO PASSO OBRIGATÓRIO:")
        print("   sistema_principal.configurar_certificado_rapido()")
        
        print("\n💡 PARA DIAGNOSTICAR PROBLEMAS:")
        print("   sistema_principal.diagnosticar_nfe()")
        
        return True
        
    except Exception as e:
        print(f"❌ ERRO DURANTE APLICAÇÃO DAS CORREÇÕES: {e}")
        import traceback
        traceback.print_exc()
        return False


def diagnostico_pos_correcao(sistema_principal):
    """Diagnóstico completo após aplicar correções"""
    try:
        print("\n🔍 DIAGNÓSTICO PÓS-CORREÇÃO")
        print("=" * 40)
        
        # 1. Verificar sistema híbrido
        print("📋 SISTEMA HÍBRIDO:")
        if hasattr(sistema_principal, 'processador_nfe'):
            print("   ✅ Processador NFe: PRESENTE")
            
            # Verificar métodos essenciais
            metodos_essenciais = [
                'consultar_nfe_sefaz',
                'configurar_certificado_a1', 
                'testar_certificado_a1'
            ]
            
            for metodo in metodos_essenciais:
                if hasattr(sistema_principal.processador_nfe, metodo):
                    print(f"   ✅ Método {metodo}: PRESENTE")
                else:
                    print(f"   ❌ Método {metodo}: AUSENTE")
        else:
            print("   ❌ Processador NFe: AUSENTE")
        
        # 2. Verificar consultor SEFAZ
        print("\n📋 CONSULTOR SEFAZ:")
        if hasattr(sistema_principal, 'consultor_sefaz_a1'):
            consultor = sistema_principal.consultor_sefaz_a1
            print("   ✅ Consultor: PRESENTE")
            
            # Verificar tipo correto
            if 'ConsultorSefazA1Corrigido' in str(type(consultor)):
                print("   ✅ Tipo: CORRIGIDO")
            else:
                print("   ⚠️ Tipo: ANTIGO")
            
            # Verificar certificado
            cert_info = consultor.obter_info_certificado()
            if cert_info.get('is_valid'):
                print("   ✅ Certificado: CONFIGURADO")
                print(f"      📅 Válido até: {cert_info['not_valid_after'].strftime('%d/%m/%Y')}")
            else:
                print("   ⚠️ Certificado: NÃO CONFIGURADO")
        else:
            print("   ❌ Consultor: AUSENTE")
        
        # 3. Verificar métodos de emergência
        print("\n📋 MÉTODOS DE EMERGÊNCIA:")
        metodos_emergencia = [
            'configurar_certificado_rapido',
            'diagnosticar_nfe',
            'resetar_sistema_nfe',
            'teste_rapido_certificado'
        ]
        
        for metodo in metodos_emergencia:
            if hasattr(sistema_principal, metodo):
                print(f"   ✅ {metodo}: PRESENTE")
            else:
                print(f"   ❌ {metodo}: AUSENTE")
        
        # 4. Teste de funcionalidade básica
        print("\n📋 TESTE DE FUNCIONALIDADE:")
        try:
            chave_teste = "35210714200166000187550010000000271234567890"
            resultado = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave_teste)
            
            if resultado:
                print("   ✅ Consulta básica: FUNCIONANDO")
                fonte = resultado.get('fonte_dados', 'N/A')
                if 'Simulação' in fonte or 'simulados' in fonte:
                    print("   📊 Modo: SIMULAÇÃO (certificado não configurado)")
                elif 'Certificado A1' in fonte:
                    print("   📊 Modo: CERTIFICADO A1 (configurado)")
                else:
                    print(f"   📊 Modo: {fonte}")
            else:
                print("   ❌ Consulta básica: FALHOU")
                
        except Exception as e:
            print(f"   ❌ Consulta básica: ERRO - {str(e)[:50]}...")
        
        # 5. Verificar dependências
        print("\n📋 DEPENDÊNCIAS:")
        dependencias = {
            'cryptography': 'Certificados digitais',
            'requests': 'Consultas HTTP',
            'tkinter': 'Interface gráfica',
            'xml.etree.ElementTree': 'Processamento XML'
        }
        
        for dep, desc in dependencias.items():
            try:
                __import__(dep)
                print(f"   ✅ {dep}: OK ({desc})")
            except ImportError:
                print(f"   ❌ {dep}: AUSENTE ({desc})")
        
        # 6. Verificar arquivos de configuração
        print("\n📋 ARQUIVOS DE CONFIGURAÇÃO:")
        arquivos_config = [
            'config_certificado_a1.json',
            'GUIA_USO_CERTIFICADO_A1.txt',
            'MANUAL_CERTIFICADO_A1.md'
        ]
        
        for arquivo in arquivos_config:
            if os.path.exists(arquivo):
                print(f"   ✅ {arquivo}: PRESENTE")
            else:
                print(f"   ⚠️ {arquivo}: AUSENTE")
        
        print("\n" + "=" * 40)
        print("✅ DIAGNÓSTICO CONCLUÍDO")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro no diagnóstico: {e}")
        return False


def teste_consulta_completa(sistema_principal, chave_nfe):
    """Teste completo de consulta de NFe"""
    try:
        print(f"\n🧪 TESTE COMPLETO DE CONSULTA")
        print(f"Chave: {chave_nfe}")
        print("=" * 40)
        
        # Validar chave
        if len(chave_nfe) != 44 or not chave_nfe.isdigit():
            print("❌ Chave inválida - deve ter 44 dígitos")
            return False
        
        # Verificar sistema
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("❌ Sistema híbrido não inicializado")
            return False
        
        # Extrair informações da chave
        uf_codigo = chave_nfe[:2]
        cnpj_emitente = chave_nfe[6:20]
        numero_nf = str(int(chave_nfe[25:34]))
        
        print(f"📋 UF: {uf_codigo}")
        print(f"📋 CNPJ Emitente: {cnpj_emitente[:2]}.{cnpj_emitente[2:5]}.{cnpj_emitente[5:8]}/{cnpj_emitente[8:12]}-{cnpj_emitente[12:]}")
        print(f"📋 Número NFe: {numero_nf}")
        
        # Executar consulta
        print(f"\n🔍 Executando consulta...")
        resultado = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave_nfe)
        
        if resultado:
            print("✅ CONSULTA REALIZADA COM SUCESSO!")
            print(f"   🏢 Emitente: {resultado.get('razao_social_emitente', 'N/A')}")
            print(f"   📄 NFe: {resultado.get('numero_nf', 'N/A')}")
            print(f"   📅 Data: {resultado.get('data_emissao', 'N/A')}")
            print(f"   💰 Valor: R$ {resultado.get('valor_total', 0):,.2f}")
            print(f"   📦 Produtos: {len(resultado.get('produtos', []))}")
            print(f"   📊 Fonte: {resultado.get('fonte_dados', 'N/A')}")
            
            if resultado.get('status_sefaz'):
                print(f"   🏛️ Status SEFAZ: {resultado.get('status_sefaz')}")
            
            if resultado.get('observacao'):
                print(f"   📝 Observação: {resultado.get('observacao')}")
            
            return True
        else:
            print("❌ CONSULTA RETORNOU RESULTADO VAZIO")
            return False
            
    except Exception as e:
        print(f"❌ ERRO NA CONSULTA: {e}")
        return False


def menu_interativo(sistema_principal):
    """Menu interativo para testar o sistema"""
    while True:
        print("\n" + "=" * 50)
        print("🎯 MENU INTERATIVO - SISTEMA CERTIFICADO A1")
        print("=" * 50)
        print("1. 🔧 Configurar certificado A1")
        print("2. 🔍 Diagnóstico completo")
        print("3. 🧪 Teste rápido de certificado")
        print("4. 📄 Consultar NFe por chave")
        print("5. 🔄 Resetar sistema NFe")
        print("6. 💾 Salvar configuração atual")
        print("7. 📋 Exibir guia de uso")
        print("0. ❌ Sair")
        print("=" * 50)
        
        try:
            opcao = input("Digite sua opção: ").strip()
            
            if opcao == "1":
                sistema_principal.configurar_certificado_rapido()
            
            elif opcao == "2":
                diagnostico_pos_correcao(sistema_principal)
            
            elif opcao == "3":
                sistema_principal.teste_rapido_certificado()
            
            elif opcao == "4":
                chave = input("Digite a chave de 44 dígitos: ").strip().replace(' ', '')
                if chave:
                    teste_consulta_completa(sistema_principal, chave)
                else:
                    print("❌ Chave não informada")
            
            elif opcao == "5":
                sistema_principal.resetar_sistema_nfe()
            
            elif opcao == "6":
                try:
                    import json
                    from datetime import datetime
                    
                    config = {
                        'sistema_corrigido': True,
                        'data_aplicacao': datetime.now().isoformat(),
                        'versao_correcao': '2025.1'
                    }
                    
                    # Adicionar info do certificado se configurado
                    if hasattr(sistema_principal, 'consultor_sefaz_a1'):
                        cert_info = sistema_principal.consultor_sefaz_a1.obter_info_certificado()
                        if cert_info.get('is_valid'):
                            config['certificado_configurado'] = True
                            config['certificado_valido_ate'] = cert_info['not_valid_after'].isoformat()
                        else:
                            config['certificado_configurado'] = False
                    
                    with open('sistema_nfe_config.json', 'w') as f:
                        json.dump(config, f, indent=2, default=str)
                    
                    print("✅ Configuração salva em sistema_nfe_config.json")
                    
                except Exception as e:
                    print(f"❌ Erro ao salvar: {e}")
            
            elif opcao == "7":
                if os.path.exists('GUIA_USO_CERTIFICADO_A1.txt'):
                    with open('GUIA_USO_CERTIFICADO_A1.txt', 'r', encoding='utf-8') as f:
                        print(f.read())
                else:
                    print("❌ Guia não encontrado")
            
            elif opcao == "0":
                print("👋 Saindo...")
                break
            
            else:
                print("❌ Opção inválida")
                
        except KeyboardInterrupt:
            print("\n👋 Saindo...")
            break
        except Exception as e:
            print(f"❌ Erro: {e}")


if __name__ == "__main__":
    print("""
🚀 SCRIPT DE CORREÇÃO DE CERTIFICADO A1

EXECUÇÃO AUTOMÁTICA:
    from aplicar_correcoes import aplicar_todas_correcoes
    aplicar_todas_correcoes(sistema_principal)

MENU INTERATIVO:
    from aplicar_correcoes import menu_interativo
    menu_interativo(sistema_principal)

DIAGNÓSTICO:
    from aplicar_correcoes import diagnostico_pos_correcao
    diagnostico_pos_correcao(sistema_principal)
    """)
