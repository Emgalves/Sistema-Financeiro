# -*- coding: utf-8 -*-
"""
Script para aplicar melhorias de certificado A1 ao sistema existente
Executa sem quebrar funcionalidades já implementadas
"""

def aplicar_melhorias_ao_sistema_existente(sistema_principal):
    """
    Aplica todas as melhorias de certificado A1 ao sistema já funcionando
    
    Args:
        sistema_principal: Instância do SistemaEntradaDados com NFe já inicializado
    """
    try:
        print("\n🚀 APLICANDO MELHORIAS DE CERTIFICADO A1")
        print("=" * 55)
        
        # 1. Verificar se sistema híbrido está funcionando
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("❌ Sistema híbrido NFe não encontrado!")
            print("💡 Execute primeiro: inicializar_sistema_nfe_hibrido(sistema_principal)")
            return False
        
        print("✅ Sistema híbrido NFe detectado")
        
        # 2. Importar e aplicar melhorias de certificado
        try:
            from src.nfe.consulta_sefaz_certificado import aplicar_melhorias_ao_sistema_existente as aplicar_melhorias_certificado_a1
            
            sucesso = aplicar_melhorias_certificado_a1(sistema_principal)
            
            if sucesso:
                print("✅ Melhorias de certificado A1 aplicadas!")
            else:
                print("⚠️ Melhorias parcialmente aplicadas")
                
        except ImportError:
            print("❌ Módulo consulta_sefaz_certificado não encontrado")
            print("💡 Certifique-se de que o arquivo está no diretório src/nfe/")
            return False
        
        # 3. Verificar dependências
        print("\n🔍 Verificando dependências...")
        dependencias_ok = verificar_dependencias()
        
        if not dependencias_ok:
            print("⚠️ Algumas dependências estão faltando")
            print("💡 Execute: pip install cryptography requests")
        
        # 4. Adicionar método de configuração rápida
        def configuracao_rapida_certificado():
            """Método de configuração rápida via console"""
            try:
                import tkinter as tk
                from tkinter import filedialog, simpledialog, messagebox
                
                # Criar janela temporária
                root = tk.Tk()
                root.withdraw()  # Ocultar janela principal
                
                # Selecionar arquivo
                cert_path = filedialog.askopenfilename(
                    title="Selecionar Certificado A1",
                    filetypes=[("Certificado", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                if not cert_path:
                    print("❌ Nenhum certificado selecionado")
                    root.destroy()
                    return False
                
                # Solicitar senha
                cert_password = simpledialog.askstring(
                    "Senha do Certificado",
                    "Digite a senha do certificado:",
                    show='*'
                )
                
                if not cert_password:
                    print("❌ Senha não informada")
                    root.destroy()
                    return False
                
                root.destroy()
                
                # Configurar certificado
                print(f"🔐 Configurando certificado: {cert_path}")
                
                if hasattr(sistema_principal.processador_nfe, 'configurar_certificado_a1'):
                    sucesso, msg = sistema_principal.processador_nfe.configurar_certificado_a1(
                        cert_path, cert_password
                    )
                    
                    if sucesso:
                        print(f"✅ {msg}")
                        
                        # Testar conexão
                        print("🧪 Testando conexão...")
                        teste_ok, teste_msg = sistema_principal.processador_nfe.testar_certificado_a1()
                        print(f"📡 {teste_msg}")
                        
                        return True
                    else:
                        print(f"❌ {msg}")
                        return False
                else:
                    print("❌ Método de configuração não disponível")
                    return False
                    
            except Exception as e:
                print(f"❌ Erro na configuração: {e}")
                return False
        
        # Adicionar método ao sistema
        sistema_principal.configurar_certificado_rapido = configuracao_rapida_certificado
        
        # 5. Criar método de diagnóstico
        def diagnosticar_sistema_nfe():
            """Diagnóstica o estado do sistema NFe"""
            print("\n🔍 DIAGNÓSTICO DO SISTEMA NFe")
            print("=" * 40)
            
            # Sistema híbrido
            if hasattr(sistema_principal, 'processador_nfe'):
                print("✅ Sistema híbrido: ATIVO")
                
                # Integrador NFe
                if hasattr(sistema_principal, 'integrador_nfe'):
                    print("✅ Integrador NFe: PRESENTE")
                else:
                    print("⚠️ Integrador NFe: AUSENTE")
                
                # Integrador completo
                if hasattr(sistema_principal, 'integrador_nfe_completo'):
                    print("✅ Integrador completo: PRESENTE")
                else:
                    print("⚠️ Integrador completo: AUSENTE")
                
                # Consultor SEFAZ A1
                if hasattr(sistema_principal, 'consultor_sefaz_a1'):
                    print("✅ Consultor SEFAZ A1: PRESENTE")
                    
                    # Verificar certificado
                    cert_info = sistema_principal.consultor_sefaz_a1.obter_info_certificado()
                    if cert_info.get('is_valid'):
                        print(f"✅ Certificado: VÁLIDO até {cert_info['not_valid_after'].strftime('%d/%m/%Y')}")
                    else:
                        print("⚠️ Certificado: NÃO CONFIGURADO")
                else:
                    print("⚠️ Consultor SEFAZ A1: AUSENTE")
                
                # Métodos disponíveis
                metodos = []
                if hasattr(sistema_principal.processador_nfe, 'configurar_certificado_a1'):
                    metodos.append("configurar_certificado_a1")
                if hasattr(sistema_principal.processador_nfe, 'testar_certificado_a1'):
                    metodos.append("testar_certificado_a1")
                if hasattr(sistema_principal, 'configurar_certificado_rapido'):
                    metodos.append("configurar_certificado_rapido")
                
                if metodos:
                    print(f"✅ Métodos disponíveis: {', '.join(metodos)}")
                else:
                    print("⚠️ Nenhum método específico A1 encontrado")
                    
            else:
                print("❌ Sistema híbrido: INATIVO")
            
            print("=" * 40)
        
        # Adicionar diagnóstico ao sistema
        sistema_principal.diagnosticar_nfe = diagnosticar_sistema_nfe
        
        # 6. Exibir resumo das melhorias
        print("\n📋 RESUMO DAS MELHORIAS APLICADAS:")
        print("=" * 45)
        print("✅ Consulta real via SEFAZ com certificado A1")
        print("✅ Interface melhorada para configuração")
        print("✅ Teste automático de conectividade")
        print("✅ Fallback para dados simulados")
        print("✅ Método de configuração rápida")
        print("✅ Diagnóstico do sistema")
        
        print("\n🎯 PRÓXIMOS PASSOS:")
        print("1. Configure o certificado A1:")
        print("   sistema_principal.configurar_certificado_rapido()")
        print("\n2. Ou use a interface gráfica:")
        print("   Botão 'Configurar Certificado' na importação de NFe")
        print("\n3. Para diagnóstico:")
        print("   sistema_principal.diagnosticar_nfe()")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro ao aplicar melhorias: {e}")
        return False


def verificar_dependencias():
    """Verifica se todas as dependências estão instaladas"""
    dependencias = {
        'cryptography': 'pip install cryptography',
        'requests': 'pip install requests',
        'xml.etree.ElementTree': 'Built-in do Python',
        'datetime': 'Built-in do Python'
    }
    
    dependencias_ok = True
    
    for dep, install_cmd in dependencias.items():
        try:
            if '.' in dep:
                # Para módulos com submodulos
                __import__(dep.split('.')[0])
            else:
                __import__(dep)
            print(f"✅ {dep}: OK")
        except ImportError:
            print(f"❌ {dep}: FALTANDO - {install_cmd}")
            dependencias_ok = False
    
    return dependencias_ok


def teste_completo_certificado_a1(sistema_principal, cert_path=None, cert_password=None):
    """
    Executa teste completo do certificado A1
    
    Args:
        sistema_principal: Sistema principal
        cert_path: Caminho do certificado (opcional, será solicitado se não informado)
        cert_password: Senha do certificado (opcional, será solicitada se não informada)
    """
    try:
        print("\n🧪 TESTE COMPLETO DE CERTIFICADO A1")
        print("=" * 50)
        
        # Verificar se sistema tem melhorias aplicadas
        if not hasattr(sistema_principal, 'consultor_sefaz_a1'):
            print("❌ Melhorias de certificado A1 não aplicadas!")
            print("💡 Execute: aplicar_melhorias_ao_sistema_existente(sistema_principal)")
            return False
        
        # Solicitar certificado se não informado
        if not cert_path or not cert_password:
            print("📋 Certificado e senha necessários para o teste")
            
            if not cert_path:
                import tkinter as tk
                from tkinter import filedialog
                
                root = tk.Tk()
                root.withdraw()
                
                cert_path = filedialog.askopenfilename(
                    title="Selecionar Certificado para Teste",
                    filetypes=[("Certificado", "*.pfx *.p12"), ("Todos", "*.*")]
                )
                
                root.destroy()
                
                if not cert_path:
                    print("❌ Teste cancelado - certificado não selecionado")
                    return False
            
            if not cert_password:
                from tkinter import simpledialog
                import tkinter as tk
                
                root = tk.Tk()
                root.withdraw()
                
                cert_password = simpledialog.askstring(
                    "Senha para Teste",
                    "Digite a senha do certificado:",
                    show='*'
                )
                
                root.destroy()
                
                if not cert_password:
                    print("❌ Teste cancelado - senha não informada")
                    return False
        
        # Executar teste usando o módulo de diagnóstico
        from src.nfe.consulta_sefaz_certificado import diagnosticar_certificado_a1, testar_certificado_a1_manualmente
        
        print("🔍 1. Diagnóstico do certificado...")
        diagnostico_ok = diagnosticar_certificado_a1(cert_path, cert_password)
        
        if not diagnostico_ok:
            print("❌ Falha no diagnóstico do certificado")
            return False
        
        print("\n🧪 2. Teste completo de funcionalidade...")
        
        # Chave de teste (formato válido, mas pode não existir)
        chave_teste = "35200114200166000187550010000000271234567890"
        
        teste_ok = testar_certificado_a1_manualmente(cert_path, cert_password, chave_teste)
        
        if teste_ok:
            print("\n✅ TESTE COMPLETO CONCLUÍDO COM SUCESSO!")
            print("📋 O certificado está funcionando corretamente")
            return True
        else:
            print("\n⚠️ TESTE APRESENTOU PROBLEMAS")
            print("📋 Verifique os logs acima para detalhes")
            return False
            
    except Exception as e:
        print(f"❌ Erro durante o teste: {e}")
        return False


def configurar_certificado_producao(sistema_principal):
    """
    Configuração dedicada para ambiente de produção
    Com validações extras e logs detalhados
    """
    try:
        print("\n🏭 CONFIGURAÇÃO PARA PRODUÇÃO")
        print("=" * 40)
        
        # Verificar se melhorias estão aplicadas
        if not hasattr(sistema_principal, 'consultor_sefaz_a1'):
            print("❌ Aplique as melhorias primeiro!")
            return False
        
        import tkinter as tk
        from tkinter import filedialog, simpledialog, messagebox
        
        # Interface dedicada para produção
        root = tk.Tk()
        root.title("Configuração Certificado A1 - Produção")
        root.geometry("500x400")
        
        # Frame principal
        main_frame = tk.Frame(root, padx=20, pady=20)
        main_frame.pack(fill='both', expand=True)
        
        # Título
        title_label = tk.Label(
            main_frame,
            text="🔐 Configuração Certificado A1",
            font=('Arial', 16, 'bold'),
            fg='#0056b3'
        )
        title_label.pack(pady=(0, 20))
        
        # Instruções
        instructions = tk.Text(main_frame, height=8, wrap='word', font=('Arial', 9))
        instructions.pack(fill='x', pady=(0, 15))
        
        instructions_text = """INSTRUÇÕES PARA CONFIGURAÇÃO EM PRODUÇÃO:

1. Use apenas certificados A1 válidos e dentro da validade
2. Mantenha a senha do certificado segura
3. O certificado será testado antes da configuração
4. Conexão com SEFAZ será validada automaticamente
5. Em caso de problemas, verifique firewall e conectividade

CERTIFICADOS ACEITOS:
• Formato: .pfx ou .p12 (PKCS#12)
• Tipo: A1 (arquivo, não token/cartão)
• Status: Válido e dentro da validade"""
        
        instructions.insert('1.0', instructions_text)
        instructions.config(state='disabled')
        
        # Frame para seleção
        selection_frame = tk.Frame(main_frame)
        selection_frame.pack(fill='x', pady=10)
        
        # Variáveis
        cert_path_var = tk.StringVar()
        cert_password_var = tk.StringVar()
        
        # Arquivo
        tk.Label(selection_frame, text="Certificado:", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        file_frame = tk.Frame(selection_frame)
        file_frame.pack(fill='x', pady=5)
        
        cert_entry = tk.Entry(file_frame, textvariable=cert_path_var, font=('Arial', 9))
        cert_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        
        def select_cert():
            file_path = filedialog.askopenfilename(
                title="Selecionar Certificado A1 para Produção",
                filetypes=[
                    ("Certificado PKCS#12", "*.pfx *.p12"),
                    ("Todos os arquivos", "*.*")
                ]
            )
            if file_path:
                cert_path_var.set(file_path)
        
        tk.Button(file_frame, text="📁", command=select_cert).pack(side='right')
        
        # Senha
        tk.Label(selection_frame, text="Senha:", font=('Arial', 10, 'bold')).pack(anchor='w', pady=(10, 0))
        
        password_entry = tk.Entry(selection_frame, textvariable=cert_password_var, 
                                show='*', font=('Arial', 10))
        password_entry.pack(fill='x', pady=5)
        
        # Status
        status_label = tk.Label(selection_frame, text="Aguardando configuração...", 
                              fg='gray', font=('Arial', 9))
        status_label.pack(pady=10)
        
        # Funções dos botões
        def configurar_producao():
            cert_path = cert_path_var.get().strip()
            cert_password = cert_password_var.get()
            
            if not cert_path:
                messagebox.showerror("Erro", "Selecione o arquivo do certificado!")
                return
            
            if not cert_password:
                messagebox.showerror("Erro", "Digite a senha do certificado!")
                return
            
            # Atualizar status
            status_label.config(text="🔄 Validando certificado...", fg='blue')
            root.update()
            
            try:
                # Primeiro: diagnóstico completo
                from src.nfe.consulta_sefaz_certificado import diagnosticar_certificado_a1
                
                diagnostico_ok = diagnosticar_certificado_a1(cert_path, cert_password)
                
                if not diagnostico_ok:
                    status_label.config(text="❌ Certificado inválido!", fg='red')
                    messagebox.showerror("Erro", "Certificado inválido! Verifique arquivo e senha.")
                    return
                
                # Segundo: configurar no sistema
                status_label.config(text="🔄 Configurando no sistema...", fg='blue')
                root.update()
                
                sucesso, msg = sistema_principal.processador_nfe.configurar_certificado_a1(
                    cert_path, cert_password
                )
                
                if not sucesso:
                    status_label.config(text=f"❌ {msg}", fg='red')
                    messagebox.showerror("Erro", f"Falha na configuração:\n{msg}")
                    return
                
                # Terceiro: testar conectividade
                status_label.config(text="🔄 Testando conectividade SEFAZ...", fg='blue')
                root.update()
                
                teste_ok, teste_msg = sistema_principal.processador_nfe.testar_certificado_a1()
                
                if teste_ok:
                    status_label.config(text="✅ Certificado configurado e testado!", fg='green')
                    messagebox.showinfo("Sucesso", 
                        f"✅ Certificado configurado com sucesso!\n\n"
                        f"📋 Status: {msg}\n"
                        f"🌐 Conectividade: {teste_msg}\n\n"
                        f"O sistema está pronto para consultar NFe via SEFAZ."
                    )
                    
                    # Salvar configuração (opcional)
                    save_config = messagebox.askyesno("Salvar", 
                        "Deseja salvar a configuração do certificado?\n"
                        "(A senha NÃO será salva por segurança)"
                    )
                    
                    if save_config:
                        # Salvar apenas o caminho do certificado
                        try:
                            import json
                            from pathlib import Path
                            
                            config_file = Path("config_certificado.json")
                            config = {
                                "cert_path": cert_path,
                                "configured_at": str(datetime.now()),
                                "status": "configured"
                            }
                            
                            with open(config_file, 'w') as f:
                                json.dump(config, f, indent=2)
                            
                            print(f"✅ Configuração salva em: {config_file}")
                        except Exception as e:
                            print(f"⚠️ Erro ao salvar configuração: {e}")
                    
                    root.destroy()
                else:
                    status_label.config(text=f"⚠️ Configurado, mas: {teste_msg}", fg='orange')
                    messagebox.showwarning("Aviso", 
                        f"Certificado configurado, mas teste apresentou problemas:\n\n{teste_msg}\n\n"
                        f"O sistema funcionará, mas pode haver limitações na conectividade."
                    )
                    
            except Exception as e:
                status_label.config(text=f"❌ Erro: {str(e)[:50]}...", fg='red')
                messagebox.showerror("Erro", f"Erro durante configuração:\n{str(e)}")
        
        def testar_atual():
            """Testa certificado atualmente configurado"""
            if hasattr(sistema_principal, 'consultor_sefaz_a1'):
                cert_info = sistema_principal.consultor_sefaz_a1.obter_info_certificado()
                
                if cert_info.get('is_valid'):
                    status_label.config(text="🔄 Testando certificado atual...", fg='blue')
                    root.update()
                    
                    sucesso, msg = sistema_principal.processador_nfe.testar_certificado_a1()
                    
                    if sucesso:
                        status_label.config(text="✅ Certificado atual OK!", fg='green')
                        messagebox.showinfo("Teste", f"✅ {msg}")
                    else:
                        status_label.config(text="❌ Problemas no certificado atual", fg='red')
                        messagebox.showerror("Teste", f"❌ {msg}")
                else:
                    messagebox.showwarning("Aviso", "Nenhum certificado configurado atualmente.")
            else:
                messagebox.showerror("Erro", "Sistema de certificado não inicializado.")
        
        # Botões
        button_frame = tk.Frame(main_frame)
        button_frame.pack(fill='x', pady=20)
        
        tk.Button(button_frame, text="🔧 Configurar", 
                 command=configurar_producao, font=('Arial', 10, 'bold'),
                 bg='#0056b3', fg='white').pack(side='left', padx=(0, 5))
        
        tk.Button(button_frame, text="🧪 Testar Atual", 
                 command=testar_atual).pack(side='left', padx=5)
        
        tk.Button(button_frame, text="❌ Cancelar", 
                 command=root.destroy).pack(side='right')
        
        # Verificar se já tem certificado configurado
        if hasattr(sistema_principal, 'consultor_sefaz_a1'):
            cert_info = sistema_principal.consultor_sefaz_a1.obter_info_certificado()
            if cert_info.get('is_valid'):
                status_label.config(
                    text=f"✅ Certificado atual válido até {cert_info['not_valid_after'].strftime('%d/%m/%Y')}", 
                    fg='green'
                )
        
        root.mainloop()
        return True
        
    except Exception as e:
        print(f"❌ Erro na configuração para produção: {e}")
        return False


def criar_manual_uso_certificado_a1():
    """Cria manual de uso do certificado A1"""
    manual = """
# MANUAL DE USO - CERTIFICADO A1 PARA CONSULTA NFe

## VISÃO GERAL
O certificado digital A1 permite consultar NFe diretamente nos servidores da SEFAZ,
obtendo dados completos e atualizados em tempo real.

## PRÉ-REQUISITOS

### 1. Certificado Digital A1
- Formato: .pfx ou .p12 (PKCS#12)
- Tipo: A1 (arquivo, NÃO token ou cartão A3)
- Status: Válido e dentro da validade
- Emissor: Autoridade Certificadora homologada

### 2. Dependências do Sistema
```bash
pip install cryptography requests
```

### 3. Conectividade
- Acesso à internet
- Portas 443 (HTTPS) liberadas no firewall
- Sem proxy restritivo (ou configurado adequadamente)

## CONFIGURAÇÃO INICIAL

### 1. Aplicar Melhorias ao Sistema
```python
# No seu sistema principal
from src.nfe.aplicar_melhorias_certificado import aplicar_melhorias_ao_sistema_existente

# Aplicar melhorias (uma vez só)
aplicar_melhorias_ao_sistema_existente(sistema_principal)
```

### 2. Configuração Rápida
```python
# Configuração via interface gráfica
sistema_principal.configurar_certificado_rapido()
```

### 3. Configuração para Produção
```python
# Interface dedicada para produção
from src.nfe.aplicar_melhorias_certificado import configurar_certificado_producao
configurar_certificado_producao(sistema_principal)
```

### 4. Configuração Programática
```python
# Configuração via código
cert_path = "/caminho/para/certificado.pfx"
cert_password = "senha_do_certificado"

sucesso, msg = sistema_principal.processador_nfe.configurar_certificado_a1(
    cert_path, cert_password
)

if sucesso:
    print(f"✅ {msg}")
else:
    print(f"❌ {msg}")
```

## USO DO SISTEMA

### 1. Consulta por Chave de Acesso
```python
# Via interface gráfica
# 1. Abra: Menu > NFe > Importar NFe
# 2. Aba: "Consultar por Chave"
# 3. Cole a chave de 44 dígitos
# 4. Clique "Consultar"

# Via código
chave = "35200114200166000187550010000000271234567890"
dados_nfe = sistema_principal.processador_nfe.consultar_nfe_sefaz(chave)
```

### 2. Processamento Completo
```python
# Após consultar, use o botão "Processar NFe Completa"
# para integrar dados financeiros e materiais
```

### 3. Consulta em Lote
```python
# Via interface: Aba "Importação em Lote"
# 1. Cole múltiplas chaves (uma por linha)
# 2. Configure opções de importação
# 3. Clique "Processar Lote"
```

## DIAGNÓSTICO E TESTES

### 1. Diagnóstico Completo do Sistema
```python
# Verifica status de todos os componentes
sistema_principal.diagnosticar_nfe()
```

### 2. Teste de Certificado
```python
# Testa certificado configurado
sucesso, msg = sistema_principal.processador_nfe.testar_certificado_a1()
print(f"Teste: {msg}")
```

### 3. Teste Manual Completo
```python
from src.nfe.aplicar_melhorias_certificado import teste_completo_certificado_a1

# Teste completo com validações
teste_completo_certificado_a1(
    sistema_principal,
    "/caminho/certificado.pfx",  # opcional
    "senha"                       # opcional
)
```

### 4. Diagnóstico de Certificado
```python
from src.nfe.consulta_sefaz_certificado import diagnosticar_certificado_a1

# Diagnóstica problemas específicos do certificado
diagnosticar_certificado_a1("/caminho/certificado.pfx", "senha")
```

## SOLUÇÃO DE PROBLEMAS

### Erro: "Certificado expirado"
- **Causa**: Certificado fora da validade
- **Solução**: Renovar certificado junto à Autoridade Certificadora

### Erro: "Senha incorreta"
- **Causa**: Senha do arquivo .pfx incorreta
- **Solução**: Verificar senha com quem emitiu o certificado

### Erro: "Timeout na consulta"
- **Causa**: Conectividade com SEFAZ
- **Soluções**: 
  - Verificar conexão com internet
  - Verificar firewall (porta 443)
  - Tentar novamente após alguns minutos

### Erro: "NFe não autorizada"
- **Causa**: NFe cancelada, rejeitada ou chave inválida
- **Soluções**:
  - Verificar se chave de acesso está correta (44 dígitos)
  - Confirmar se NFe não foi cancelada
  - Verificar se NFe existe no SEFAZ

### Erro: "Arquivo de certificado não encontrado"
- **Causa**: Caminho do arquivo incorreto
- **Solução**: Verificar caminho e existência do arquivo .pfx

## FLUXO RECOMENDADO

### 1. Primeira Configuração
1. Instalar dependências: `pip install cryptography requests`
2. Aplicar melhorias: `aplicar_melhorias_ao_sistema_existente(sistema)`
3. Configurar certificado: `configurar_certificado_producao(sistema)`
4. Testar conectividade: `sistema.processador_nfe.testar_certificado_a1()`

### 2. Uso Diário
1. Abrir sistema NFe
2. Consultar NFe por chave ou XML
3. Usar "Processar NFe Completa" para integração
4. Verificar dados importados

### 3. Manutenção
1. Verificar validade do certificado mensalmente
2. Renovar certificado antes do vencimento
3. Manter backups do arquivo .pfx
4. Executar diagnósticos em caso de problemas

## SEGURANÇA

### Proteção do Certificado
- Mantenha o arquivo .pfx em local seguro
- Não compartilhe a senha do certificado
- Faça backup do certificado
- Renove antes do vencimento

### Logs e Monitoramento
- Verifique logs de consulta regularmente
- Monitor connectividade com SEFAZ
- Acompanhe mudanças nas URLs dos webservices

## SUPORTE TÉCNICO

### Logs Detalhados
O sistema gera logs detalhados das operações. Em caso de problemas:
1. Execute diagnóstico completo
2. Verifique mensagens de erro
3. Teste conectividade
4. Consulte este manual

### Contatos Úteis
- SEFAZ de cada estado para problemas de conectividade
- Autoridade Certificadora para problemas de certificado
- Suporte técnico do sistema para problemas de integração
"""
    
    # Salvar manual em arquivo
    try:
        with open("MANUAL_CERTIFICADO_A1.md", "w", encoding="utf-8") as f:
            f.write(manual)
        print("✅ Manual salvo em: MANUAL_CERTIFICADO_A1.md")
    except Exception as e:
        print(f"⚠️ Erro ao salvar manual: {e}")
    
    return manual


# FUNÇÃO PRINCIPAL PARA EXECUTAR TUDO
def setup_completo_certificado_a1(sistema_principal):
    """
    Setup completo do certificado A1 - executa tudo de uma vez
    
    Args:
        sistema_principal: Sistema principal com NFe inicializado
    """
    try:
        print("\n🚀 SETUP COMPLETO - CERTIFICADO A1 PARA NFe")
        print("=" * 60)
        
        # 1. Verificar sistema base
        print("🔍 1. Verificando sistema base...")
        if not hasattr(sistema_principal, 'processador_nfe'):
            print("❌ Sistema híbrido NFe não inicializado!")
            print("💡 Execute primeiro:")
            print("   from src.nfe.sistema_hibrido_nfe import inicializar_sistema_nfe_hibrido")
            print("   inicializar_sistema_nfe_hibrido(sistema_principal)")
            return False
        
        print("✅ Sistema híbrido NFe encontrado")
        
        # 2. Verificar dependências
        print("\n🔍 2. Verificando dependências...")
        deps_ok = verificar_dependencias()
        if not deps_ok:
            print("⚠️ Instale as dependências faltantes antes de continuar")
            return False
        
        # 3. Aplicar melhorias
        print("\n🔧 3. Aplicando melhorias de certificado A1...")
        melhorias_ok = aplicar_melhorias_ao_sistema_existente(sistema_principal)
        if not melhorias_ok:
            print("❌ Falha ao aplicar melhorias")
            return False
        
        # 4. Criar manual
        print("\n📖 4. Criando manual de uso...")
        criar_manual_uso_certificado_a1()
        
        # 5. Executar diagnóstico inicial
        print("\n🔍 5. Diagnóstico inicial do sistema...")
        sistema_principal.diagnosticar_nfe()
        
        # 6. Oferecer configuração imediata
        print("\n🎯 6. Configuração do certificado...")
        
        try:
            import tkinter as tk
            from tkinter import messagebox
            
            root = tk.Tk()
            root.withdraw()
            
            configurar_agora = messagebox.askyesno(
                "Configurar Certificado",
                "Setup completo! Deseja configurar o certificado A1 agora?\n\n"
                "✅ Sim: Abrir interface de configuração\n"
                "❌ Não: Configurar depois"
            )
            
            root.destroy()
            
            if configurar_agora:
                configurar_certificado_producao(sistema_principal)
            else:
                print("💡 Para configurar depois, use:")
                print("   sistema_principal.configurar_certificado_rapido()")
                print("   # ou #")
                print("   configurar_certificado_producao(sistema_principal)")
        
        except:
            print("💡 Para configurar certificado:")
            print("   sistema_principal.configurar_certificado_rapido()")
        
        print("\n✅ SETUP COMPLETO CONCLUÍDO!")
        print("📋 Recursos disponíveis:")
        print("   • Consulta NFe via SEFAZ com certificado A1")
        print("   • Interface melhorada de configuração")
        print("   • Testes automáticos de conectividade")
        print("   • Diagnóstico completo do sistema")
        print("   • Manual de uso detalhado")
        print("   • Fallback automático para dados simulados")
        
        return True
        
    except Exception as e:
        print(f"❌ Erro no setup completo: {e}")
        return False


# EXEMPLO DE USO FINAL
if __name__ == "__main__":
    print("""
EXEMPLO DE USO - CERTIFICADO A1 PARA NFe

# 1. Setup completo (recomendado)
from src.nfe.aplicar_melhorias_certificado import setup_completo_certificado_a1
setup_completo_certificado_a1(sistema_principal)

# 2. Apenas aplicar melhorias
from src.nfe.aplicar_melhorias_certificado import aplicar_melhorias_ao_sistema_existente
aplicar_melhorias_ao_sistema_existente(sistema_principal)

# 3. Configurar certificado
sistema_principal.configurar_certificado_rapido()

# 4. Diagnosticar sistema
sistema_principal.diagnosticar_nfe()

# 5. Testar certificado
sucesso, msg = sistema_principal.processador_nfe.testar_certificado_a1()
print(f"Teste: {msg}")
""")
