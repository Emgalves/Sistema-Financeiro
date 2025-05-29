#!/usr/bin/env python3
"""
Correção definitiva dos imports no sistema_principal.py
"""

import os
import re

def corrigir_sistema_principal():
    """Corrige todos os imports no sistema_principal.py"""
    
    arquivo = "src/sistema_principal.py"
    
    if not os.path.exists(arquivo):
        print(f"❌ Arquivo {arquivo} não encontrado!")
        return False
    
    print(f"🔧 Corrigindo imports em: {arquivo}")
    
    # Ler arquivo atual
    with open(arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    # Fazer backup
    backup_file = arquivo + ".backup_imports"
    with open(backup_file, 'w', encoding='utf-8') as f:
        f.write(conteudo)
    print(f"📄 Backup criado: {backup_file}")
    
    # Lista de correções a fazer
    correcoes = [
        # Método reload_module - primeiro tentativa
        ("modulo = self.reload_module('relatorios_interface')", 
         "modulo = self.reload_module('src.relatorios_interface')"),
        
        # Método reload_module - segunda tentativa  
        ("modulo = self.reload_module('relatorio_despesas_aprimorado')",
         "modulo = self.reload_module('src.relatorio_despesas_aprimorado')"),
        
        # Outros métodos
        ("modulo = self.reload_module('despesas_rateadas')",
         "modulo = self.reload_module('src.despesas_rateadas')"),
        
        ("modulo = self.reload_module('gestao_medicoes')",
         "modulo = self.reload_module('src.gestao_medicoes')"),
        
        # Import direto no abrir_configuracoes
        ("from configuracoes_sistema import GerenciadorConfiguracoes",
         "from src.configuracoes_sistema import GerenciadorConfiguracoes"),
    ]
    
    conteudo_novo = conteudo
    correcoes_aplicadas = []
    
    # Aplicar cada correção
    for buscar, substituir in correcoes:
        if buscar in conteudo_novo:
            conteudo_novo = conteudo_novo.replace(buscar, substituir)
            correcoes_aplicadas.append(f"✅ {buscar} → {substituir}")
            print(f"✅ Corrigido: {buscar}")
        else:
            print(f"⚠️  Não encontrado: {buscar}")
    
    # Correções adicionais usando regex para pegar variações
    
    # Corrigir qualquer reload_module que não tenha src.
    pattern_reload = r"self\.reload_module\('([^']+)'\)"
    
    def corrigir_reload(match):
        modulo = match.group(1)
        if modulo in ['relatorios_interface', 'relatorio_despesas_aprimorado', 
                     'despesas_rateadas', 'gestao_medicoes', 'configuracoes_sistema']:
            if not modulo.startswith('src.'):
                return f"self.reload_module('src.{modulo}')"
        return match.group(0)
    
    conteudo_novo = re.sub(pattern_reload, corrigir_reload, conteudo_novo)
    
    # Salvar arquivo corrigido
    if conteudo_novo != conteudo:
        with open(arquivo, 'w', encoding='utf-8') as f:
            f.write(conteudo_novo)
        
        print(f"\n✅ Correções aplicadas com sucesso!")
        print(f"Total de correções: {len(correcoes_aplicadas)}")
        
        for correcao in correcoes_aplicadas:
            print(f"   {correcao}")
        
        return True
    else:
        print(f"⚠️  Nenhuma correção foi necessária")
        return False

def verificar_correcoes():
    """Verifica se as correções foram aplicadas"""
    
    arquivo = "src/sistema_principal.py"
    
    with open(arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    print(f"\n🔍 Verificando correções:")
    
    # Verificar se os imports corretos estão presentes
    imports_corretos = [
        "self.reload_module('src.relatorios_interface')",
        "self.reload_module('src.relatorio_despesas_aprimorado')",
        "self.reload_module('src.despesas_rateadas')", 
        "self.reload_module('src.gestao_medicoes')",
        "from src.configuracoes_sistema import GerenciadorConfiguracoes"
    ]
    
    for import_correto in imports_corretos:
        if import_correto in conteudo:
            print(f"✅ {import_correto}")
        else:
            print(f"❌ {import_correto}")
    
    # Verificar se ainda há imports incorretos
    imports_incorretos = [
        "self.reload_module('relatorios_interface')",
        "self.reload_module('relatorio_despesas_aprimorado')",
        "self.reload_module('despesas_rateadas')",
        "self.reload_module('gestao_medicoes')",
        "from configuracoes_sistema import GerenciadorConfiguracoes"
    ]
    
    problemas = []
    for import_incorreto in imports_incorretos:
        if import_incorreto in conteudo:
            problemas.append(import_incorreto)
            print(f"⚠️  AINDA PRESENTE: {import_incorreto}")
    
    return len(problemas) == 0

def main():
    print("=" * 70)
    print("CORREÇÃO DEFINITIVA DOS IMPORTS")
    print("=" * 70)
    
    print("\n🎯 PROBLEMA IDENTIFICADO:")
    print("Os módulos só funcionam com prefixo 'src.':")
    print("- relatorios_interface → src.relatorios_interface")
    print("- relatorio_despesas_aprimorado → src.relatorio_despesas_aprimorado")
    print("- despesas_rateadas → src.despesas_rateadas")
    print("- gestao_medicoes → src.gestao_medicoes")
    print("- configuracoes_sistema → src.configuracoes_sistema")
    
    # Aplicar correções
    if corrigir_sistema_principal():
        # Verificar se foram aplicadas corretamente
        if verificar_correcoes():
            print(f"\n🎉 CORREÇÃO CONCLUÍDA COM SUCESSO!")
            print(f"\n🚀 PRÓXIMOS PASSOS:")
            print("1. Execute o build:")
            print("   python build_simples.py")
            print("2. Teste o executável")
            print("3. TODOS os módulos devem funcionar na PRIMEIRA tentativa!")
            print("4. Não será mais necessário abrir 'Entrada de Dados' primeiro!")
            
        else:
            print(f"\n⚠️  Algumas correções podem não ter sido aplicadas")
            print("Verifique manualmente o arquivo src/sistema_principal.py")
    else:
        print(f"\n❌ Não foi possível aplicar as correções automaticamente")

if __name__ == "__main__":
    main()