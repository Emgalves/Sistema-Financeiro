#!/usr/bin/env python3
"""
Aplica fix diretamente no sistema_principal.py para resolver o problema de imports
"""

import os
import re

def aplicar_fix_sistema_principal():
    """Aplica o fix diretamente no sistema_principal.py"""
    
    arquivo = "src/sistema_principal.py"
    
    if not os.path.exists(arquivo):
        print(f"❌ Arquivo {arquivo} não encontrado!")
        return False
    
    print(f"🔧 Aplicando fix em: {arquivo}")
    
    # Ler arquivo atual
    with open(arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    # Código do fix para adicionar no início da função __init__ da classe SistemaPrincipal
    fix_code = '''    def __init__(self):
        # FIX: Configurar paths antes de qualquer import
        self._configurar_paths_sistema()
        
        self.usuario_atual = None
        self.root = tk.Tk()'''
    
    # Código da função para adicionar na classe SistemaPrincipal
    funcao_fix = '''
    def _configurar_paths_sistema(self):
        """Configura os paths do sistema para garantir que todos os módulos sejam encontrados"""
        import sys
        from pathlib import Path
        
        try:
            # Obter diretório atual e raiz do projeto
            current_dir = Path(__file__).resolve().parent
            project_root = current_dir.parent
            
            # Lista de diretórios para adicionar ao path
            paths_adicionar = [
                str(current_dir),      # src/
                str(project_root),     # raiz do projeto
            ]
            
            # Adicionar paths se não estiverem já incluídos
            for path in paths_adicionar:
                if path not in sys.path:
                    sys.path.insert(0, path)
                    print(f"Path adicionado: {path}")
            
            # Limpar cache de módulos problemáticos para forçar reload
            modulos_problematicos = [
                'relatorios_interface',
                'relatorio_despesas_aprimorado',
                'despesas_rateadas',
                'gestao_medicoes', 
                'configuracoes_sistema'
            ]
            
            for modulo in modulos_problematicos:
                # Remover versão direta
                if modulo in sys.modules:
                    del sys.modules[modulo]
                    print(f"Cache limpo: {modulo}")
                
                # Remover versão com src
                modulo_src = f"src.{modulo}"
                if modulo_src in sys.modules:
                    del sys.modules[modulo_src] 
                    print(f"Cache limpo: {modulo_src}")
                    
        except Exception as e:
            print(f"Erro ao configurar paths: {str(e)}")
'''
    
    # Fazer backup
    backup_file = arquivo + ".backup_fix"
    with open(backup_file, 'w', encoding='utf-8') as f:
        f.write(conteudo)
    print(f"📄 Backup criado: {backup_file}")
    
    # Aplicar correções
    conteudo_novo = conteudo
    
    # 1. Adicionar a função _configurar_paths_sistema na classe SistemaPrincipal
    # Procurar pela classe SistemaPrincipal
    pattern_class = r'(class SistemaPrincipal:.*?\n)(.*?def __init__\(self\):)'
    
    def substituir_class(match):
        class_declaration = match.group(1)
        funcao_init_start = match.group(2)
        return class_declaration + funcao_fix + "\n" + funcao_init_start
    
    conteudo_novo = re.sub(pattern_class, substituir_class, conteudo_novo, flags=re.DOTALL)
    
    # 2. Modificar o __init__ para chamar _configurar_paths_sistema
    pattern_init = r'def __init__\(self\):\s*\n\s*self\.usuario_atual = None\s*\n\s*self\.root = tk\.Tk\(\)'
    
    replacement_init = '''def __init__(self):
        # FIX: Configurar paths antes de qualquer operação
        self._configurar_paths_sistema()
        
        self.usuario_atual = None
        self.root = tk.Tk()'''
    
    conteudo_novo = re.sub(pattern_init, replacement_init, conteudo_novo)
    
    # Verificar se as mudanças foram aplicadas
    if conteudo_novo != conteudo:
        # Salvar arquivo modificado
        with open(arquivo, 'w', encoding='utf-8') as f:
            f.write(conteudo_novo)
        
        print("✅ Fix aplicado com sucesso!")
        print("\nModificações feitas:")
        print("- Adicionada função _configurar_paths_sistema()")
        print("- Modificado __init__ para chamar a configuração de paths")
        print("- Sistema agora configura paths antes de carregar módulos")
        
        return True
    else:
        print("⚠️  Nenhuma modificação foi necessária ou possível")
        return False

def verificar_fix():
    """Verifica se o fix foi aplicado corretamente"""
    
    arquivo = "src/sistema_principal.py"
    
    with open(arquivo, 'r', encoding='utf-8') as f:
        conteudo = f.read()
    
    print(f"\n🔍 Verificando fix aplicado:")
    
    # Verificar se a função foi adicionada
    if '_configurar_paths_sistema' in conteudo:
        print("✅ Função _configurar_paths_sistema encontrada")
    else:
        print("❌ Função _configurar_paths_sistema NÃO encontrada")
    
    # Verificar se o __init__ foi modificado
    if 'self._configurar_paths_sistema()' in conteudo:
        print("✅ Chamada da função no __init__ encontrada")
    else:
        print("❌ Chamada da função no __init__ NÃO encontrada")
    
    return '_configurar_paths_sistema' in conteudo

def main():
    print("=" * 70)
    print("FIX DIRETO NO SISTEMA PRINCIPAL")
    print("=" * 70)
    
    print("\n🎯 ESTRATÉGIA:")
    print("Se 'Entrada de Dados' faz os outros módulos funcionarem,")
    print("vamos aplicar a mesma lógica diretamente no sistema principal.")
    
    # Aplicar fix
    if aplicar_fix_sistema_principal():
        # Verificar se foi aplicado corretamente
        if verificar_fix():
            print(f"\n🚀 PRÓXIMOS PASSOS:")
            print("1. Execute o build novamente:")
            print("   python build_simples.py")
            print("2. Teste o executável")
            print("3. Todos os módulos devem funcionar na primeira tentativa!")
            
        else:
            print(f"\n⚠️  Fix pode não ter sido aplicado corretamente")
            print("Verifique manualmente o arquivo src/sistema_principal.py")
    else:
        print(f"\n❌ Não foi possível aplicar o fix automaticamente")
        print("Você pode aplicar manualmente:")
        print("1. Abrir src/sistema_principal.py")
        print("2. Adicionar configuração de paths no início do __init__")

if __name__ == "__main__":
    main()