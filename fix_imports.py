
def fix_imports():
    """Fix temporário para resolver imports - baseado no que Sistema_Entrada_Dados faz"""
    import sys
    import os
    from pathlib import Path
    
    print("Aplicando fix de imports...")
    
    # Adicionar diretório atual
    current_dir = Path(__file__).resolve().parent
    if str(current_dir) not in sys.path:
        sys.path.insert(0, str(current_dir))
        print(f"Adicionado ao path: {current_dir}")
    
    # Adicionar diretório pai (raiz do projeto)
    project_root = current_dir.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
        print(f"Adicionado ao path: {project_root}")
    
    # Adicionar src especificamente
    src_dir = project_root / "src"
    if src_dir.exists() and str(src_dir) not in sys.path:
        sys.path.insert(0, str(src_dir))
        print(f"Adicionado ao path: {src_dir}")
    
    # Forçar reload de módulos problemáticos se já estiverem carregados
    modulos_problematicos = [
        'relatorios_interface',
        'relatorio_despesas_aprimorado', 
        'despesas_rateadas',
        'gestao_medicoes',
        'configuracoes_sistema'
    ]
    
    for modulo in modulos_problematicos:
        if modulo in sys.modules:
            print(f"Removendo {modulo} do cache para forçar reload")
            del sys.modules[modulo]
        
        # Também versões com src
        modulo_src = f"src.{modulo}"
        if modulo_src in sys.modules:
            print(f"Removendo {modulo_src} do cache para forçar reload")
            del sys.modules[modulo_src]

if __name__ == "__main__":
    fix_imports()
