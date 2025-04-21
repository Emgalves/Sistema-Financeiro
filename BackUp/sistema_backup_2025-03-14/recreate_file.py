# recreate_file.py
from pathlib import Path

def recreate_file(filename):
    # Encontrar o diretório raiz do projeto
    current_dir = Path(__file__).resolve().parent
    
    # Procurar o arquivo em todo o projeto
    target_file = None
    for file_path in current_dir.rglob(filename):
        target_file = file_path
        break
    
    if not target_file:
        print(f"Arquivo {filename} não encontrado!")
        return
    
    print(f"Arquivo encontrado: {target_file}")
    
    try:
        # Lê o conteúdo ignorando erros de encoding
        with open(target_file, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Cria um novo arquivo com extensão .new
        new_file = target_file.with_suffix('.new.py')
        with open(new_file, 'w', encoding='utf-8', newline='\n') as f:
            f.write(content)
        
        print(f"Novo arquivo criado: {new_file}")
        print("Por favor, verifique o novo arquivo e se estiver correto, substitua o original.")
        
    except Exception as e:
        print(f"Erro ao processar arquivo: {str(e)}")

if __name__ == "__main__":
    # Agora apenas passa o nome do arquivo
    file_to_fix = "gestao_taxas.py"
    recreate_file(file_to_fix)