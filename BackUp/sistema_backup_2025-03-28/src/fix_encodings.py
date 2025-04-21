# save as fix_encodings.py
from pathlib import Path

def fix_file_encoding(file_path):
    try:
        # Primeiro lê o conteúdo ignorando erros
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Depois reescreve o arquivo com encoding correto
        with open(file_path, 'w', encoding='utf-8', newline='\n') as f:
            f.write(content)
            
        print(f"Fixed: {file_path}")
    except Exception as e:
        print(f"Error fixing {file_path}: {str(e)}")

def main():
    project_dir = Path(__file__).parent
    python_files = list(project_dir.rglob('*.py'))
    
    for file_path in python_files:
        fix_file_encoding(file_path)

if __name__ == "__main__":
    main()