# save as check_null_bytes.py
from pathlib import Path

def check_file_for_null_bytes(file_path):
    with open(file_path, 'rb') as f:
        content = f.read()
        null_positions = []
        for i, byte in enumerate(content):
            if byte == 0:
                null_positions.append(i)
        return null_positions

def main():
    project_dir = Path(__file__).parent
    python_files = list(project_dir.rglob('*.py'))
    
    for file_path in python_files:
        null_positions = check_file_for_null_bytes(file_path)
        if null_positions:
            print(f"\nNULL bytes encontrados em {file_path}:")
            print(f"Posições: {null_positions}")
            
            # Ler o conteúdo como string para mostrar contexto
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
                for pos in null_positions:
                    start = max(0, pos - 20)
                    end = min(len(content), pos + 20)
                    print(f"\nContexto próximo à posição {pos}:")
                    print(content[start:end])
        else:
            print(f"OK: {file_path}")

if __name__ == "__main__":
    main()