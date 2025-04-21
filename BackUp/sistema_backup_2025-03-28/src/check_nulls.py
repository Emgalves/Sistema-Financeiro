from pathlib import Path

def check_file_for_nulls(filepath):
    print(f"Verificando arquivo: {filepath}")
    try:
        with open(filepath, "rb") as f:
            content = f.read()
            null_positions = [i for i, byte in enumerate(content) if byte == 0]
            
        # Tentar diferentes encodings
        encodings = ['utf-8', 'ascii', 'iso-8859-1', 'cp1252']
        for encoding in encodings:
            try:
                with open(filepath, "r", encoding=encoding) as f:
                    content = f.read()
                print(f"Arquivo pode ser lido com encoding {encoding}")
            except UnicodeDecodeError:
                print(f"Arquivo não pode ser lido com encoding {encoding}")
            
        if null_positions:
            print(f"Encontrados {len(null_positions)} caracteres nulos nas posições: {null_positions}")
        else:
            print("Nenhum caractere nulo encontrado!")
            
        print(f"Tamanho do arquivo: {len(content)} bytes")
            
    except Exception as e:
        print(f"Erro ao ler arquivo: {e}")

# Verifica os arquivos principais
files_to_check = [
     "sistema_principal.py",
    "Sistema_Entrada_Dados.py",
    "relatorio_despesas_aprimorado.py",
    "configuracoes_sistema.py",
    "config/config.py",
    "config/utils.py",
    "config/logger_config.py",
    "config/__init__.py"
]

base_path = Path("C:/Users/Obras/sistema_gestao_testes/src")


for file in files_to_check:
    filepath = base_path / file
    check_file_for_nulls(filepath)
    print("\n" + "="*50 + "\n")
