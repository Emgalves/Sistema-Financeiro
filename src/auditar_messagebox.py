"""
Auditoria: encontra chamadas de messagebox (showinfo/showwarning/showerror/askyesno)
que NAO possuem o parametro parent= explicito.

Usa o modulo ast (parser oficial do Python) em vez de contagem de parenteses
por texto -- isso evita falsos positivos quando a mensagem exibida ao
usuario contem parenteses dentro de si mesma (ex.: uma frase que termina
com ")" como parte do texto, nao como fechamento de codigo real).

Uso:
    python auditar_messagebox.py caminho/para/arquivo.py [outro_arquivo.py ...]
    python auditar_messagebox.py caminho/para/pasta
    python auditar_messagebox.py caminho/para/pasta arquivo_extra.py

Aceita arquivos .py individuais e/ou pastas (pastas sao varridas
recursivamente, incluindo subpastas).
"""
import ast
import sys
from pathlib import Path

METODOS_MSGBOX = {"showinfo", "showwarning", "showerror", "askyesno"}


def coletar_arquivos_py(caminhos):
    """Expande a lista de argumentos em uma lista de arquivos .py,
    varrendo pastas recursivamente quando necessario."""
    arquivos = []
    for caminho_str in caminhos:
        caminho = Path(caminho_str)
        if not caminho.exists():
            print(f"[AVISO] Caminho nao encontrado, ignorando: {caminho}")
            continue
        if caminho.is_dir():
            encontrados = sorted(caminho.rglob('*.py'))
            if not encontrados:
                print(f"[AVISO] Nenhum arquivo .py encontrado em: {caminho}")
            arquivos.extend(encontrados)
        elif caminho.is_file():
            arquivos.append(caminho)
        else:
            print(f"[AVISO] Caminho nao e arquivo nem pasta, ignorando: {caminho}")
    return arquivos


def auditar(caminho):
    try:
        src = caminho.read_text(encoding='utf-8')
    except Exception as e:
        print(f"[ERRO] Nao foi possivel ler {caminho}: {e}")
        return 0

    try:
        arvore = ast.parse(src, filename=str(caminho))
    except SyntaxError as e:
        print(f"[ERRO DE SINTAXE] {caminho}: {e}")
        return 0

    total = 0
    sem_parent = []
    for node in ast.walk(arvore):
        if not (isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute)
                and node.func.attr in METODOS_MSGBOX
                and isinstance(node.func.value, ast.Name) and node.func.value.id == "messagebox"):
            continue
        total += 1
        tem_parent = any(kw.arg == "parent" for kw in node.keywords)
        if not tem_parent:
            sem_parent.append((node.lineno, f"messagebox.{node.func.attr}("))

    print(f"\n=== {caminho} ===")
    print(f"Total de chamadas messagebox: {total}")
    print(f"Chamadas SEM parent=: {len(sem_parent)}")
    for linha, txt in sorted(sem_parent):
        print(f"  linha {linha}: {txt}")

    return len(sem_parent)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python auditar_messagebox.py arquivo1.py [arquivo2.py ...]")
        print("     python auditar_messagebox.py caminho/para/pasta")
        sys.exit(1)

    arquivos = coletar_arquivos_py(sys.argv[1:])
    if not arquivos:
        print("\nNenhum arquivo .py para analisar.")
        sys.exit(1)

    total_problemas = 0
    for arquivo in arquivos:
        total_problemas += auditar(arquivo)

    print(f"\n{'=' * 50}")
    print(f"Arquivos analisados: {len(arquivos)}")
    print(f"TOTAL GERAL de chamadas sem parent explicito: {total_problemas}")
    sys.exit(1 if total_problemas > 0 else 0)
