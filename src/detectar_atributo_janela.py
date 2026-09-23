"""
Para arquivos onde 'self' existe (a classe tem metodos com self) mas
'self.root' nao e atribuido em lugar nenhum, este script escaneia o
__init__ de cada classe procurando candidatos ao atributo que guarda a
janela principal/Toplevel da classe.

Candidatos, em ordem de confianca:
  ALTA   : self.X = tk.Toplevel(...)   ou   self.X = tk.Tk(...)
  MEDIA  : self.X = parent             (armazena o parametro recebido)
  BAIXA  : self.X = master             (idem, nome de parametro alternativo)

Uso:
    python detectar_atributo_janela.py caminho\\para\\pasta_ou_arquivo.py
"""
import ast
import sys
from pathlib import Path


def coletar_arquivos_py(caminhos):
    arquivos = []
    for caminho_str in caminhos:
        caminho = Path(caminho_str)
        if not caminho.exists():
            continue
        if caminho.is_dir():
            arquivos.extend(sorted(caminho.rglob("*.py")))
        elif caminho.is_file():
            arquivos.append(caminho)
    return arquivos


def arquivo_tem_self_root(texto):
    import re
    return bool(re.search(r'\bself\.root\s*=', texto))


def nomes_parametros_init(class_node):
    """Retorna o conjunto de nomes de parametros do(s) __init__ desta classe
    (exceto 'self'), para reconhecer 'self.X = <parametro>' como candidato,
    qualquer que seja o nome escolhido pelo desenvolvedor (parent, parent_frame,
    janela_pai, master, container, etc.)."""
    nomes = set()
    for node in class_node.body:
        if isinstance(node, ast.FunctionDef) and node.name == "__init__":
            args = node.args
            todos = list(args.posonlyargs) + list(args.args) + list(args.kwonlyargs)
            nomes.update(a.arg for a in todos if a.arg != "self")
    return nomes


def analisar_classe(class_node):
    """Retorna lista de (nome_atributo, confianca, linha, descricao) para
    esta classe, olhando atribuicoes self.X = ... em qualquer metodo dela
    (no __init__ e tambem em metodos tipo setup_gui, chamados pelo __init__)."""
    candidatos = []
    params_init = nomes_parametros_init(class_node)

    for node in ast.walk(class_node):
        if not isinstance(node, ast.Assign):
            continue
        if len(node.targets) != 1:
            continue
        target = node.targets[0]
        if not (isinstance(target, ast.Attribute) and isinstance(target.value, ast.Name)
                and target.value.id == "self"):
            continue
        nome_attr = target.attr

        valor = node.value
        if isinstance(valor, ast.Call) and isinstance(valor.func, ast.Attribute):
            if isinstance(valor.func.value, ast.Name) and valor.func.value.id == "tk" \
                    and valor.func.attr in ("Toplevel", "Tk"):
                candidatos.append((nome_attr, "ALTA", node.lineno, f"self.{nome_attr} = tk.{valor.func.attr}(...)"))
                continue

        if isinstance(valor, ast.Name):
            if valor.id in params_init:
                # guarda um parametro recebido pelo __init__ -- muito provavel
                # que seja a referencia a janela pai/mestre
                confianca = "ALTA" if valor.id in ("parent", "master", "root", "janela_pai", "parent_window") else "MEDIA"
                candidatos.append((nome_attr, confianca, node.lineno, f"self.{nome_attr} = {valor.id}  (parametro do __init__)"))
            elif valor.id in ("parent", "master", "janela_pai", "root", "parent_window"):
                candidatos.append((nome_attr, "BAIXA", node.lineno, f"self.{nome_attr} = {valor.id}"))

    return candidatos


def processar_arquivo(caminho):
    try:
        texto = caminho.read_text(encoding="utf-8")
    except Exception as e:
        print(f"[ERRO] Nao foi possivel ler {caminho}: {e}")
        return

    if arquivo_tem_self_root(texto):
        return  # ja tem self.root, nao precisa de sugestao

    try:
        arvore = ast.parse(texto, filename=str(caminho))
    except SyntaxError as e:
        print(f"[ERRO DE SINTAXE] {caminho}: {e}")
        return

    classes = [n for n in ast.walk(arvore) if isinstance(n, ast.ClassDef)]
    if not classes:
        return  # arquivo sem classes, provavelmente Grupo B (funcoes soltas)

    saida_arquivo = []
    for classe in classes:
        candidatos = analisar_classe(classe)
        if candidatos:
            # ordena por confianca (ALTA > MEDIA > BAIXA) e remove duplicados por nome de atributo
            ordem_confianca = {"ALTA": 0, "MEDIA": 1, "BAIXA": 2}
            candidatos_unicos = {}
            for nome_attr, confianca, lineno, desc in candidatos:
                chave = nome_attr
                if chave not in candidatos_unicos or ordem_confianca[confianca] < ordem_confianca[candidatos_unicos[chave][1]]:
                    candidatos_unicos[chave] = (nome_attr, confianca, lineno, desc)
            candidatos_ordenados = sorted(candidatos_unicos.values(), key=lambda c: ordem_confianca[c[1]])
            saida_arquivo.append((classe.name, candidatos_ordenados))

    if saida_arquivo:
        print(f"\n=== {caminho} ===")
        for nome_classe, candidatos in saida_arquivo:
            print(f"  classe {nome_classe}:")
            for nome_attr, confianca, lineno, desc in candidatos:
                print(f"    [{confianca}] self.{nome_attr}  (linha {lineno}: {desc})")
    else:
        print(f"\n=== {caminho} ===")
        print("  [SEM CANDIDATOS] Nenhuma classe deste arquivo atribui um Toplevel/Tk/parent a self.X.")
        print("  Provavelmente e um modulo de funcoes soltas (Grupo B) -- precisa de refatoracao,")
        print("  nao de mapeamento de atributo. Revise manualmente.")


def main():
    if len(sys.argv) < 2:
        print("Uso: python detectar_atributo_janela.py caminho\\para\\pasta_ou_arquivo.py")
        sys.exit(1)

    arquivos = coletar_arquivos_py(sys.argv[1:])
    if not arquivos:
        print("Nenhum arquivo .py encontrado.")
        sys.exit(1)

    for arquivo in arquivos:
        processar_arquivo(arquivo)


if __name__ == "__main__":
    main()
