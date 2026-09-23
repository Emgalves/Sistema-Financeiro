"""
Corrige automaticamente chamadas messagebox.(showinfo|showwarning|showerror|askyesno)
que nao possuem parent= explicito, inserindo o parent correto quando isso pode
ser determinado com seguranca.

VERSAO 2 - reconhece closures (funcoes aninhadas dentro de metodos) e
generaliza a deteccao de janelas modais (qualquer variavel = tk.Toplevel(...),
nao so 'janela').

Heuristica, percorrendo a PILHA DE FUNCOES do mais interno ao mais externo
(para respeitar closures do Python: uma funcao aninhada enxerga 'self' e
variaveis locais das funcoes que a envolvem):

  PASSO A (prioridade alta - janela modal e sempre o parent mais correto,
  por ser o ancestral mais proximo da janela que esta de fato visivel):
    Em cada nivel da pilha, do mais interno para o mais externo:
      A1. Se esse nivel tem um parametro chamado 'janela' -> usa parent=janela
      A2. Senao, se no corpo desse nivel (do inicio dele ate a linha da
          chamada) existe 'NOME = tk.Toplevel(...)' -> usa parent=NOME
    Para no primeiro nivel onde uma das duas bater.

  PASSO B (so roda se o Passo A nao achou nada em nenhum nivel):
    Em cada nivel da pilha, do mais interno para o mais externo:
      B1. Se esse nivel (ou algum nivel mais externo) tem 'self' como
          primeiro parametro -> usa parent=self.root, MAS SOMENTE SE o
          arquivo realmente atribui 'self.root = ' em algum lugar (checagem
          de seguranca; sem isso, self.root poderia nem existir na classe).

  Se nada disso resolver -> NAO MEXE. Fica em "pendencias manuais".

Nunca reescreve uma chamada que ja tem parent=. Nunca adivinha quando nao
tem certeza -- fica para revisao humana, que e o mais seguro dado o volume
de ocorrencias.

Uso:
    # 1) sempre rode em modo dry-run primeiro (nao altera nada, so mostra o que faria)
    python corrigir_parent_messagebox.py --dry-run caminho\\para\\pasta_ou_arquivo.py

    # 2) quando estiver satisfeito, aplique de fato (cria .bak antes de sobrescrever)
    python corrigir_parent_messagebox.py --aplicar caminho\\para\\pasta_ou_arquivo.py
"""
import argparse
import ast
import json
import re
import sys
from pathlib import Path

METODOS_MSGBOX = {"showinfo", "showwarning", "showerror", "askyesno"}
PADRAO_TOPLEVEL = re.compile(r'\b([A-Za-z_][A-Za-z0-9_]*)\s*=\s*tk\.Toplevel\s*\(')
PADRAO_SELF_ROOT_ASSIGN = re.compile(r'\bself\.root\s*=')


def coletar_arquivos_py(caminhos):
    arquivos = []
    for caminho_str in caminhos:
        caminho = Path(caminho_str)
        if not caminho.exists():
            print(f"[AVISO] Caminho nao encontrado, ignorando: {caminho}")
            continue
        if caminho.is_dir():
            arquivos.extend(sorted(caminho.rglob("*.py")))
        elif caminho.is_file():
            arquivos.append(caminho)
    return arquivos


def eh_chamada_messagebox(node):
    """Retorna o nome do metodo (showinfo/showerror/...) se o node for uma
    chamada messagebox.<metodo>(...), senao None."""
    if not isinstance(node, ast.Call):
        return None
    func = node.func
    if not isinstance(func, ast.Attribute):
        return None
    if func.attr not in METODOS_MSGBOX:
        return None
    if not isinstance(func.value, ast.Name):
        return None
    if func.value.id != "messagebox":
        return None
    return func.attr


def tem_parent_kw(node):
    return any(kw.arg == "parent" for kw in node.keywords)


def nome_parametros(func_node):
    args = func_node.args
    nomes = []
    for a in list(args.posonlyargs) + list(args.args) + list(args.kwonlyargs):
        nomes.append(a.arg)
    return nomes


class Analisador(ast.NodeVisitor):
    """Percorre a AST mantendo uma pilha de FunctionDef (para closures) e uma
    pilha de ClassDef (para saber a que classe uma chamada pertence, e assim
    aplicar o mapeamento por classe quando self.root nao existe)."""

    def __init__(self, linhas_fonte, self_root_existe, mapa_classes):
        self.linhas_fonte = linhas_fonte
        self.self_root_existe = self_root_existe
        self.mapa_classes = mapa_classes or {}  # {nome_classe: "self.X" ou "self.X.Y"}
        self.pilha_funcoes = []
        self.pilha_classes = []
        self.edicoes = []
        self.pendencias = []

    def visit_ClassDef(self, node):
        self.pilha_classes.append(node.name)
        self.generic_visit(node)
        self.pilha_classes.pop()

    def visit_FunctionDef(self, node):
        self.pilha_funcoes.append(node)
        self.generic_visit(node)
        self.pilha_funcoes.pop()

    visit_AsyncFunctionDef = visit_FunctionDef

    def visit_Call(self, node):
        metodo = eh_chamada_messagebox(node)
        if metodo and not tem_parent_kw(node):
            self._processar_chamada(node)
        self.generic_visit(node)

    def _processar_chamada(self, node):
        niveis = list(reversed(self.pilha_funcoes))
        fim_chamada = node.lineno
        classe_atual = self.pilha_classes[-1] if self.pilha_classes else None

        alvo = None
        motivo = None

        # PASSO A: janela modal (parametro 'janela' OU variavel local Toplevel)
        for nivel in niveis:
            parametros = nome_parametros(nivel)
            if "janela" in parametros:
                alvo = "janela"
                motivo = f"parametro 'janela' da funcao '{nivel.name}'"
                break
            inicio = nivel.lineno
            trecho = "\n".join(self.linhas_fonte[inicio - 1:fim_chamada])
            matches = PADRAO_TOPLEVEL.findall(trecho)
            if matches:
                nome_var = matches[-1]
                alvo = nome_var
                motivo = f"variavel local '{nome_var} = tk.Toplevel(...)' em '{nivel.name}'"
                break

        # PASSO B: self.root, ou mapeamento por classe, se self.root nao existir
        if alvo is None:
            for nivel in niveis:
                parametros = nome_parametros(nivel)
                if parametros and parametros[0] == "self":
                    if self.self_root_existe:
                        alvo = "self.root"
                        motivo = f"'self' acessivel via closure em '{nivel.name}' (self.root confirmado no arquivo)"
                    elif classe_atual and classe_atual in self.mapa_classes:
                        alvo = self.mapa_classes[classe_atual]
                        motivo = f"'self' acessivel via closure em '{nivel.name}' (classe '{classe_atual}' mapeada para {alvo})"
                    else:
                        motivo_pendencia = (
                            f"encontrado 'self' via closure (classe '{classe_atual}'), mas NAO ha 'self.root = ' "
                            "neste arquivo e a classe nao esta no mapa -- confirme o atributo correto"
                        )
                        linha_txt = self.linhas_fonte[node.lineno - 1].strip()
                        self.pendencias.append((node.lineno, linha_txt, motivo_pendencia))
                        return
                    break

        if alvo is None:
            linha_txt = self.linhas_fonte[node.lineno - 1].strip()
            self.pendencias.append((node.lineno, linha_txt, "nao foi possivel determinar o parent com seguranca"))
            return

        self.edicoes.append((node.end_lineno, node.end_col_offset, alvo, motivo, node.lineno))


def aplicar_edicoes_no_arquivo(caminho, linhas, edicoes):
    """Insere ', parent=<alvo>' antes do ')' de fechamento de cada chamada.

    IMPORTANTE: o modulo ast do Python reporta end_col_offset em BYTES UTF-8,
    nao em caracteres (e documentado, mas surpreende). Como o codigo tem
    acentuacao (nao, gestao, codigo, etc.), cada caractere acentuado ocupa
    2 bytes mas conta como 1 caractere na string Python -- por isso a posicao
    precisa ser convertida de bytes para indice de caractere antes de fatiar
    a string, ou a insercao cai no lugar errado e quebra a sintaxe.

    Processa por linha, em ordem decrescente de posicao, para nao invalidar
    offsets dentro da mesma linha."""
    por_linha = {}
    for end_lineno, end_col_offset, alvo, motivo, lineno_inicio in edicoes:
        por_linha.setdefault(end_lineno, []).append((end_col_offset, alvo))

    for end_lineno, itens in por_linha.items():
        linha = linhas[end_lineno - 1]
        linha_bytes = linha.encode('utf-8')

        # converte cada end_col_offset (bytes) para indice de caractere,
        # decodificando o prefixo de bytes correspondente
        itens_convertidos = []
        for end_col_offset_bytes, alvo in itens:
            prefixo_bytes = linha_bytes[:end_col_offset_bytes]
            prefixo_str = prefixo_bytes.decode('utf-8')
            pos_caractere = len(prefixo_str)
            itens_convertidos.append((pos_caractere, alvo))

        itens_convertidos.sort(key=lambda x: x[0], reverse=True)
        for pos_caractere, alvo in itens_convertidos:
            pos_fechamento = pos_caractere - 1  # posicao do ')' de fechamento
            linha = linha[:pos_fechamento] + f", parent={alvo}" + linha[pos_fechamento:]
        linhas[end_lineno - 1] = linha
    return linhas


def processar_arquivo(caminho, aplicar, mapa_atributos, raiz):
    try:
        texto_original = caminho.read_text(encoding="utf-8")
    except Exception as e:
        print(f"[ERRO] Nao foi possivel ler {caminho}: {e}")
        return 0, 0

    try:
        arvore = ast.parse(texto_original, filename=str(caminho))
    except SyntaxError as e:
        print(f"[ERRO DE SINTAXE] {caminho}: {e}")
        return 0, 0

    linhas_fonte = texto_original.splitlines(keepends=True)
    self_root_existe = bool(PADRAO_SELF_ROOT_ASSIGN.search(texto_original))

    # chave de lookup no mapa: caminho relativo a raiz, com barras normais
    try:
        chave_arquivo = str(caminho.relative_to(raiz)).replace("\\", "/")
    except ValueError:
        chave_arquivo = caminho.name
    mapa_classes = mapa_atributos.get(chave_arquivo, {})

    analisador = Analisador([l.rstrip("\n").rstrip("\r") for l in linhas_fonte], self_root_existe, mapa_classes)
    analisador.visit(arvore)

    if not analisador.edicoes and not analisador.pendencias:
        return 0, 0

    print(f"\n=== {caminho} ===")
    if not self_root_existe:
        print("  [AVISO] Este arquivo NAO possui 'self.root = ' -- correcoes que dependeriam de self.root foram movidas para pendencia.")
    if analisador.edicoes:
        print(f"  Corrigiveis automaticamente: {len(analisador.edicoes)}")
        for _, _, alvo, motivo, lineno_inicio in sorted(analisador.edicoes, key=lambda x: x[4]):
            print(f"    linha {lineno_inicio}: parent={alvo}  ({motivo})")
    if analisador.pendencias:
        print(f"  PENDENCIA MANUAL: {len(analisador.pendencias)}")
        for lineno, txt, motivo_pend in analisador.pendencias:
            print(f"    linha {lineno}: {txt}   [{motivo_pend}]")

    if aplicar and analisador.edicoes:
        linhas_editaveis = list(linhas_fonte)
        linhas_editaveis = aplicar_edicoes_no_arquivo(caminho, linhas_editaveis, analisador.edicoes)
        novo_texto = "".join(linhas_editaveis)

        # valida que o resultado ainda e Python valido antes de gravar
        try:
            ast.parse(novo_texto, filename=str(caminho))
        except SyntaxError as e:
            print(f"  [ERRO] Correcao geraria sintaxe invalida, arquivo NAO alterado: {e}")
            return 0, len(analisador.pendencias)

        backup = caminho.with_suffix(caminho.suffix + ".bak")
        if not backup.exists():
            backup.write_text(texto_original, encoding="utf-8")
        caminho.write_text(novo_texto, encoding="utf-8")
        print(f"  [OK] Arquivo corrigido. Backup salvo em: {backup}")

    return len(analisador.edicoes), len(analisador.pendencias)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("caminhos", nargs="+", help="Arquivos .py ou pastas")
    modo = parser.add_mutually_exclusive_group(required=True)
    modo.add_argument("--dry-run", action="store_true", help="Apenas mostra o que seria feito, nao altera nada")
    modo.add_argument("--aplicar", action="store_true", help="Aplica as correcoes de fato (gera .bak antes)")
    parser.add_argument("--mapa", help="Arquivo JSON com mapeamento {arquivo: {classe: 'self.X'}} para quando self.root nao existe")
    args = parser.parse_args()

    arquivos = coletar_arquivos_py(args.caminhos)
    if not arquivos:
        print("Nenhum arquivo .py encontrado.")
        sys.exit(1)

    mapa_atributos = {}
    if args.mapa:
        try:
            with open(args.mapa, encoding="utf-8") as f:
                mapa_atributos = json.load(f)
        except Exception as e:
            print(f"[ERRO] Nao foi possivel ler o mapa {args.mapa}: {e}")
            sys.exit(1)

    # raiz para calcular caminho relativo -- usa o primeiro caminho de entrada
    # que seja uma pasta, ou o pai do primeiro arquivo
    primeiro = Path(args.caminhos[0])
    raiz = primeiro if primeiro.is_dir() else primeiro.parent

    total_corrigiveis = 0
    total_pendencias = 0
    for arquivo in arquivos:
        c, p = processar_arquivo(arquivo, aplicar=args.aplicar, mapa_atributos=mapa_atributos, raiz=raiz)
        total_corrigiveis += c
        total_pendencias += p

    print(f"\n{'=' * 60}")
    print(f"Arquivos verificados: {len(arquivos)}")
    print(f"Chamadas corrigidas automaticamente: {total_corrigiveis}" if args.aplicar
          else f"Chamadas que SERIAM corrigidas automaticamente: {total_corrigiveis}")
    print(f"Chamadas que precisam de revisao manual: {total_pendencias}")
    if args.dry_run:
        print("\nNada foi alterado (modo --dry-run). Revise a lista acima e rode com --aplicar quando estiver de acordo.")


if __name__ == "__main__":
    main()
