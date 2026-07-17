# src/teste_manual.py
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

# Garante que a raiz do projeto está no caminho de import,
# necessário para "from src.config.config import ..." funcionar
# quando rodamos este arquivo diretamente de dentro da pasta src/
RAIZ_PROJETO = Path(__file__).resolve().parent.parent
if str(RAIZ_PROJETO) not in sys.path:
    sys.path.insert(0, str(RAIZ_PROJETO))

from leitura_guias import extrair_dados_guia, GuiaNaoReconhecida

try:
    from src.config.config import PASTA_CLIENTES
    pasta_inicial = str(PASTA_CLIENTES)
except Exception as e:
    print(f"Aviso: não foi possível carregar PASTA_CLIENTES de config.py ({e}).")
    pasta_inicial = None

root = tk.Tk()
root.withdraw()  # esconde a janela principal vazia — só queremos a caixa de diálogo

caminho_pdf = filedialog.askopenfilename(
    title="Selecione o PDF da guia (FGTS ou DARF)",
    initialdir=pasta_inicial,
    filetypes=[("PDF", "*.pdf")]
)

root.destroy()

if not caminho_pdf:
    print("Nenhum arquivo selecionado.")
else:
    print(f"Arquivo selecionado: {caminho_pdf}\n")
    try:
        resultado = extrair_dados_guia(caminho_pdf)
        print("Extração bem-sucedida:")
        for chave, valor in resultado.items():
            print(f"  {chave}: {valor}")
    except GuiaNaoReconhecida as e:
        print(f"Erro: {e}")
    except Exception as e:
        print(f"Erro inesperado: {type(e).__name__}: {e}")