"""
Runtime hook para python-docx
Busca o docx na pasta do executável (copiado manualmente para dist/)

Versão sem o diagnóstico: a gravação de docx_debug.txt na Área de Trabalho
foi removida (o problema do docx já está resolvido). A lógica de caminhos
foi mantida exatamente igual, para o executável continuar carregando o
mesmo docx de hoje.
"""
import sys
import os

if hasattr(sys, '_MEIPASS'):
    meipass = sys._MEIPASS
    
    # Garantir _MEIPASS no path
    if meipass not in sys.path:
        sys.path.insert(0, meipass)
    
    # Buscar docx na pasta do executável (dist/)
    exe_dir = os.path.dirname(sys.executable)
    docx_em_exe_dir = os.path.join(exe_dir, 'docx')
    
    if os.path.exists(docx_em_exe_dir):
        if exe_dir not in sys.path:
            sys.path.insert(0, exe_dir)
