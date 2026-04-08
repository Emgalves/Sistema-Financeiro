"""
Runtime hook para python-docx
Busca o docx na pasta do executável (copiado manualmente para dist/)
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
    
    # Gravar diagnóstico
    try:
        log_path = os.path.join(os.path.expanduser('~'), 'Desktop', 'docx_debug.txt')
        with open(log_path, 'w') as f:
            f.write(f"_MEIPASS: {meipass}\n")
            f.write(f"exe_dir: {exe_dir}\n")
            f.write(f"docx em _MEIPASS: {os.path.exists(os.path.join(meipass, 'docx'))}\n")
            f.write(f"docx em exe_dir: {os.path.exists(docx_em_exe_dir)}\n")
            f.write(f"sys.path: {sys.path[:5]}\n")
            try:
                import docx
                f.write(f"import docx OK: {docx.__file__}\n")
            except Exception as e:
                f.write(f"import docx FALHOU: {e}\n")
    except:
        pass
