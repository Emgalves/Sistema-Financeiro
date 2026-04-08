import zipfile

exe = r'dist\Sistema_Gestao_Financeira_PRODUCAO.exe'

try:
    with zipfile.ZipFile(exe, 'r') as z:
        docx_files = [f for f in z.namelist() if 'docx' in f.lower()]
        for f in sorted(docx_files)[:40]:
            print(f)
        print(f'\n... total entradas docx: {len(docx_files)}')
except zipfile.BadZipFile:
    print("Executável não é um zip — usando método alternativo")
    # Buscar pasta _MEIPASS extraída no Temp
    import os, glob
    temp = os.environ.get('TEMP', '')
    meis = glob.glob(os.path.join(temp, '_MEI*'))
    for mei in meis:
        docx_path = os.path.join(mei, 'docx')
        if os.path.exists(docx_path):
            print(f"\nPasta docx encontrada em: {docx_path}")
            for root, dirs, files in os.walk(docx_path):
                for f in files:
                    rel = os.path.relpath(os.path.join(root, f), mei)
                    print(rel)
            break
    else:
        print("Nenhuma pasta _MEI com docx encontrada no Temp")
        print("Execute o executável primeiro e rode este script novamente")
