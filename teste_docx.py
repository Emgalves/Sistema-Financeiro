from PyInstaller.utils.hooks import collect_all

datas, binaries, hiddenimports = collect_all('docx')

print('=== DATAS ===')
for d in datas[:15]:
    print(d)
print(f'... total: {len(datas)}')

print()
print('=== HIDDENIMPORTS ===')
for h in sorted(hiddenimports)[:15]:
    print(h)
print(f'... total: {len(hiddenimports)}')
