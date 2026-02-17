# Fix the garbled unicode bullet character
with open('sidebar_main.py', 'rb') as f:
    data = f.read()

old = b'\xc3\xa2\xe2\x80\x94\xc2\x8f'
new = '\u25CF'.encode('utf-8')

if old in data:
    data = data.replace(old, new)
    with open('sidebar_main.py', 'wb') as f:
        f.write(data)
    print('Fixed! Replaced garbled bytes with bullet character.')
else:
    print('Target bytes not found.')
