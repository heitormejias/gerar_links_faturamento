# Projeto


# Build using cxfreeze
instal:
cx_Freeze:8.1.0

comand:
cxfreeze build

cxfreeze build --windowed

# Build using pyinstaller
install: 
pyinstaller:6.12.0

pyinstaller gerar_links_faturamento.py -n gerar_links_faturamento --onefile

