import os
import json
import re
from openpyxl import Workbook

def drop_duplicate(x):
    return list(dict.fromkeys(x))


diretorio_raiz = os.getcwd()   # substitua pelo caminho do diretório raiz que contém os subdiretórios com os arquivos note.json

lista_de_imports = []
Conta_dir = 0
for pasta, subpastas, arquivos in os.walk(diretorio_raiz):
    for arquivo in arquivos:
        if arquivo == 'note.json':
            caminho_arquivo = os.path.join(pasta, arquivo)
            with open(caminho_arquivo, encoding='utf-8') as f:
                notebook = json.load(f)
            nomes_notebook = []
            for paragraph in notebook['paragraphs']:
                if "text" in paragraph:
                    nomes_notebook.append(paragraph['text'])
            nomes_notebook_ordenados = sorted(nomes_notebook)
            for nome in nomes_notebook_ordenados:
                if nome not in lista_de_imports:
                    #imports = re.findall(r'(?:import|from).+(?=)', nome)
                    # imports = re.findall(r"(?:import|:from\s+\w+\s+import\s+.+$).+(?=)", nome)
                    # lista_de_imports.append(imports)
                    #imports = re.findall(r"(?:import\s+.+)|(?:from\s+.+\s+import\s+.+)", nome) FUNCIONAL
                    imports = re.findall(r"(?:import\s+.+)|(?:from\s+.+\s+import)", nome)
                    lista_de_imports.extend(imports)
                    # for i in imports:
                    #     lista_de_imports.append(i)
new_list = []
for item in lista_de_imports:
    new_list.append(item.strip())

print(len(new_list))
lista_de_imports = drop_duplicate(new_list)
print(len(lista_de_imports))
lista_de_imports = sorted(lista_de_imports)
#print(lista_de_imports)
# lista_set = set()
# lista_set.update(lista_de_imports)
#print(lista_de_imports)
# for item in lista_de_imports:
#     print(item)

#VC VAI JOGAR PRO EXCEL ESSA LISTA = lista_set

workbook = Workbook()

# Selecione a planilha ativa
sheet = workbook.active

# Crie uma variável de lista
minha_lista = lista_de_imports

# Adicione cada elemento da lista em uma linha separada
for i, item in enumerate(minha_lista, start=1):
    sheet.cell(row=i, column=1, value=item)

# Salve o arquivo Excel
workbook.save(filename="lista_imports_zeppelin_6.xlsx")
