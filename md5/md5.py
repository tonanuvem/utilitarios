# Import hashlib library (md5 method is part of it)
import hashlib
import os
import json
import sys

# File to check

file_dir = input("Digite a pasta onde estão os arquivos:  ")
file_dir = file_dir.replace("\\", "/")

'''
def getJSON(file):
    with open(file) as jsonFile:
        dados = json.load(jsonFile)
        jsonFile.close()
    return dados

arquivo_config = "config.json"
# passar o diretorio como parametro
if len(sys.argv) > 1:
    file_dir = sys.argv[1]
else :
    dados = getJSON(arquivo_config)
    file_dir = dados['file_dir']
'''
print("Pasta configurada: " + file_dir + "\n")

# Open,close, read file and calculate MD5 on its contents 
for arquivo in os.listdir(file_dir):
    with open(file_dir + "\\" + arquivo, "rb") as f:
    # leitura total de uma vez
        data = f.read()
        md5_returned = hashlib.md5(data).hexdigest()

    print(arquivo)
    print(": \t"+md5_returned)

input("\n\nDigite qualquer tecla para finalizar ..." )
